import "./styles.css";
import packageMetadata from "../package.json";
import { getConfig } from "./config.js";
import { getDataverseAccessToken, getGraphAccessToken } from "./auth.js";
import {
  beginStartupDiagnostic,
  clearInterruptedStartupDiagnostic,
  completeStartupDiagnostic,
  getCurrentStartupDiagnostic,
  recordStartupFailure,
  recordStartupPhase,
} from "./diagnostics.js";
import {
  APPOINTMENT_QUALIFICATION_FIELD,
  chooseUnambiguousCandidate,
  createRegardingNote,
  findAppointmentCandidates,
  getAppointmentQualificationOptions,
  getAppointmentNotes,
  getRegardingReference,
  getRegardingLabel,
  searchRegardingRecords,
  updateAppointmentQualification,
  updateAppointmentRegarding,
  resolveAppointmentEditingTarget,
  verifyAppointmentNoteTarget,
} from "./dataverse.js";
import { TRACKED_TO_DYNAMICS_CATEGORY, isRecurringEvent, resolveOutlookEvent, trackOutlookAppointment } from "./graph.js";
import { assertAppointmentContext, getAppointmentContext, markAppointmentAsLogged } from "./outlook.js";
import { mobileAppointmentLink } from "./mobile-link.js";
import { withoutAppointmentReference } from "./appointment-notes.js";
import { delay } from "./resilience.js";
import { redirectToLatestVersion } from "./version-update.js";

const app = document.querySelector("#app");
let config;
let context;
let accessToken;
let candidates = [];
let selectedAppointment = null;
let qualificationOptions = null;
let regardingSearchResults = [];
let interruptedDiagnostic = null;
let actionBusy = false;
let notesRequest = 0;

function confirmAction(title, message, label = "Confirmer") {
  return new Promise((resolve) => {
    const previous = document.activeElement;
    const overlay = document.createElement("div");
    overlay.className = "confirmation-overlay";
    overlay.innerHTML = `<section class="confirmation-dialog" role="alertdialog" aria-modal="true" aria-labelledby="confirmation-title">
      <h2 id="confirmation-title">${escapeHtml(title)}</h2><p>${escapeHtml(message)}</p>
      <button class="primary-button" data-confirm="yes" type="button">${escapeHtml(label)}</button>
      <button class="secondary-button" data-confirm="no" type="button">Annuler</button></section>`;
    const done = (value) => { overlay.remove(); previous?.focus(); resolve(value); };
    overlay.querySelector('[data-confirm="yes"]').addEventListener("click", () => done(true));
    const cancel = overlay.querySelector('[data-confirm="no"]');
    cancel.addEventListener("click", () => done(false));
    overlay.addEventListener("keydown", (event) => {
      if (event.key === "Escape") { event.preventDefault(); done(false); }
      if (event.key === "Tab") {
        event.preventDefault();
        (document.activeElement === cancel ? overlay.querySelector('[data-confirm="yes"]') : cancel).focus();
      }
    });
    document.body.append(overlay); cancel.focus();
  });
}

function updateActionControls() {
  const target = selectedAppointment?.editTarget;
  const enabled = !actionBusy && Boolean(selectedAppointment && context.graphEvent && target?.kind !== "blocked");
  const regarding = getRegardingReference(selectedAppointment);
  const noteEnabled = enabled && Boolean(regarding) && !selectedAppointment?.noteBlockedReason;
  for (const id of ["noteSubject", "noteText", "saveButton"]) document.querySelector("#" + id).disabled = !noteEnabled;
  const qualified = enabled && !target?.qualificationError && Boolean(qualificationOptions?.length);
  document.querySelector("#qualificationSelect").disabled = !qualified;
  document.querySelector("#saveCrmFieldsButton").disabled = !qualified || document.querySelector("#qualificationSelect").value === "";
  for (const id of ["regardingTypeSelect", "regardingSearchText", "searchRegardingButton"]) document.querySelector("#" + id).disabled = !enabled;
  document.querySelector("#saveRegardingButton").disabled = !enabled || !document.querySelector("#regardingResultSelect").value;
  document.querySelector("#trackAppointmentButton").disabled = actionBusy;
  document.querySelector("#refreshAppointmentButton").disabled = actionBusy;
}

const PHASE_LABELS = {
  application_start: "Démarrage de l’application",
  configuration: "Lecture de la configuration",
  office_library: "Chargement de la bibliothèque Office",
  outlook_initialization: "Initialisation d’Outlook",
  outlook_context_direct: "Contexte du rendez-vous déjà disponible",
  outlook_item: "Accès au rendez-vous Outlook",
  outlook_fields: "Lecture du rendez-vous Outlook",
  authentication: "Connexion Microsoft",
  dataverse_search: "Recherche du rendez-vous CRM",
  dataverse_notes: "Chargement des notes CRM",
  ready: "Application prête",
};

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

const ALLOWED_NOTE_TAGS = new Set([
  "A",
  "B",
  "BLOCKQUOTE",
  "BR",
  "DIV",
  "EM",
  "H1",
  "H2",
  "H3",
  "H4",
  "I",
  "LI",
  "OL",
  "P",
  "SPAN",
  "STRONG",
  "TABLE",
  "TBODY",
  "TD",
  "TH",
  "THEAD",
  "TR",
  "U",
  "UL",
]);
const REMOVED_NOTE_TAGS = new Set([
  "EMBED",
  "FORM",
  "IFRAME",
  "IMG",
  "INPUT",
  "LINK",
  "MATH",
  "META",
  "OBJECT",
  "SCRIPT",
  "STYLE",
  "SVG",
]);

function renderSafeNoteHtml(value) {
  const source = String(value || "");
  if (!/<[a-z][\s\S]*>/i.test(source)) {
    return escapeHtml(source).replaceAll("\n", "<br>");
  }

  const template = document.createElement("template");
  template.innerHTML = source;

  for (const element of template.content.querySelectorAll("*")) {
    if (REMOVED_NOTE_TAGS.has(element.tagName)) {
      element.remove();
      continue;
    }
    if (!ALLOWED_NOTE_TAGS.has(element.tagName)) {
      element.replaceWith(...element.childNodes);
      continue;
    }

    for (const attribute of element.getAttributeNames()) {
      if (element.tagName !== "A" || attribute.toLowerCase() !== "href") {
        element.removeAttribute(attribute);
      }
    }

    if (element.tagName === "A") {
      const href = element.getAttribute("href") || "";
      if (!/^(https?:|mailto:|tel:)/i.test(href)) {
        element.removeAttribute("href");
      } else {
        element.setAttribute("target", "_blank");
        element.setAttribute("rel", "noopener noreferrer");
      }
    }
  }

  return template.innerHTML;
}

function formatDate(date) {
  return new Intl.DateTimeFormat("fr-FR", {
    weekday: "short",
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(date));
}

function setLoadingPhase(message, detail = "") {
  const status = document.querySelector("#loadingStatus");
  const phaseDetail = document.querySelector("#loadingDetail");
  const version = document.querySelector("#loadingVersion");
  if (status) status.textContent = message;
  if (phaseDetail) phaseDetail.textContent = detail;
  if (version) version.textContent = `v${packageMetadata.version}`;
}

function diagnosticMarkup(diagnostic, title = "Diagnostic") {
  if (!diagnostic) return "";
  const phase = PHASE_LABELS[diagnostic.phase] || diagnostic.phase || "Étape inconnue";
  const date = diagnostic.updatedAt
    ? new Intl.DateTimeFormat("fr-FR", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      }).format(new Date(diagnostic.updatedAt))
    : "Date inconnue";
  return `
    <details class="diagnostic-details">
      <summary>${escapeHtml(title)}</summary>
      <p><strong>Version :</strong> ${escapeHtml(diagnostic.appVersion || packageMetadata.version)}</p>
      <p><strong>Étape :</strong> ${escapeHtml(phase)}</p>
      <p><strong>Tentative :</strong> ${escapeHtml(diagnostic.attempt || 1)}</p>
      <p><strong>Date :</strong> ${escapeHtml(date)}</p>
      ${diagnostic.errorCode ? `<p><strong>Code :</strong> ${escapeHtml(diagnostic.errorCode)}</p>` : ""}
    </details>`;
}

function showFatalError(error) {
  console.error(error);
  recordStartupFailure(error);
  const diagnostic = getCurrentStartupDiagnostic();
  app.innerHTML = `
    <section class="card error-card">
      <h1>Impossible d’ouvrir RDV &amp; Notes</h1>
      <p>${escapeHtml(error?.message || error)}</p>
      <p class="error-code">Code : ${escapeHtml(error?.code || error?.name || "UNKNOWN_ERROR")} · v${escapeHtml(packageMetadata.version)}</p>
      <p class="hint">Aucune donnée n’a été enregistrée.</p>
      <button id="retryButton" class="primary-button" type="button">Réessayer</button>
      ${diagnosticMarkup(diagnostic, "Détails techniques")}
      ${diagnosticMarkup(interruptedDiagnostic, "Tentative précédente interrompue")}
    </section>`;
  document.querySelector("#retryButton").addEventListener("click", () => globalThis.location.reload());
}

function renderShell() {
  const environmentLabel = config.dataverseUrl.toLowerCase() === "https://roth.crm4.dynamics.com"
    ? "Production"
    : "Test";

  app.innerHTML = `
    <header class="app-header">
      <div>
        <p class="eyebrow">RDV &amp; Notes</p>
        <h1>${escapeHtml(context.subject)}</h1>
      </div>
      <div class="header-status">
        <span class="version-chip">v${escapeHtml(packageMetadata.version)} · ${escapeHtml(environmentLabel)}</span>
        <span class="status-chip" id="connectionStatus">Connexion…</span>
      </div>
    </header>

    ${diagnosticMarkup(interruptedDiagnostic, "La tentative précédente a été interrompue par Outlook")}

    <section class="card appointment-summary">
      <dl>
        <div><dt>Début</dt><dd>${escapeHtml(formatDate(context.start))}</dd></div>
        <div><dt>Fin</dt><dd>${escapeHtml(formatDate(context.end))}</dd></div>
        ${context.location ? `<div><dt>Lieu</dt><dd>${escapeHtml(context.location)}</dd></div>` : ""}
        ${context.organizerName ? `<div><dt>Organisateur</dt><dd>${escapeHtml(context.organizerName)}</dd></div>` : ""}
      </dl>
      <p id="recurrenceStatus" class="hint" role="status"></p>
      ${context.isMobile ? `<a id="openMobileApp" class="secondary-button app-link hidden" target="_blank" rel="noopener noreferrer" referrerpolicy="no-referrer">Ouvrir dans RDV &amp; Notes</a>
        <p id="mobileLinkStatus" class="hint" role="status">Identification du rendez-vous pour l’application mobile…</p>` : ""}
    </section>

    <section class="card" id="matchCard">
      <h2>Rendez-vous associé dans Dynamics</h2>
      <p id="matchStatus" class="muted">Recherche en cours…</p>
      <p id="seriesScopeStatus" class="hint" role="status"></p>
      <button id="refreshAppointmentButton" class="secondary-button" type="button">Actualiser le rendez-vous</button>
      <button id="trackAppointmentButton" class="primary-button hidden" type="button">
        Activer « ${TRACKED_TO_DYNAMICS_CATEGORY} »
      </button>
      <p id="trackingStatus" class="status-message" role="status"></p>
    </section>

    <section class="card" id="crmFieldsCard">
      <h2>Champs Dynamics</h2>
      <label class="field">
        <span id="qualificationScopeLabel">Qualification du RDV</span>
        <select id="qualificationSelect" disabled>
          <option value="">Sélectionnez d’abord un rendez-vous CRM</option>
        </select>
      </label>
      <button id="saveCrmFieldsButton" class="primary-button" type="button" disabled>
        Mettre à jour la qualification
      </button>
      <p id="crmFieldsStatus" class="status-message" role="status"></p>

      <div class="field-separator"></div>
      <h3 id="regardingScopeLabel">Concernant</h3>
      <p id="currentRegardingStatus" class="muted">Aucun rendez-vous CRM sélectionné.</p>
      <label class="field">
        <span>Type</span>
        <select id="regardingTypeSelect" disabled>
          <option value="account">Compte</option>
          <option value="contact">Contact</option>
          <option value="opportunity">Opportunité</option>
        </select>
      </label>
      <label class="field">
        <span>Rechercher</span>
        <input id="regardingSearchText" placeholder="Nom ou numéro de compte" disabled />
        <small class="field-hint">Jokers : * ou %. L’ordre des termes est respecté (ex. *ma*26*rexel*).</small>
      </label>
      <button id="searchRegardingButton" class="secondary-button" type="button" disabled>
        Rechercher dans le CRM
      </button>
      <label id="regardingResultsLabel" class="field hidden">
        <span>Résultat</span>
        <select id="regardingResultSelect"></select>
      </label>
      <button id="saveRegardingButton" class="primary-button" type="button" disabled>
        Affecter ce Concernant
      </button>
      <p id="regardingStatus" class="status-message" role="status"></p>
    </section>

    <section class="card" id="noteCard">
      <h2>Ajouter une note simple</h2>
      <p id="noteTarget" class="muted">Sélectionnez d’abord un rendez-vous CRM.</p>
      <label class="field">
        <span>Titre</span>
        <input id="noteSubject" maxlength="500" value="Note — ${escapeHtml(context.subject)}" disabled />
      </label>
      <label class="field">
        <span>Note</span>
        <textarea id="noteText" rows="7" maxlength="100000" placeholder="Saisissez votre note manuelle…" disabled></textarea>
      </label>
      <button id="saveButton" class="primary-button" type="button" disabled>Enregistrer la note dans Dynamics</button>
      <p id="saveStatus" class="status-message" role="status"></p>
    </section>

    <section class="card" id="historyCard">
      <h2>Notes de ce rendez-vous</h2>
      <p class="muted">Cette séance uniquement · 20 dernières notes maximum. Les anciennes notes sans référence de RDV restent accessibles dans Dynamics.</p>
      <div id="notesList"><p class="muted">Sélectionnez d’abord un rendez-vous CRM.</p></div>
    </section>`;

  document.querySelector("#trackAppointmentButton").addEventListener("click", trackAndWaitForAppointment);
  document.querySelector("#refreshAppointmentButton").addEventListener("click", refreshCurrentMeeting);
  document.querySelector("#qualificationSelect").addEventListener("change", updateActionControls);
  document.querySelector("#openMobileApp")?.addEventListener("click", (event) => {
    try {
      assertAppointmentContext(context);
      if (actionBusy || document.querySelector("#noteText").value.trim()) throw new Error("Terminez l’action et enregistrez votre note avant d’ouvrir l’application mobile.");
      document.querySelector("#mobileLinkStatus").textContent = "Demande d’ouverture transmise. Utilisez le même compte Microsoft dans RDV & Notes.";
    } catch (error) { event.preventDefault(); document.querySelector("#mobileLinkStatus").textContent = error.message; }
  });
  document.querySelector("#saveCrmFieldsButton").addEventListener("click", saveCrmFields);
  document.querySelector("#regardingTypeSelect").addEventListener("change", () => {
    resetRegardingSearch();
    updateRegardingSearchPlaceholder();
  });
  document.querySelector("#regardingSearchText").addEventListener("keydown", (event) => {
    if (event.key === "Enter") searchRegarding();
  });
  document.querySelector("#searchRegardingButton").addEventListener("click", searchRegarding);
  document.querySelector("#regardingResultSelect").addEventListener("change", updateActionControls);
  document.querySelector("#saveRegardingButton").addEventListener("click", saveRegarding);
  document.querySelector("#saveButton").addEventListener("click", saveNote);
  if (interruptedDiagnostic) clearInterruptedStartupDiagnostic();
}

function candidateText(candidate) {
  const start = formatDate(candidate.scheduledstart);
  return `${candidate.subject || "Sans objet"} — ${start} — ${getRegardingLabel(candidate)}`;
}

async function renderCandidates() {
  const status = document.querySelector("#matchStatus");
  const trackButton = document.querySelector("#trackAppointmentButton");
  const candidate = chooseUnambiguousCandidate(candidates);
  selectedAppointment = null;
  qualificationOptions = null;
  updateSelectedAppointment();
  const strongCandidates = candidates.filter((candidate) => candidate.matchScore >= 85);

  if (candidate) {
    selectedAppointment = await resolveAppointmentEditingTarget(config, accessToken, candidate, isRecurringEvent(context.graphEvent));
    status.textContent = `Correspondance fiable : ${candidateText(selectedAppointment)}`;
    status.className = "success";
    trackButton.classList.toggle("hidden", !isRecurringEvent(context.graphEvent));
    trackButton.textContent = "Vérifier / activer le suivi de toute la série";
  } else if (strongCandidates.length > 1) {
    status.textContent =
      "Plusieurs copies CRM ont exactement le même sujet et les mêmes horaires. " +
      "Aucune n’est sélectionnée : les doublons doivent être corrigés dans le CRM.";
    status.className = "warning";
    trackButton.classList.add("hidden");
  } else {
    status.textContent =
      "Aucun rendez-vous CRM ne correspond exactement au sujet et aux horaires Outlook.";
    status.className = "warning";
    trackButton.classList.remove("hidden");
    trackButton.textContent = isRecurringEvent(context.graphEvent) ? "Activer le suivi de toute la série" : "Activer le suivi Dynamics 365";
  }

  updateSelectedAppointment();
}

async function refreshGraphContext() {
  context.graphEvent = null;
  const link = document.querySelector("#openMobileApp");
  link?.classList.add("hidden");
  try {
    const token = await getGraphAccessToken(config);
    const event = await resolveOutlookEvent(token, context);
    assertAppointmentContext(context);
    context.graphEvent = event;
    document.querySelector("#recurrenceStatus").textContent = isRecurringEvent(event)
      ? "↻ Rendez-vous récurrent — suivi et informations Dynamics sur toute la série." : "Rendez-vous non récurrent.";
    if (link) {
      link.href = mobileAppointmentLink(event.id); link.classList.remove("hidden");
      document.querySelector("#mobileLinkStatus").textContent = "Ouvre cette séance à l’étape Dynamics, sans démarrer d’enregistrement. L’association Android au domaine doit être activée.";
    }
  } catch (error) {
    document.querySelector("#recurrenceStatus").textContent = "Portée Outlook non vérifiée : modifications bloquées. " + error.message;
    if (link) document.querySelector("#mobileLinkStatus").textContent = "Ouverture mobile indisponible : le rendez-vous n’a pas pu être identifié avec certitude.";
  }
}

async function refreshCurrentMeeting() {
  if (actionBusy) return;
  actionBusy = true; updateActionControls();
  try {
    assertAppointmentContext(context);
    await refreshGraphContext();
    candidates = await findAppointmentCandidates(config, accessToken, context);
    await renderCandidates();
    if (selectedAppointment) await Promise.all([refreshCrmFields(), refreshNotes()]);
  } catch (error) {
    document.querySelector("#matchStatus").textContent = error.message;
  } finally { actionBusy = false; updateActionControls(); }
}

async function trackAndWaitForAppointment() {
  if (actionBusy) return;
  const button = document.querySelector("#trackAppointmentButton");
  const status = document.querySelector("#trackingStatus");
  actionBusy = true; updateActionControls();
  status.textContent = "Connexion au calendrier Outlook…";
  status.className = "status-message muted";

  try {
    const graphToken = await getGraphAccessToken(config);
    const event = await resolveOutlookEvent(graphToken, context);
    assertAppointmentContext(context);
    const series = isRecurringEvent(event);
    if (!await confirmAction(series ? "Suivre toute la série ?" : "Activer le suivi Dynamics ?",
      series ? "La catégorie sera activée sur toute la série dans VOTRE calendrier, organisateur ou invité. Autres catégories et exceptions conservées. La synchronisation Dynamics sera ensuite vérifiée."
        : "La catégorie sera ajoutée uniquement à ce rendez-vous dans votre calendrier. Les autres catégories seront conservées.", series ? "Suivre toute la série" : "Activer le suivi")) return;
    const tracking = await trackOutlookAppointment(graphToken, context, event, () => assertAppointmentContext(context));
    context.graphEvent = tracking.event;
    status.textContent = tracking.alreadyTracked
      ? `La catégorie « ${TRACKED_TO_DYNAMICS_CATEGORY} » est déjà active. Attente de la synchronisation CRM…`
      : `Catégorie « ${TRACKED_TO_DYNAMICS_CATEGORY} » activée. Attente de la synchronisation CRM…`;
    status.textContent = tracking.notice + " Attente de la confirmation Dynamics…";
    status.className = "status-message success";

    for (let attempt = 1; attempt <= 12; attempt += 1) {
      candidates = await findAppointmentCandidates(config, accessToken, context);
      if (chooseUnambiguousCandidate(candidates)) {
        await renderCandidates();
        status.textContent = "Le rendez-vous est maintenant disponible dans le CRM.";
        status.className = "status-message success";
        await Promise.all([refreshCrmFields(), refreshNotes()]);
        return;
      }

      status.textContent =
        `Synchronisation CRM en cours — vérification ${attempt}/12. ` +
        "La synchronisation serveur peut prendre jusqu’à environ 15 minutes.";
      status.className = "status-message muted";
      if (attempt < 12) await delay(10_000);
    }

    status.textContent =
      "La catégorie est active, mais le CRM n’a pas encore terminé la synchronisation. " +
      "Vous pouvez fermer puis rouvrir ce panneau plus tard, ou relancer la recherche.";
    status.className = "status-message warning";
    button.textContent = "Rechercher à nouveau dans le CRM";
  } catch (error) {
    console.error(error);
    status.textContent =
      `${error.message} L’autorisation Microsoft Graph Calendars.ReadWrite peut devoir être accordée par un administrateur.`;
    status.className = "status-message warning";
  } finally {
    actionBusy = false; updateActionControls();
  }
}

function resetRegardingSearch(clearText = true) {
  regardingSearchResults = [];
  const text = document.querySelector("#regardingSearchText");
  const result = document.querySelector("#regardingResultSelect");
  const label = document.querySelector("#regardingResultsLabel");
  const button = document.querySelector("#saveRegardingButton");
  if (clearText && text) text.value = "";
  if (result) result.innerHTML = "";
  label?.classList.add("hidden");
  if (button) button.disabled = true;
  const status = document.querySelector("#regardingStatus");
  if (status) status.textContent = "";
}

function updateRegardingSearchPlaceholder() {
  const type = document.querySelector("#regardingTypeSelect")?.value;
  const input = document.querySelector("#regardingSearchText");
  if (!input) return;
  input.placeholder = type === "contact"
    ? "Nom, prénom ou adresse e-mail"
    : type === "opportunity"
      ? "Numéro d’opportunité ou affaire"
      : "Nom ou numéro de compte";
}

async function searchRegarding() {
  const logicalName = document.querySelector("#regardingTypeSelect").value;
  const searchText = document.querySelector("#regardingSearchText").value.trim();
  const searchButton = document.querySelector("#searchRegardingButton");
  const resultSelect = document.querySelector("#regardingResultSelect");
  const resultLabel = document.querySelector("#regardingResultsLabel");
  const status = document.querySelector("#regardingStatus");

  if (searchText.replace(/[%*]/g, "").trim().length < 2) {
    status.textContent = "Saisissez au moins 2 caractères hors jokers.";
    status.className = "status-message warning";
    return;
  }

  searchButton.disabled = true;
  searchButton.textContent = "Recherche…";
  status.textContent = "";
  resultLabel.classList.add("hidden");
  document.querySelector("#saveRegardingButton").disabled = true;

  try {
    regardingSearchResults = await searchRegardingRecords(
      config,
      accessToken,
      logicalName,
      searchText,
    );
    if (!regardingSearchResults.length) {
      status.textContent = "Aucun résultat trouvé.";
      status.className = "status-message warning";
      return;
    }

    resultSelect.innerHTML = '<option value="">— Sélectionner —</option>';
    for (const record of regardingSearchResults) {
      const option = document.createElement("option");
      option.value = record.id;
      option.textContent = record.detail ? `${record.label} — ${record.detail}` : record.label;
      resultSelect.append(option);
    }
    resultLabel.classList.remove("hidden");
    status.textContent = `${regardingSearchResults.length} résultat(s).`;
    status.className = "status-message muted";
  } catch (error) {
    console.error(error);
    status.textContent = error.message || "La recherche a échoué.";
    status.className = "status-message warning";
  } finally {
    searchButton.disabled = !selectedAppointment;
    searchButton.textContent = "Rechercher dans le CRM";
  }
}

async function saveRegarding() {
  if (actionBusy) return;
  const resultSelect = document.querySelector("#regardingResultSelect");
  const button = document.querySelector("#saveRegardingButton");
  const status = document.querySelector("#regardingStatus");
  const regarding = regardingSearchResults.find((record) => record.id === resultSelect.value);

  if (!selectedAppointment || !regarding) {
    status.textContent = "Sélectionnez un résultat.";
    status.className = "status-message warning";
    return;
  }

  actionBusy = true; updateActionControls();
  button.textContent = "Mise à jour…";
  status.textContent = "";

  try {
    assertAppointmentContext(context);
    const series = selectedAppointment.editTarget?.kind === "series";
    if (series && !await confirmAction("Modifier le Concernant de toute la série ?",
      `La série Dynamics entière sera rattachée à « ${regarding.label} ». Les exceptions sont conservées ; aucune note existante n’est déplacée.`, "Modifier toute la série")) return;
    accessToken = await getDataverseAccessToken(config);
    const updatedAppointment = await updateAppointmentRegarding(
      config,
      accessToken,
      selectedAppointment.activityid,
      regarding,
      selectedAppointment,
      () => assertAppointmentContext(context),
    );
    selectedAppointment = updatedAppointment;
    updateSelectedAppointment();
    status.textContent = `Concernant ${series ? "de toute la série " : ""}confirmé : ${getRegardingLabel(selectedAppointment)}.`;
    status.className = "status-message success";
    await refreshNotes();
  } catch (error) {
    console.error(error);
    status.textContent = error.message || "La mise à jour du Concernant a échoué.";
    status.className = "status-message warning";
  } finally {
    actionBusy = false; updateActionControls();
    button.textContent = selectedAppointment?.editTarget?.kind === "series" ? "Affecter à toute la série" : "Affecter ce Concernant";
  }
}

function updateSelectedAppointment() {
  const regarding = getRegardingReference(selectedAppointment);
  const noteEnabled = Boolean(selectedAppointment && regarding);
  document.querySelector("#noteSubject").disabled = !noteEnabled;
  document.querySelector("#noteText").disabled = !noteEnabled;
  document.querySelector("#saveButton").disabled = !noteEnabled;
  document.querySelector("#saveCrmFieldsButton").disabled =
    !selectedAppointment || document.querySelector("#qualificationSelect").value === "";
  const regardingTypeSelect = document.querySelector("#regardingTypeSelect");
  const regardingSearchText = document.querySelector("#regardingSearchText");
  const searchRegardingButton = document.querySelector("#searchRegardingButton");
  regardingTypeSelect.disabled = !selectedAppointment;
  regardingSearchText.disabled = !selectedAppointment;
  searchRegardingButton.disabled = !selectedAppointment;

  const currentRegardingStatus = document.querySelector("#currentRegardingStatus");
  if (regarding) {
    currentRegardingStatus.textContent = `Actuel : ${regarding.label}`;
    currentRegardingStatus.className = "success";
    if (["account", "contact", "opportunity"].includes(regarding.logicalName)) {
      regardingTypeSelect.value = regarding.logicalName;
    }
  } else if (selectedAppointment) {
    currentRegardingStatus.textContent = "Aucun Concernant — renseignez-en un pour activer les notes.";
    currentRegardingStatus.className = "warning";
  } else {
    currentRegardingStatus.textContent = "Aucun rendez-vous CRM sélectionné.";
    currentRegardingStatus.className = "muted";
  }
  resetRegardingSearch(false);
  updateRegardingSearchPlaceholder();

  const noteTarget = document.querySelector("#noteTarget");
  if (regarding) {
    noteTarget.textContent = `La note sera enregistrée sur le Concernant : ${regarding.label}.`;
    noteTarget.className = "success";
  } else if (selectedAppointment) {
    noteTarget.textContent =
      "Ce rendez-vous n’a aucun Concernant. L’enregistrement d’une note est désactivé.";
    noteTarget.className = "warning";
  } else {
    noteTarget.textContent = "Sélectionnez d’abord un rendez-vous CRM.";
    noteTarget.className = "muted";
  }

  if (!selectedAppointment) {
    document.querySelector("#notesList").innerHTML = '<p class="muted">Aucun rendez-vous CRM sélectionné.</p>';
    const qualificationSelect = document.querySelector("#qualificationSelect");
    qualificationSelect.disabled = true;
    qualificationSelect.innerHTML =
      '<option value="">Sélectionnez d’abord un rendez-vous CRM</option>';
    document.querySelector("#crmFieldsStatus").textContent = "";
  } else if (!regarding) {
    document.querySelector("#notesList").innerHTML =
      '<p class="warning">Aucune note disponible : le rendez-vous n’a pas de Concernant.</p>';
  }
  const target = selectedAppointment?.editTarget;
  const series = target?.kind === "series";
  document.querySelector("#qualificationScopeLabel").textContent = series ? "Qualification de la série" : "Qualification du RDV";
  document.querySelector("#regardingScopeLabel").textContent = series ? "Concernant de la série" : "Concernant";
  document.querySelector("#saveCrmFieldsButton").textContent = series ? "Modifier la qualification de toute la série" : "Mettre à jour la qualification";
  document.querySelector("#saveRegardingButton").textContent = series ? "Affecter à toute la série" : "Affecter ce Concernant";
  document.querySelector("#seriesScopeStatus").textContent = target?.kind === "blocked" ? target.reason : selectedAppointment?.seriesNotice || "";
  if (selectedAppointment?.noteBlockedReason) {
    noteTarget.textContent = selectedAppointment.noteBlockedReason;
    noteTarget.className = "warning";
  }
  updateActionControls();
}

function qualificationLabel(value) {
  return qualificationOptions?.find((option) => option.value === Number(value))?.label || "";
}

async function refreshCrmFields() {
  const appointmentId = selectedAppointment?.activityid;
  const select = document.querySelector("#qualificationSelect");
  const button = document.querySelector("#saveCrmFieldsButton");
  const status = document.querySelector("#crmFieldsStatus");

  if (!appointmentId) return;

  select.disabled = true;
  button.disabled = true;
  select.innerHTML = '<option value="">Chargement des choix CRM…</option>';
  status.textContent = "";

  try {
    qualificationOptions = await getAppointmentQualificationOptions(config, accessToken, selectedAppointment);
    if (selectedAppointment?.activityid !== appointmentId) return;

    select.innerHTML = [
      '<option value="">— Sélectionner —</option>',
      ...qualificationOptions.map(
        (option) =>
          `<option value="${escapeHtml(option.value)}">${escapeHtml(option.label)}</option>`,
      ),
    ].join("");

    const currentValue = selectedAppointment[APPOINTMENT_QUALIFICATION_FIELD];
    select.value = currentValue == null ? "" : String(currentValue);
    select.disabled = false;
    button.disabled = select.value === "";

    const formattedValue =
      selectedAppointment[
        `${APPOINTMENT_QUALIFICATION_FIELD}@OData.Community.Display.V1.FormattedValue`
      ] || qualificationLabel(currentValue);
    status.textContent = formattedValue ? `Valeur actuelle : ${formattedValue}` : "Aucune qualification renseignée.";
    status.className = "status-message muted";
  } catch (error) {
    console.error(error);
    status.textContent = `Impossible de charger les qualifications : ${error.message}`;
    status.className = "status-message warning";
  } finally {
    updateActionControls();
  }
}

async function saveCrmFields() {
  if (actionBusy) return;
  const select = document.querySelector("#qualificationSelect");
  const button = document.querySelector("#saveCrmFieldsButton");
  const status = document.querySelector("#crmFieldsStatus");
  const value = select.value;

  if (!selectedAppointment || value === "") {
    status.textContent = "Sélectionnez une qualification.";
    status.className = "status-message warning";
    return;
  }

  actionBusy = true; updateActionControls();
  button.textContent = "Mise à jour…";
  status.textContent = "";

  try {
    assertAppointmentContext(context);
    const series = selectedAppointment.editTarget?.kind === "series";
    if (series && !await confirmAction("Modifier la qualification de toute la série ?",
      "La modification vise le maître Dynamics de toute la série. Les exceptions existantes sont conservées ; la propagation à chaque séance n’est pas garantie.", "Modifier toute la série")) return;
    accessToken = await getDataverseAccessToken(config);
    const updatedAppointment = await updateAppointmentQualification(
      config,
      accessToken,
      selectedAppointment.activityid,
      value,
      selectedAppointment,
      () => assertAppointmentContext(context),
    );
    selectedAppointment = updatedAppointment;
    updateSelectedAppointment();
    status.textContent = `Qualification ${series ? "de toute la série " : ""}confirmée : ${qualificationLabel(value)}. ${selectedAppointment.seriesNotice || ""}`;
    status.className = "status-message success";
  } catch (error) {
    console.error(error);
    status.textContent = error.message || "La mise à jour de la qualification a échoué.";
    status.className = "status-message warning";
  } finally {
    actionBusy = false; updateActionControls();
    button.textContent = selectedAppointment?.editTarget?.kind === "series" ? "Modifier la qualification de toute la série" : "Mettre à jour la qualification";
  }
}

async function refreshNotes() {
  const request = ++notesRequest;
  const appointment = selectedAppointment;
  const originalContext = context;
  const notesList = document.querySelector("#notesList");
  const regarding = getRegardingReference(selectedAppointment);
  if (selectedAppointment?.noteBlockedReason) {
    notesList.innerHTML = `<p class="warning">${escapeHtml(selectedAppointment.noteBlockedReason)}</p>`;
    return;
  }
  if (!regarding) {
    notesList.innerHTML =
      '<p class="warning">Aucune note disponible : le rendez-vous n’a pas de Concernant.</p>';
    return;
  }
  notesList.innerHTML = '<p class="muted">Chargement des notes…</p>';
  try {
    const notes = await getAppointmentNotes(config, accessToken, regarding, appointment.activityid);
    if (request !== notesRequest || selectedAppointment !== appointment || context !== originalContext || !notesList.isConnected) return;
    assertAppointmentContext(originalContext);
    if (!notes.length) {
      notesList.innerHTML = '<p class="muted">Aucune note identifiée pour cette séance.</p>';
      return;
    }
    notesList.innerHTML = notes
      .map(
        (note) => `
          <article class="note-item">
            <h3>${escapeHtml(note.subject || "Note")}</h3>
            <time>${escapeHtml(formatDate(note.createdon))}</time>
            <div class="note-content">${renderSafeNoteHtml(withoutAppointmentReference(note.notetext))}</div>
          </article>`,
      )
      .join("");
  } catch (error) {
    if (request !== notesRequest || selectedAppointment !== appointment || context !== originalContext || !notesList.isConnected) return;
    notesList.innerHTML = `<p class="warning">${escapeHtml(error.message)}</p>`;
  }
}

async function saveNote() {
  if (actionBusy) return;
  const button = document.querySelector("#saveButton");
  const status = document.querySelector("#saveStatus");
  const subject = document.querySelector("#noteSubject").value.trim();
  const text = document.querySelector("#noteText").value.trim();
  const regarding = getRegardingReference(selectedAppointment);

  if (!selectedAppointment) {
    status.textContent = "Sélectionnez le rendez-vous CRM.";
    status.className = "status-message warning";
    return;
  }
  if (!regarding) {
    status.textContent =
      "Ce rendez-vous n’a aucun Concernant. La note ne peut pas être enregistrée.";
    status.className = "status-message warning";
    return;
  }
  if (!text) {
    status.textContent = "La note est vide.";
    status.className = "status-message warning";
    document.querySelector("#noteText").focus();
    return;
  }

  actionBusy = true; updateActionControls();
  button.textContent = "Enregistrement…";
  status.textContent = "";

  try {
    const appointment = selectedAppointment;
    const originalContext = context;
    assertAppointmentContext(context);
    if (!context.graphEvent) throw new Error("La portée Outlook doit être vérifiée avant de créer une note.");
    if (!await confirmAction("Enregistrer cette note manuelle ?", `La note sera rattachée à « ${regarding.label} », pas à toutes les séances d’une série. Aucun audio ni génération automatique.`, "Enregistrer la note")) return;
    accessToken = await getDataverseAccessToken(config);
    if (selectedAppointment !== appointment || context !== originalContext) throw new Error("Le rendez-vous sélectionné a changé. Aucun nouvel envoi.");
    const verified = await verifyAppointmentNoteTarget(config, accessToken, appointment);
    assertAppointmentContext(context);
    if (selectedAppointment !== appointment || context !== originalContext) throw new Error("Le rendez-vous sélectionné a changé. Aucun nouvel envoi.");
    await createRegardingNote(config, accessToken, verified, {
      subject: subject || `Note — ${context.subject}`,
      text,
    }, appointment.activityid);
    let markerWarning = "";
    try {
      await markAppointmentAsLogged();
    } catch (error) {
      console.warn("La note est créée, mais le libellé Outlook n’a pas pu être actualisé.", error);
      markerWarning = " La note est bien enregistrée, mais Outlook n’a pas actualisé le libellé Journal/Afficher.";
    }
    document.querySelector("#noteText").value = "";
    status.textContent = `Note enregistrée sur ${regarding.label}.${markerWarning}`;
    status.className = "status-message success";
    await refreshNotes();
  } catch (error) {
    console.error(error);
    status.textContent = error.message || "L’enregistrement a échoué.";
    status.className = "status-message warning";
  } finally {
    actionBusy = false; updateActionControls();
    button.textContent = "Enregistrer la note dans Dynamics";
  }
}

async function start() {
  try {
    interruptedDiagnostic = beginStartupDiagnostic(packageMetadata.version);
    setLoadingPhase("Initialisation d’Outlook…", `Diagnostic v${packageMetadata.version}`);

    recordStartupPhase("configuration");
    config = getConfig();
    context = await getAppointmentContext({
      onPhase: (phase, attempt) => {
        recordStartupPhase(phase, { attempt });
        const suffix = attempt > 1 ? ` — tentative ${attempt}/2` : "";
        setLoadingPhase(`${PHASE_LABELS[phase] || "Lecture du rendez-vous"}${suffix}…`);
      },
      onRetry: (_error, nextAttempt) => {
        setLoadingPhase(
          "Outlook tarde à répondre…",
          `Nouvelle tentative automatique ${nextAttempt}/2`,
        );
      },
    });
    renderShell();

    recordStartupPhase("authentication");
    accessToken = await getDataverseAccessToken(config);
    document.querySelector("#connectionStatus").textContent = "CRM connecté";
    document.querySelector("#connectionStatus").classList.add("connected");

    await refreshGraphContext();

    recordStartupPhase("dataverse_search");
    candidates = await findAppointmentCandidates(config, accessToken, context);
    await renderCandidates();
    if (selectedAppointment) {
      recordStartupPhase("dataverse_notes");
      await Promise.all([refreshCrmFields(), refreshNotes()]);
    }
    completeStartupDiagnostic();
  } catch (error) {
    showFatalError(error);
  }
}

async function bootstrap() {
  if (await redirectToLatestVersion(packageMetadata.version)) return;
  await start();
}

bootstrap();
