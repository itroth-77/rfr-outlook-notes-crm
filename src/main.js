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
import { TRACKED_TO_DYNAMICS_CATEGORY, isRecurringEvent, resolveOutlookEvent, readTrackingState, trackOutlookAppointment } from "./graph.js";
import { getAppointmentContext, watchAppointmentChanges, markAppointmentAsLogged } from "./outlook.js";
import { mobileAppointmentLink, mobileLaunchLink } from "./mobile-link.js";
import { AppointmentSession } from "./appointment-session.js";
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
let trackingState = null;
let regardingRequest = 0;
const session = new AppointmentSession();

function rememberDraft() {
  session.rememberDraft(document.querySelector("#noteSubject")?.value || "", document.querySelector("#noteText")?.value || "");
}

function showContextChange(error) {
  if (error?.code !== "OUTLOOK_CONTEXT_CHANGED") return false;
  const banner = document.querySelector("#contextStatus");
  if (!banner) return false; // Before the first usable Office snapshot: startup retry screen.
  session.invalidate();
  if (banner) {
    banner.textContent = "Le rendez-vous ouvert dans Outlook a changé. Actualisez pour afficher le rendez-vous courant. Votre brouillon reste conservé dans ce volet.";
    banner.classList.remove("hidden");
  }
  const matchStatus = document.querySelector("#matchStatus");
  matchStatus.textContent = "Informations à actualiser.";
  matchStatus.className = "muted";
  document.querySelector("#openMobileApp")?.classList.add("hidden");
  updateActionControls();
  return true;
}

function currentScope() { const scope = session.capture(); scope.ensure(); return scope; }
function contextError(error, status) {
  if (showContextChange(error)) return;
  status.textContent = error.message || "L’opération a échoué. Actualisez avant de réessayer.";
  status.className = "status-message warning";
}

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
  const enabled = !actionBusy && !session.invalid && Boolean(selectedAppointment && context.graphEvent && target?.kind !== "blocked");
  const regarding = getRegardingReference(selectedAppointment);
  const noteEnabled = enabled && Boolean(regarding) && !selectedAppointment?.noteBlockedReason;
  for (const id of ["noteSubject", "noteText", "saveButton"]) document.querySelector("#" + id).disabled = !noteEnabled;
  const qualified = enabled && !target?.qualificationError && Boolean(qualificationOptions?.length);
  document.querySelector("#qualificationSelect").disabled = !qualified;
  const value = document.querySelector("#qualificationSelect").value;
  document.querySelector("#saveCrmFieldsButton").disabled = !qualified || value === "" || value === String(selectedAppointment?.[APPOINTMENT_QUALIFICATION_FIELD] ?? "");
  for (const id of ["regardingTypeSelect", "regardingSearchText", "searchRegardingButton"]) document.querySelector("#" + id).disabled = !enabled;
  document.querySelector("#saveRegardingButton").disabled = !enabled || !document.querySelector("#regardingResultSelect").value;
  document.querySelector("#trackAppointmentButton").disabled = actionBusy || session.invalid || !context.graphEvent;
  document.querySelector("#refreshAppointmentButton").disabled = actionBusy;
  document.querySelector("#refreshNotesButton").disabled = !enabled;
  document.querySelector("#openMobileApp")?.setAttribute("aria-disabled", String(actionBusy || session.invalid));
}

const PHASE_LABELS = {
  application_start: "Démarrage de l’application",
  configuration: "Lecture de la configuration",
  office_library: "Chargement de la bibliothèque Office",
  outlook_initialization: "Initialisation d’Outlook",
  outlook_context_direct: "Contexte du rendez-vous déjà disponible",
  outlook_item: "Accès au rendez-vous Outlook",
  outlook_item_id: "Lecture de l’identifiant Outlook (lecture ou modification)",
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
    <header class="app-header sticky-header">
      <div>
        <p class="eyebrow">RDV &amp; Notes</p>
        <span class="status-chip" id="connectionStatus">Connexion…</span>
      </div>
      <div class="header-status">
        <button id="refreshAppointmentButton" class="compact-button" type="button">↻ Actualiser</button>
        <details class="about-menu"><summary>À propos</summary><p>v${escapeHtml(packageMetadata.version)} · ${escapeHtml(environmentLabel)}</p>
          <p>${escapeHtml(context.mailboxEmail)}</p></details>
      </div>
    </header>

    <p id="contextStatus" class="context-banner warning hidden" role="alert"></p>

    ${diagnosticMarkup(interruptedDiagnostic, "La tentative précédente a été interrompue par Outlook")}

    <section class="card appointment-summary">
      <p class="eyebrow">Rendez-vous Outlook</p>
      <h1>${escapeHtml(context.subject)}</h1>
      <dl>
        <div><dt>Début</dt><dd>${escapeHtml(formatDate(context.start))}</dd></div>
        <div><dt>Fin</dt><dd>${escapeHtml(formatDate(context.end))}</dd></div>
        ${context.location ? `<div><dt>Lieu</dt><dd>${escapeHtml(context.location)}</dd></div>` : ""}
        ${context.organizerName ? `<div><dt>Organisateur</dt><dd>${escapeHtml(context.organizerName)}</dd></div>` : ""}
      </dl>
      <p id="recurrenceStatus" class="recurrence-badge" role="status"></p>
      ${context.isMobile ? `<a id="openMobileApp" class="secondary-button app-link hidden" target="_blank" rel="noopener noreferrer" referrerpolicy="no-referrer">Ouvrir dans RDV &amp; Notes</a>
        <p id="mobileLinkStatus" class="hint" role="status">Identification du rendez-vous…</p>
        <details class="mobile-help"><summary>Si Outlook bloque l’ouverture</summary>
          <p>Certains Outlook mobiles empêchent de quitter leur fenêtre intégrée. Copiez ce lien HTTPS et ouvrez-le avec le navigateur du téléphone. Le lien des mails reste utilisable.</p>
          <input id="mobileCopyLink" aria-label="Lien du rendez-vous pour l’application" readonly />
          <button id="copyMobileLink" class="compact-button" type="button">Copier le lien</button>
        </details>` : ""}
    </section>

    <section class="card" id="matchCard">
      <h2>Dynamics</h2>
      <p id="matchStatus" class="muted">Recherche en cours…</p>
      <details id="seriesScopeDetails" class="hidden"><summary>Portée et exceptions de la série</summary><p id="seriesScopeStatus" class="hint" role="status"></p></details>
      <button id="trackAppointmentButton" class="primary-button hidden" type="button">
        Activer « ${TRACKED_TO_DYNAMICS_CATEGORY} »
      </button>
      <p id="trackingStatus" class="status-message" role="status"></p>
    </section>

    <section class="card" id="crmFieldsCard">
      <h2>Qualification et Concernant</h2>
      <label class="field">
        <span id="qualificationScopeLabel">Qualification du RDV</span>
        <select id="qualificationSelect" disabled>
          <option value="">Sélectionnez d’abord un rendez-vous CRM</option>
        </select>
      </label>
      <button id="saveCrmFieldsButton" class="primary-button" type="button" disabled>
        Enregistrer la qualification
      </button>
      <p id="crmFieldsStatus" class="status-message" role="status"></p>

      <div class="field-separator"></div>
      <h3 id="regardingScopeLabel">Concernant</h3>
      <p id="currentRegardingStatus" class="muted">Aucun rendez-vous CRM sélectionné.</p>
      <details id="regardingEditor"><summary>Modifier le Concernant</summary>
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
      </details>
    </section>

    <section class="card" id="noteCard">
      <details id="noteEditor"><summary>Ajouter une note manuelle</summary>
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
      </details>
    </section>

    <section class="card" id="historyCard">
      <details id="historyDetails"><summary>Notes et CR de ce rendez-vous</summary>
      <p class="hint">Cette séance uniquement · 20 dernières notes maximum.</p>
      <button id="refreshNotesButton" class="compact-button" type="button">↻ Actualiser les notes</button>
      <div id="notesList"><p class="muted">Sélectionnez d’abord un rendez-vous CRM.</p></div>
      <p class="hint">Les anciennes notes sans référence à cette séance restent accessibles dans Dynamics.</p>
      </details>
    </section>`;

  document.querySelector("#trackAppointmentButton").addEventListener("click", trackAndWaitForAppointment);
  document.querySelector("#refreshAppointmentButton").addEventListener("click", refreshCurrentMeeting);
  document.querySelector("#refreshNotesButton").addEventListener("click", refreshNotes);
  document.querySelector("#historyDetails").addEventListener("toggle", () => {
    if (document.querySelector("#historyDetails").open && !actionBusy && !session.invalid) refreshNotes();
  });
  for (const id of ["noteSubject", "noteText"]) document.querySelector("#" + id).addEventListener("input", rememberDraft);
  document.querySelector("#qualificationSelect").addEventListener("change", updateActionControls);
  document.querySelector("#openMobileApp")?.addEventListener("click", (event) => {
    try {
      currentScope();
      if (actionBusy || document.querySelector("#noteText").value.trim()) throw new Error("Terminez l’action et enregistrez votre note avant d’ouvrir l’application mobile.");
      document.querySelector("#mobileLinkStatus").textContent = "Ouverture demandée. Si Outlook la bloque, utilisez l’aide ci-dessous ou le lien reçu par mail.";
    } catch (error) { event.preventDefault(); contextError(error, document.querySelector("#mobileLinkStatus")); }
  });
  document.querySelector("#copyMobileLink")?.addEventListener("click", async () => {
    try {
      currentScope();
      const input = document.querySelector("#mobileCopyLink");
      if (!input.value) return;
      input.select();
      await navigator.clipboard.writeText(input.value);
      document.querySelector("#mobileLinkStatus").textContent = "Lien copié. Ouvrez-le avec le navigateur du téléphone.";
    } catch (error) {
      if (!showContextChange(error)) document.querySelector("#mobileLinkStatus").textContent = "Sélectionnez le lien ci-dessous puis choisissez Copier.";
    }
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

async function renderCandidates(scope = currentScope()) {
  await scope.ensureCurrent();
  const status = document.querySelector("#matchStatus");
  const trackButton = document.querySelector("#trackAppointmentButton");
  const candidate = chooseUnambiguousCandidate(candidates);
  selectedAppointment = null;
  qualificationOptions = null;
  updateSelectedAppointment();
  const strongCandidates = candidates.filter((candidate) => candidate.matchScore >= 85);

  if (candidate) {
    const resolved = await resolveAppointmentEditingTarget(config, accessToken, candidate, isRecurringEvent(context.graphEvent));
    await scope.ensureCurrent();
    selectedAppointment = resolved;
    status.textContent = "✓ Rendez-vous synchronisé avec Dynamics";
    status.className = "success";
    trackButton.classList.add("hidden");
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
    trackButton.textContent = trackingState?.tracked ? "Vérifier la synchronisation Dynamics"
      : isRecurringEvent(context.graphEvent) ? "Activer le suivi de toute la série" : "Activer le suivi Dynamics 365";
    if (trackingState?.tracked) status.textContent = "Suivi Outlook activé · synchronisation Dynamics non encore confirmée.";
  }

  updateSelectedAppointment();
}

async function refreshGraphContext(scope = currentScope()) {
  await scope.ensureCurrent();
  context.graphEvent = null;
  trackingState = null;
  const link = document.querySelector("#openMobileApp");
  link?.classList.add("hidden");
  try {
    const token = await getGraphAccessToken(config);
    const event = await resolveOutlookEvent(token, scope.context);
    await scope.ensureCurrent();
    context.graphEvent = event;
    document.querySelector("#recurrenceStatus").textContent = isRecurringEvent(event)
      ? "↻ Rendez-vous récurrent · modifications Dynamics sur la série" : "Rendez-vous simple";
    if (link) {
      const isAndroid = Office.context.platform === "Android" || /OutlookAndroid/i.test(Office.context.mailbox.diagnostics?.hostName || "");
      link.href = mobileLaunchLink(event.id, { isAndroid }); link.classList.remove("hidden");
      document.querySelector("#mobileCopyLink").value = mobileAppointmentLink(event.id);
      document.querySelector("#mobileLinkStatus").textContent = "Ouvre cette séance à l’étape Dynamics dans l’application Android.";
    }
    try { trackingState = await readTrackingState(token, event); await scope.ensureCurrent(); }
    catch (error) { await scope.ensureCurrent(); document.querySelector("#trackingStatus").textContent = "Statut de catégorie non vérifié. Une activation relira la série avant toute modification."; }
  } catch (error) {
    if (showContextChange(error)) throw error;
    await scope.ensureCurrent();
    document.querySelector("#recurrenceStatus").textContent = "Portée Outlook non vérifiée : modifications bloquées. " + error.message;
    if (link) document.querySelector("#mobileLinkStatus").textContent = "Ouverture mobile indisponible : le rendez-vous n’a pas pu être identifié avec certitude.";
  }
}

async function refreshCurrentMeeting() {
  if (actionBusy) return;
  rememberDraft();
  actionBusy = true; updateActionControls();
  try {
    // Always reacquire Office's current item. Testing the stale snapshot here
    // previously made every subsequent refresh fail permanently.
    session.invalidate();
    const freshContext = await getAppointmentContext();
    context = freshContext;
    session.use(context);
    notesRequest++;
    candidates = []; selectedAppointment = null; qualificationOptions = null;
    renderShell();
    restoreDraft();
    updateActionControls();
    const scope = currentScope();
    accessToken = await getDataverseAccessToken(config);
    await scope.ensureCurrent();
    showConnected();
    await loadCurrentMeeting(scope);
  } catch (error) {
    contextError(error, document.querySelector("#matchStatus"));
  } finally { actionBusy = false; updateActionControls(); }
}

function restoreDraft() {
  const draft = session.draft();
  if (!draft) return;
  document.querySelector("#noteSubject").value = draft.subject;
  document.querySelector("#noteText").value = draft.text;
  document.querySelector("#noteEditor").open = true;
}

function showConnected() {
  document.querySelector("#connectionStatus").textContent = "Microsoft connecté";
  document.querySelector("#connectionStatus").classList.add("connected");
}

async function loadCurrentMeeting(scope) {
  await refreshGraphContext(scope);
  await scope.ensureCurrent();
  const found = await findAppointmentCandidates(config, accessToken, scope.context);
  await scope.ensureCurrent();
  candidates = found;
  await renderCandidates(scope);
  if (selectedAppointment) {
    await refreshCrmFields(scope);
    if (document.querySelector("#historyDetails").open) await refreshNotes();
  }
}

async function trackAndWaitForAppointment() {
  if (actionBusy) return;
  if (trackingState?.tracked) { await refreshCurrentMeeting(); return; }
  const button = document.querySelector("#trackAppointmentButton");
  const status = document.querySelector("#trackingStatus");
  actionBusy = true; updateActionControls();
  status.textContent = "Connexion au calendrier Outlook…";
  status.className = "status-message muted";

  try {
    const scope = currentScope();
    const graphToken = await getGraphAccessToken(config);
    const event = await resolveOutlookEvent(graphToken, scope.context);
    await scope.ensureCurrent();
    const series = isRecurringEvent(event);
    if (!await confirmAction(series ? "Suivre toute la série ?" : "Activer le suivi Dynamics ?",
      series ? "La catégorie sera activée sur toute la série dans VOTRE calendrier, organisateur ou invité. Autres catégories et exceptions conservées. La synchronisation Dynamics sera ensuite vérifiée."
        : "La catégorie sera ajoutée uniquement à ce rendez-vous dans votre calendrier. Les autres catégories seront conservées.", series ? "Suivre toute la série" : "Activer le suivi")) return;
    await scope.ensureCurrent();
    const tracking = await trackOutlookAppointment(graphToken, scope.context, event, scope.ensureCurrent);
    await scope.ensureCurrent();
    context.graphEvent = tracking.event;
    trackingState = { tracked: true };
    status.textContent = tracking.alreadyTracked
      ? `La catégorie « ${TRACKED_TO_DYNAMICS_CATEGORY} » est déjà active. Attente de la synchronisation CRM…`
      : `Catégorie « ${TRACKED_TO_DYNAMICS_CATEGORY} » activée. Attente de la synchronisation CRM…`;
    status.textContent = tracking.notice + " Attente de la confirmation Dynamics…";
    status.className = "status-message success";

    for (let attempt = 1; attempt <= 12; attempt += 1) {
      await scope.ensureCurrent();
      const found = await findAppointmentCandidates(config, accessToken, scope.context);
      await scope.ensureCurrent();
      candidates = found;
      if (chooseUnambiguousCandidate(candidates)) {
        await renderCandidates(scope);
        status.textContent = "Le rendez-vous est maintenant disponible dans le CRM.";
        status.className = "status-message success";
        await refreshCrmFields(scope);
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
      "Cliquez sur Actualiser pour vérifier à nouveau plus tard.";
    status.className = "status-message warning";
    button.textContent = "Vérifier la synchronisation Dynamics";
  } catch (error) {
    console.error(error);
    contextError(error, status);
  } finally {
    actionBusy = false; updateActionControls();
  }
}

function resetRegardingSearch(clearText = true) {
  regardingRequest++;
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
  if (actionBusy || session.invalid) return;
  const request = ++regardingRequest;
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
    const scope = currentScope();
    const results = await searchRegardingRecords(
      config,
      accessToken,
      logicalName,
      searchText,
    );
    if (request !== regardingRequest || !resultSelect.isConnected || logicalName !== document.querySelector("#regardingTypeSelect").value) return;
    await scope.ensureCurrent();
    regardingSearchResults = results;
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
    if (request === regardingRequest && status.isConnected) contextError(error, status);
  } finally {
    updateActionControls();
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
    const scope = currentScope();
    const appointment = selectedAppointment;
    const series = selectedAppointment.editTarget?.kind === "series";
    if (series && !await confirmAction("Modifier le Concernant de toute la série ?",
      `La série Dynamics entière sera rattachée à « ${regarding.label} ». Les exceptions sont conservées ; aucune note existante n’est déplacée.`, "Modifier toute la série")) return;
    accessToken = await getDataverseAccessToken(config);
    await scope.ensureCurrent();
    const updatedAppointment = await updateAppointmentRegarding(
      config,
      accessToken,
      appointment.activityid,
      regarding,
      appointment,
      scope.ensureCurrent,
    );
    await scope.ensureCurrent();
    selectedAppointment = updatedAppointment;
    updateSelectedAppointment();
    status.textContent = `Concernant ${series ? "de toute la série " : ""}confirmé : ${getRegardingLabel(selectedAppointment)}.`;
    status.className = "status-message success";
    if (document.querySelector("#historyDetails").open) await refreshNotes();
  } catch (error) {
    console.error(error);
    contextError(error, status);
  } finally {
    actionBusy = false; updateActionControls();
    button.textContent = selectedAppointment?.editTarget?.kind === "series" ? "Affecter à toute la série" : "Affecter ce Concernant";
  }
}

function updateSelectedAppointment() {
  for (const id of ["crmFieldsCard", "noteCard", "historyCard"]) document.querySelector("#" + id).classList.toggle("hidden", !selectedAppointment);
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
  document.querySelector("#seriesScopeDetails").classList.toggle("hidden", !document.querySelector("#seriesScopeStatus").textContent);
  if (target?.kind === "blocked") document.querySelector("#seriesScopeDetails").open = true;
  if (selectedAppointment?.noteBlockedReason) {
    noteTarget.textContent = selectedAppointment.noteBlockedReason;
    noteTarget.className = "warning";
  }
  updateActionControls();
}

function qualificationLabel(value) {
  return qualificationOptions?.find((option) => option.value === Number(value))?.label || "";
}

async function refreshCrmFields(scope = currentScope()) {
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
    const options = await getAppointmentQualificationOptions(config, accessToken, selectedAppointment);
    await scope.ensureCurrent();
    if (selectedAppointment?.activityid !== appointmentId || !select.isConnected) return;
    qualificationOptions = options;

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
    if (select.isConnected) contextError(error, status);
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
    const scope = currentScope();
    const appointment = selectedAppointment;
    const series = selectedAppointment.editTarget?.kind === "series";
    if (series && !await confirmAction("Modifier la qualification de toute la série ?",
      "La modification vise le maître Dynamics de toute la série. Les exceptions existantes sont conservées ; la propagation à chaque séance n’est pas garantie.", "Modifier toute la série")) return;
    accessToken = await getDataverseAccessToken(config);
    await scope.ensureCurrent();
    const updatedAppointment = await updateAppointmentQualification(
      config,
      accessToken,
      appointment.activityid,
      value,
      appointment,
      scope.ensureCurrent,
    );
    await scope.ensureCurrent();
    selectedAppointment = updatedAppointment;
    updateSelectedAppointment();
    status.textContent = `Qualification ${series ? "de toute la série " : ""}confirmée : ${qualificationLabel(value)}. ${selectedAppointment.seriesNotice || ""}`;
    status.className = "status-message success";
  } catch (error) {
    console.error(error);
    contextError(error, status);
  } finally {
    actionBusy = false; updateActionControls();
    button.textContent = selectedAppointment?.editTarget?.kind === "series" ? "Modifier la qualification de toute la série" : "Mettre à jour la qualification";
  }
}

async function refreshNotes() {
  if (!selectedAppointment || session.invalid) return;
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
    const scope = currentScope();
    const notes = await getAppointmentNotes(config, accessToken, regarding, appointment.activityid);
    if (request !== notesRequest || selectedAppointment !== appointment || context !== originalContext || !notesList.isConnected) return;
    await scope.ensureCurrent();
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
            <details><summary>Lire la note / le CR</summary><div class="note-content">${renderSafeNoteHtml(withoutAppointmentReference(note.notetext))}</div></details>
          </article>`,
      )
      .join("");
  } catch (error) {
    if (request !== notesRequest || selectedAppointment !== appointment || context !== originalContext || !notesList.isConnected) return;
    if (!showContextChange(error)) notesList.innerHTML = `<p class="warning">${escapeHtml(error.message)}</p>`;
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
    const scope = currentScope();
    if (!context.graphEvent) throw new Error("La portée Outlook doit être vérifiée avant de créer une note.");
    if (!await confirmAction("Enregistrer cette note manuelle ?", `La note sera rattachée à « ${regarding.label} », pas à toutes les séances d’une série. Aucun audio ni génération automatique.`, "Enregistrer la note")) return;
    accessToken = await getDataverseAccessToken(config);
    if (selectedAppointment !== appointment || context !== originalContext) throw new Error("Le rendez-vous sélectionné a changé. Aucun nouvel envoi.");
    const verified = await verifyAppointmentNoteTarget(config, accessToken, appointment);
    await scope.ensureCurrent();
    if (selectedAppointment !== appointment || context !== originalContext) throw new Error("Le rendez-vous sélectionné a changé. Aucun nouvel envoi.");
    await createRegardingNote(config, accessToken, verified, {
      subject: subject || `Note — ${context.subject}`,
      text,
    }, appointment.activityid, scope.ensureCurrent);
    let markerWarning = "";
    try {
      await markAppointmentAsLogged(originalContext);
    } catch (error) {
      console.warn("La note est créée, mais le libellé Outlook n’a pas pu être actualisé.", error);
      markerWarning = " La note est bien enregistrée, mais Outlook n’a pas actualisé le libellé Journal/Afficher.";
    }
    // The POST may have completed just before an Outlook item change. Never
    // duplicate a successfully created note, nor clear another item's draft.
    session.drafts.delete(session.key(originalContext));
    if (context !== originalContext) return;
    document.querySelector("#noteText").value = "";
    status.textContent = `Note enregistrée sur ${regarding.label}.${markerWarning}`;
    status.className = "status-message success";
    if (document.querySelector("#historyDetails").open && !session.invalid) await refreshNotes();
  } catch (error) {
    console.error(error);
    contextError(error, status);
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
    session.use(context);
    renderShell();
    actionBusy = true; updateActionControls();
    watchAppointmentChanges(onOutlookItemChanged);
    globalThis.addEventListener("focus", onOutlookItemChanged);
    document.addEventListener("visibilitychange", () => { if (!document.hidden) onOutlookItemChanged(); });

    recordStartupPhase("authentication");
    accessToken = await getDataverseAccessToken(config);
    showConnected();
    await loadCurrentMeeting(currentScope());
    completeStartupDiagnostic();
  } catch (error) {
    if (!showContextChange(error)) { showFatalError(error); return; }
  } finally { actionBusy = false; if (document.querySelector("#trackAppointmentButton")) updateActionControls(); }
}

async function onOutlookItemChanged() {
  if (!context || session.invalid) return;
  const snapshot = context;
  try { await session.capture().ensureCurrent(); }
  catch (error) {
    if (context !== snapshot || error.code !== "OUTLOOK_CONTEXT_CHANGED") return;
    rememberDraft();
    showContextChange(error);
    // Do not replace the form while a confirmation or a write is in progress.
    if (!actionBusy) void refreshCurrentMeeting();
  }
}

async function bootstrap() {
  if (await redirectToLatestVersion(packageMetadata.version)) return;
  await start();
}

bootstrap();
