import "./styles.css";
import packageMetadata from "../package.json";
import { getConfig } from "./config.js";
import { getDataverseAccessToken } from "./auth.js";
import {
  beginStartupDiagnostic,
  clearInterruptedStartupDiagnostic,
  completeStartupDiagnostic,
  getCurrentStartupDiagnostic,
  recordStartupFailure,
  recordStartupPhase,
} from "./diagnostics.js";
import {
  chooseUnambiguousCandidate,
  createAppointmentNote,
  findAppointmentCandidates,
  getAppointmentNotes,
  getRegardingLabel,
} from "./dataverse.js";
import { getAppointmentContext, markAppointmentAsLogged } from "./outlook.js";

const app = document.querySelector("#app");
let config;
let context;
let accessToken;
let candidates = [];
let selectedAppointment = null;
let interruptedDiagnostic = null;

const PHASE_LABELS = {
  application_start: "Démarrage de l’application",
  configuration: "Lecture de la configuration",
  office_library: "Chargement de la bibliothèque Office",
  outlook_initialization: "Initialisation d’Outlook",
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
      <h1>Impossible d’ouvrir les notes CRM</h1>
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
        <p class="eyebrow">RENDEZ-VOUS CRM</p>
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
    </section>

    <section class="card" id="matchCard">
      <h2>Rendez-vous correspondant dans le CRM</h2>
      <p id="matchStatus" class="muted">Recherche en cours…</p>
      <label id="candidateLabel" class="field hidden">
        <span>Choisir le rendez-vous</span>
        <select id="candidateSelect"></select>
      </label>
    </section>

    <section class="card" id="noteCard">
      <h2>Ajouter une note</h2>
      <label class="field">
        <span>Titre</span>
        <input id="noteSubject" maxlength="500" value="Compte rendu — ${escapeHtml(context.subject)}" />
      </label>
      <label class="field">
        <span>Note</span>
        <textarea id="noteText" rows="7" maxlength="100000" placeholder="Saisissez le compte rendu, les décisions et les actions à réaliser…"></textarea>
      </label>
      <button id="saveButton" class="primary-button" type="button" disabled>Enregistrer dans le CRM</button>
      <p id="saveStatus" class="status-message" role="status"></p>
    </section>

    <section class="card" id="historyCard">
      <h2>Notes existantes</h2>
      <div id="notesList"><p class="muted">Sélectionnez d’abord un rendez-vous CRM.</p></div>
    </section>`;

  document.querySelector("#candidateSelect").addEventListener("change", async (event) => {
    selectedAppointment = candidates.find((candidate) => candidate.activityid === event.target.value) || null;
    updateSelectedAppointment();
    if (selectedAppointment) await refreshNotes();
  });
  document.querySelector("#saveButton").addEventListener("click", saveNote);
  if (interruptedDiagnostic) clearInterruptedStartupDiagnostic();
}

function candidateText(candidate) {
  const start = formatDate(candidate.scheduledstart);
  return `${candidate.subject || "Sans objet"} — ${start} — ${getRegardingLabel(candidate)}`;
}

function renderCandidates() {
  const status = document.querySelector("#matchStatus");
  const label = document.querySelector("#candidateLabel");
  const select = document.querySelector("#candidateSelect");
  selectedAppointment = chooseUnambiguousCandidate(candidates);

  if (!candidates.length) {
    status.textContent = "Aucun rendez-vous CRM trouvé autour de cet horaire. La note ne peut pas être enregistrée.";
    status.className = "warning";
    label.classList.add("hidden");
    updateSelectedAppointment();
    return;
  }

  select.innerHTML = [
    `<option value="">— Choisir —</option>`,
    ...candidates.map(
      (candidate) =>
        `<option value="${escapeHtml(candidate.activityid)}">${escapeHtml(candidateText(candidate))}</option>`,
    ),
  ].join("");

  if (selectedAppointment) {
    select.value = selectedAppointment.activityid;
    status.textContent = `Correspondance fiable : ${candidateText(selectedAppointment)}`;
    status.className = "success";
    label.classList.toggle("hidden", candidates.length === 1);
  } else {
    status.textContent = "Plusieurs rendez-vous sont possibles. Choisissez le bon avant d’enregistrer.";
    status.className = "warning";
    label.classList.remove("hidden");
  }

  updateSelectedAppointment();
}

function updateSelectedAppointment() {
  document.querySelector("#saveButton").disabled = !selectedAppointment;
  if (!selectedAppointment) {
    document.querySelector("#notesList").innerHTML = '<p class="muted">Aucun rendez-vous CRM sélectionné.</p>';
  }
}

async function refreshNotes() {
  const notesList = document.querySelector("#notesList");
  notesList.innerHTML = '<p class="muted">Chargement des notes…</p>';
  try {
    const notes = await getAppointmentNotes(config, accessToken, selectedAppointment.activityid);
    if (!notes.length) {
      notesList.innerHTML = '<p class="muted">Aucune note pour ce rendez-vous.</p>';
      return;
    }
    notesList.innerHTML = notes
      .map(
        (note) => `
          <article class="note-item">
            <h3>${escapeHtml(note.subject || "Note")}</h3>
            <time>${escapeHtml(formatDate(note.createdon))}</time>
            <p>${escapeHtml(note.notetext || "").replaceAll("\n", "<br>")}</p>
          </article>`,
      )
      .join("");
  } catch (error) {
    notesList.innerHTML = `<p class="warning">${escapeHtml(error.message)}</p>`;
  }
}

async function saveNote() {
  const button = document.querySelector("#saveButton");
  const status = document.querySelector("#saveStatus");
  const subject = document.querySelector("#noteSubject").value.trim();
  const text = document.querySelector("#noteText").value.trim();

  if (!selectedAppointment) {
    status.textContent = "Sélectionnez le rendez-vous CRM.";
    status.className = "status-message warning";
    return;
  }
  if (!text) {
    status.textContent = "La note est vide.";
    status.className = "status-message warning";
    document.querySelector("#noteText").focus();
    return;
  }

  button.disabled = true;
  button.textContent = "Enregistrement…";
  status.textContent = "";

  try {
    await createAppointmentNote(config, accessToken, selectedAppointment.activityid, {
      subject: subject || `Compte rendu — ${context.subject}`,
      text,
    });
    let markerWarning = "";
    try {
      await markAppointmentAsLogged();
    } catch (error) {
      console.warn("La note est créée, mais le libellé Outlook n’a pas pu être actualisé.", error);
      markerWarning = " La note est bien enregistrée, mais Outlook n’a pas actualisé le libellé Journal/Afficher.";
    }
    document.querySelector("#noteText").value = "";
    status.textContent = `Note enregistrée dans le rendez-vous CRM.${markerWarning}`;
    status.className = "status-message success";
    await refreshNotes();
  } catch (error) {
    console.error(error);
    status.textContent = error.message || "L’enregistrement a échoué.";
    status.className = "status-message warning";
  } finally {
    button.disabled = !selectedAppointment;
    button.textContent = "Enregistrer dans le CRM";
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

    recordStartupPhase("dataverse_search");
    candidates = await findAppointmentCandidates(config, accessToken, context);
    renderCandidates();
    if (selectedAppointment) {
      recordStartupPhase("dataverse_notes");
      await refreshNotes();
    }
    completeStartupDiagnostic();
  } catch (error) {
    showFatalError(error);
  }
}

start();
