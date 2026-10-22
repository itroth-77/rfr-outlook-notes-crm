const CURRENT_KEY = "rfr-notes-crm.current-startup";
const INTERRUPTED_KEY = "rfr-notes-crm.last-interrupted-startup";

function read(key) {
  try {
    return JSON.parse(globalThis.localStorage?.getItem(key) || "null");
  } catch {
    return null;
  }
}

function write(key, value) {
  try {
    globalThis.localStorage?.setItem(key, JSON.stringify(value));
  } catch {
    // Le diagnostic ne doit jamais empêcher l'ouverture du complément.
  }
}

function remove(key) {
  try {
    globalThis.localStorage?.removeItem(key);
  } catch {
    // Le diagnostic ne doit jamais empêcher l'ouverture du complément.
  }
}

function now() {
  return new Date().toISOString();
}

export function beginStartupDiagnostic(appVersion) {
  const previous = read(CURRENT_KEY);
  if (previous?.status === "running") write(INTERRUPTED_KEY, previous);

  const diagnostic = {
    appVersion,
    startedAt: now(),
    updatedAt: now(),
    status: "running",
    phase: "application_start",
    attempt: 1,
  };
  write(CURRENT_KEY, diagnostic);
  return read(INTERRUPTED_KEY);
}

export function recordStartupPhase(phase, options = {}) {
  const diagnostic = read(CURRENT_KEY) || {};
  write(CURRENT_KEY, {
    ...diagnostic,
    status: "running",
    phase,
    attempt: options.attempt ?? diagnostic.attempt ?? 1,
    updatedAt: now(),
  });
}

export function recordStartupFailure(error) {
  const diagnostic = read(CURRENT_KEY) || {};
  write(CURRENT_KEY, {
    ...diagnostic,
    status: "failed",
    errorCode: error?.code || error?.name || "UNKNOWN_ERROR",
    errorMessage: String(error?.message || error || "Erreur inconnue"),
    updatedAt: now(),
  });
}

export function completeStartupDiagnostic() {
  const diagnostic = read(CURRENT_KEY) || {};
  write(CURRENT_KEY, {
    ...diagnostic,
    status: "complete",
    phase: "ready",
    updatedAt: now(),
  });
}

export function getCurrentStartupDiagnostic() {
  return read(CURRENT_KEY);
}

export function getInterruptedStartupDiagnostic() {
  return read(INTERRUPTED_KEY);
}

export function clearInterruptedStartupDiagnostic() {
  remove(INTERRUPTED_KEY);
}
