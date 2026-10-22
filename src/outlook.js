import { delay, retry, withTimeout } from "./resilience.js";

const OFFICE_READY_TIMEOUT_MS = 8_000;
const ITEM_WAIT_TIMEOUT_MS = 4_000;
const FIELD_READ_TIMEOUT_MS = 5_000;

function outlookError(message, code, cause) {
  const error = new Error(message, cause ? { cause } : undefined);
  error.code = code;
  return error;
}

function officeAsync(invoke, options = {}) {
  const operation = new Promise((resolve, reject) => {
    invoke((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(
          outlookError(
            result.error?.message || "Erreur Outlook inconnue.",
            options.code || "OUTLOOK_ASYNC_ERROR",
          ),
        );
      }
    });
  });

  return withTimeout(
    operation,
    options.timeoutMilliseconds ?? FIELD_READ_TIMEOUT_MS,
    options.timeoutMessage || "Outlook ne répond pas à la lecture du rendez-vous.",
    options.timeoutCode || "OUTLOOK_ASYNC_TIMEOUT",
  );
}

function readDate(value, fieldName) {
  if (value instanceof Date) return Promise.resolve(value);
  if (value && typeof value.getAsync === "function") {
    return officeAsync((done) => value.getAsync(done), {
      code: `OUTLOOK_${fieldName.toUpperCase()}_ERROR`,
      timeoutCode: `OUTLOOK_${fieldName.toUpperCase()}_TIMEOUT`,
      timeoutMessage: `Outlook n’a pas répondu lors de la lecture du champ « ${fieldName} ».`,
    });
  }
  return Promise.resolve(value ? new Date(value) : null);
}

function readText(value, fieldName) {
  if (typeof value === "string") return Promise.resolve(value);
  if (value && typeof value.getAsync === "function") {
    return officeAsync((done) => value.getAsync(done), {
      code: `OUTLOOK_${fieldName.toUpperCase()}_ERROR`,
      timeoutCode: `OUTLOOK_${fieldName.toUpperCase()}_TIMEOUT`,
      timeoutMessage: `Outlook n’a pas répondu lors de la lecture du champ « ${fieldName} ».`,
    });
  }
  return Promise.resolve(value ?? "");
}

async function waitForAppointmentItem() {
  const startedAt = Date.now();
  while (Date.now() - startedAt < ITEM_WAIT_TIMEOUT_MS) {
    const item = globalThis.Office?.context?.mailbox?.item;
    if (item) return item;
    await delay(200);
  }
  throw outlookError(
    "Outlook n’a pas rendu le rendez-vous accessible au complément.",
    "OUTLOOK_ITEM_UNAVAILABLE",
  );
}

async function readOptional(readOperation) {
  try {
    return await readOperation;
  } catch (error) {
    console.warn("Champ Outlook facultatif indisponible.", error);
    return "";
  }
}

async function getAppointmentContextOnce(onPhase, attempt) {
  onPhase?.("outlook_initialization", attempt);
  await withTimeout(
    globalThis.Office.onReady(),
    OFFICE_READY_TIMEOUT_MS,
    "Outlook n’a pas terminé l’initialisation du complément.",
    "OUTLOOK_READY_TIMEOUT",
  );

  if (Office.context.host !== Office.HostType.Outlook) {
    throw outlookError(
      "Ce formulaire doit être ouvert depuis Outlook.",
      "OUTLOOK_WRONG_HOST",
    );
  }

  onPhase?.("outlook_item", attempt);
  const item = await waitForAppointmentItem();

  onPhase?.("outlook_fields", attempt);
  const [subject, start, end, location] = await Promise.all([
    readText(item.subject, "subject"),
    readDate(item.start, "start"),
    readDate(item.end, "end"),
    readOptional(readText(item.location, "location")),
  ]);

  const parsedStart = start ? new Date(start) : null;
  const parsedEnd = end ? new Date(end) : null;
  if (
    !subject ||
    !parsedStart ||
    Number.isNaN(parsedStart.getTime()) ||
    !parsedEnd ||
    Number.isNaN(parsedEnd.getTime())
  ) {
    throw outlookError(
      "Outlook n’a pas fourni le sujet ou les horaires du rendez-vous.",
      "OUTLOOK_INCOMPLETE_APPOINTMENT",
    );
  }

  return {
    itemId: item.itemId || "",
    subject: String(subject),
    start: parsedStart,
    end: parsedEnd,
    location: String(location || ""),
    organizerName: item.organizer?.displayName || "",
    organizerEmail: item.organizer?.emailAddress || "",
  };
}

export async function getAppointmentContext(options = {}) {
  if (!globalThis.Office?.onReady) {
    throw outlookError(
      "La bibliothèque Office n’a pas été chargée.",
      "OFFICE_JS_UNAVAILABLE",
    );
  }

  return retry(
    (attempt) => getAppointmentContextOnce(options.onPhase, attempt),
    {
      attempts: 2,
      delayMilliseconds: 700,
      onRetry: (error, nextAttempt) => options.onRetry?.(error, nextAttempt),
    },
  );
}

export async function markAppointmentAsLogged() {
  const item = Office.context.mailbox.item;
  if (!item?.loadCustomPropertiesAsync) {
    throw outlookError(
      "Outlook ne permet pas de marquer ce rendez-vous comme journalisé.",
      "OUTLOOK_CUSTOM_PROPERTIES_UNAVAILABLE",
    );
  }

  const properties = await officeAsync((done) => item.loadCustomPropertiesAsync(done), {
    code: "OUTLOOK_CUSTOM_PROPERTIES_LOAD_ERROR",
    timeoutCode: "OUTLOOK_CUSTOM_PROPERTIES_LOAD_TIMEOUT",
  });
  properties.set("EventLogged", true);
  await officeAsync((done) => properties.saveAsync(done), {
    code: "OUTLOOK_CUSTOM_PROPERTIES_SAVE_ERROR",
    timeoutCode: "OUTLOOK_CUSTOM_PROPERTIES_SAVE_TIMEOUT",
  });
}
