import { delay, retry, withTimeout } from "./resilience.js";

const OFFICE_READY_TIMEOUT_MS = 8_000;
const OFFICE_JS_TIMEOUT_MS = 10_000;
const ITEM_WAIT_TIMEOUT_MS = 4_000;
const FIELD_READ_TIMEOUT_MS = 5_000;
const OFFICE_JS_URL = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";

let officeJsPromise;

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

function loadOfficeJs() {
  if (globalThis.Office?.onReady) return Promise.resolve();
  if (officeJsPromise) return officeJsPromise;

  officeJsPromise = new Promise((resolve, reject) => {
    const existingScript = document.querySelector('script[data-rfr-office-js="true"]');
    const script = existingScript || document.createElement("script");

    script.addEventListener("load", () => resolve(), { once: true });
    script.addEventListener(
      "error",
      () => reject(outlookError("La bibliothèque Office n’a pas pu être chargée.", "OFFICE_JS_LOAD_ERROR")),
      { once: true },
    );

    if (!existingScript) {
      script.src = OFFICE_JS_URL;
      script.async = true;
      script.dataset.rfrOfficeJs = "true";
      document.head.append(script);
    }
  });

  return officeJsPromise;
}

async function ensureOfficeJs() {
  try {
    await withTimeout(
      loadOfficeJs(),
      OFFICE_JS_TIMEOUT_MS,
      "La bibliothèque Office n’a pas répondu dans le délai prévu.",
      "OFFICE_JS_TIMEOUT",
    );
  } catch (error) {
    document.querySelector('script[data-rfr-office-js="true"]')?.remove();
    officeJsPromise = null;
    throw error;
  }
  if (!globalThis.Office?.onReady) {
    throw outlookError(
      "La bibliothèque Office est chargée mais son API n’est pas disponible.",
      "OFFICE_JS_API_UNAVAILABLE",
    );
  }
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

async function waitForAppointmentItem(timeoutMilliseconds = ITEM_WAIT_TIMEOUT_MS) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMilliseconds) {
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

export function isOutlookMobile(office = globalThis.Office) {
  const platform = office?.context?.platform;
  return ["Android", "iOS"].includes(platform) ||
    /^(OutlookIOS|OutlookAndroid)$/i.test(office?.context?.mailbox?.diagnostics?.hostName || "");
}

export function getRestItemId(item, office = globalThis.Office, { ews = false } = {}) {
  if (!item?.itemId) return "";
  if (isOutlookMobile(office) && !ews) return item.itemId;
  const mailbox = office?.context?.mailbox;
  if (typeof mailbox?.convertToRestId !== "function") return "";
  try { return mailbox.convertToRestId(item.itemId, office.MailboxEnums.RestVersion.v2_0); }
  catch { return ""; }
}

async function readAppointmentId(item) {
  if (typeof item?.itemId === "string" && item.itemId) return item.itemId;
  if (typeof item?.getItemIdAsync !== "function") {
    throw outlookError("Enregistrez le rendez-vous dans Outlook, puis cliquez sur Réessayer ou Actualiser. Ce client ne fournit pas son identifiant en mode modification.", "OUTLOOK_ITEM_ID_UNAVAILABLE");
  }
  // Read only: saveAsync would persist pending Outlook edits and is never used.
  try {
    const id = await officeAsync((done) => item.getItemIdAsync(done), {
      code: "OUTLOOK_ITEM_ID_ERROR",
      timeoutCode: "OUTLOOK_ITEM_ID_TIMEOUT",
      timeoutMessage: "Outlook ne répond pas à la lecture de l’identifiant du rendez-vous.",
    });
    if (typeof id !== "string" || !id) throw new Error("empty_item_id");
    return id;
  } catch (error) {
    if (error.code === "OUTLOOK_ITEM_ID_TIMEOUT") throw error;
    throw outlookError("L’identifiant du rendez-vous n’est pas encore disponible. Enregistrez-le dans Outlook, puis cliquez sur Réessayer ou Actualiser.", "OUTLOOK_ITEM_ID_UNAVAILABLE", error);
  }
}

function sameAppointmentId(context, id, office) {
  return Boolean(context?.itemId && id && (id === context.itemId ||
    (context.restItemId && getRestItemId({ itemId: id }, office, { ews: context.identitySource === "getItemIdAsync" }) === context.restItemId)));
}

export function assertAppointmentContext(context, office = globalThis.Office) {
  const current = office?.context?.mailbox?.item;
  const mailboxEmail = office?.context?.mailbox?.userProfile?.emailAddress || "";
  const sameMailbox = !context?.mailboxEmail || context.mailboxEmail.toLowerCase() === mailboxEmail.toLowerCase();
  // Compare opaque IDs, never subject/time. Office may expose an equivalent REST ID.
  const sameId = current?.itemId ? sameAppointmentId(context, current.itemId, office)
    : Boolean(context?.itemId && context.identitySource === "getItemIdAsync" && context.officeItem === current && typeof current?.getItemIdAsync === "function");
  if (!sameMailbox || !sameId) {
    throw outlookError("Le rendez-vous Outlook a changé ou n’est plus accessible. Cliquez sur Actualiser pour le recharger.", "OUTLOOK_CONTEXT_CHANGED");
  }
}

/** A synchronous check can only bind a compose wrapper, not reread its ID.
 * Use this async check before displaying retrieved data and before each write.
 */
export async function assertAppointmentContextAsync(context, office = globalThis.Office) {
  assertAppointmentContext(context, office);
  const item = office.context.mailbox.item;
  const id = await readAppointmentId(item);
  assertAppointmentContext(context, office);
  if (office.context.mailbox.item !== item || !sameAppointmentId(context, id, office)) {
    throw outlookError("Le rendez-vous Outlook a changé. Cliquez sur Actualiser pour le recharger.", "OUTLOOK_CONTEXT_CHANGED");
  }
}

export function watchAppointmentChanges(onChange, office = globalThis.Office) {
  const mailbox = office?.context?.mailbox;
  if (!office?.context?.requirements?.isSetSupported("Mailbox", "1.5") ||
      !office?.EventType?.ItemChanged || !mailbox?.addHandlerAsync) return;
  // Some mobile hosts expose the method but reject the event: refresh remains usable.
  try { mailbox.addHandlerAsync(office.EventType.ItemChanged, onChange, () => {}); }
  catch { /* Optional host capability. Explicit refresh does not depend on it. */ }
}

async function getAppointmentContextOnce(onPhase, attempt) {
  onPhase?.("office_library", attempt);
  await ensureOfficeJs();

  let item = globalThis.Office?.context?.mailbox?.item;
  if (item) {
    onPhase?.("outlook_context_direct", attempt);
  } else {
    onPhase?.("outlook_initialization", attempt);
    const itemAvailable = waitForAppointmentItem(OFFICE_READY_TIMEOUT_MS);
    await withTimeout(
      Promise.race([
        globalThis.Office.onReady(),
        itemAvailable.then(() => undefined),
      ]),
      OFFICE_READY_TIMEOUT_MS,
      "Outlook n’a pas terminé l’initialisation du complément.",
      "OUTLOOK_READY_TIMEOUT",
    );
    item = globalThis.Office?.context?.mailbox?.item;
    if (!item) {
      onPhase?.("outlook_item", attempt);
      item = await itemAvailable;
    }
  }

  const host = Office.context?.host;
  if (host && host !== Office.HostType.Outlook) {
    throw outlookError(
      "Ce formulaire doit être ouvert depuis Outlook.",
      "OUTLOOK_WRONG_HOST",
    );
  }

  onPhase?.("outlook_item_id", attempt);
  const itemId = await readAppointmentId(item);
  if (globalThis.Office?.context?.mailbox?.item !== item) {
    throw outlookError("Le rendez-vous Outlook a changé pendant sa lecture. Cliquez sur Actualiser.", "OUTLOOK_CONTEXT_CHANGED");
  }
  const identitySource = item.itemId ? "itemId" : "getItemIdAsync";
  const identity = {
    itemId,
    restItemId: getRestItemId({ itemId }, globalThis.Office, { ews: identitySource === "getItemIdAsync" }),
    identitySource,
    officeItem: item,
    mailboxEmail: Office.context?.mailbox?.userProfile?.emailAddress || "",
  };
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

  await assertAppointmentContextAsync(identity);
  return {
    ...identity,
    itemIdFormat: identitySource === "getItemIdAsync" ? "ewsId" : isOutlookMobile() ? "restId" : "ewsId",
    isMobile: isOutlookMobile(),
    subject: String(subject),
    start: parsedStart,
    end: parsedEnd,
    location: String(location || ""),
    organizerName: item.organizer?.displayName || "",
    organizerEmail: item.organizer?.emailAddress || "",
    mailboxEmail: Office.context?.mailbox?.userProfile?.emailAddress || "",
  };
}

export async function getAppointmentContext(options = {}) {
  return retry(
    (attempt) => getAppointmentContextOnce(options.onPhase, attempt),
    {
      attempts: 2,
      delayMilliseconds: 700,
      onRetry: (error, nextAttempt) => options.onRetry?.(error, nextAttempt),
    },
  );
}

export async function markAppointmentAsLogged(context) {
  await assertAppointmentContextAsync(context);
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
  await assertAppointmentContextAsync(context);
  properties.set("EventLogged", true);
  await officeAsync((done) => properties.saveAsync(done), {
    code: "OUTLOOK_CUSTOM_PROPERTIES_SAVE_ERROR",
    timeoutCode: "OUTLOOK_CUSTOM_PROPERTIES_SAVE_TIMEOUT",
  });
}
