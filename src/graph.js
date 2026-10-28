import { withTimeout } from "./resilience.js";

const GRAPH_URL = "https://graph.microsoft.com/v1.0";
const GRAPH_TIMEOUT_MS = 20_000;
export const TRACKED_TO_DYNAMICS_CATEGORY = "Tracked To Dynamics 365";

function graphDate(value) {
  const dateTime = String(value?.dateTime || "");
  if (!dateTime) return new Date(Number.NaN);
  return new Date(/[zZ]|[+-]\d\d:\d\d$/.test(dateTime) ? dateTime : `${dateTime}Z`);
}

function normalizedSubject(value) {
  return String(value || "").trim().toLocaleLowerCase("fr-FR");
}

export function chooseExactGraphEvent(events, context) {
  const expectedSubject = normalizedSubject(context.subject);
  const expectedStart = new Date(context.start).getTime();
  const expectedEnd = new Date(context.end).getTime();
  const exactMatches = events.filter((event) => {
    const startDifference = Math.abs(graphDate(event.start).getTime() - expectedStart);
    const endDifference = Math.abs(graphDate(event.end).getTime() - expectedEnd);
    return (
      normalizedSubject(event.subject) === expectedSubject &&
      startDifference <= 60_000 &&
      endDifference <= 60_000
    );
  });
  return exactMatches.length === 1 ? exactMatches[0] : null;
}

async function graphFetch(token, path, options = {}) {
  const response = await withTimeout(
    fetch(`${GRAPH_URL}${path}`, {
      ...options,
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json; charset=utf-8",
        Authorization: `Bearer ${token}`,
        ...(options.headers || {}),
      },
    }),
    GRAPH_TIMEOUT_MS,
    "Microsoft Graph n’a pas répondu dans le délai prévu.",
    "GRAPH_TIMEOUT",
  );

  if (!response.ok) {
    let detail = "";
    try {
      const body = await response.json();
      detail = body?.error?.message || JSON.stringify(body);
    } catch {
      detail = await response.text();
    }
    const error = new Error([409, 412].includes(response.status)
      ? "Le rendez-vous a changé dans Outlook. Actualisez et confirmez à nouveau ; aucune relance automatique."
      : `Microsoft Graph ${response.status} : ${detail || response.statusText}`);
    error.status = response.status;
    throw error;
  }

  if (response.status === 204) return null;
  return response.json();
}

const EVENT_SELECT = "id,subject,start,end,type,seriesMasterId,categories,isOrganizer,isCancelled";
const eventPath = (id) => `/me/events/${encodeURIComponent(id)}`;
export const isRecurringEvent = (event) => Boolean(event && event.type !== "singleInstance");
const tracked = (event) => (event.categories || []).some((value) =>
  value.trim().toLocaleLowerCase("fr-FR") === TRACKED_TO_DYNAMICS_CATEGORY.toLocaleLowerCase("fr-FR"));

export async function readTrackingState(token, event) {
  if (!isRecurringEvent(event)) return { tracked: tracked(event), scope: "appointment" };
  const master = event.type === "seriesMaster" ? event : await readEvent(token, event.seriesMasterId);
  if (master.type !== "seriesMaster") throw new Error("Le maître de série Outlook n’est pas vérifiable.");
  return { tracked: tracked(master), scope: "series" };
}

async function readEvent(token, id) {
  if (typeof id !== "string" || !id || id.length > 2048 || /[\x00-\x1f\x7f]/.test(id)) throw new Error("Identifiant Outlook invalide.");
  const event = await graphFetch(token, eventPath(id) + "?$select=" + EVENT_SELECT, { headers: { Prefer: 'outlook.timezone="UTC"' } });
  if (event.id !== id || !["singleInstance", "occurrence", "exception", "seriesMaster"].includes(event.type) || event.isCancelled ||
    (event.categories !== undefined && (!Array.isArray(event.categories) || event.categories.some((value) => typeof value !== "string")))) {
    throw new Error("Le rendez-vous Outlook est annulé ou son identité n’est pas vérifiable. Aucune autre copie ne sera utilisée.");
  }
  return event;
}

/** Identite native Office -> Graph, pas d'appariement de secours sur sujet/horaires. */
export async function resolveOutlookEvent(token, context) {
  let id = context.restItemId || (context.itemIdFormat === "restId" ? context.itemId : "");
  if (!id && context.itemId && context.itemIdFormat === "ewsId") {
    const result = await graphFetch(token, "/me/translateExchangeIds", { method: "POST", body: JSON.stringify({
      inputIds: [context.itemId], sourceIdType: "ewsId", targetIdType: "restId",
    }) });
    if (result.value?.length !== 1 || result.value[0].sourceId !== context.itemId) throw new Error("Conversion de l’identifiant Outlook ambiguë.");
    id = result.value[0].targetId;
  }
  if (!id) throw new Error("Enregistrez puis rouvrez ce rendez-vous pour obtenir son identifiant Outlook.");
  const event = await readEvent(token, id);
  if (!chooseExactGraphEvent([event], context)) throw new Error("Le sujet ou les horaires Outlook ont changé. Enregistrez le rendez-vous puis cliquez sur Actualiser.");
  return event;
}

export async function trackOutlookAppointment(token, context, expected, ensureCurrent = () => {}) {
  const event = await resolveOutlookEvent(token, context);
  const recurring = isRecurringEvent(event);
  if (!expected || expected.id !== event.id || expected.type !== event.type || expected.seriesMasterId !== event.seriesMasterId) {
    throw new Error("La portée du rendez-vous a changé. Actualisez puis confirmez le suivi du rendez-vous ou de toute la série.");
  }
  const masterId = event.type === "seriesMaster" ? event.id : event.seriesMasterId;
  if (recurring && !masterId) throw new Error("Le maître de série Outlook est introuvable. Aucune occurrence seule ne sera modifiée.");
  const target = recurring && event.type !== "seriesMaster" ? await readEvent(token, masterId) : event;
  if (recurring && target.type !== "seriesMaster") throw new Error("Le maître de série Outlook n’est pas vérifiable.");
  const alreadyTracked = tracked(target);
  if (!alreadyTracked) {
    const etag = target["@odata.etag"];
    if (!etag || /[\r\n]/.test(etag)) throw new Error("La version Outlook manque pour préserver les autres catégories. Actualisez.");
    await ensureCurrent();
    await graphFetch(token, eventPath(target.id), { method: "PATCH", headers: { "If-Match": etag },
      body: JSON.stringify({ categories: [...(target.categories || []), TRACKED_TO_DYNAMICS_CATEGORY] }) });
  }
  const verified = await readEvent(token, target.id);
  if (!tracked(verified)) throw new Error("Le suivi n’est pas confirmé sur la cible. Actualisez avant de réessayer.");
  const selected = target.id === event.id ? verified : await readEvent(token, event.id);
  if (recurring && selected.id !== target.id && selected.seriesMasterId !== target.id) throw new Error("Le rattachement Outlook a changé après la modification. Vérifiez la série avant toute relance.");
  return { alreadyTracked, eventId: event.id, event: selected, seriesTrackingVerified: recurring,
    notice: recurring && !tracked(selected)
      ? "Catégorie confirmée sur toute la série. Cette occurrence conserve une exception de catégorie ; elle n’a pas été modifiée séparément."
      : "Catégorie de suivi confirmée dans votre calendrier. La synchronisation Dynamics reste à vérifier." };
}
