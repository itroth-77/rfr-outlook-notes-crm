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
    throw new Error(`Microsoft Graph ${response.status} : ${detail || response.statusText}`);
  }

  if (response.status === 204) return null;
  return response.json();
}

export async function trackOutlookAppointment(token, context) {
  const start = new Date(new Date(context.start).getTime() - 60_000).toISOString();
  const end = new Date(new Date(context.end).getTime() + 60_000).toISOString();
  const query = new URLSearchParams({
    startDateTime: start,
    endDateTime: end,
    $select: "id,subject,start,end,categories",
    $top: "20",
  });
  const result = await graphFetch(token, `/me/calendarView?${query}`, {
    headers: { Prefer: 'outlook.timezone="UTC"' },
  });
  const event = chooseExactGraphEvent(result.value || [], context);
  if (!event) {
    throw new Error(
      "Impossible d’identifier de façon unique ce rendez-vous dans le calendrier Outlook.",
    );
  }

  const currentCategories = Array.isArray(event.categories) ? event.categories : [];
  const alreadyTracked = currentCategories.some(
    (category) => category.toLocaleLowerCase("fr-FR") === TRACKED_TO_DYNAMICS_CATEGORY.toLocaleLowerCase("fr-FR"),
  );
  if (alreadyTracked) return { alreadyTracked: true, eventId: event.id };

  await graphFetch(token, `/me/events/${encodeURIComponent(event.id)}`, {
    method: "PATCH",
    body: JSON.stringify({
      categories: [...currentCategories, TRACKED_TO_DYNAMICS_CATEGORY],
    }),
  });
  return { alreadyTracked: false, eventId: event.id };
}
