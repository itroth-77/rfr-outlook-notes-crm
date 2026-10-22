const API_PATH = "/api/data/v9.2";

export function escapeODataString(value) {
  return String(value ?? "").replaceAll("'", "''");
}

export function buildAppointmentQuery(context) {
  const start = new Date(context.start);
  const windowStart = new Date(start.getTime() - 15 * 60 * 1000).toISOString();
  const windowEnd = new Date(start.getTime() + 15 * 60 * 1000).toISOString();

  const select = [
    "activityid",
    "subject",
    "scheduledstart",
    "scheduledend",
    "location",
    "globalobjectid",
    "_regardingobjectid_value",
  ].join(",");
  const filter = `scheduledstart ge ${windowStart} and scheduledstart le ${windowEnd}`;

  return `/appointments?$select=${select}&$filter=${encodeURIComponent(filter)}&$orderby=scheduledstart asc&$top=20`;
}

export function rankAppointmentCandidates(candidates, context) {
  const subject = context.subject.trim().toLocaleLowerCase("fr-FR");
  const targetStart = new Date(context.start).getTime();
  const targetEnd = new Date(context.end).getTime();

  return candidates
    .map((candidate) => {
      const candidateSubject = String(candidate.subject || "").trim().toLocaleLowerCase("fr-FR");
      const startDifference = Math.abs(new Date(candidate.scheduledstart).getTime() - targetStart);
      const endDifference = Math.abs(new Date(candidate.scheduledend).getTime() - targetEnd);
      let score = 0;
      if (candidateSubject === subject) score += 60;
      else if (candidateSubject.includes(subject) || subject.includes(candidateSubject)) score += 25;
      if (startDifference <= 60_000) score += 25;
      else if (startDifference <= 5 * 60_000) score += 15;
      if (endDifference <= 60_000) score += 15;
      else if (endDifference <= 5 * 60_000) score += 8;
      return { ...candidate, matchScore: score };
    })
    .sort((left, right) => right.matchScore - left.matchScore);
}

export function chooseUnambiguousCandidate(candidates) {
  if (!candidates.length || candidates[0].matchScore < 85) return null;
  if (candidates.length > 1 && candidates[0].matchScore === candidates[1].matchScore) return null;
  return candidates[0];
}

async function dataverseFetch(config, token, path, options = {}) {
  const response = await fetch(`${config.dataverseUrl}${API_PATH}${path}`, {
    ...options,
    headers: {
      Accept: "application/json",
      "Content-Type": "application/json; charset=utf-8",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue,Microsoft.Dynamics.CRM.lookuplogicalname"',
      Authorization: `Bearer ${token}`,
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    let detail = "";
    try {
      const body = await response.json();
      detail = body?.error?.message || JSON.stringify(body);
    } catch {
      detail = await response.text();
    }
    throw new Error(`Dataverse ${response.status} : ${detail || response.statusText}`);
  }

  if (response.status === 204) return null;
  return response.json();
}

export async function findAppointmentCandidates(config, token, context) {
  const result = await dataverseFetch(config, token, buildAppointmentQuery(context));
  return rankAppointmentCandidates(result.value || [], context);
}

export async function getAppointmentNotes(config, token, appointmentId) {
  const id = String(appointmentId).replace(/[{}]/g, "");
  const path = `/appointments(${id})/Appointment_Annotation?$select=annotationid,subject,notetext,createdon&$orderby=createdon desc&$top=20`;
  const result = await dataverseFetch(config, token, path);
  return result.value || [];
}

export async function createAppointmentNote(config, token, appointmentId, note) {
  const id = String(appointmentId).replace(/[{}]/g, "");
  const response = await fetch(`${config.dataverseUrl}${API_PATH}/annotations`, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/json; charset=utf-8",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      Prefer: "return=representation",
      Authorization: `Bearer ${token}`,
    },
    body: JSON.stringify({
      subject: note.subject,
      notetext: note.text,
      "objectid_appointment@odata.bind": `/appointments(${id})`,
    }),
  });

  if (!response.ok) {
    let detail = "";
    try {
      const body = await response.json();
      detail = body?.error?.message || JSON.stringify(body);
    } catch {
      detail = await response.text();
    }
    throw new Error(`Création de la note impossible (${response.status}) : ${detail || response.statusText}`);
  }

  return response.json();
}

export function getRegardingLabel(candidate) {
  return (
    candidate["_regardingobjectid_value@OData.Community.Display.V1.FormattedValue"] ||
    "Aucun concernant"
  );
}
