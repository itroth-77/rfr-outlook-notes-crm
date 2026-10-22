import { withTimeout } from "./resilience.js";

const API_PATH = "/api/data/v9.2";
const DATAVERSE_TIMEOUT_MS = 20_000;
export const APPOINTMENT_QUALIFICATION_FIELD = "rfr_os_appointmentqualification";

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
    APPOINTMENT_QUALIFICATION_FIELD,
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
  const response = await withTimeout(
    fetch(`${config.dataverseUrl}${API_PATH}${path}`, {
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
    }),
    DATAVERSE_TIMEOUT_MS,
    "Le CRM n’a pas répondu dans le délai prévu.",
    "DATAVERSE_TIMEOUT",
  );

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

function choiceLabel(option, languageCode) {
  const localizedLabels = option?.Label?.LocalizedLabels || [];
  return (
    localizedLabels.find((label) => label.LanguageCode === languageCode)?.Label ||
    option?.Label?.UserLocalizedLabel?.Label ||
    localizedLabels[0]?.Label ||
    String(option?.Value ?? "")
  );
}

export function extractQualificationOptions(attributeMetadata, languageCode = 1036) {
  return (attributeMetadata?.OptionSet?.Options || [])
    .filter((option) => Number.isInteger(option?.Value))
    .map((option) => ({
      value: option.Value,
      label: choiceLabel(option, languageCode),
    }))
    .sort((left, right) => left.label.localeCompare(right.label, "fr", { sensitivity: "base" }));
}

export async function getAppointmentQualificationOptions(config, token) {
  const path =
    `/EntityDefinitions(LogicalName='appointment')` +
    `/Attributes(LogicalName='${APPOINTMENT_QUALIFICATION_FIELD}')` +
    "/Microsoft.Dynamics.CRM.PicklistAttributeMetadata" +
    "?$select=LogicalName&$expand=OptionSet";
  const metadata = await dataverseFetch(config, token, path);
  return extractQualificationOptions(metadata);
}

export async function updateAppointmentQualification(
  config,
  token,
  appointmentId,
  qualificationValue,
) {
  const id = String(appointmentId).replace(/[{}]/g, "");
  return dataverseFetch(config, token, `/appointments(${id})`, {
    method: "PATCH",
    headers: {
      Prefer:
        'return=representation,odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
    },
    body: JSON.stringify({
      [APPOINTMENT_QUALIFICATION_FIELD]: Number(qualificationValue),
    }),
  });
}

export function getRegardingReference(candidate) {
  const id = String(candidate?._regardingobjectid_value || "").replace(/[{}]/g, "");
  const logicalName = String(
    candidate?.[
      "_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname"
    ] || "",
  ).toLowerCase();

  if (!id || !logicalName || !/^[a-z][a-z0-9_]*$/.test(logicalName)) return null;

  return {
    id,
    logicalName,
    label: getRegardingLabel(candidate),
  };
}

export async function getRegardingNotes(config, token, regarding) {
  const filter = encodeURIComponent(`_objectid_value eq ${regarding.id}`);
  const path = `/annotations?$select=annotationid,subject,notetext,createdon&$filter=${filter}&$orderby=createdon desc&$top=20`;
  const result = await dataverseFetch(config, token, path);
  return result.value || [];
}

export function buildRegardingNotePayload(regarding, entitySetName, note) {
  return {
    subject: note.subject,
    notetext: note.text,
    [`objectid_${regarding.logicalName}@odata.bind`]: `/${entitySetName}(${regarding.id})`,
  };
}

async function getEntitySetName(config, token, logicalName) {
  if (!/^[a-z][a-z0-9_]*$/.test(logicalName)) {
    throw new Error("Le type du champ Concernant n’est pas valide.");
  }
  const metadata = await dataverseFetch(
    config,
    token,
    `/EntityDefinitions(LogicalName='${logicalName}')?$select=EntitySetName`,
  );
  if (!metadata?.EntitySetName) {
    throw new Error(`Impossible de déterminer la table Dataverse pour ${logicalName}.`);
  }
  return metadata.EntitySetName;
}

export async function createRegardingNote(config, token, regarding, note) {
  const entitySetName = await getEntitySetName(config, token, regarding.logicalName);
  return dataverseFetch(config, token, "/annotations", {
    method: "POST",
    headers: {
      Prefer: "return=representation",
    },
    body: JSON.stringify(buildRegardingNotePayload(regarding, entitySetName, note)),
  });
}

export function getRegardingLabel(candidate) {
  return (
    candidate["_regardingobjectid_value@OData.Community.Display.V1.FormattedValue"] ||
    "Aucun concernant"
  );
}
