import { withTimeout } from "./resilience.js";
import { appointmentNoteReference, belongsToAppointment, withAppointmentReference } from "./appointment-notes.js";

const API_PATH = "/api/data/v9.2";
const DATAVERSE_TIMEOUT_MS = 20_000;
const APPOINTMENT_PARTICIPATION_TYPES = new Set([5, 6, 7]);
const OPPORTUNITY_BUSINESS_ATTRIBUTE = "rfr_lk_business";
const OPPORTUNITY_BUSINESS_LOOKUP_VALUE = "_rfr_lk_business_value";
export const APPOINTMENT_QUALIFICATION_FIELD = "rfr_os_appointmentqualification";
let opportunityBusinessDefinitionPromise;
const REGARDING_TYPES = {
  account: {
    entitySet: "accounts",
    idField: "accountid",
    nameField: "name",
    detailField: "accountnumber",
    navigationProperty: "regardingobjectid_account_appointment",
  },
  contact: {
    entitySet: "contacts",
    idField: "contactid",
    nameField: "fullname",
    detailField: "emailaddress1",
    navigationProperty: "regardingobjectid_contact_appointment",
  },
  opportunity: {
    entitySet: "opportunities",
    idField: "opportunityid",
    nameField: "name",
    detailField: OPPORTUNITY_BUSINESS_LOOKUP_VALUE,
    detailIsFormattedLookup: true,
    navigationProperty: "regardingobjectid_opportunity_appointment",
  },
};

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
    "instancetypecode",
    "seriesid",
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
    const error = new Error([409, 412].includes(response.status)
      ? "Les données Dynamics ont changé. Actualisez puis confirmez à nouveau ; aucune relance automatique."
      : `Dataverse ${response.status} : ${detail || response.statusText}`);
    error.status = response.status;
    throw error;
  }

  if (response.status === 204) return null;
  return response.json();
}

export function normalizeGuid(value) {
  const guid = String(value || "").replace(/[{}]/g, "").toLowerCase();
  // Dataverse identifiers are GUIDs, but their version and variant bits are not
  // guaranteed to follow the RFC 4122 restrictions. Only validate the canonical
  // 8-4-4-4-12 hexadecimal shape so valid Dataverse record IDs are not discarded.
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(guid)
    ? guid
    : "";
}

export function buildCurrentUserActivityPartyQuery(appointments, userId, userEmail = "") {
  const normalizedUserId = normalizeGuid(userId);
  const normalizedEmail = String(userEmail || "").trim();
  const appointmentIds = appointments
    .map((appointment) => normalizeGuid(appointment.activityid))
    .filter(Boolean);

  if ((!normalizedUserId && !normalizedEmail) || !appointmentIds.length) return "";

  const appointmentFilter = appointmentIds
    .map((id) => `_activityid_value eq ${id}`)
    .join(" or ");
  const roleFilter = [...APPOINTMENT_PARTICIPATION_TYPES]
    .map((role) => `participationtypemask eq ${role}`)
    .join(" or ");
  const userFilter = [
    normalizedUserId ? `_partyid_value eq ${normalizedUserId}` : "",
    normalizedEmail ? `addressused eq '${escapeODataString(normalizedEmail)}'` : "",
  ]
    .filter(Boolean)
    .join(" or ");
  const filter = `(${userFilter}) and (${appointmentFilter}) and (${roleFilter})`;

  return `/activityparties?$select=_activityid_value,participationtypemask&$filter=${encodeURIComponent(filter)}&$top=100`;
}

export function filterAppointmentsForCurrentUser(appointments, activityParties) {
  const allowedAppointmentIds = new Set(
    activityParties
      .filter((party) => APPOINTMENT_PARTICIPATION_TYPES.has(Number(party.participationtypemask)))
      .map((party) => normalizeGuid(party._activityid_value))
      .filter(Boolean),
  );
  return appointments.filter((appointment) =>
    allowedAppointmentIds.has(normalizeGuid(appointment.activityid)),
  );
}

export async function findAppointmentCandidates(config, token, context) {
  const result = await dataverseFetch(config, token, buildAppointmentQuery(context));
  const appointments = result.value || [];
  if (!appointments.length) return [];

  const identity = await dataverseFetch(config, token, "/WhoAmI");
  const activityPartyQuery = buildCurrentUserActivityPartyQuery(
    appointments,
    identity?.UserId,
    context.mailboxEmail,
  );
  if (!activityPartyQuery) {
    throw new Error("Impossible d’identifier l’utilisateur CRM connecté.");
  }
  const activityParties = await dataverseFetch(config, token, activityPartyQuery);
  const visibleAppointments = filterAppointmentsForCurrentUser(
    appointments,
    activityParties.value || [],
  );
  return rankAppointmentCandidates(visibleAppointments, context);
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

export async function getAppointmentQualificationOptions(config, token, appointment) {
  if (appointment?.editTarget?.kind === "blocked") throw new Error(appointment.editTarget.reason);
  if (appointment?.editTarget?.kind === "series") {
    if (appointment.editTarget.qualificationError) throw new Error(appointment.editTarget.qualificationError);
    return extractQualificationOptions(await getSeriesQualificationMetadata(config, token));
  }
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
  expected,
  ensureCurrent = () => {},
) {
  const value = Number(qualificationValue);
  if (!Number.isInteger(value)) throw new Error("Qualification invalide.");
  const current = await readEditingTarget(config, token, appointmentId, expected);
  const options = await getAppointmentQualificationOptions(config, token, current);
  if (!options.some((option) => option.value === value)) throw new Error("Cette qualification n’est pas disponible sur la cible Dynamics.");
  await ensureCurrent();
  await dataverseFetch(config, token, editingPath(current), {
    method: "PATCH",
    headers: { "If-Match": current.editTarget.etag },
    body: JSON.stringify({
      [APPOINTMENT_QUALIFICATION_FIELD]: value,
    }),
  });
  const saved = await dataverseFetch(config, token, editingPath(current) + `?$select=${APPOINTMENT_QUALIFICATION_FIELD}`);
  if (saved[APPOINTMENT_QUALIFICATION_FIELD] !== value) throw new Error("La qualification enregistrée n’est pas confirmée. Actualisez avant toute relance.");
  const updated = await readCurrentTarget(config, token, appointmentId, current);
  if (current.editTarget.kind === "series" && updated.occurrenceQualification !== value) {
    updated.seriesNotice = "Qualification enregistrée sur la série. Cette séance n’a pas repris la valeur : exception ou propagation à vérifier par l’administrateur Dynamics.";
  }
  return updated;
}

export function parseRegardingSearchFragments(searchText) {
  const text = String(searchText || "").trim();
  const fragments = text
    .split(/[%*]+/)
    .map((fragment) => fragment.trim())
    .filter(Boolean);
  if (fragments.join("").length < 2) {
    throw new Error("Saisissez au moins 2 caractères hors jokers.");
  }
  return fragments;
}

function normalizeSearchText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("fr-FR");
}

export function matchesRegardingSearchPattern(record, searchText) {
  const fragments = parseRegardingSearchFragments(searchText).map(normalizeSearchText);
  const searchableText = normalizeSearchText(`${record?.label || ""} ${record?.detail || ""}`);
  let position = 0;

  for (const fragment of fragments) {
    const matchPosition = searchableText.indexOf(fragment, position);
    if (matchPosition < 0) return false;
    position = matchPosition + fragment.length;
  }
  return true;
}

export function buildRegardingSearchQuery(logicalName, searchText) {
  const type = REGARDING_TYPES[logicalName];
  if (!type) throw new Error("Le type de Concernant n’est pas pris en charge.");
  const fragments = parseRegardingSearchFragments(searchText);

  const searchFields = [
    type.nameField,
    type.detailIsFormattedLookup ? "" : type.detailField,
  ].filter(Boolean);
  const textFilter = fragments
    .map((fragment) => {
      const escapedFragment = escapeODataString(fragment);
      return `(${searchFields
        .map((field) => `contains(${field},'${escapedFragment}')`)
        .join(" or ")})`;
    })
    .join(" and ");
  const select = [type.idField, type.nameField, type.detailField].filter(Boolean).join(",");
  const filter = encodeURIComponent(`statecode eq 0 and (${textFilter})`);
  return `/${type.entitySet}?$select=${select}&$filter=${filter}&$orderby=${type.nameField} asc&$top=100`;
}

export function buildOpportunitySearchQuery(
  searchText,
  businessIdsByFragment = [],
  businessIdsMatchingAllFragments = [],
) {
  const type = REGARDING_TYPES.opportunity;
  const fragments = parseRegardingSearchFragments(searchText);
  const fragmentFilters = fragments.map((fragment, index) => {
    const conditions = [`contains(${type.nameField},'${escapeODataString(fragment)}')`];
    for (const value of businessIdsByFragment[index] || []) {
      const id = normalizeGuid(value);
      if (id) conditions.push(`${OPPORTUNITY_BUSINESS_LOOKUP_VALUE} eq ${id}`);
    }
    return `(${conditions.join(" or ")})`;
  });
  const combinedConditions = [`(${fragmentFilters.join(" and ")})`];
  for (const value of businessIdsMatchingAllFragments) {
    const id = normalizeGuid(value);
    if (id) combinedConditions.push(`${OPPORTUNITY_BUSINESS_LOOKUP_VALUE} eq ${id}`);
  }
  const select = [type.idField, type.nameField, type.detailField].join(",");
  const filter = encodeURIComponent(`statecode eq 0 and (${combinedConditions.join(" or ")})`);
  return `/${type.entitySet}?$select=${select}&$filter=${filter}&$orderby=${type.nameField} asc&$top=100`;
}

function validateMetadataName(value, label) {
  const name = String(value || "");
  if (!/^[A-Za-z][A-Za-z0-9_]*$/.test(name)) {
    throw new Error(`Les métadonnées Dataverse de ${label} sont invalides.`);
  }
  return name;
}

async function getOpportunityBusinessDefinition(config, token) {
  if (!opportunityBusinessDefinitionPromise) {
    opportunityBusinessDefinitionPromise = (async () => {
      const lookup = await dataverseFetch(
        config,
        token,
        "/EntityDefinitions(LogicalName='opportunity')" +
          `/Attributes(LogicalName='${OPPORTUNITY_BUSINESS_ATTRIBUTE}')` +
          "/Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=Targets",
      );
      const logicalName = validateMetadataName(lookup?.Targets?.[0], "l’Affaire");
      const entity = await dataverseFetch(
        config,
        token,
        `/EntityDefinitions(LogicalName='${logicalName}')` +
          "?$select=LogicalName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute",
      );
      return {
        logicalName,
        entitySet: validateMetadataName(entity?.EntitySetName, "la table Affaire"),
        idField: validateMetadataName(entity?.PrimaryIdAttribute, "l’identifiant de l’Affaire"),
        nameField: validateMetadataName(entity?.PrimaryNameAttribute, "le nom de l’Affaire"),
      };
    })().catch((error) => {
      opportunityBusinessDefinitionPromise = null;
      throw error;
    });
  }
  return opportunityBusinessDefinitionPromise;
}

function buildBusinessFragmentsQuery(definition, fragments) {
  const textFilter = fragments
    .map(
      (fragment) =>
        `contains(${definition.nameField},'${escapeODataString(fragment)}')`,
    )
    .join(" and ");
  const filter = encodeURIComponent(textFilter);
  return (
    `/${definition.entitySet}` +
    `?$select=${definition.idField},${definition.nameField}` +
    `&$filter=${filter}&$orderby=${definition.nameField} asc&$top=100`
  );
}

async function searchOpportunityRecords(config, token, searchText) {
  const fragments = parseRegardingSearchFragments(searchText);
  const businessDefinition = await getOpportunityBusinessDefinition(config, token);
  const searchBusinessIds = async (searchFragments) => {
    const result = await dataverseFetch(
      config,
      token,
      buildBusinessFragmentsQuery(businessDefinition, searchFragments),
    );
    return (result.value || [])
      .map((record) => normalizeGuid(record[businessDefinition.idField]))
      .filter(Boolean);
  };
  const searches = fragments.map((fragment) => searchBusinessIds([fragment]));
  if (fragments.length > 1) searches.push(searchBusinessIds(fragments));
  const searchResults = await Promise.all(searches);
  const businessIdsByFragment = searchResults.slice(0, fragments.length);
  const businessIdsMatchingAllFragments = fragments.length > 1
    ? searchResults.at(-1)
    : businessIdsByFragment[0];
  return dataverseFetch(
    config,
    token,
    buildOpportunitySearchQuery(
      searchText,
      businessIdsByFragment,
      businessIdsMatchingAllFragments,
    ),
  );
}

export async function searchRegardingRecords(config, token, logicalName, searchText) {
  const type = REGARDING_TYPES[logicalName];
  if (!type) throw new Error("Le type de Concernant n’est pas pris en charge.");
  const result = logicalName === "opportunity"
    ? await searchOpportunityRecords(config, token, searchText)
    : await dataverseFetch(config, token, buildRegardingSearchQuery(logicalName, searchText));

  const records = (result.value || []).map((record) =>
    mapRegardingSearchRecord(logicalName, record),
  );
  return /[%*]/.test(searchText)
    ? records.filter((record) => matchesRegardingSearchPattern(record, searchText))
    : records;
}

export function mapRegardingSearchRecord(logicalName, record) {
  const type = REGARDING_TYPES[logicalName];
  if (!type) throw new Error("Le type de Concernant n’est pas pris en charge.");
  return {
    id: record[type.idField],
    logicalName,
    label: record[type.nameField] || "Sans nom",
    detail: type.detailIsFormattedLookup
      ? record[`${type.detailField}@OData.Community.Display.V1.FormattedValue`] || ""
      : type.detailField
        ? record[type.detailField] || ""
        : "",
  };
}

export function buildAppointmentRegardingPayload(logicalName, recordId) {
  const type = REGARDING_TYPES[logicalName];
  const id = normalizeGuid(recordId);
  if (!type || !id) throw new Error("Le Concernant sélectionné n’est pas valide.");
  return {
    [`${type.navigationProperty}@odata.bind`]: `/${type.entitySet}(${id})`,
  };
}

export async function updateAppointmentRegarding(config, token, appointmentId, regarding, expected, ensureCurrent = () => {}) {
  const id = normalizeGuid(appointmentId);
  if (!id) throw new Error("L’identifiant du rendez-vous CRM n’est pas valide.");
  const current = await readEditingTarget(config, token, id, expected);
  let payload = buildAppointmentRegardingPayload(regarding.logicalName, regarding.id);
  if (current.editTarget.kind === "series") {
    const filter = `ReferencingEntity eq 'recurringappointmentmaster' and ReferencingAttribute eq 'regardingobjectid' and ReferencedEntity eq '${regarding.logicalName}'`;
    const metadata = await dataverseFetch(config, token,
      "/EntityDefinitions(LogicalName='recurringappointmentmaster')/ManyToOneRelationships" +
      `?$select=ReferencingEntityNavigationPropertyName&$filter=${encodeURIComponent(filter)}`);
    const names = [...new Set((metadata.value || []).map((row) => row.ReferencingEntityNavigationPropertyName).filter(Boolean))];
    if (names.length !== 1 || !/^regardingobjectid_[a-z0-9_]+$/.test(names[0])) throw new Error("La relation Concernant de la série est ambiguë. Aucune occurrence ne sera modifiée.");
    payload = { [names[0] + "@odata.bind"]: `/${REGARDING_TYPES[regarding.logicalName].entitySet}(${normalizeGuid(regarding.id)})` };
  }
  await ensureCurrent();
  await dataverseFetch(config, token, editingPath(current), {
    method: "PATCH",
    headers: { "If-Match": current.editTarget.etag },
    body: JSON.stringify(payload),
  });
  const saved = await dataverseFetch(config, token, editingPath(current) + "?$select=_regardingobjectid_value");
  const verified = getRegardingReference(saved);
  if (verified?.id !== normalizeGuid(regarding.id) || verified?.logicalName !== regarding.logicalName) throw new Error("Le Concernant enregistré n’est pas confirmé. Actualisez avant toute relance.");
  return readCurrentTarget(config, token, id, current);
}

const targetSelect = ["activityid", "subject", "scheduledstart", "scheduledend", "instancetypecode", "seriesid",
  "_regardingobjectid_value", APPOINTMENT_QUALIFICATION_FIELD].join(",");

async function getSeriesQualificationMetadata(config, token) {
  const field = `/Attributes(LogicalName='${APPOINTMENT_QUALIFICATION_FIELD}')`;
  const [master, occurrence] = await Promise.all([
    dataverseFetch(config, token, "/EntityDefinitions(LogicalName='recurringappointmentmaster')" + field +
      "/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=MetadataId,LogicalName,LinkedAttributeId,IsValidForRead,IsValidForUpdate,IsSecured&$expand=OptionSet"),
    dataverseFetch(config, token, "/EntityDefinitions(LogicalName='appointment')" + field + "?$select=MetadataId"),
  ]);
  if (master.LogicalName !== APPOINTMENT_QUALIFICATION_FIELD || master.IsValidForRead !== true ||
    master.IsValidForUpdate !== true || master.IsSecured === true || !extractQualificationOptions(master).length) {
    throw new Error("Qualification de série indisponible : champ absent, sécurisé, non modifiable ou sans choix. Aucun schéma ne sera modifié.");
  }
  return { ...master, propagationLinked: Boolean(normalizeGuid(master.LinkedAttributeId) &&
    normalizeGuid(master.LinkedAttributeId) === normalizeGuid(occurrence.MetadataId)) };
}

/** Garde l'identifiant de la seance pour son contexte, mais affiche/edite le maitre CRM. */
export async function resolveAppointmentEditingTarget(config, token, candidate, expectSeries = false) {
  const id = normalizeGuid(candidate.activityid);
  if (!id) throw new Error("Identifiant de rendez-vous Dynamics invalide.");
  const type = candidate.instancetypecode;
  const recurring = expectSeries || Boolean(candidate.seriesid) || [1, 2, 3, 4].includes(type);
  const blocked = (reason) => ({ ...candidate, editTarget: { kind: "blocked", id: "", reason }, noteBlockedReason: reason });
  if (!recurring) return type === 0
    ? { ...candidate, editTarget: { kind: "single", id, etag: candidate["@odata.etag"] } }
    : blocked("Le type Dynamics du rendez-vous n’est pas vérifiable. Actualisez avant de modifier.");
  const masterId = type === 1 ? id : normalizeGuid(candidate.seriesid);
  if (!masterId || ![1, 2, 3, 4].includes(type)) return blocked("La série Dynamics n’est pas identifiée. Aucune occurrence seule ne sera modifiée.");
  let qualificationError, metadata;
  try { metadata = await getSeriesQualificationMetadata(config, token); }
  catch { qualificationError = "Qualification de série indisponible : faire vérifier le champ rfr_os_appointmentqualification et ses droits sur RecurringAppointmentMaster. Aucun schéma modifié."; }
  const select = ["activityid", "subject", "scheduledstart", "scheduledend", "instancetypecode", "_regardingobjectid_value",
    ...(!qualificationError ? [APPOINTMENT_QUALIFICATION_FIELD] : [])].join(",");
  let master;
  try { master = await dataverseFetch(config, token, `/recurringappointmentmasters(${masterId})?$select=${select}`); }
  catch (error) {
    if ([403, 404].includes(error.status)) return blocked("Le maître Dynamics est absent ou inaccessible avec vos droits. Aucune occurrence ne sera modifiée.");
    throw error;
  }
  if (normalizeGuid(master.activityid) !== masterId || master.instancetypecode !== 1) return blocked("Le maître Dynamics n’est pas vérifiable. Aucune occurrence ne sera modifiée.");
  const regarding = getRegardingReference(master), occurrenceRegarding = getRegardingReference(candidate);
  const differs = regarding?.id !== occurrenceRegarding?.id || regarding?.logicalName !== occurrenceRegarding?.logicalName;
  return { ...candidate,
    _regardingobjectid_value: master._regardingobjectid_value,
    "_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname": master["_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname"],
    "_regardingobjectid_value@OData.Community.Display.V1.FormattedValue": master["_regardingobjectid_value@OData.Community.Display.V1.FormattedValue"],
    [APPOINTMENT_QUALIFICATION_FIELD]: qualificationError ? undefined : master[APPOINTMENT_QUALIFICATION_FIELD],
    [`${APPOINTMENT_QUALIFICATION_FIELD}@OData.Community.Display.V1.FormattedValue`]: master[`${APPOINTMENT_QUALIFICATION_FIELD}@OData.Community.Display.V1.FormattedValue`],
    occurrenceQualification: candidate[APPOINTMENT_QUALIFICATION_FIELD],
    editTarget: { kind: "series", id: masterId, etag: master["@odata.etag"], qualificationError },
    seriesNotice: metadata && !metadata.propagationLinked
      ? "Toute la série Dynamics sera modifiée. La copie de la qualification aux séances n’est pas confirmée ; les exceptions sont conservées."
      : "Toute la série Dynamics sera modifiée. Les exceptions déjà personnalisées sont conservées.",
    noteBlockedReason: differs ? "Le Concernant de cette séance diffère de celui de la série (exception ou synchronisation en attente). Vérifiez dans Dynamics avant de créer une note ; aucune note n’est déplacée." : null,
  };
}

function editingPath(candidate) {
  const target = candidate.editTarget;
  return `/${target.kind === "series" ? "recurringappointmentmasters" : "appointments"}(${target.id})`;
}

async function readCurrentTarget(config, token, appointmentId, expected) {
  const id = normalizeGuid(appointmentId);
  if (!id) throw new Error("Identifiant Dynamics invalide.");
  const master = expected?.editTarget?.kind === "series" && expected.editTarget.id === id;
  const select = master ? targetSelect.split(",").filter((field) => field !== "seriesid" &&
    !(field === APPOINTMENT_QUALIFICATION_FIELD && expected.editTarget.qualificationError)).join(",") : targetSelect;
  const row = await dataverseFetch(config, token, `/${master ? "recurringappointmentmasters" : "appointments"}(${id})?$select=${select}`);
  if (normalizeGuid(row.activityid) !== id) throw new Error("Le rendez-vous Dynamics a changé.");
  return resolveAppointmentEditingTarget(config, token, row, expected?.editTarget?.kind === "series");
}

async function readEditingTarget(config, token, id, expected) {
  const current = await readCurrentTarget(config, token, id, expected), target = current.editTarget;
  if (target.kind === "blocked") throw new Error(target.reason);
  if (target.kind === "series" && expected?.editTarget?.kind !== "series") throw new Error("Confirmez explicitement la modification de toute la série.");
  if (expected?.editTarget && (expected.activityid !== current.activityid || expected.editTarget.id !== target.id ||
    expected.editTarget.kind !== target.kind || expected.editTarget.etag !== target.etag)) throw new Error("Les informations Dynamics ont changé. Actualisez et confirmez à nouveau.");
  if (!target.etag || /[\r\n]/.test(target.etag)) throw new Error("Version Dynamics absente : aucune modification effectuée.");
  return current;
}

export async function verifyAppointmentNoteTarget(config, token, expected) {
  const current = await readCurrentTarget(config, token, expected.activityid, expected);
  if (current.noteBlockedReason || current.editTarget.kind === "blocked") throw new Error(current.noteBlockedReason || current.editTarget.reason);
  const regarding = getRegardingReference(current), before = getRegardingReference(expected);
  if (!regarding || !before || regarding.id !== before.id || regarding.logicalName !== before.logicalName ||
    current.editTarget.id !== expected.editTarget.id || current.editTarget.kind !== expected.editTarget.kind) throw new Error("Le Concernant a changé. Actualisez puis confirmez à nouveau avant de créer la note.");
  return regarding;
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

export async function getAppointmentNotes(config, token, regarding, appointmentId) {
  const marker = appointmentNoteReference(appointmentId);
  const id = normalizeGuid(appointmentId), target = normalizeGuid(regarding?.id);
  if (!target) throw new Error("Le Concernant du rendez-vous doit être vérifié.");
  const filter = encodeURIComponent(`_objectid_value eq ${id} or (_objectid_value eq ${target} and contains(notetext,'${marker}'))`);
  const path = `/annotations?$select=annotationid,subject,notetext,createdon,_objectid_value&$filter=${filter}&$orderby=createdon desc&$top=20`;
  const result = await dataverseFetch(config, token, path);
  return (result.value || []).filter((note) => normalizeGuid(note._objectid_value) === id ||
    (normalizeGuid(note._objectid_value) === target && belongsToAppointment(note.notetext, id))).slice(0, 20);
}

export function buildRegardingNotePayload(regarding, entitySetName, note, appointmentId) {
  return {
    subject: note.subject,
    notetext: appointmentId ? withAppointmentReference(note.text, appointmentId) : note.text,
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

export async function createRegardingNote(config, token, regarding, note, appointmentId, ensureCurrent = () => {}) {
  appointmentNoteReference(appointmentId); // Bloquer avant tout appel si la séance n’est pas identifiée.
  const entitySetName = await getEntitySetName(config, token, regarding.logicalName);
  await ensureCurrent();
  return dataverseFetch(config, token, "/annotations", {
    method: "POST",
    headers: {
      Prefer: "return=representation",
    },
    body: JSON.stringify(buildRegardingNotePayload(regarding, entitySetName, note, appointmentId)),
  });
}

export function getRegardingLabel(candidate) {
  return (
    candidate["_regardingobjectid_value@OData.Community.Display.V1.FormattedValue"] ||
    "Aucun concernant"
  );
}
