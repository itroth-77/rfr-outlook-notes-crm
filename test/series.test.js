import test from "node:test";
import assert from "node:assert/strict";
import { resolveOutlookEvent, trackOutlookAppointment } from "../src/graph.js";
import { isOutlookMobile, getRestItemId, assertAppointmentContext } from "../src/outlook.js";
import { mobileAppointmentLink } from "../src/mobile-link.js";
import { resolveAppointmentEditingTarget, updateAppointmentQualification, updateAppointmentRegarding,
  verifyAppointmentNoteTarget, APPOINTMENT_QUALIFICATION_FIELD as FIELD } from "../src/dataverse.js";

async function withFetch(fetcher, task) {
  const before = globalThis.fetch; globalThis.fetch = fetcher;
  try { await task(); } finally { globalThis.fetch = before; }
}
const clone = (value) => structuredClone(value);
const context = { subject: "Rendez-vous test", start: new Date("2026-10-27T09:00:00Z"), end: new Date("2026-10-27T10:00:00Z"),
  itemId: "instance/1+", restItemId: "instance/1+", itemIdFormat: "restId", isMobile: true };
const category = "Tracked To Dynamics 365";
function graphFixture({ type = "occurrence", organizer = true, tracked = false, status = 0 } = {}) {
  const calls = [];
  const event = { id: context.restItemId, type, seriesMasterId: type === "occurrence" || type === "exception" ? "series/ABC+" : undefined,
    subject: context.subject, start: { dateTime: context.start.toISOString() }, end: { dateTime: context.end.toISOString() },
    categories: ["Client", category], isOrganizer: organizer, "@odata.etag": 'W/"instance-1"' };
  const master = { ...clone(event), id: "series/ABC+", type: "seriesMaster", seriesMasterId: undefined,
    categories: tracked ? ["Commercial", category] : ["Commercial"], "@odata.etag": 'W/"master-1"' };
  return { calls, event, master, fetch: async (url, options) => {
    calls.push({ url, ...options });
    const path = decodeURIComponent(new URL(url).pathname);
    const row = path.endsWith("/" + event.id) ? event : path.endsWith("/" + master.id) ? master : null;
    if (!row) return Response.json({}, { status: 404 });
    if (options.method === "PATCH") {
      if (status) return Response.json({ error: { message: "refused" } }, { status });
      Object.assign(row, JSON.parse(options.body));
    }
    return Response.json(row);
  } };
}

for (const organizer of [true, false]) for (const type of ["occurrence", "exception"]) {
  test(`suivi Outlook ${type}, ${organizer ? "organisateur" : "invite"} : maitre /me et categories preservees`, async () => {
    const f = graphFixture({ organizer, type });
    await withFetch(f.fetch, async () => {
      const result = await trackOutlookAppointment("fake", context, clone(f.event));
      const writes = f.calls.filter((call) => call.method === "PATCH");
      assert.equal(writes.length, 1); assert.ok(writes[0].url.endsWith("/me/events/series%2FABC%2B"));
      assert.equal(writes[0].headers["If-Match"], 'W/"master-1"');
      assert.deepEqual(JSON.parse(writes[0].body), { categories: ["Commercial", category] });
      assert.equal(result.eventId, context.restItemId, "le lien mobile conserve la seance");
      assert.ok(f.calls.every((call) => new URL(call.url).pathname.startsWith("/v1.0/me/")));
    });
  });
}

test("serie deja suivie : aucune ecriture, occurrence non categorisee preservee", async () => {
  const f = graphFixture({ type: "exception", tracked: true }); f.event.categories = ["Client"];
  await withFetch(f.fetch, async () => {
    const result = await trackOutlookAppointment("fake", context, clone(f.event));
    assert.equal(f.calls.filter((call) => call.method === "PATCH").length, 0);
    assert.match(result.notice, /exception/); assert.deepEqual(f.event.categories, ["Client"]);
  });
});

test("Outlook non recurrent : seul evenement modifie et reponse verifiee", async () => {
  const f = graphFixture({ type: "singleInstance" }); f.event.categories = ["Client"];
  await withFetch(f.fetch, async () => {
    await trackOutlookAppointment("fake", context, clone(f.event));
    const write = f.calls.find((call) => call.method === "PATCH");
    assert.ok(write.url.endsWith("instance%2F1%2B")); assert.deepEqual(JSON.parse(write.body).categories, ["Client", category]);
  });
});

for (const status of [403, 409, 412]) test(`Outlook ${status} : un refus, aucun fallback occurrence`, async () => {
  const f = graphFixture({ status });
  await withFetch(f.fetch, async () => {
    await assert.rejects(trackOutlookAppointment("fake", context, clone(f.event)));
    assert.equal(f.calls.filter((call) => call.method === "PATCH").length, 1);
  });
});

for (const scenario of ["master", "etag", "scope", "selection"]) test(`Outlook refuse ${scenario} avant PATCH`, async () => {
  const f = graphFixture(); const expected = clone(f.event);
  if (scenario === "master") f.master.type = "occurrence";
  if (scenario === "etag") delete f.master["@odata.etag"];
  if (scenario === "scope") expected.seriesMasterId = "another";
  await withFetch(f.fetch, async () => {
    await assert.rejects(trackOutlookAppointment("fake", context, expected, () => { if (scenario === "selection") throw new Error("changed"); }));
    assert.equal(f.calls.filter((call) => call.method === "PATCH").length, 0);
  });
});

test("identifiant natif requis : pas de fallback sur sujet/horaires ni de lien ambigu", async () => {
  await withFetch(() => { throw new Error("aucun reseau attendu"); }, async () => {
    await assert.rejects(resolveOutlookEvent("fake", { ...context, itemId: "", restItemId: "" }));
  });
  const url = new URL(mobileAppointmentLink("AAMk/ABC+=="));
  assert.equal(url.origin, "https://cr-api.roth-france.fr"); assert.equal(url.searchParams.get("eventId"), "AAMk/ABC+==");
  assert.equal(url.searchParams.get("view"), "crm"); assert.deepEqual([...url.searchParams.keys()], ["eventId", "view"]);
  assert.throws(() => mobileAppointmentLink("x\r\n"));
});

test("Office mobile : REST ID direct ; bureau : conversion EWS ; changement selection refuse", () => {
  for (const platform of ["Android", "iOS"]) {
    const office = { context: { platform, mailbox: { item: { itemId: "id" } } } };
    assert.equal(isOutlookMobile(office), true); assert.equal(getRestItemId({ itemId: "id" }, office), "id");
    assertAppointmentContext({ itemId: "id" }, office); assert.throws(() => assertAppointmentContext({ itemId: "old" }, office));
  }
  const office = { context: { platform: "PC", mailbox: { convertToRestId: (id, version) => { assert.equal(version, "v2.0"); return "rest-" + id; } } }, MailboxEnums: { RestVersion: { v2_0: "v2.0" } } };
  assert.equal(getRestItemId({ itemId: "ews" }, office), "rest-ews");
});

const config = { dataverseUrl: "https://example.crm4.dynamics.com" };
const id = "11111111-1111-4111-8111-111111111111", masterId = "22222222-2222-4222-8222-222222222222";
const targetId = "33333333-3333-4333-8333-333333333333", metadataId = "44444444-4444-4444-8444-444444444444";
const annotation = "_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname";
function crmFixture({ type = 2, missingMaster = 0, noField = false, secured = false, linked = true, relationships = 1, denyWrite = 0 } = {}) {
  const calls = [];
  const occurrence = { activityid: id, subject: "Test", instancetypecode: type, seriesid: type ? masterId : null,
    _regardingobjectid_value: targetId, [annotation]: "account", [FIELD]: 1, "@odata.etag": 'W/"instance-1"' };
  const master = { ...clone(occurrence), activityid: masterId, instancetypecode: 1, seriesid: undefined, "@odata.etag": 'W/"master-1"' };
  const metadata = { MetadataId: metadataId, LogicalName: FIELD, LinkedAttributeId: linked ? metadataId : null,
    IsValidForRead: true, IsValidForUpdate: true, IsSecured: secured, OptionSet: { Options: [1, 2].map((Value) => ({ Value, Label: { UserLocalizedLabel: { Label: "Choix " + Value } } })) } };
  return { calls, occurrence, master, fetch: async (url, options) => {
    calls.push({ url, ...options });
    const path = decodeURIComponent(new URL(url).pathname + new URL(url).search);
    if (path.includes("EntityDefinitions")) {
      if (path.includes("ManyToOneRelationships")) {
        const type = path.match(/ReferencedEntity eq '([^']+)'/)[1];
        return Response.json({ value: Array.from({ length: relationships }, (_, n) => ({ ReferencingEntityNavigationPropertyName: `regardingobjectid_${type}_${n}` })) });
      }
      if (path.includes("recurringappointmentmaster") && noField) return Response.json({}, { status: 404 });
      return Response.json(metadata);
    }
    const row = path.includes("/recurringappointmentmasters(") ? master : path.includes("/appointments(") ? occurrence : null;
    if (!row) throw new Error("Endpoint inattendu " + path);
    if (row === master && missingMaster) return Response.json({}, { status: missingMaster });
    if (options.method === "PATCH") {
      if (denyWrite) return Response.json({}, { status: denyWrite });
      const body = JSON.parse(options.body);
      if (FIELD in body) row[FIELD] = body[FIELD];
      else {
        const key = Object.keys(body)[0];
        row._regardingobjectid_value = body[key].match(/\(([^)]+)\)/)[1];
        row[annotation] = key.includes("contact") ? "contact" : key.includes("opportunity") ? "opportunity" : "account";
      }
      row["@odata.etag"] = 'W/"updated"'; return new Response(null, { status: 204 });
    }
    return Response.json(row);
  } };
}

test("Dynamics qualif : maitre resolu via seriesid, ETag, relecture et exception intacte", async () => {
  const f = crmFixture();
  await withFetch(f.fetch, async () => {
    const expected = await resolveAppointmentEditingTarget(config, "fake", f.occurrence, true);
    const saved = await updateAppointmentQualification(config, "fake", id, 2, expected);
    const writes = f.calls.filter((call) => call.method === "PATCH");
    assert.equal(writes.length, 1); assert.ok(writes[0].url.endsWith(`/recurringappointmentmasters(${masterId})`));
    assert.equal(writes[0].headers["If-Match"], 'W/"master-1"');
    assert.deepEqual(JSON.parse(writes[0].body), { [FIELD]: 2 }); assert.equal(f.occurrence[FIELD], 1);
    assert.equal(saved[FIELD], 2); assert.equal(saved.activityid, id); assert.match(saved.seriesNotice, /séance/);
  });
});

for (const type of ["account", "contact", "opportunity"]) test(`Dynamics Concernant ${type} : relation du maitre issue des metadonnees`, async () => {
  const f = crmFixture();
  await withFetch(f.fetch, async () => {
    const expected = await resolveAppointmentEditingTarget(config, "fake", f.occurrence, true);
    const nextId = "55555555-5555-4555-8555-555555555555";
    const result = await updateAppointmentRegarding(config, "fake", id, { id: nextId, logicalName: type }, expected);
    const writes = f.calls.filter((call) => call.method === "PATCH");
    assert.equal(writes.length, 1); assert.ok(writes[0].url.includes("/recurringappointmentmasters("));
    assert.equal(Object.keys(JSON.parse(writes[0].body))[0], `regardingobjectid_${type}_0@odata.bind`);
    assert.equal(f.occurrence._regardingobjectid_value, targetId); assert.match(result.noteBlockedReason, /diffère/);
    await assert.rejects(verifyAppointmentNoteTarget(config, "fake", result), /diffère/);
  });
});

for (const options of [{ missingMaster: 403 }, { missingMaster: 404 }, { type: 0 }]) test("Dynamics : maitre absent/inaccessible/incoherent, aucun fallback occurrence", async () => {
  const f = crmFixture(options);
  await withFetch(f.fetch, async () => {
    const expected = await resolveAppointmentEditingTarget(config, "fake", f.occurrence, true);
    assert.equal(expected.editTarget.kind, "blocked");
    await assert.rejects(updateAppointmentQualification(config, "fake", id, 2, expected));
    assert.equal(f.calls.filter((call) => call.method === "PATCH").length, 0);
  });
});

for (const options of [{ noField: true }, { secured: true }]) test("qualification indisponible : pas d'ecriture ni de schema, Concernant encore modifiable", async () => {
  const f = crmFixture(options);
  await withFetch(f.fetch, async () => {
    const expected = await resolveAppointmentEditingTarget(config, "fake", f.occurrence, true);
    assert.equal(expected.editTarget.kind, "series"); assert.ok(expected.editTarget.qualificationError);
    await assert.rejects(updateAppointmentQualification(config, "fake", id, 2, expected));
    assert.equal(f.calls.filter((call) => call.method === "PATCH").length, 0);
    await updateAppointmentRegarding(config, "fake", id, { id: targetId, logicalName: "account" }, expected);
    assert.equal(f.calls.filter((call) => call.method === "PATCH").length, 1);
  });
});

for (const scenario of ["etag", "confirmation", "relationship", "selection", "value"]) test(`Dynamics ${scenario} : refus avant PATCH`, async () => {
  const f = crmFixture({ relationships: scenario === "relationship" ? 2 : 1 });
  await withFetch(f.fetch, async () => {
    const expected = await resolveAppointmentEditingTarget(config, "fake", f.occurrence, true);
    if (scenario === "etag") f.master["@odata.etag"] = 'W/"another"';
    if (scenario === "relationship") await assert.rejects(updateAppointmentRegarding(config, "fake", id, { id: targetId, logicalName: "account" }, expected));
    else await assert.rejects(updateAppointmentQualification(config, "fake", id, scenario === "value" ? 999 : 2,
      scenario === "confirmation" ? undefined : expected, () => { if (scenario === "selection") throw new Error("changed"); }));
    assert.equal(f.calls.filter((call) => call.method === "PATCH").length, 0);
  });
});

test("Dynamics sans lien qualification : maitre modifiable, propagation non promise", async () => {
  const f = crmFixture({ linked: false });
  await withFetch(f.fetch, async () => {
    const expected = await resolveAppointmentEditingTarget(config, "fake", f.occurrence, true);
    assert.match(expected.seriesNotice, /pas confirmée/);
    await updateAppointmentQualification(config, "fake", id, 2, expected);
    assert.equal(f.master[FIELD], 2); assert.equal(f.occurrence[FIELD], 1);
  });
});

for (const denyWrite of [403, 412]) test(`Dynamics ${denyWrite} : refus conserve, aucune relance/fallback`, async () => {
  const f = crmFixture({ denyWrite });
  await withFetch(f.fetch, async () => {
    const expected = await resolveAppointmentEditingTarget(config, "fake", f.occurrence, true);
    await assert.rejects(updateAppointmentQualification(config, "fake", id, 2, expected));
    assert.equal(f.calls.filter((call) => call.method === "PATCH").length, 1);
  });
});

test("Dynamics RDV simple : PATCH individuel conserve", async () => {
  const f = crmFixture({ type: 0 });
  await withFetch(f.fetch, async () => {
    const expected = await resolveAppointmentEditingTarget(config, "fake", f.occurrence);
    await updateAppointmentQualification(config, "fake", id, 2, expected);
    assert.equal(f.calls.filter((call) => call.method === "PATCH")[0].url, `${config.dataverseUrl}/api/data/v9.2/appointments(${id})`);
    assert.ok(f.calls.every((call) => !call.url.includes("recurringappointmentmasters")));
  });
});
