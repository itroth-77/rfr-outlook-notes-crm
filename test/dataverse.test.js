import test from "node:test";
import assert from "node:assert/strict";
import {
  buildAppointmentQuery,
  buildAppointmentRegardingPayload,
  buildCurrentUserActivityPartyQuery,
  buildOpportunitySearchQuery,
  buildRegardingSearchQuery,
  buildRegardingNotePayload,
  chooseUnambiguousCandidate,
  escapeODataString,
  extractQualificationOptions,
  filterAppointmentsForCurrentUser,
  getRegardingReference,
  mapRegardingSearchRecord,
  normalizeGuid,
  rankAppointmentCandidates,
} from "../src/dataverse.js";

test("escapeODataString doubles apostrophes", () => {
  assert.equal(escapeODataString("L'équipe d'Aline"), "L''équipe d''Aline");
});

test("Dataverse GUIDs are accepted regardless of RFC version and variant bits", () => {
  assert.equal(
    normalizeGuid("{11111111-2222-7333-6444-555555555555}"),
    "11111111-2222-7333-6444-555555555555",
  );
  assert.equal(normalizeGuid("not-a-guid"), "");
});

test("regarding account search accepts percent and star wildcards", () => {
  const percentQuery = decodeURIComponent(
    buildRegardingSearchQuery("account", "%EDEO%GNY%"),
  );
  const starQuery = decodeURIComponent(
    buildRegardingSearchQuery("account", "*EDEO*GNY*"),
  );

  for (const query of [percentQuery, starQuery]) {
    assert.match(query, /\/accounts\?/);
    assert.match(query, /contains\(name,'EDEO'\)/);
    assert.match(query, /contains\(accountnumber,'EDEO'\)/);
    assert.match(query, / and /);
    assert.match(query, /contains\(name,'GNY'\)/);
    assert.match(query, /contains\(accountnumber,'GNY'\)/);
  }
});

test("regarding contact wildcards search both name and email", () => {
  const query = decodeURIComponent(
    buildRegardingSearchQuery("contact", "%Dupont%roth-france%"),
  );

  assert.match(query, /contains\(fullname,'Dupont'\)/);
  assert.match(query, /contains\(emailaddress1,'Dupont'\)/);
  assert.match(query, / and /);
  assert.match(query, /contains\(fullname,'roth-france'\)/);
  assert.match(query, /contains\(emailaddress1,'roth-france'\)/);
});

test("regarding opportunity search combines its number and related business", () => {
  const query = decodeURIComponent(
    buildOpportunitySearchQuery(
      "%403151%MA REXEL%",
      [[], ["11111111-2222-7333-6444-555555555555"]],
    ),
  );

  assert.match(query, /\/opportunities\?/);
  assert.match(query, /\$select=opportunityid,name,_rfr_lk_business_value/);
  assert.match(query, /\(contains\(name,'403151'\)\)/);
  assert.match(query, / and /);
  assert.match(query, /contains\(name,'MA REXEL'\)/);
  assert.match(
    query,
    /_rfr_lk_business_value eq 11111111-2222-7333-6444-555555555555/,
  );
});

test("regarding opportunity result displays number and formatted business name", () => {
  const result = mapRegardingSearchRecord("opportunity", {
    opportunityid: "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee",
    name: "403151",
    _rfr_lk_business_value: "11111111-2222-7333-6444-555555555555",
    "_rfr_lk_business_value@OData.Community.Display.V1.FormattedValue": "MA REXEL",
  });

  assert.deepEqual(result, {
    id: "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee",
    logicalName: "opportunity",
    label: "403151",
    detail: "MA REXEL",
  });
});

test("regarding wildcard search requires two literal characters", () => {
  assert.throws(
    () => buildRegardingSearchQuery("account", "%%%*"),
    /2 caractères hors jokers/,
  );
});

test("buildAppointmentQuery creates a 30-minute UTC window", () => {
  const query = buildAppointmentQuery({ start: new Date("2026-10-22T10:00:00.000Z") });
  const decoded = decodeURIComponent(query);
  assert.match(decoded, /2026-10-22T09:45:00.000Z/);
  assert.match(decoded, /2026-10-22T10:15:00.000Z/);
  assert.match(decoded, /\$top=20/);
  assert.match(decoded, /rfr_os_appointmentqualification/);
});

test("qualification metadata is converted to French sorted choices", () => {
  const options = extractQualificationOptions({
    OptionSet: {
      Options: [
        {
          Value: 2,
          Label: {
            UserLocalizedLabel: { Label: "Other" },
            LocalizedLabels: [
              { LanguageCode: 1033, Label: "Other" },
              { LanguageCode: 1036, Label: "Divers" },
            ],
          },
        },
        {
          Value: 1,
          Label: {
            UserLocalizedLabel: { Label: "Commercial" },
            LocalizedLabels: [{ LanguageCode: 1036, Label: "Commercial" }],
          },
        },
      ],
    },
  });

  assert.deepEqual(options, [
    { value: 1, label: "Commercial" },
    { value: 2, label: "Divers" },
  ]);
});

test("activity party query only requests organizer and attendee roles", () => {
  const query = decodeURIComponent(
    buildCurrentUserActivityPartyQuery(
      [
        { activityid: "11111111-2222-3333-8444-555555555555" },
        { activityid: "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee" },
      ],
      "99999999-8888-4777-8666-555555555555",
      "sebastien.ribou@roth-france.fr",
    ),
  );

  assert.match(query, /_partyid_value eq 99999999-8888-4777-8666-555555555555/);
  assert.match(query, /addressused eq 'sebastien\.ribou@roth-france\.fr'/);
  assert.match(query, /participationtypemask eq 5/);
  assert.match(query, /participationtypemask eq 6/);
  assert.match(query, /participationtypemask eq 7/);
  assert.doesNotMatch(query, /participationtypemask eq 9/);
});

test("unrelated appointments are removed from the candidate list", () => {
  const appointments = [
    { activityid: "11111111-2222-3333-8444-555555555555", subject: "Visible" },
    { activityid: "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee", subject: "Masqué" },
  ];
  const parties = [
    {
      _activityid_value: "11111111-2222-3333-8444-555555555555",
      participationtypemask: 6,
    },
    {
      _activityid_value: "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee",
      participationtypemask: 9,
    },
  ];

  assert.deepEqual(filterAppointmentsForCurrentUser(appointments, parties), [appointments[0]]);
});

test("the Dataverse regarding lookup is read with its polymorphic type", () => {
  const regarding = getRegardingReference({
    _regardingobjectid_value: "{11111111-2222-3333-4444-555555555555}",
    "_regardingobjectid_value@Microsoft.Dynamics.CRM.lookuplogicalname": "account",
    "_regardingobjectid_value@OData.Community.Display.V1.FormattedValue": "ROTH FRANCE",
  });

  assert.deepEqual(regarding, {
    id: "11111111-2222-3333-4444-555555555555",
    logicalName: "account",
    label: "ROTH FRANCE",
  });
});

test("a note is bound to the regarding record instead of the appointment", () => {
  const payload = buildRegardingNotePayload(
    {
      id: "11111111-2222-3333-4444-555555555555",
      logicalName: "account",
    },
    "accounts",
    { subject: "Compte rendu", text: "Décision prise" },
  );

  assert.deepEqual(payload, {
    subject: "Compte rendu",
    notetext: "Décision prise",
    "objectid_account@odata.bind":
      "/accounts(11111111-2222-3333-4444-555555555555)",
  });
  assert.equal("objectid_appointment@odata.bind" in payload, false);
});

test("the appointment regarding payload uses the correct polymorphic relationship", () => {
  assert.deepEqual(
    buildAppointmentRegardingPayload(
      "account",
      "11111111-2222-4333-8444-555555555555",
    ),
    {
      "regardingobjectid_account_appointment@odata.bind":
        "/accounts(11111111-2222-4333-8444-555555555555)",
    },
  );
  assert.deepEqual(
    buildAppointmentRegardingPayload(
      "contact",
      "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee",
    ),
    {
      "regardingobjectid_contact_appointment@odata.bind":
        "/contacts(aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee)",
    },
  );
});

test("an exact subject and exact times are selected automatically", () => {
  const context = {
    subject: "Visite client",
    start: new Date("2026-10-22T10:00:00.000Z"),
    end: new Date("2026-10-22T11:00:00.000Z"),
  };
  const ranked = rankAppointmentCandidates(
    [
      {
        activityid: "a",
        subject: "Visite client",
        scheduledstart: "2026-10-22T10:00:00.000Z",
        scheduledend: "2026-10-22T11:00:00.000Z",
      },
      {
        activityid: "b",
        subject: "Autre rendez-vous",
        scheduledstart: "2026-10-22T10:00:00.000Z",
        scheduledend: "2026-10-22T11:00:00.000Z",
      },
    ],
    context,
  );
  assert.equal(chooseUnambiguousCandidate(ranked)?.activityid, "a");
});

test("two equally strong CRM matches remain unselected", () => {
  const context = {
    subject: "Visite client",
    start: new Date("2026-10-22T10:00:00.000Z"),
    end: new Date("2026-10-22T11:00:00.000Z"),
  };
  const ranked = rankAppointmentCandidates(
    ["a", "b"].map((activityid) => ({
      activityid,
      subject: "Visite client",
      scheduledstart: "2026-10-22T10:00:00.000Z",
      scheduledend: "2026-10-22T11:00:00.000Z",
    })),
    context,
  );
  assert.equal(chooseUnambiguousCandidate(ranked), null);
});
