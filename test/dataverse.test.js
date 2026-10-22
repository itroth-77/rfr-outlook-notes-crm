import test from "node:test";
import assert from "node:assert/strict";
import {
  buildAppointmentQuery,
  chooseUnambiguousCandidate,
  escapeODataString,
  rankAppointmentCandidates,
} from "../src/dataverse.js";

test("escapeODataString doubles apostrophes", () => {
  assert.equal(escapeODataString("L'équipe d'Aline"), "L''équipe d''Aline");
});

test("buildAppointmentQuery creates a 30-minute UTC window", () => {
  const query = buildAppointmentQuery({ start: new Date("2026-10-22T10:00:00.000Z") });
  const decoded = decodeURIComponent(query);
  assert.match(decoded, /2026-10-22T09:45:00.000Z/);
  assert.match(decoded, /2026-10-22T10:15:00.000Z/);
  assert.match(decoded, /\$top=20/);
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

test("two equally strong matches require an explicit choice", () => {
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
