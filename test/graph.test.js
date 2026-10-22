import test from "node:test";
import assert from "node:assert/strict";
import { chooseExactGraphEvent } from "../src/graph.js";

test("the Outlook event requires the same subject and times", () => {
  const context = {
    subject: "Visite client",
    start: new Date("2026-10-22T10:00:00.000Z"),
    end: new Date("2026-10-22T11:00:00.000Z"),
  };
  const events = [
    {
      id: "exact",
      subject: "Visite client",
      start: { dateTime: "2026-10-22T10:00:00.0000000", timeZone: "UTC" },
      end: { dateTime: "2026-10-22T11:00:00.0000000", timeZone: "UTC" },
    },
    {
      id: "other",
      subject: "Autre réunion",
      start: { dateTime: "2026-10-22T10:00:00.0000000", timeZone: "UTC" },
      end: { dateTime: "2026-10-22T11:00:00.0000000", timeZone: "UTC" },
    },
  ];

  assert.equal(chooseExactGraphEvent(events, context)?.id, "exact");
});

test("two identical Outlook events remain ambiguous", () => {
  const context = {
    subject: "Visite client",
    start: new Date("2026-10-22T10:00:00.000Z"),
    end: new Date("2026-10-22T11:00:00.000Z"),
  };
  const event = {
    subject: "Visite client",
    start: { dateTime: "2026-10-22T10:00:00Z", timeZone: "UTC" },
    end: { dateTime: "2026-10-22T11:00:00Z", timeZone: "UTC" },
  };

  assert.equal(
    chooseExactGraphEvent(
      [
        { ...event, id: "a" },
        { ...event, id: "b" },
      ],
      context,
    ),
    null,
  );
});
