import test from "node:test";
import assert from "node:assert/strict";
import { retry, withTimeout } from "../src/resilience.js";

test("withTimeout returns a completed operation", async () => {
  assert.equal(await withTimeout(Promise.resolve("ok"), 50, "trop long"), "ok");
});

test("withTimeout returns a stable diagnostic code", async () => {
  await assert.rejects(
    withTimeout(new Promise(() => {}), 5, "Outlook ne répond pas", "OUTLOOK_TIMEOUT"),
    (error) => error.code === "OUTLOOK_TIMEOUT" && /Outlook/.test(error.message),
  );
});

test("retry retries once before succeeding", async () => {
  let calls = 0;
  const result = await retry(
    async () => {
      calls += 1;
      if (calls === 1) throw new Error("premier échec");
      return "ok";
    },
    { attempts: 2, delayMilliseconds: 0 },
  );
  assert.equal(result, "ok");
  assert.equal(calls, 2);
});
