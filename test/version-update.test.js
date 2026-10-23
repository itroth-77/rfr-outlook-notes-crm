import test from "node:test";
import assert from "node:assert/strict";
import { compareVersions, redirectToLatestVersion } from "../src/version-update.js";

test("semantic versions are compared numerically", () => {
  assert.equal(compareVersions("2.3.1", "2.3.0"), 1);
  assert.equal(compareVersions("2.3.0", "2.3.1"), -1);
  assert.equal(compareVersions("2.3", "2.3.0"), 0);
});

test("an outdated task pane redirects to the latest uncached version", async () => {
  let redirectedUrl = "";
  const redirected = await redirectToLatestVersion("2.3.0", {
    baseUrl: "https://example.test/index.html?v=2.3.0",
    fetchImpl: async () => ({
      ok: true,
      json: async () => ({ version: "2.3.2" }),
    }),
    replace: (url) => {
      redirectedUrl = url;
    },
    now: () => 12345,
  });

  assert.equal(redirected, true);
  assert.equal(redirectedUrl, "https://example.test/index.html?v=2.3.2&_=12345");
});

test("a current task pane is not redirected", async () => {
  const redirected = await redirectToLatestVersion("2.3.2", {
    baseUrl: "https://example.test/index.html?v=2.3.2",
    fetchImpl: async () => ({
      ok: true,
      json: async () => ({ version: "2.3.2" }),
    }),
    replace: () => assert.fail("No redirect expected"),
    now: () => 12345,
  });

  assert.equal(redirected, false);
});
