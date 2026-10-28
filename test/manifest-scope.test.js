import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

test("RDV & Notes v2.3.4 : meme identite, surfaces invite/organisateur, notes simples uniquement", () => {
  const manifest = fs.readFileSync(new URL("../manifest.template.xml", import.meta.url), "utf8");
  const main = fs.readFileSync(new URL("../src/main.js", import.meta.url), "utf8");
  const metadata = JSON.parse(fs.readFileSync(new URL("../package.json", import.meta.url), "utf8"));
  assert.equal(metadata.version, "2.3.4");
  assert.match(manifest, /<Id>0a316d2c-0556-4628-9bb3-bdc1d4735cba<\/Id>/);
  assert.match(manifest, /<DisplayName DefaultValue="RDV &amp; Notes"/);
  for (const surface of ["AppointmentAttendeeCommandSurface", "AppointmentOrganizerCommandSurface", "MobileLogEventAppointmentAttendee"]) {
    assert.match(manifest, new RegExp(`xsi:type="${surface}"`));
  }
  assert.match(main, /Ajouter une note manuelle/);
  assert.match(main, /Ouvrir dans RDV &amp; Notes/);
  assert.ok(!manifest.includes('<AppDomain>https://cr-api.roth-france.fr</AppDomain>'), 'le domaine mobile n’est pas un domaine de navigation interne');
  for (const feature of ["MediaRecorder", "getUserMedia", "/v1/jobs", "report_queued", "generating_report"]) {
    assert.ok(!main.includes(feature));
  }
});
