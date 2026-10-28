import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const manifest = fs.readFileSync(new URL('../manifest.template.xml', import.meta.url), 'utf8');
function dimensions(relative) {
  const data = fs.readFileSync(new URL('../public/' + relative, import.meta.url));
  assert.equal(data.subarray(0, 8).toString('hex'), '89504e470d0a1a0a');
  return [data.readUInt32BE(16), data.readUInt32BE(20), data[24], data[25]];
}
test('logo Outlook : tous les emplacements utilisent le nouveau logo PNG RGBA et les bonnes tailles', () => {
  const resources = new Map([...manifest.matchAll(/<bt:Image id="([^"]+)" DefaultValue="\{\{APP_URL\}\}\/([^"?]+)"/g)]
    .map((match) => [match[1], match[2]]));
  const images = [...manifest.matchAll(/<bt:Image size="(\d+)"(?: scale="(\d+)")? resid="([^"]+)"/g)];
  assert.ok(images.length >= 15);
  for (const [, size, scale, resource] of images) {
    const path = resources.get(resource);
    assert.ok(path?.startsWith('assets/rdv-notes-logo-'));
    const expected = Number(size) * Number(scale || 1);
    assert.deepEqual(dimensions(path), [expected, expected, 8, 6]);
  }
  for (const [element, size] of [['IconUrl', 64], ['HighResolutionIconUrl', 128]]) {
    const match = manifest.match(new RegExp(`<${element} DefaultValue="\\{\\{APP_URL\\}\\}/([^"?]+)"`));
    assert.ok(match); assert.deepEqual(dimensions(match[1]), [size, size, 8, 6]);
  }
  assert.ok(!manifest.includes('/assets/icon-'));
});
test('logo Outlook : meme GUID, version actualisee sans nouvelle permission', () => {
  const metadata = JSON.parse(fs.readFileSync(new URL('../package.json', import.meta.url), 'utf8'));
  assert.equal(metadata.version, '2.3.4'); assert.equal(metadata.manifestRevision, 0);
  assert.match(manifest, /<Id>0a316d2c-0556-4628-9bb3-bdc1d4735cba<\/Id>/);
  assert.match(manifest, /<Permissions>ReadItem<\/Permissions>/);
});
