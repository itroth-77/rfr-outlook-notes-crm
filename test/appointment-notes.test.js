import test from 'node:test';
import assert from 'node:assert/strict';
import { appointmentNoteReference, withAppointmentReference, belongsToAppointment, withoutAppointmentReference } from '../src/appointment-notes.js';
import { getAppointmentNotes, buildRegardingNotePayload, createRegardingNote } from '../src/dataverse.js';

const id = '22222222-2222-4222-8222-222222222222';
const other = '33333333-3333-4333-8333-333333333333';
const target = { id: '11111111-1111-4111-8111-111111111111', logicalName: 'account' };
const config = { dataverseUrl: 'https://example.crm4.dynamics.com' };
async function withFetch(fetcher, task) {
  const before = globalThis.fetch; globalThis.fetch = fetcher;
  try { await task(); } finally { globalThis.fetch = before; }
}
test('notes : reference partagee APK/Outlook, GUID de seance valide et footer discret retire de l’affichage', () => {
  assert.equal(appointmentNoteReference('{' + id.toUpperCase() + '}'), `[RFR-RDV:${id}]`);
  assert.throws(() => appointmentNoteReference('Exchange-series-id'));
  const text = withAppointmentReference('Observation', id);
  assert.equal(belongsToAppointment(text, id), true);
  assert.equal(belongsToAppointment(text, other), false);
  assert.equal(belongsToAppointment(text + '\nSuite sans référence', id), false);
  assert.equal(withoutAppointmentReference(text), 'Observation');
  assert.equal(belongsToAppointment('<p>' + text + '</p>', id), true);
  assert.equal(withoutAppointmentReference('<p>' + text + '</p>'), '<p>Observation</p>');
});
test('notes : requete serveur sur seance, pas le maitre ni tout le Concernant ; revalidation et limite 20', async () => {
  const rows = [
    { annotationid: 'session', _objectid_value: id, notetext: 'Note directement sur le RDV' },
    { annotationid: 'good', _objectid_value: target.id, notetext: withAppointmentReference('Observation', id) },
    { annotationid: 'other-session', _objectid_value: target.id, notetext: withAppointmentReference('Autre séance, même titre', other) },
    { annotationid: 'unmarked', _objectid_value: target.id, notetext: 'Même titre, même date' },
    { annotationid: 'wrong-target', _objectid_value: other, notetext: withAppointmentReference('Autre concernant', id) },
  ];
  const calls = [];
  await withFetch(async (url) => { calls.push(decodeURIComponent(url)); return Response.json({ value: rows }); }, async () => {
    const notes = await getAppointmentNotes(config, 'fake', target, id);
    assert.deepEqual(notes.map((note) => note.annotationid), ['session', 'good']);
    assert.match(calls[0], new RegExp(`_objectid_value eq ${id} or \\(_objectid_value eq ${target.id} and contains`));
    assert.match(calls[0], /\$top=20$/); assert.equal(calls.length, 1);
  });
  await withFetch(async () => Response.json({ value: Array(25).fill(rows[1]) }), async () => {
    assert.equal((await getAppointmentNotes(config, 'fake', target, id)).length, 20);
  });
});
test('notes : permissions refusees sans repli sur toutes les notes ; ID absent bloque avant reseau', async () => {
  let calls = 0;
  await withFetch(async () => { calls++; return Response.json({}, { status: 403 }); }, async () => {
    await assert.rejects(getAppointmentNotes(config, 'fake', target, id)); assert.equal(calls, 1);
    await assert.rejects(getAppointmentNotes(config, 'fake', target, 'exchange-id')); assert.equal(calls, 1);
    await assert.rejects(createRegardingNote(config, 'fake', target, { subject: 'Note', text: 'Texte' })); assert.equal(calls, 1);
  });
});
test('nouvelle note : reste sur le Concernant, avec reference de la seule seance', () => {
  const payload = buildRegardingNotePayload(target, 'accounts', { subject: 'Note', text: 'Observation' }, id);
  assert.equal(payload.notetext, withAppointmentReference('Observation', id));
  assert.equal(payload['objectid_account@odata.bind'], `/accounts(${target.id})`);
});

test('note : contexte revalide apres metadonnees et immediatement avant POST', async () => {
  let posts = 0;
  await withFetch(async (_url, options) => {
    if (options?.method === 'POST') { posts++; return Response.json({}); }
    return Response.json({ EntitySetName: 'accounts' });
  }, async () => {
    await assert.rejects(createRegardingNote(config, 'fake', target, { subject: 'Note', text: 'Texte' }, id,
      () => { throw new Error('appointment changed'); }), /appointment changed/);
    assert.equal(posts, 0);
  });
});
