import test from 'node:test';
import assert from 'node:assert/strict';
import { assertAppointmentContext, getAppointmentContext, watchAppointmentChanges, markAppointmentAsLogged } from '../src/outlook.js';
import { AppointmentSession } from '../src/appointment-session.js';
import { mobileAppointmentLink, androidAppointmentIntent, mobileLaunchLink } from '../src/mobile-link.js';

const item = (id) => ({ itemId: id, subject: 'Réunion', start: new Date('2026-10-28T09:00Z'), end: new Date('2026-10-28T10:00Z') });
const officeFixture = (id = 'one') => ({
  onReady: async () => {}, HostType: { Outlook: 'Outlook' }, AsyncResultStatus: { Succeeded: 'succeeded' },
  MailboxEnums: { RestVersion: { v2_0: 'v2.0' } },
  context: { host: 'Outlook', platform: 'Android', mailbox: { item: item(id), userProfile: { emailAddress: 'user@roth-france.fr' } } },
});
async function withOffice(office, task) {
  const previous = globalThis.Office;
  globalThis.Office = office;
  try { await task(); } finally { globalThis.Office = previous; }
}

test('Office : identite inchangée et copie equivalent REST acceptees, jamais par sujet', () => {
  const office = officeFixture();
  assertAppointmentContext({ itemId: 'one', mailboxEmail: 'USER@roth-france.fr' }, office);
  office.context.platform = 'PC';
  office.context.mailbox.convertToRestId = () => 'stable-rest';
  assertAppointmentContext({ itemId: 'ews-old-format', restItemId: 'stable-rest' }, office);
  assert.throws(() => assertAppointmentContext({ itemId: 'other' }, office), { code: 'OUTLOOK_CONTEXT_CHANGED' });
  assert.throws(() => assertAppointmentContext({ itemId: 'one', mailboxEmail: 'other@roth-france.fr' }, office), { code: 'OUTLOOK_CONTEXT_CHANGED' });
});

test('Office : identifiant vide/transitoire refuse avec actualisation possible', () => {
  const office = officeFixture('');
  assert.throws(() => assertAppointmentContext({ itemId: 'one' }, office), /Actualiser/);
  assert.throws(() => assertAppointmentContext({ itemId: '' }, office), { code: 'OUTLOOK_CONTEXT_CHANGED' });
});

test('Office : les champs lus pendant un changement ne creent pas un contexte hybride', async () => {
  const office = officeFixture();
  let reads = 0;
  office.context.mailbox.item.subject = { getAsync: (done) => {
    reads++;
    office.context.mailbox.item = item('two');
    done({ status: 'succeeded', value: 'Ancien sujet' });
  } };
  await withOffice(office, async () => {
    const result = await getAppointmentContext();
    assert.equal(result.itemId, 'two'); assert.equal(result.subject, 'Réunion'); assert.equal(reads, 1);
  });
});

test('Office : ItemChanged optionnel, pas de panne sur un host sans API', () => {
  watchAppointmentChanges(() => {}, officeFixture());
  const office = officeFixture(); let handler;
  office.context.requirements = { isSetSupported: () => true };
  office.EventType = { ItemChanged: 'itemChanged' };
  office.context.mailbox.addHandlerAsync = (event, callback, done) => { assert.equal(event, 'itemChanged'); handler = callback; done({ status: 'failed' }); };
  watchAppointmentChanges(() => {}, office); assert.equal(typeof handler, 'function');
  office.context.mailbox.addHandlerAsync = () => { throw new Error('unsupported host'); };
  assert.doesNotThrow(() => watchAppointmentChanges(() => {}, office));
});

test('session : retour A -> B -> A invalide aussi une action en vol', async () => {
  await withOffice(officeFixture(), async () => {
    const session = new AppointmentSession(); const context = { itemId: 'one' };
    session.use(context); const scope = session.capture(); scope.ensure();
    session.invalidate(); session.use(context);
    assert.throws(scope.ensure, { code: 'OUTLOOK_CONTEXT_CHANGED' });
    session.capture().ensure();
  });
});

test('session : brouillons en memoire separes par compte et seance, pas de stockage persistant', () => {
  const session = new AppointmentSession();
  const a = { itemId: 'one', mailboxEmail: 'user@roth-france.fr' };
  session.use(a); session.rememberDraft('Titre', 'Brouillon');
  session.use({ ...a, itemId: 'two' }); assert.equal(session.draft(), undefined);
  session.use({ ...a, mailboxEmail: 'other@roth-france.fr' }); assert.equal(session.draft(), undefined);
  session.use(a); assert.deepEqual(session.draft(), { subject: 'Titre', text: 'Brouillon' });
  session.rememberDraft('', ''); assert.equal(session.draft(), undefined);
});

test('Office : ne jamais marquer le nouvel item apres une creation de note sur l’ancien', async () => {
  const office = officeFixture(); let saves = 0;
  office.context.mailbox.item.loadCustomPropertiesAsync = (done) => {
    office.context.mailbox.item = item('two');
    done({ status: 'succeeded', value: { set: () => {}, saveAsync: () => { saves++; } } });
  };
  await withOffice(office, async () => {
    await assert.rejects(markAppointmentAsLogged({ itemId: 'one' }), { code: 'OUTLOOK_CONTEXT_CHANGED' });
    assert.equal(saves, 0);
  });
});

test('lien Android : intent au clic, paquet fixe, route HTTPS de l’APK, aucun jeton ou redirection', () => {
  const id = 'AAMk/+%&x#;=TEST';
  const intent = androidAppointmentIntent(id);
  assert.equal(intent, mobileLaunchLink(id, { isAndroid: true }));
  assert.ok(intent.endsWith('#Intent;scheme=https;package=fr.roth.compterendu;end'));
  const https = intent.split('#Intent;')[0].replace(/^intent:/, 'https:');
  assert.equal(https, mobileAppointmentLink(id));
  const url = new URL(https); assert.equal(url.pathname, '/open'); assert.equal(url.searchParams.get('eventId'), id);
  assert.equal(mobileLaunchLink(id), https);
  assert.throws(() => androidAppointmentIntent('evil\n'));
});
