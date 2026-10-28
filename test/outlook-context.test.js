import test from 'node:test';
import assert from 'node:assert/strict';
import { assertAppointmentContext, assertAppointmentContextAsync, getAppointmentContext, watchAppointmentChanges, markAppointmentAsLogged } from '../src/outlook.js';
import { resolveOutlookEvent } from '../src/graph.js';
import { AppointmentSession } from '../src/appointment-session.js';
import { mobileAppointmentLink } from '../src/mobile-link.js';

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

test('lien mobile : HTTPS seulement, ID encode, route de l’APK, aucun jeton ni redirection', () => {
  const id = 'AAMk/+%&x#;=TEST';
  const https = mobileAppointmentLink(id);
  const url = new URL(https); assert.equal(url.pathname, '/open'); assert.equal(url.searchParams.get('eventId'), id);
  assert.equal(url.protocol, 'https:'); assert.equal(url.origin, 'https://cr-api.roth-france.fr');
  assert.equal(url.searchParams.get('view'), 'crm'); assert.equal(url.hash, '');
  assert.deepEqual([...url.searchParams.keys()], ['eventId', 'view']);
  assert.throws(() => mobileAppointmentLink('evil\n'));
});

function composeFixture() {
  const office = officeFixture(); office.context.platform = 'PC';
  let nativeId = 'ews-occurrence'; let reads = 0; let saves = 0;
  const asyncValue = (value) => ({ getAsync: (done) => done({ status: 'succeeded', value }) });
  office.context.mailbox.convertToRestId = (id) => 'rest-' + id;
  office.context.mailbox.item = {
    getItemIdAsync: (done) => { reads++; done({ status: 'succeeded', value: nativeId }); },
    subject: asyncValue('IT Meeting'), start: asyncValue(new Date('2026-10-27T12:30Z')), end: asyncValue(new Date('2026-10-27T13:30Z')),
    location: asyncValue('Teams'), saveAsync: () => { saves++; throw new Error('must never save'); },
  };
  return { office, id(value) { nativeId = value; }, reads: () => reads, saves: () => saves };
}

test('regression capture 2.3.4 : occurrence en modification sans itemId utilise getItemIdAsync et la conversion REST', async () => {
  const f = composeFixture();
  await withOffice(f.office, async () => {
    const context = await getAppointmentContext();
    assert.equal(context.itemId, 'ews-occurrence'); assert.equal(context.restItemId, 'rest-ews-occurrence');
    assert.equal(context.identitySource, 'getItemIdAsync'); assert.equal(context.subject, 'IT Meeting');
    assert.equal(context.start.toISOString(), '2026-10-27T12:30:00.000Z');
    assert.equal(f.reads(), 2, 'lecture avant et apres les champs'); assert.equal(f.saves(), 0);
    const session = new AppointmentSession(); session.use(context);
    await session.capture().ensureCurrent(); assert.equal(f.saves(), 0);
  });
});

test('compose : reutilisation du meme objet Office avec un autre ID bloque la session async', async () => {
  const f = composeFixture();
  await withOffice(f.office, async () => {
    const context = await getAppointmentContext();
    const session = new AppointmentSession(); session.use(context); const scope = session.capture();
    f.id('different-occurrence');
    assert.doesNotThrow(scope.ensure, 'le check synchrone seul ne peut pas relire Office');
    await assert.rejects(scope.ensureCurrent, { code: 'OUTLOOK_CONTEXT_CHANGED' });
    assert.equal(f.saves(), 0);
  });
});

test('compose : nouveau wrapper pendant lecture async refuse, meme si le callback ancien retourne le bon ID', async () => {
  const f = composeFixture();
  await withOffice(f.office, async () => {
    const context = await getAppointmentContext();
    f.office.context.mailbox.item.getItemIdAsync = (done) => {
      f.office.context.mailbox.item = item('another'); done({ status: 'succeeded', value: context.itemId });
    };
    await assert.rejects(assertAppointmentContextAsync(context), { code: 'OUTLOOK_CONTEXT_CHANGED' });
  });
});

test('compose non enregistre : message utile, pas de faux changement de contexte ni saveAsync automatique', async () => {
  const f = composeFixture();
  f.office.context.mailbox.item.getItemIdAsync = (done) => done({ status: 'failed', error: { code: 'ItemNotSaved' } });
  await withOffice(f.office, async () => {
    await assert.rejects(getAppointmentContext(), (error) => error.code === 'OUTLOOK_ITEM_ID_UNAVAILABLE' && /Enregistrez/.test(error.message));
    assert.equal(f.saves(), 0);
  });
});

test('ancien client compose sans getItemIdAsync : bloque sans sauvegarder ni modifier les permissions', async () => {
  const f = composeFixture(); delete f.office.context.mailbox.item.getItemIdAsync;
  await withOffice(f.office, async () => {
    await assert.rejects(getAppointmentContext(), { code: 'OUTLOOK_ITEM_ID_UNAVAILABLE' }); assert.equal(f.saves(), 0);
  });
});

test('lecture invite/mobile : itemId conserve, getItemIdAsync inutile jamais appele', async () => {
  const office = officeFixture('rest-read');
  office.context.mailbox.item.getItemIdAsync = () => { throw new Error('read mode must not call compose API'); };
  await withOffice(office, async () => {
    const context = await getAppointmentContext();
    assert.equal(context.itemId, 'rest-read'); assert.equal(context.restItemId, 'rest-read');
    await assertAppointmentContextAsync(context);
  });
});

test('compose : Graph recoit uniquement l’identifiant REST de l’occurrence enregistree, pas le maitre ou une recherche par titre', async () => {
  const f = composeFixture(); const oldFetch = globalThis.fetch; const calls = [];
  await withOffice(f.office, async () => {
    const context = await getAppointmentContext();
    globalThis.fetch = async (url, options) => {
      calls.push({url,options});
      return Response.json({ id: context.restItemId, subject: context.subject, type: 'occurrence', seriesMasterId: 'rest-master',
        start: {dateTime:context.start.toISOString()}, end: {dateTime:context.end.toISOString()} });
    };
    try {
      assert.equal((await resolveOutlookEvent('fake', context)).id, 'rest-ews-occurrence');
      assert.equal(calls.length, 1); assert.match(calls[0].url, /\/me\/events\/rest-ews-occurrence\?/);
      assert.equal(calls[0].options.method, undefined); assert.equal(f.saves(), 0);
    } finally { globalThis.fetch = oldFetch; }
  });
});
