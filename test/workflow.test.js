import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { AppointmentSession } from '../src/appointment-session.js';
import { assertAppointmentContext, getAppointmentContext } from '../src/outlook.js';
import { mobileAppointmentLink } from '../src/mobile-link.js';

// Run the actual UI orchestration without Office, network, a CRM or a browser.
// This small DOM adapter deliberately tracks detached controls after re-render.
class Element {
  constructor(document, id = '') { this.document = document; this.id = id; this.value = ''; this.className = ''; this.isConnected = true; this.open = false; this.textContent = ''; this.listeners = {}; }
  classList = {
    add: (...names) => { this.className = [...new Set([...this.className.split(' '), ...names])].join(' '); },
    remove: (...names) => { this.className = this.className.split(' ').filter((name) => !names.includes(name)).join(' '); },
    contains: (name) => this.className.split(' ').includes(name),
    toggle: (name, force) => { const add = force ?? !this.classList.contains(name); this.classList[add ? 'add' : 'remove'](name); },
  };
  set innerHTML(html) {
    this.html = html;
    if (this.id === 'app') {
      for (const element of this.document.elements.values()) if (element !== this) element.isConnected = false;
      this.document.elements = new Map([['app', this]]);
      for (const [, id, attributes] of html.matchAll(/<[^>]+\bid="([^"]+)"([^>]*)>/g)) {
        const element = new Element(this.document, id);
        element.value = attributes.match(/\bvalue="([^"]*)"/)?.[1] || '';
        element.className = attributes.match(/\bclass="([^"]*)"/)?.[1] || '';
        this.document.elements.set(id, element);
      }
    }
    if (this.id.endsWith('Select')) this.value = html.match(/value="([^"]*)"/)?.[1] || '';
  }
  get innerHTML() { return this.html || ''; }
  addEventListener(event, handler) { this.listeners[event] = handler; }
  setAttribute() {}
  append() {}
  focus() {}
}
function harness() {
  const document = { elements: new Map(), querySelector: (id) => document.elements.get(id.slice(1)) || null, addEventListener() {}, createElement: () => new Element(document) };
  document.elements.set('app', new Element(document, 'app'));
  const meeting = (id) => ({ itemId: id, restItemId: id, mailboxEmail: 'user@roth-france.fr', subject: 'RDV ' + id, start: new Date(), end: new Date(), isMobile: false });
  let current = meeting('one');
  const counts = { context: 0, notes: 0, patch: 0, tracking: 0, copies: 0, copiedLink: '' };
  const regarding = { id: 'account', label: 'Compte test', logicalName: 'account' };
  const appointment = { activityid: 'crm-one', editTarget: { kind: 'appointment' }, qualification: 1 };
  let search = async () => [appointment];
  let regardingSearch = async () => [];
  let graphCategories = [];
  let realOfficeContext = false;
  const office = { context: { mailbox: { item: { itemId: current.itemId }, userProfile: { emailAddress: current.mailboxEmail } } } };
  const sandbox = {
    console: { error() {}, warn() {} }, document, Office: office, Intl, Date, URLSearchParams,
    packageMetadata: { version: '2.3.8' }, AppointmentSession, assertAppointmentContext,
    mobileAppointmentLink,
    navigator: { clipboard: { writeText: async (text) => { counts.copies++; counts.copiedLink = text; } } },
    APPOINTMENT_QUALIFICATION_FIELD: 'qualification', TRACKED_TO_DYNAMICS_CATEGORY: 'Tracked To Dynamics 365',
    getAppointmentContext: async () => { counts.context++; return realOfficeContext ? getAppointmentContext() : current; },
    getDataverseAccessToken: async () => 'fake', getGraphAccessToken: async () => 'fake',
    resolveOutlookEvent: async (_token, context) => ({ id: context.itemId, type: 'singleInstance', categories: graphCategories }),
    readTrackingState: async (_token, event) => ({ tracked: event.categories.length > 0 }),
    isRecurringEvent: (event) => Boolean(event && event.type !== 'singleInstance'),
    findAppointmentCandidates: async (...args) => search(...args), chooseUnambiguousCandidate: (rows) => rows[0],
    resolveAppointmentEditingTarget: async (_config, _token, row) => row,
    getRegardingReference: (row) => row ? regarding : null, getRegardingLabel: () => regarding.label,
    getAppointmentQualificationOptions: async () => [{ value: 1, label: 'Client' }, { value: 2, label: 'Autre' }],
    getAppointmentNotes: async () => { counts.notes++; return []; },
    searchRegardingRecords: async () => regardingSearch(),
    updateAppointmentQualification: async (_config, _token, _id, _value, row, ensure) => { await ensure(); counts.patch++; return row; },
    trackOutlookAppointment: async () => { counts.tracking++; throw new Error('no write expected'); },
  };
  const source = fs.readFileSync(new URL('../src/main.js', import.meta.url), 'utf8')
    .replace(/^import[\s\S]*?;\r?\n/gm, '').replace(/bootstrap\(\);\s*$/, '');
  const runtime = vm.createContext(sandbox);
  vm.runInContext(source + `
    globalThis.ui = {
      setup(value) { config = {dataverseUrl: 'https://roth.crm4.dynamics.com'}; context = value; session.use(value); accessToken='fake'; renderShell(); updateActionControls(); },
      refresh: refreshCurrentMeeting, notes: refreshNotes, search: searchRegarding, track: trackAndWaitForAppointment,
      async saveQualification() { await saveCrmFields(); },
      state() { return {context, selectedAppointment, invalid:session.invalid, trackingState}; },
      invalidate() {session.invalidate();},
      confirm(callback) {confirmAction=callback;},
      boot: () => loadCurrentMeeting(currentScope()),
    };`, runtime);
  const api = sandbox.ui;
  return { api, counts, document, office, meeting, appointment,
    mobile() { current.isMobile = true; office.context.platform = 'Android'; },
    clipboardDenied() { sandbox.navigator.clipboard.writeText = async () => { throw new Error('clipboard denied'); }; },
    compose() {
      realOfficeContext = true;
      office.onReady = async () => {}; office.HostType = { Outlook: 'Outlook' }; office.AsyncResultStatus = { Succeeded: 'succeeded' };
      office.MailboxEnums = { RestVersion: { v2_0: 'v2.0' } }; office.context.platform = 'PC';
      office.context.mailbox.convertToRestId = (id) => 'rest-' + id;
      office.context.mailbox.item = { getItemIdAsync: (done) => done({ status:'succeeded',value:current.itemId }),
        subject:current.subject,start:current.start,end:current.end };
    },
    changeComposeId(id) { current = meeting(id); },
    select(id) { current = meeting(id); office.context.mailbox.item.itemId = id; },
    search(task) { search = task; }, searchRegarding(task) { regardingSearch = task; }, tracked() { graphCategories = ['Tracked To Dynamics 365']; },
    async run(task) { const previous = globalThis.Office; globalThis.Office = office; try { api.setup(current); await task(); } finally { globalThis.Office = previous; } },
  };
}

test('UI : Actualiser relit Office, recupere un contexte invalide sans fermeture et garde les brouillons par RDV', async () => {
  const h = harness(); await h.run(async () => {
    h.document.querySelector('#noteText').value = 'Brouillon A';
    h.select('two'); h.api.invalidate(); await h.api.refresh();
    assert.equal(h.counts.context, 1); assert.equal(h.api.state().context.itemId, 'two'); assert.equal(h.api.state().invalid, false);
    assert.equal(h.document.querySelector('#noteText').value, '');
    h.select('one'); await h.api.refresh();
    assert.equal(h.document.querySelector('#noteText').value, 'Brouillon A');
    assert.equal(h.document.querySelector('#noteEditor').open, true);
  });
});

test('UI : RDV synchronise sans activation inutile, historique seulement a la demande', async () => {
  const h = harness(); await h.run(async () => {
    await h.api.boot();
    assert.equal(h.document.querySelector('#trackAppointmentButton').classList.contains('hidden'), true);
    assert.equal(h.document.querySelector('#saveCrmFieldsButton').disabled, true, 'valeur inchangee');
    assert.equal(h.counts.notes, 0); await h.api.notes(); assert.equal(h.counts.notes, 1);
  });
});

test('UI : categorie deja active et CRM absent => Verifier, aucun PATCH Outlook', async () => {
  const h = harness(); h.tracked(); h.search(async () => []);
  await h.run(async () => {
    await h.api.boot(); assert.match(h.document.querySelector('#trackAppointmentButton').textContent, /Vérifier/);
    await h.api.track(); assert.equal(h.counts.tracking, 0); assert.equal(h.counts.context, 1);
  });
});

test('UI : reponse CRM tardive apres changement => aucune selection de l’ancien RDV, une seule alerte', async () => {
  const h = harness(); h.search(async () => { h.select('two'); return [h.appointment]; });
  await h.run(async () => {
    await h.api.refresh();
    assert.equal(h.api.state().selectedAppointment, null);
    assert.equal(h.api.state().invalid, true);
    assert.equal(h.document.querySelector('#contextStatus').classList.contains('hidden'), false);
    assert.ok(!h.document.querySelector('#matchStatus').classList.contains('success'));
  });
});

test('UI : changement pendant confirmation de qualification => aucun PATCH, controles bloques', async () => {
  const h = harness(); h.appointment.editTarget.kind = 'series';
  await h.run(async () => {
    await h.api.boot(); h.document.querySelector('#qualificationSelect').value = '2';
    h.api.confirm(async () => { h.select('two'); return true; });
    await h.api.saveQualification();
    assert.equal(h.counts.patch, 0); assert.equal(h.api.state().invalid, true);
    assert.equal(h.document.querySelector('#saveCrmFieldsButton').disabled, true);
  });
});

test('UI : ancienne recherche Concernant ignoree apres rechargement, sans invalider le nouveau contexte', async () => {
  const h = harness(); let complete;
  h.searchRegarding(() => new Promise((resolve) => { complete = resolve; }));
  await h.run(async () => {
    await h.api.boot(); h.document.querySelector('#regardingSearchText').value = 'Roth';
    const pending = h.api.search();
    h.select('two'); await h.api.refresh();
    complete([{id:'old',label:'Ancien compte'}]); await pending;
    assert.equal(h.api.state().context.itemId, 'two'); assert.equal(h.api.state().invalid, false);
  });
});

test('UI regression compose : Actualiser atteint Dynamics sur un item sans propriete itemId', async () => {
  const h = harness(); await h.run(async () => {
    h.compose(); await h.api.refresh();
    assert.equal(h.api.state().invalid, false); assert.equal(h.api.state().context.itemId, 'one');
    assert.equal(h.api.state().selectedAppointment.activityid, 'crm-one');
    assert.match(h.document.querySelector('#matchStatus').textContent, /synchronisé/);
  });
});

test('UI compose : changement ID dans le meme wrapper pendant confirmation interdit tout PATCH', async () => {
  const h = harness(); h.appointment.editTarget.kind = 'series';
  await h.run(async () => {
    h.compose(); await h.api.refresh();
    h.document.querySelector('#qualificationSelect').value = '2';
    h.api.confirm(async () => { h.changeComposeId('two'); return true; });
    await h.api.saveQualification();
    assert.equal(h.counts.patch, 0); assert.equal(h.api.state().invalid, true);
  });
});

test('UI serie : aucun rappel en double dans Dynamics et aucun volet technique a deplier', async () => {
  const h = harness(); h.appointment.editTarget.kind = 'series';
  await h.run(async () => {
    await h.api.boot();
    const status = h.document.querySelector('#seriesScopeStatus');
    assert.equal(status.textContent, '');
    assert.equal(status.classList.contains('hidden'), true);
    assert.equal(h.document.querySelector('#seriesScopeDetails'), null);
  });
});

test('UI RDV simple : pas de rappel ; motif de blocage toujours visible si serie non verifiee', async () => {
  const h = harness();
  await h.run(async () => {
    await h.api.boot();
    assert.equal(h.document.querySelector('#seriesScopeStatus').classList.contains('hidden'), true);
    h.appointment.editTarget = {kind:'blocked', reason:'La série est inaccessible avec vos droits.'};
    await h.api.refresh();
    const status = h.document.querySelector('#seriesScopeStatus');
    assert.equal(status.textContent, h.appointment.editTarget.reason);
    assert.equal(status.classList.contains('warning'), true);
    assert.equal(status.classList.contains('hidden'), false);
    assert.equal(h.document.querySelector('#saveCrmFieldsButton').disabled, true);
  });
});

test('UI mobile : le clic navigue directement vers le meme HTTPS que Copier, sans intent ni volet aide', async () => {
  const h = harness(); h.mobile();
  await h.run(async () => {
    await h.api.boot();
    const link = h.document.querySelector('#openMobileApp');
    const copy = h.document.querySelector('#mobileCopyLink');
    assert.equal(link.href, mobileAppointmentLink('one'));
    assert.equal(link.href, copy.href);
    assert.equal(h.document.querySelector('#copyMobileLink').disabled, false);
    assert.ok(!h.document.querySelector('#app').innerHTML.includes('mobile-help'));
    assert.ok(!h.document.querySelector('#app').innerHTML.includes('<input id="mobileCopyLink"'));
    let prevented = false;
    link.listeners.click({ preventDefault() { prevented = true; } });
    assert.equal(prevented, false, 'la navigation HTML au clic est conservee');
    await h.document.querySelector('#copyMobileLink').listeners.click();
    assert.equal(h.counts.copies, 1); assert.equal(h.counts.copiedLink, link.href);
  });
});

test('UI mobile : clipboard refuse => vrai lien HTTPS cliquable et copiable par appui long', async () => {
  const h = harness(); h.mobile(); h.clipboardDenied();
  await h.run(async () => {
    await h.api.boot();
    await h.document.querySelector('#copyMobileLink').listeners.click();
    const link = h.document.querySelector('#mobileCopyLink');
    assert.equal(link.classList.contains('hidden'), false);
    assert.equal(link.href, mobileAppointmentLink('one'));
    assert.equal(link.textContent, link.href);
    assert.equal(typeof link.listeners.click, 'function');
    assert.match(h.document.querySelector('#mobileLinkStatus').textContent, /appui long/);
  });
});

test('UI mobile : clic refuse pour un autre RDV ou une note non enregistree, pas de copie silencieuse', async () => {
  const h = harness(); h.mobile();
  await h.run(async () => {
    await h.api.boot(); let prevented = false;
    const event = { preventDefault() { prevented = true; } };
    h.document.querySelector('#noteText').value = 'Note non enregistree';
    h.document.querySelector('#openMobileApp').listeners.click(event);
    assert.equal(prevented, true);
    h.select('two'); prevented = false;
    h.document.querySelector('#mobileCopyLink').listeners.click(event);
    assert.equal(prevented, true);
    assert.equal(h.counts.copies, 0);
    assert.equal(h.api.state().invalid, true);
  });
});
