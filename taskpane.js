/*
 * Request Tracker - Outlook add-in (v0)
 * Tracks client requests (tasks) and the supplier quote requests sent for them (subtasks),
 * with follow-up dates, reminder drafts and reply detection on the opened email.
 * Storage: Outlook roaming settings (follows the mailbox, ~32 KB limit).
 * Outside Outlook the page runs in demo mode with localStorage and a mock email.
 */
'use strict';

const STORE_KEY = 'rt.v1';
const ROAMING_LIMIT = 32 * 1024;
const LOCAL_ITEM_MAP = 'rt.itemIds';
const LOCAL_ARCHIVE = 'rt.archive';

const SUP_STATUS = {
  sent: { label: 'Waiting', tone: 'pending' },
  replied: { label: 'Replied', tone: 'info' },
  quoted: { label: 'Quoted', tone: 'ok' },
  declined: { label: 'Declined', tone: 'muted' },
};
const REQ_STATUS = {
  open: { label: 'Open', tone: 'info' },
  done: { label: 'Done', tone: 'ok' },
  cancelled: { label: 'Cancelled', tone: 'muted' },
};

const DEFAULT_SETTINGS = {
  followUpDays: 2,
  workdaysOnly: true,
  subjectPrefix: 'Reminder: ',
  template:
    'Buna ziua,\n\n' +
    'Revin cu solicitarea de oferta transmisa in data de {sentDate} privind "{subject}".\n' +
    'Va rog sa ne transmiteti oferta sau un termen estimat de raspuns.\n\n' +
    'Multumesc,',
};

const state = {
  inOutlook: false,
  data: { requests: [], settings: { ...DEFAULT_SETTINGS } },
  view: 'email',
  reqId: null,
  info: null,
  listFilter: 'open',
  listQuery: '',
};

/* ---------------- Utilities ---------------- */

const $ = (sel, root = document) => root.querySelector(sel);

function h(tag, attrs, ...children) {
  const el = document.createElement(tag);
  if (attrs) {
    for (const [k, v] of Object.entries(attrs)) {
      if (v === null || v === undefined || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
      else if (k === 'html') el.innerHTML = v;
      else el.setAttribute(k, v === true ? '' : v);
    }
  }
  for (const c of children.flat()) {
    if (c === null || c === undefined || c === false) continue;
    el.append(c.nodeType ? c : document.createTextNode(String(c)));
  }
  return el;
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function uid() {
  return Math.random().toString(36).slice(2, 8) + Date.now().toString(36).slice(-3);
}

// Short stable hash (cyrb53) so long Exchange conversation IDs do not eat the roaming quota.
function hashKey(str) {
  if (!str) return '';
  let h1 = 0xdeadbeef, h2 = 0x41c6ce57;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}

const pad = (n) => String(n).padStart(2, '0');
function toDay(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function today() { return toDay(new Date()); }
function parseDay(s) { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); }
function fmtDay(s) { if (!s) return '-'; const d = parseDay(s); return `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()}`; }
function parseUserDate(s) {
  const m = String(s).trim().match(/^(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?$/);
  if (!m) return null;
  let y = m[3] ? Number(m[3]) : new Date().getFullYear();
  if (y < 100) y += 2000;
  const d = new Date(y, Number(m[2]) - 1, Number(m[1]));
  if (d.getDate() !== Number(m[1]) || d.getMonth() !== Number(m[2]) - 1) return null;
  return toDay(d);
}
function addDays(day, n, workdaysOnly) {
  const d = parseDay(day);
  if (!workdaysOnly) { d.setDate(d.getDate() + n); return toDay(d); }
  let left = n;
  while (left > 0) {
    d.setDate(d.getDate() + 1);
    if (d.getDay() !== 0 && d.getDay() !== 6) left--;
  }
  return toDay(d);
}
function daysBetween(a, b) { return Math.round((parseDay(b) - parseDay(a)) / 86400000); }
function relDay(day) {
  const n = daysBetween(today(), day);
  if (n === 0) return 'today';
  if (n === 1) return 'tomorrow';
  if (n === -1) return 'yesterday';
  return n > 0 ? `in ${n} days` : `${-n} days ago`;
}
function cleanSubject(s) { return String(s || '').replace(/^\s*((re|fw|fwd|tr|raspuns|redirectionare)\s*:\s*)+/i, '').trim(); }
function nameFromEmail(email) {
  const domain = (email.split('@')[1] || email).split('.');
  return domain.length > 1 ? domain[domain.length - 2] : domain[0];
}

function toast(msg, isErr) {
  const t = $('#toast');
  t.textContent = msg;
  t.className = 'toast' + (isErr ? ' err' : '');
  t.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { t.hidden = true; }, isErr ? 5000 : 2200);
}

/* ---------------- Storage ---------------- */

function localGet(key, fallback) {
  try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch { return fallback; }
}
function localSet(key, value) {
  try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* ignore */ }
}

function loadData() {
  let raw = null;
  if (state.inOutlook) raw = Office.context.roamingSettings.get(STORE_KEY);
  else raw = localGet(STORE_KEY, null);
  if (typeof raw === 'string') { try { raw = JSON.parse(raw); } catch { raw = null; } }
  if (raw && Array.isArray(raw.requests)) {
    state.data = { requests: raw.requests, settings: { ...DEFAULT_SETTINGS, ...(raw.settings || {}) } };
  }
}

function dataSize() { return new Blob([JSON.stringify(state.data)]).size; }

function save() {
  const size = dataSize();
  if (size > ROAMING_LIMIT * 0.97) {
    toast('Storage is full. Archive finished requests in Settings.', true);
    return Promise.resolve(false);
  }
  if (!state.inOutlook) { localSet(STORE_KEY, state.data); updateBadge(); return Promise.resolve(true); }
  Office.context.roamingSettings.set(STORE_KEY, state.data);
  return new Promise((resolve) => {
    Office.context.roamingSettings.saveAsync((res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        toast('Save failed: ' + (res.error && res.error.message), true);
        resolve(false);
      } else { updateBadge(); resolve(true); }
    });
  });
}

// Device-local map convKey -> itemId, used only to reopen emails on this device.
function rememberItem(convKey, itemId) {
  if (!convKey || !itemId) return;
  const map = localGet(LOCAL_ITEM_MAP, {});
  map[convKey] = itemId;
  localSet(LOCAL_ITEM_MAP, map);
}
function itemIdFor(convKey) { return convKey ? localGet(LOCAL_ITEM_MAP, {})[convKey] : null; }

/* ---------------- Domain logic ---------------- */

const reqById = (id) => state.data.requests.find((r) => r.id === id);
const isSupOverdue = (r, s) => r.status === 'open' && s.status === 'sent' && s.followUpAt && s.followUpAt <= today();
const isReqOverdue = (r) => r.status === 'open' && r.dueAt && r.dueAt < today();
function reqOverdueCount(r) { return r.suppliers.filter((s) => isSupOverdue(r, s)).length; }
function totalOverdue() {
  return state.data.requests.reduce((n, r) => n + reqOverdueCount(r) + (isReqOverdue(r) ? 1 : 0), 0);
}
function answeredCount(r) { return r.suppliers.filter((s) => s.status !== 'sent').length; }

function newSupplier({ name, email, convKey, subject, sentAt }) {
  const st = state.data.settings;
  const sent = sentAt || today();
  return {
    id: uid(),
    name: name || nameFromEmail(email),
    email: (email || '').toLowerCase(),
    status: 'sent',
    sentAt: sent,
    followUpAt: addDays(sent, st.followUpDays, st.workdaysOnly),
    reminders: 0,
    convKey: convKey || '',
    subject: subject || '',
  };
}

function analyzeItem(info) {
  const out = { clientOf: [], supplierHits: [] };
  if (!info) return out;
  for (const r of state.data.requests) {
    if (info.convKey && r.convKey === info.convKey) out.clientOf.push(r);
    for (const s of r.suppliers) {
      const thread = info.convKey && s.convKey === info.convKey;
      const sender = !info.fromMe && info.from && s.email && s.email === info.from.email;
      if (thread || (sender && r.status === 'open')) out.supplierHits.push({ r, s, by: thread ? 'thread' : 'sender', sender });
    }
  }
  // One email often goes to several suppliers: on an incoming reply keep only the supplier who actually wrote it.
  if (!info.fromMe && out.supplierHits.some((x) => x.sender)) out.supplierHits = out.supplierHits.filter((x) => x.sender);
  return out;
}

/* ---------------- Outlook item ---------------- */

function readItem() {
  if (!state.inOutlook) return mockItem();
  const it = Office.context.mailbox.item;
  if (!it || it.itemType !== Office.MailboxEnums.ItemType.Message) return null;
  const me = (Office.context.mailbox.userProfile.emailAddress || '').toLowerCase();
  const mapAddr = (a) => ({ name: a.displayName || '', email: (a.emailAddress || '').toLowerCase() });
  const from = it.from ? mapAddr(it.from) : null;
  const info = {
    subject: it.subject || '',
    convKey: hashKey(it.conversationId || ''),
    itemId: it.itemId,
    from,
    to: (it.to || []).map(mapAddr),
    cc: (it.cc || []).map(mapAddr),
    date: it.dateTimeCreated ? toDay(new Date(it.dateTimeCreated)) : today(),
    fromMe: !!(from && from.email === me),
  };
  rememberItem(info.convKey, info.itemId);
  return info;
}

function mockItem() {
  const sample = localGet('rt.mock', 0);
  const mocks = [
    { subject: 'Cerere oferta - 2x server rack 2U + licente', from: { name: 'Ion Popescu (Client SRL)', email: 'ion.popescu@client.ro' }, to: [{ name: 'Andrei', email: 'andrei@demo.ro' }], fromMe: false, convKey: 'c1' },
    { subject: 'RE: Cerere oferta - 2x server rack 2U + licente', from: { name: 'Andrei', email: 'andrei@demo.ro' }, to: [{ name: 'Maria - Distribuitor A', email: 'maria@distrib-a.ro' }, { name: 'Vlad - Distribuitor B', email: 'vlad@distrib-b.ro' }], fromMe: true, convKey: 's1' },
    { subject: 'RE: Cerere oferta - 2x server rack 2U + licente', from: { name: 'Maria - Distribuitor A', email: 'maria@distrib-a.ro' }, to: [{ name: 'Andrei', email: 'andrei@demo.ro' }], fromMe: false, convKey: 's1' },
  ];
  const m = mocks[sample % mocks.length];
  return { ...m, cc: [], itemId: 'mock-' + sample, date: today() };
}

function openEmail(convKey) {
  const id = itemIdFor(convKey);
  if (!id) { toast('Email not opened on this device yet. Search for it in Outlook.', true); return; }
  if (state.inOutlook) Office.context.mailbox.displayMessageForm(id);
  else toast('Demo: would open ' + id);
}

function draftReminder(r, s) {
  const st = state.data.settings;
  const subject = s.subject ? cleanSubject(s.subject) : r.title;
  const text = st.template
    .replaceAll('{supplier}', s.name)
    .replaceAll('{subject}', subject)
    .replaceAll('{sentDate}', fmtDay(s.sentAt))
    .replaceAll('{client}', r.client || '');
  const html = text.split('\n').map((l) => esc(l) || '&nbsp;').join('<br>');
  const onThread = state.info && s.convKey && state.info.convKey === s.convKey;
  if (state.inOutlook) {
    const it = Office.context.mailbox.item;
    if (onThread && it && typeof it.displayReplyForm === 'function') it.displayReplyForm({ htmlBody: html });
    else Office.context.mailbox.displayNewMessageForm({ toRecipients: [s.email], subject: st.subjectPrefix + subject, htmlBody: html });
  } else {
    toast('Demo: reminder drafted to ' + s.email);
  }
  s.reminders = (s.reminders || 0) + 1;
  s.lastReminderAt = today();
  s.followUpAt = addDays(today(), st.followUpDays, st.workdaysOnly);
  save().then(render);
}

function addToCalendar(title, day, body) {
  const start = parseDay(day < today() ? today() : day);
  start.setHours(9, 0, 0, 0);
  const end = new Date(start.getTime() + 15 * 60000);
  if (state.inOutlook) Office.context.mailbox.displayNewAppointmentForm({ start, end, subject: title, body, requiredAttendees: [] });
  else toast('Demo: calendar entry on ' + fmtDay(day));
}

/* ---------------- Controls (custom, no native selects/number/date) ---------------- */

let openPopup = null;
document.addEventListener('mousedown', (e) => {
  if (openPopup && !openPopup.contains(e.target)) closePopup();
});
function closePopup() { if (openPopup) { openPopup.querySelector('.cselect-pop')?.remove(); openPopup = null; } }

function customSelect({ options, value, onChange, placeholder = 'Select...', searchable = false, compact = false, renderValue }) {
  const wrap = h('div', { class: 'cselect' + (compact ? ' compact' : '') });
  const current = options.find((o) => o.value === value);
  const btn = h('button', { type: 'button', class: 'cselect-btn' + (current && current.className ? ' ' + current.className : '') },
    h('span', { class: 'grow' }, renderValue ? renderValue(current) : (current ? current.label : h('span', { class: 'muted' }, placeholder))),
    h('span', { class: 'caret' }, '▾'));
  wrap.append(btn);

  function open() {
    if (openPopup === wrap) { closePopup(); return; }
    closePopup();
    const pop = h('div', { class: 'cselect-pop', role: 'listbox' });
    let hl = 0;
    let filtered = options;
    const list = h('div');
    const search = searchable ? h('input', { class: 'input cselect-search', placeholder: 'Search...', autocomplete: 'off' }) : null;
    function paint() {
      const q = search ? search.value.trim().toLowerCase() : '';
      filtered = options.filter((o) => !q || (o.label + ' ' + (o.sub || '')).toLowerCase().includes(q));
      hl = Math.min(hl, Math.max(filtered.length - 1, 0));
      list.replaceChildren(...(filtered.length ? filtered.map((o, i) =>
        h('div', { class: 'cselect-opt' + (o.value === value ? ' sel' : '') + (i === hl ? ' hl' : ''), onmousedown: (e) => { e.preventDefault(); pick(o); } },
          o.label, o.sub ? h('div', { class: 'sub' }, o.sub) : null)) : [h('div', { class: 'cselect-opt muted' }, 'No results')]));
    }
    function pick(o) { closePopup(); onChange(o.value); }
    if (search) {
      search.addEventListener('input', () => { hl = 0; paint(); });
      search.addEventListener('keydown', (e) => {
        if (e.key === 'ArrowDown') { hl = Math.min(hl + 1, filtered.length - 1); paint(); e.preventDefault(); }
        if (e.key === 'ArrowUp') { hl = Math.max(hl - 1, 0); paint(); e.preventDefault(); }
        if (e.key === 'Enter' && filtered[hl]) { pick(filtered[hl]); e.preventDefault(); }
        if (e.key === 'Escape') closePopup();
      });
      pop.append(search);
    }
    pop.append(list);
    paint();
    wrap.append(pop);
    openPopup = wrap;
    if (search) search.focus();
  }
  btn.addEventListener('click', open);
  return wrap;
}

function statusSelect(map, value, onChange) {
  const options = Object.entries(map).map(([k, v]) => ({ value: k, label: v.label }));
  return customSelect({
    options, value, onChange, compact: true,
    renderValue: (o) => h('span', { class: 'pill ' + map[o.value].tone }, o.label),
  });
}

function checkbox(label, checked, onToggle, sub) {
  const el = h('div', { class: 'check' + (checked ? ' on' : ''), role: 'checkbox', tabindex: '0', 'aria-checked': String(checked) },
    h('span', { class: 'box' }, checked ? '✓' : ''), h('div', { class: 'grow' }, label, sub ? h('div', { class: 'sub' }, sub) : null));
  const toggle = () => onToggle(!checked);
  el.addEventListener('click', toggle);
  el.addEventListener('keydown', (e) => { if (e.key === ' ' || e.key === 'Enter') { e.preventDefault(); toggle(); } });
  return el;
}

function toggleSwitch(label, on, onToggle) {
  const el = h('div', { class: 'switch' + (on ? ' on' : ''), role: 'switch', tabindex: '0', 'aria-checked': String(on) },
    h('span', null, label), h('span', { class: 'track' }));
  const toggle = () => onToggle(!on);
  el.addEventListener('click', toggle);
  el.addEventListener('keydown', (e) => { if (e.key === ' ' || e.key === 'Enter') { e.preventDefault(); toggle(); } });
  return el;
}

function chipGroup(items, activeValue, onPick) {
  return h('div', { class: 'chips' }, items.map((it) =>
    h('button', { type: 'button', class: 'chip' + (it.value === activeValue ? ' on' : ''), onclick: () => onPick(it.value) }, it.label)));
}

// Date field: quick chips + free text in dd.mm.yyyy
function dateField(value, onChange, { allowEmpty = true } = {}) {
  const st = state.data.settings;
  const wrap = h('div', { class: 'stack' });
  const input = h('input', { class: 'input', placeholder: 'dd.mm.yyyy', value: value ? fmtDay(value) : '', autocomplete: 'off' });
  input.addEventListener('change', () => {
    if (!input.value.trim() && allowEmpty) { input.classList.remove('invalid'); onChange(''); return; }
    const d = parseUserDate(input.value);
    if (!d) { input.classList.add('invalid'); return; }
    input.classList.remove('invalid');
    input.value = fmtDay(d);
    onChange(d);
  });
  const set = (d) => { input.value = d ? fmtDay(d) : ''; input.classList.remove('invalid'); onChange(d); };
  const quick = [
    { label: 'Today', value: () => today() },
    { label: '+1d', value: () => addDays(today(), 1, st.workdaysOnly) },
    { label: '+2d', value: () => addDays(today(), 2, st.workdaysOnly) },
    { label: '+3d', value: () => addDays(today(), 3, st.workdaysOnly) },
    { label: '+1w', value: () => addDays(today(), 7, false) },
  ];
  if (allowEmpty) quick.push({ label: 'None', value: () => '' });
  wrap.append(input, h('div', { class: 'chips' }, quick.map((q) => h('button', { type: 'button', class: 'chip', onclick: () => set(q.value()) }, q.label))));
  return wrap;
}

function field(label, control) { return h('div', { class: 'field' }, h('span', { class: 'label' }, label), control); }
function textInput(value, onInput, placeholder) {
  const el = h('input', { class: 'input', value: value || '', placeholder: placeholder || '', autocomplete: 'off' });
  el.addEventListener('input', () => onInput(el.value));
  return el;
}
function textArea(value, onInput, placeholder, rows) {
  const el = h('textarea', { class: 'textarea', placeholder: placeholder || '', rows: rows || 3 });
  el.value = value || '';
  el.addEventListener('input', () => onInput(el.value));
  return el;
}

/* ---------------- Views ---------------- */

function render() {
  closePopup();
  document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.dataset.view === (state.view === 'request' ? 'list' : state.view)));
  const app = $('#app');
  const view = { email: viewEmail, list: viewList, request: viewRequest, settings: viewSettings }[state.view] || viewEmail;
  app.replaceChildren(view());
  updateBadge();
}

function updateBadge() {
  const n = totalOverdue();
  const b = $('#overdueBadge');
  b.hidden = n === 0;
  b.textContent = n;
}

function go(view, reqId) { state.view = view; if (reqId !== undefined) state.reqId = reqId; render(); window.scrollTo(0, 0); }

function emailHeader(info) {
  return h('div', { class: 'card flat' },
    h('div', { class: 'title' }, info.subject || '(no subject)'),
    h('div', { class: 'sub' }, info.fromMe ? 'Sent by you to ' + (info.to.map((t) => t.name || t.email).join(', ') || '-') : 'From ' + (info.from ? (info.from.name || info.from.email) : '-')),
    !state.inOutlook ? h('div', { class: 'row', style: 'margin-top:6px' },
      h('span', { class: 'sub grow' }, 'Demo mode (outside Outlook)'),
      h('button', { class: 'btn sm', onclick: () => { localSet('rt.mock', (localGet('rt.mock', 0) + 1) % 3); state.info = readItem(); render(); } }, 'Next mock email')) : null);
}

function viewEmail() {
  const info = state.info;
  if (!info) return h('div', { class: 'empty' }, 'Select an email to start.');
  const root = h('div');
  root.append(emailHeader(info));
  const a = analyzeItem(info);

  // Supplier reply detected
  for (const { r, s, by } of a.supplierHits) {
    if (!info.fromMe && s.status === 'sent') {
      root.append(h('div', { class: 'card alert' },
        h('div', { class: 'title' }, 'Reply from ' + s.name),
        h('div', { class: 'sub' }, 'Request: ' + r.title + (by === 'sender' ? ' (matched by sender)' : '')),
        h('div', { class: 'row wrap', style: 'margin-top:8px' },
          h('button', { class: 'btn primary', onclick: () => setSupStatus(r, s, 'replied', by) }, 'Mark replied'),
          h('button', { class: 'btn', onclick: () => setSupStatus(r, s, 'quoted', by) }, 'Offer received'),
          h('button', { class: 'btn ghost', onclick: () => go('request', r.id) }, 'Open'))));
    } else {
      root.append(h('div', { class: 'card good link', onclick: () => go('request', r.id) },
        h('div', { class: 'row between' }, h('div', { class: 'title grow' }, s.name), h('span', { class: 'pill ' + SUP_STATUS[s.status].tone }, SUP_STATUS[s.status].label)),
        h('div', { class: 'sub' }, 'Supplier on: ' + r.title)));
    }
  }
  for (const r of a.clientOf) root.append(requestCard(r, 'Client request'));

  const tracked = a.clientOf.length || a.supplierHits.length;
  if (!tracked) {
    if (info.fromMe) root.append(formLinkSuppliers(info));
    else root.append(formCreateRequest(info));
  }

  // Secondary actions
  const more = h('details');
  more.append(h('summary', null, tracked ? 'Other actions for this email' : (info.fromMe ? 'Or: create a new request from this email' : 'Or: this is a supplier email')));
  if (info.fromMe || tracked) more.append(formCreateRequest(info, true));
  if (!info.fromMe || tracked) more.append(formLinkSuppliers(info, true));
  root.append(h('div', { class: 'divider' }), more);
  return root;
}

function setSupStatus(r, s, status, by) {
  s.status = status;
  s.answeredAt = today();
  if (by === 'sender' && state.info && !s.convKey) s.convKey = state.info.convKey;
  save().then(() => { toast(s.name + ': ' + SUP_STATUS[status].label); render(); });
}

function formCreateRequest(info, secondary) {
  const st = state.data.settings;
  const draft = {
    title: cleanSubject(info.subject),
    client: info.fromMe ? (info.to[0] ? info.to[0].name || info.to[0].email : '') : (info.from ? info.from.name || info.from.email : ''),
    clientEmail: info.fromMe ? (info.to[0] ? info.to[0].email : '') : (info.from ? info.from.email : ''),
    dueAt: addDays(today(), 5, st.workdaysOnly),
    notes: '',
  };
  return h('div', { class: secondary ? '' : 'card' },
    secondary ? null : h('div', { class: 'section-title' }, 'New client request'),
    field('Title', textInput(draft.title, (v) => (draft.title = v))),
    field('Client', textInput(draft.client, (v) => (draft.client = v))),
    field('Deadline to client', dateField(draft.dueAt, (v) => (draft.dueAt = v))),
    field('Notes', textArea('', (v) => (draft.notes = v), 'Optional', 2)),
    h('button', { class: 'btn primary block', onclick: () => {
      if (!draft.title.trim()) { toast('Title is required', true); return; }
      const r = { id: uid(), title: draft.title.trim(), client: draft.client.trim(), clientEmail: draft.clientEmail, convKey: info.convKey, createdAt: today(), dueAt: draft.dueAt, status: 'open', notes: draft.notes.trim(), suppliers: [] };
      state.data.requests.unshift(r);
      save().then((ok) => { if (ok) { toast('Request created'); go('request', r.id); } });
    } }, 'Create request'));
}

function formLinkSuppliers(info, secondary) {
  const open = state.data.requests.filter((r) => r.status === 'open');
  const recipients = info.fromMe ? [...info.to, ...info.cc] : (info.from ? [info.from] : []);
  const pick = { reqId: state.reqId && reqById(state.reqId) && reqById(state.reqId).status === 'open' ? state.reqId : (open[0] ? open[0].id : null), chosen: new Set(info.fromMe ? info.to.map((t) => t.email) : recipients.map((t) => t.email)) };
  const root = h('div', { class: secondary ? '' : 'card' });

  function paint() {
    root.replaceChildren();
    if (!secondary) root.append(h('div', { class: 'section-title' }, 'Track as supplier request'));
    if (!open.length) {
      root.append(h('div', { class: 'sub' }, 'No open requests yet. Create one from the client email first.'));
      return;
    }
    root.append(field('Request', customSelect({
      options: open.map((r) => ({ value: r.id, label: r.title, sub: (r.client || '') + ' · ' + r.suppliers.length + ' suppliers' })),
      value: pick.reqId, searchable: open.length > 5, onChange: (v) => { pick.reqId = v; paint(); },
    })));
    root.append(h('span', { class: 'label', style: 'font-size:11px;font-weight:700;color:var(--text-2)' }, info.fromMe ? 'Suppliers (recipients)' : 'Supplier (sender)'));
    for (const rc of recipients) {
      root.append(checkbox(rc.name || rc.email, pick.chosen.has(rc.email), (on) => { on ? pick.chosen.add(rc.email) : pick.chosen.delete(rc.email); paint(); }, rc.name ? rc.email : null));
    }
    root.append(h('button', { class: 'btn primary block', style: 'margin-top:8px', onclick: () => {
      const r = reqById(pick.reqId);
      if (!r) { toast('Choose a request', true); return; }
      const chosen = recipients.filter((rc) => pick.chosen.has(rc.email));
      if (!chosen.length) { toast('Choose at least one supplier', true); return; }
      let added = 0;
      for (const rc of chosen) {
        const existing = r.suppliers.find((s) => s.email === rc.email);
        if (existing) {
          existing.convKey = info.convKey;
          if (!info.fromMe && existing.status === 'sent') { existing.status = 'replied'; existing.answeredAt = today(); }
          continue;
        }
        const s = newSupplier({ name: rc.name, email: rc.email, convKey: info.convKey, subject: info.subject, sentAt: info.date });
        if (!info.fromMe) { s.status = 'replied'; s.answeredAt = today(); }
        r.suppliers.push(s);
        added++;
      }
      save().then((ok) => { if (ok) { toast(added ? `${added} supplier(s) added` : 'Linked'); go('request', r.id); } });
    } }, info.fromMe ? 'Add as suppliers' : 'Link supplier'));
  }
  paint();
  return root;
}

function requestCard(r, kicker) {
  const od = reqOverdueCount(r);
  const total = r.suppliers.length;
  const done = answeredCount(r);
  return h('div', { class: 'card link', onclick: () => go('request', r.id) },
    kicker ? h('div', { class: 'section-title' }, kicker) : null,
    h('div', { class: 'row between' }, h('div', { class: 'title grow' }, r.title), h('span', { class: 'pill ' + REQ_STATUS[r.status].tone }, REQ_STATUS[r.status].label)),
    h('div', { class: 'sub' }, [r.client, r.dueAt ? 'due ' + fmtDay(r.dueAt) : null].filter(Boolean).join(' · ')),
    total ? h('div', { class: 'progress' }, h('span', { style: `width:${Math.round((done / total) * 100)}%` })) : null,
    h('div', { class: 'row wrap', style: 'margin-top:4px' },
      h('span', { class: 'sub' }, total ? `${done}/${total} answered` : 'No suppliers yet'),
      od ? h('span', { class: 'pill danger' }, `${od} to follow up`) : null,
      isReqOverdue(r) ? h('span', { class: 'pill danger' }, 'deadline passed') : null));
}

function viewList() {
  const root = h('div');
  const q = state.listQuery.toLowerCase();
  const filters = [
    { value: 'open', label: 'Open' },
    { value: 'followup', label: 'Follow up' },
    { value: 'all', label: 'All' },
  ];
  const search = h('input', { class: 'input', placeholder: 'Search title, client, supplier...', value: state.listQuery, autocomplete: 'off' });
  search.addEventListener('input', () => { state.listQuery = search.value; paintList(); });
  root.append(h('div', { class: 'stack', style: 'margin-bottom:10px' }, search, chipGroup(filters, state.listFilter, (v) => { state.listFilter = v; render(); })));
  const list = h('div');
  root.append(list);

  function paintList() {
    const qq = state.listQuery.toLowerCase();
    let items = state.data.requests.slice();
    if (state.listFilter === 'open') items = items.filter((r) => r.status === 'open');
    if (state.listFilter === 'followup') items = items.filter((r) => reqOverdueCount(r) || isReqOverdue(r));
    if (qq) items = items.filter((r) => [r.title, r.client, ...r.suppliers.map((s) => s.name + ' ' + s.email)].join(' ').toLowerCase().includes(qq));
    const urgency = (r) => (r.status !== 'open' ? 2 : (reqOverdueCount(r) || isReqOverdue(r) ? 0 : 1));
    items.sort((a, b) => urgency(a) - urgency(b) || (a.dueAt || '9999').localeCompare(b.dueAt || '9999'));
    list.replaceChildren(...(items.length ? items.map((r) => requestCard(r)) : [h('div', { class: 'empty' },
      state.data.requests.length ? 'Nothing matches.' : 'No requests yet. Open a client email and create one from the "This email" tab.')]));
  }
  void q;
  paintList();
  return root;
}

function viewRequest() {
  const r = reqById(state.reqId);
  if (!r) { state.view = 'list'; return viewList(); }
  const st = state.data.settings;
  const root = h('div');
  root.append(h('button', { class: 'btn ghost back', onclick: () => go('list') }, '← All requests'));

  // Header
  const head = h('div', { class: 'card' },
    h('div', { class: 'row between' }, h('div', { class: 'title grow' }, r.title), statusSelect(REQ_STATUS, r.status, (v) => { r.status = v; save().then(render); })),
    h('div', { class: 'sub' }, [r.client, 'created ' + fmtDay(r.createdAt)].filter(Boolean).join(' · ')),
    h('div', { class: 'row wrap', style: 'margin-top:6px' },
      r.dueAt ? h('span', { class: 'pill ' + (isReqOverdue(r) ? 'danger' : 'info') }, 'Deadline ' + fmtDay(r.dueAt) + ' (' + relDay(r.dueAt) + ')') : h('span', { class: 'sub' }, 'No deadline'),
      r.convKey ? h('button', { class: 'btn sm', onclick: () => openEmail(r.convKey) }, 'Open client email') : null,
      r.dueAt ? h('button', { class: 'btn sm', onclick: () => addToCalendar('Deadline: ' + r.title, r.dueAt, 'Client: ' + (r.client || '-')) }, 'Calendar') : null));
  root.append(head);

  // Suppliers
  const total = r.suppliers.length;
  root.append(h('div', { class: 'section-title' }, `Suppliers (${answeredCount(r)}/${total} answered)`));
  const supCard = h('div', { class: 'card' });
  if (!total) supCard.append(h('div', { class: 'sub' }, 'No suppliers yet. Open each email you sent to a supplier and use "Add as suppliers", or add them below.'));
  const order = { sent: 0, replied: 1, quoted: 2, declined: 3 };
  const sups = r.suppliers.slice().sort((a, b) => (isSupOverdue(r, b) - isSupOverdue(r, a)) || order[a.status] - order[b.status] || (a.followUpAt || '').localeCompare(b.followUpAt || ''));
  for (const s of sups) supCard.append(supplierRow(r, s));
  root.append(supCard);

  // Add suppliers manually
  const add = { text: '' };
  root.append(h('details', null, h('summary', null, '+ Add suppliers manually'),
    h('div', { class: 'card flat', style: 'margin-top:6px' },
      field('One per line: Name <email> or just email', textArea('', (v) => (add.text = v), 'Maria Distribuitor A <maria@distrib-a.ro>\noferte@distrib-b.ro', 3)),
      h('button', { class: 'btn primary', onclick: () => {
        const parsed = parseRecipients(add.text);
        if (!parsed.length) { toast('No valid email found', true); return; }
        let n = 0;
        for (const p of parsed) if (!r.suppliers.some((s) => s.email === p.email)) { r.suppliers.push(newSupplier(p)); n++; }
        save().then((ok) => { if (ok) { toast(`${n} supplier(s) added`); render(); } });
      } }, 'Add'))));

  // Notes
  root.append(h('div', { class: 'section-title' }, 'Notes'));
  const notes = textArea(r.notes, (v) => { r.notes = v; }, 'Notes...', 3);
  notes.addEventListener('blur', () => save());
  root.append(notes);

  root.append(h('div', { class: 'divider' }), h('button', { class: 'btn danger', onclick: () => {
    if (!confirmTwice(root, 'del')) return;
    state.data.requests = state.data.requests.filter((x) => x.id !== r.id);
    save().then(() => { toast('Request deleted'); go('list'); });
  } }, 'Delete request'));
  void st;
  return root;
}

// Two-click confirm without native dialogs
function confirmTwice(scope, key) {
  const k = '_confirm_' + key;
  if (scope[k] && Date.now() - scope[k] < 3000) return true;
  scope[k] = Date.now();
  toast('Click again to confirm');
  return false;
}

function parseRecipients(text) {
  const out = [];
  for (const line of String(text).split(/[\n;,]+/)) {
    const m = line.match(/([^\s<>"']+@[^\s<>"']+\.[a-z]{2,})/i);
    if (!m) continue;
    const email = m[1].toLowerCase();
    const name = line.replace(m[0], '').replace(/[<>"]/g, '').trim();
    out.push({ name, email });
  }
  return out;
}

function supplierRow(r, s) {
  const st = state.data.settings;
  const overdue = isSupOverdue(r, s);
  const row = h('div', { class: 'sup' });
  const meta = [];
  meta.push('sent ' + fmtDay(s.sentAt));
  if (s.reminders) meta.push(`${s.reminders} reminder${s.reminders > 1 ? 's' : ''}`);
  if (s.answeredAt && s.status !== 'sent') meta.push('answered ' + fmtDay(s.answeredAt));
  row.append(
    h('div', { class: 'row between' },
      h('div', { class: 'grow' }, h('div', { style: 'font-weight:700' }, s.name), h('div', { class: 'sub' }, s.email)),
      statusSelect(SUP_STATUS, s.status, (v) => {
        s.status = v;
        if (v !== 'sent') s.answeredAt = s.answeredAt || today();
        else { s.answeredAt = null; s.followUpAt = s.followUpAt || addDays(today(), st.followUpDays, st.workdaysOnly); }
        save().then(render);
      })),
    h('div', { class: 'sub', style: 'margin-top:2px' }, meta.join(' · ')));
  if (s.status === 'sent') {
    row.append(h('div', { class: 'row wrap', style: 'margin-top:4px' },
      h('span', { class: 'pill ' + (overdue ? 'danger' : 'pending') }, (overdue ? 'Follow up ' : 'Follow up ') + relDay(s.followUpAt))));
    const actions = h('div', { class: 'actions' },
      h('button', { class: 'btn sm' + (overdue ? ' primary' : ''), onclick: () => draftReminder(r, s) }, 'Draft reminder'),
      snoozeMenu(s),
      h('button', { class: 'btn sm', onclick: () => addToCalendar('Follow up: ' + s.name + ' - ' + r.title, s.followUpAt, 'Supplier: ' + s.name + ' <' + s.email + '>\nRequest: ' + r.title) }, 'Calendar'),
      s.convKey ? h('button', { class: 'btn sm', onclick: () => openEmail(s.convKey) }, 'Open email') : null,
      h('button', { class: 'btn sm ghost danger', onclick: () => {
        if (!confirmTwice(row, 'rm')) return;
        r.suppliers = r.suppliers.filter((x) => x.id !== s.id);
        save().then(render);
      } }, 'Remove'));
    row.append(actions);
  } else {
    row.append(h('div', { class: 'actions' },
      s.convKey ? h('button', { class: 'btn sm', onclick: () => openEmail(s.convKey) }, 'Open email') : null,
      h('button', { class: 'btn sm ghost danger', onclick: () => {
        if (!confirmTwice(row, 'rm')) return;
        r.suppliers = r.suppliers.filter((x) => x.id !== s.id);
        save().then(render);
      } }, 'Remove')));
  }
  return row;
}

function snoozeMenu(s) {
  const st = state.data.settings;
  return customSelect({
    compact: true,
    options: [
      { value: '1', label: '+1 day' },
      { value: '2', label: '+2 days' },
      { value: '3', label: '+3 days' },
      { value: '5', label: '+1 week (workdays)' },
    ],
    value: null,
    renderValue: () => 'Snooze',
    onChange: (v) => {
      const base = s.followUpAt && s.followUpAt > today() ? s.followUpAt : today();
      s.followUpAt = addDays(base, Number(v), st.workdaysOnly);
      save().then(() => { toast('Follow up ' + fmtDay(s.followUpAt)); render(); });
    },
  });
}

function viewSettings() {
  const st = state.data.settings;
  const root = h('div');
  root.append(h('div', { class: 'section-title' }, 'Follow-up'));
  root.append(h('div', { class: 'card' },
    field('Default follow-up after', chipGroup([1, 2, 3, 5, 7].map((n) => ({ value: n, label: n + (n === 1 ? ' day' : ' days') })), st.followUpDays, (v) => { st.followUpDays = v; save().then(render); })),
    toggleSwitch('Count working days only', st.workdaysOnly, (v) => { st.workdaysOnly = v; save().then(render); })));

  root.append(h('div', { class: 'section-title' }, 'Reminder email'));
  const subj = textInput(st.subjectPrefix, (v) => (st.subjectPrefix = v));
  const tpl = textArea(st.template, (v) => (st.template = v), '', 7);
  subj.addEventListener('blur', () => save());
  tpl.addEventListener('blur', () => save());
  root.append(h('div', { class: 'card' },
    field('Subject prefix', subj),
    field('Body', tpl),
    h('div', { class: 'sub' }, 'Placeholders: {supplier} {subject} {sentDate} {client}'),
    h('button', { class: 'btn sm', style: 'margin-top:6px', onclick: () => { st.template = DEFAULT_SETTINGS.template; st.subjectPrefix = DEFAULT_SETTINGS.subjectPrefix; save().then(render); } }, 'Reset to default')));

  root.append(h('div', { class: 'section-title' }, 'Storage'));
  const size = dataSize();
  const pct = Math.min(100, Math.round((size / ROAMING_LIMIT) * 100));
  const closed = state.data.requests.filter((r) => r.status !== 'open');
  const archive = localGet(LOCAL_ARCHIVE, []);
  const fileIn = h('input', { type: 'file', accept: 'application/json,.json', hidden: true });
  fileIn.addEventListener('change', () => importJson(fileIn.files[0]));
  root.append(h('div', { class: 'card' },
    h('div', { class: 'row between' }, h('span', null, state.inOutlook ? 'Mailbox storage' : 'Local storage (demo)'), h('span', { class: 'sub' }, `${(size / 1024).toFixed(1)} / 32 KB`)),
    h('div', { class: 'meter' + (pct > 80 ? ' hot' : '') }, h('span', { style: `width:${pct}%` })),
    h('div', { class: 'sub' }, `${state.data.requests.length} requests · ${archive.length} archived on this device`),
    h('div', { class: 'row wrap', style: 'margin-top:8px' },
      h('button', { class: 'btn sm', disabled: !closed.length, onclick: () => archiveClosed() }, `Archive closed (${closed.length})`),
      h('button', { class: 'btn sm', onclick: () => exportJson() }, 'Export JSON'),
      h('button', { class: 'btn sm', onclick: () => fileIn.click() }, 'Import JSON'),
      fileIn)));
  root.append(h('div', { class: 'sub', style: 'margin-top:10px' }, 'Request Tracker v0.1 · data stays in your mailbox'));
  return root;
}

function downloadJson(obj, name) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
  const a = h('a', { href: URL.createObjectURL(blob), download: name });
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}

function exportJson() {
  downloadJson({ exportedAt: new Date().toISOString(), ...state.data, archive: localGet(LOCAL_ARCHIVE, []) }, `request-tracker_${today()}.json`);
}

function archiveClosed() {
  const closed = state.data.requests.filter((r) => r.status !== 'open');
  if (!closed.length) return;
  const archive = localGet(LOCAL_ARCHIVE, []);
  localSet(LOCAL_ARCHIVE, archive.concat(closed));
  downloadJson({ archivedAt: new Date().toISOString(), requests: closed }, `request-tracker_archive_${today()}.json`);
  state.data.requests = state.data.requests.filter((r) => r.status === 'open');
  save().then(() => { toast(`${closed.length} request(s) archived`); render(); });
}

function importJson(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const obj = JSON.parse(reader.result);
      if (!Array.isArray(obj.requests)) throw new Error('missing requests');
      const known = new Set(state.data.requests.map((r) => r.id));
      let n = 0;
      for (const r of obj.requests) if (!known.has(r.id)) { state.data.requests.push(r); n++; }
      if (obj.settings) state.data.settings = { ...state.data.settings, ...obj.settings };
      save().then((ok) => { if (ok) { toast(`${n} request(s) imported`); render(); } });
    } catch (e) { toast('Invalid file: ' + e.message, true); }
  };
  reader.readAsText(file);
}

/* ---------------- Boot ---------------- */

function boot() {
  document.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => go(t.dataset.view)));
  loadData();
  state.info = readItem();
  // Default view: email context if it is tracked or new, list if nothing is selected
  state.view = state.info ? 'email' : 'list';
  render();
  if (state.inOutlook && Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      state.info = readItem();
      if (state.view === 'email' || !state.info) state.view = state.info ? 'email' : 'list';
      render();
    });
  }
}

if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady((info) => {
    state.inOutlook = !!(info && info.host === Office.HostType.Outlook);
    boot();
  });
} else {
  document.addEventListener('DOMContentLoaded', boot);
}
