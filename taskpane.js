/*
 * Request Tracker - Outlook add-in (v0.2)
 * Structure: Request (e.g. tender SCN xxx) > Equipment > Supplier quote requests.
 * Each supplier entry has its own follow-up interval and a link to the email, so the
 * "To do" list shows what needs chasing today and opens the right email in one click.
 * Nothing is ever sent automatically.
 * Storage: Outlook roaming settings (follows the mailbox, ~32 KB limit).
 * Outside Outlook the page runs in demo mode with localStorage and mock emails.
 */
'use strict';

const STORE_KEY = 'rt.v1';
const ROAMING_LIMIT = 32 * 1024;
const LOCAL_ITEM_MAP = 'rt.itemIds';
const LOCAL_ARCHIVE = 'rt.archive';
const FOLLOW_UP_CHOICES = [1, 2, 3, 5, 7, 10];

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
    'Revin cu solicitarea de oferta transmisa in data de {sentDate} privind {equipment} ("{subject}").\n' +
    'Va rog sa ne transmiteti oferta sau un termen estimat de raspuns.\n\n' +
    'Multumesc,',
};

const state = {
  inOutlook: false,
  data: { requests: [], settings: { ...DEFAULT_SETTINGS } },
  view: 'todo',
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
      else el.setAttribute(k, v === true ? '' : v);
    }
  }
  for (const c of children.flat()) {
    if (c === null || c === undefined || c === false) continue;
    el.append(c.nodeType ? c : document.createTextNode(String(c)));
  }
  return el;
}

function uid() { return Math.random().toString(36).slice(2, 8) + Date.now().toString(36).slice(-3); }

// Short stable hash (cyrb53) of the Exchange conversation ID, used for matching emails to entries.
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
  if (n === -1) return '1 day late';
  return n > 0 ? `in ${n} days` : `${-n} days late`;
}
function cleanSubject(s) { return String(s || '').replace(/^\s*((re|fw|fwd|tr|raspuns|redirectionare)\s*:\s*)+/i, '').trim(); }
function nameFromEmail(email) {
  const domain = (String(email).split('@')[1] || email).split('.');
  const n = domain.length > 1 ? domain[domain.length - 2] : domain[0];
  return n.charAt(0).toUpperCase() + n.slice(1);
}
const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`;

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

// Upgrade v0.1 data (no equipment, settings with template) to the current shape.
function migrate(data) {
  for (const r of data.requests) {
    r.items = r.items || [];
    for (const s of r.suppliers || []) {
      s.itemIds = s.itemIds || [];
      if (!s.followDays) s.followDays = data.settings.followUpDays || 2;
    }
  }
  if (data.settings.template && !data.settings.template.includes('{equipment}')) data.settings.template = DEFAULT_SETTINGS.template;
  return data;
}

function loadData() {
  let raw = state.inOutlook ? Office.context.roamingSettings.get(STORE_KEY) : localGet(STORE_KEY, null);
  if (typeof raw === 'string') { try { raw = JSON.parse(raw); } catch { raw = null; } }
  if (raw && Array.isArray(raw.requests)) {
    state.data = migrate({ requests: raw.requests, settings: { ...DEFAULT_SETTINGS, ...(raw.settings || {}) } });
  }
}

function dataSize() { return new Blob([JSON.stringify(state.data)]).size; }

function save() {
  if (dataSize() > ROAMING_LIMIT * 0.97) {
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

// Device-local fallback map convKey -> itemId (for entries saved before mailId was stored).
function rememberItem(convKey, itemId) {
  if (!convKey || !itemId) return;
  const map = localGet(LOCAL_ITEM_MAP, {});
  map[convKey] = itemId;
  localSet(LOCAL_ITEM_MAP, map);
}

/* ---------------- Domain ---------------- */

const reqById = (id) => state.data.requests.find((r) => r.id === id);
const isSupDue = (r, s) => r.status === 'open' && s.status === 'sent' && s.followUpAt && s.followUpAt <= today();
const isReqOverdue = (r) => r.status === 'open' && r.dueAt && r.dueAt < today();
const dueCount = (r) => r.suppliers.filter((s) => isSupDue(r, s)).length;
const answeredCount = (r) => r.suppliers.filter((s) => s.status !== 'sent').length;
function totalDue() { return state.data.requests.reduce((n, r) => n + dueCount(r), 0); }
function itemNames(r, s) {
  const names = s.itemIds.map((id) => (r.items.find((i) => i.id === id) || {}).name).filter(Boolean);
  return names.length ? names.join(', ') : 'General';
}

function newSupplier({ name, email, convKey, mailId, sentAt, followDays, itemIds }) {
  const st = state.data.settings;
  const sent = sentAt || today();
  const days = followDays || st.followUpDays;
  return {
    id: uid(),
    name: name || nameFromEmail(email || ''),
    email: (email || '').toLowerCase(),
    itemIds: itemIds || [],
    status: 'sent',
    sentAt: sent,
    followDays: days,
    followUpAt: addDays(sent > today() ? sent : today(), days, st.workdaysOnly),
    convKey: convKey || '',
    mailId: mailId || '',
  };
}

function parseLines(text) { return String(text || '').split('\n').map((l) => l.trim()).filter(Boolean); }

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
  // One email often goes to several suppliers: on an incoming reply keep only the supplier who wrote it.
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

const MOCKS = [
  { subject: 'LICITATIA SCN1180225 - echipamente retea', from: { name: 'Client Primarie', email: 'achizitii@primarie.ro' }, to: [{ name: 'Andrei', email: 'andrei@demo.ro' }], fromMe: false, convKey: 'c1' },
  { subject: 'Cerere oferta SCN1180225 - switch 48p PoE', from: { name: 'Andrei', email: 'andrei@demo.ro' }, to: [{ name: 'ASBIS Romania', email: 'oferte@asbis.ro' }], fromMe: true, convKey: 's1' },
  { subject: 'Cerere oferta SCN1180225 - firewall', from: { name: 'Andrei', email: 'andrei@demo.ro' }, to: [{ name: 'Ingram Micro', email: 'sales@ingrammicro.ro' }], fromMe: true, convKey: 's2' },
  { subject: 'RE: Cerere oferta SCN1180225 - switch 48p PoE', from: { name: 'ASBIS Romania', email: 'oferte@asbis.ro' }, to: [{ name: 'Andrei', email: 'andrei@demo.ro' }], fromMe: false, convKey: 's1' },
];
function mockItem() {
  const m = MOCKS[localGet('rt.mock', 0) % MOCKS.length];
  return { ...m, cc: [], itemId: 'mock-' + m.convKey, date: today() };
}

function openMail(mailId, convKey) {
  const id = mailId || (convKey ? localGet(LOCAL_ITEM_MAP, {})[convKey] : null);
  if (!id) { toast('No email linked. Open the email and use "Track" to link it.', true); return; }
  if (!state.inOutlook) { toast('Demo: would open ' + id); return; }
  try { Office.context.mailbox.displayMessageForm(id); }
  catch (e) { toast('Could not open the email (moved or deleted?)', true); }
}

function addToCalendar(title, day, body) {
  const start = parseDay(day < today() ? today() : day);
  start.setHours(9, 0, 0, 0);
  const end = new Date(start.getTime() + 15 * 60000);
  if (state.inOutlook) Office.context.mailbox.displayNewAppointmentForm({ start, end, subject: title, body, requiredAttendees: [] });
  else toast('Demo: calendar entry on ' + fmtDay(day));
}

function esc(t) { return String(t).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

// Opens a pre-filled draft (reply on the supplier thread when it is the open email). Never sends.
function draftReminder(r, s) {
  const st = state.data.settings;
  const subject = s.subject ? cleanSubject(s.subject) : r.title;
  const text = st.template
    .replaceAll('{supplier}', s.name)
    .replaceAll('{subject}', subject)
    .replaceAll('{equipment}', itemNames(r, s))
    .replaceAll('{sentDate}', fmtDay(s.sentAt))
    .replaceAll('{client}', r.client || '');
  const html = text.split('\n').map((l) => esc(l) || '&nbsp;').join('<br>');
  const onThread = state.info && s.convKey && state.info.convKey === s.convKey;
  if (state.inOutlook) {
    const it = Office.context.mailbox.item;
    if (onThread && it && typeof it.displayReplyForm === 'function') it.displayReplyForm({ htmlBody: html });
    else Office.context.mailbox.displayNewMessageForm({ toRecipients: s.email ? [s.email] : [], subject: st.subjectPrefix + subject, htmlBody: html });
  } else {
    toast('Demo: draft opened for ' + (s.email || s.name));
  }
  // Drafting counts as chasing: push the next check by the supplier interval
  snooze(s, s.followDays);
}

/* ---------------- Custom controls (no native select/number/date) ---------------- */

let openPopup = null;
document.addEventListener('mousedown', (e) => { if (openPopup && !openPopup.contains(e.target)) closePopup(); });
function closePopup() { if (openPopup) { openPopup.querySelector('.cselect-pop')?.remove(); openPopup = null; } }

function customSelect({ options, value, onChange, placeholder = 'Select...', searchable = false, compact = false, renderValue }) {
  const wrap = h('div', { class: 'cselect' + (compact ? ' compact' : '') });
  const current = options.find((o) => o.value === value);
  const btn = h('button', { type: 'button', class: 'cselect-btn' },
    h('span', { class: 'grow' }, renderValue ? renderValue(current) : (current ? current.label : h('span', { class: 'muted' }, placeholder))),
    h('span', { class: 'caret' }, '▾'));
  wrap.append(btn);

  function open() {
    if (openPopup === wrap) { closePopup(); return; }
    closePopup();
    const pop = h('div', { class: 'cselect-pop', role: 'listbox' });
    const list = h('div');
    const search = searchable ? h('input', { class: 'input cselect-search', placeholder: 'Search...', autocomplete: 'off' }) : null;
    let hl = 0, filtered = options;
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
  return customSelect({
    options: Object.entries(map).map(([k, v]) => ({ value: k, label: v.label })),
    value, onChange, compact: true,
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
function daysChips(active, onPick) {
  return chipGroup(FOLLOW_UP_CHOICES.map((n) => ({ value: n, label: n === 1 ? '1 day' : n + ' days' })), active, onPick);
}

function dateField(value, onChange) {
  const st = state.data.settings;
  const wrap = h('div', { class: 'stack' });
  const input = h('input', { class: 'input', placeholder: 'dd.mm.yyyy', value: value ? fmtDay(value) : '', autocomplete: 'off' });
  input.addEventListener('change', () => {
    if (!input.value.trim()) { input.classList.remove('invalid'); onChange(''); return; }
    const d = parseUserDate(input.value);
    if (!d) { input.classList.add('invalid'); return; }
    input.classList.remove('invalid');
    input.value = fmtDay(d);
    onChange(d);
  });
  const set = (d) => { input.value = d ? fmtDay(d) : ''; input.classList.remove('invalid'); onChange(d); };
  const quick = [
    { label: '+1d', value: () => addDays(today(), 1, st.workdaysOnly) },
    { label: '+3d', value: () => addDays(today(), 3, st.workdaysOnly) },
    { label: '+5d', value: () => addDays(today(), 5, st.workdaysOnly) },
    { label: '+1w', value: () => addDays(today(), 7, false) },
    { label: '+2w', value: () => addDays(today(), 14, false) },
    { label: 'None', value: () => '' },
  ];
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

// Two-click confirm without native dialogs
function confirmTwice(scope, key) {
  const k = '_confirm_' + key;
  if (scope[k] && Date.now() - scope[k] < 3000) return true;
  scope[k] = Date.now();
  toast('Click again to confirm');
  return false;
}

/* ---------------- Views ---------------- */

function render() {
  closePopup();
  const tabView = state.view === 'request' ? 'list' : state.view;
  document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.dataset.view === tabView));
  const view = { todo: viewTodo, email: viewEmail, list: viewList, request: viewRequest, settings: viewSettings }[state.view] || viewTodo;
  $('#app').replaceChildren(view());
  updateBadge();
}

function updateBadge() {
  const n = totalDue();
  const b = $('#dueBadge');
  b.hidden = n === 0;
  b.textContent = n;
}

function go(view, reqId) { state.view = view; if (reqId !== undefined) state.reqId = reqId; render(); window.scrollTo(0, 0); }

function setSupStatus(r, s, status, linkCurrent) {
  s.status = status;
  if (status !== 'sent') s.answeredAt = today();
  else s.answeredAt = null;
  if (linkCurrent && state.info) {
    if (!s.convKey) s.convKey = state.info.convKey;
    s.mailId = state.info.itemId; // point to the latest message from the supplier
  }
  save().then(() => { toast(s.name + ': ' + SUP_STATUS[status].label); render(); });
}

function snooze(s, days) {
  const st = state.data.settings;
  s.followUpAt = addDays(today(), days, st.workdaysOnly);
  s.followups = (s.followups || 0) + 1;
  save().then(() => { toast(`${s.name}: next check ${fmtDay(s.followUpAt)}`); render(); });
}

/* To do: everything that needs a follow-up, grouped by request */
function viewTodo() {
  const root = h('div');
  const due = [];
  const upcoming = [];
  const horizon = addDays(today(), 3, false);
  for (const r of state.data.requests) {
    if (r.status !== 'open') continue;
    for (const s of r.suppliers) {
      if (s.status !== 'sent' || !s.followUpAt) continue;
      if (s.followUpAt <= today()) due.push({ r, s });
      else if (s.followUpAt <= horizon) upcoming.push({ r, s });
    }
  }
  const byDate = (a, b) => a.s.followUpAt.localeCompare(b.s.followUpAt);
  due.sort(byDate);
  upcoming.sort(byDate);
  const deadlines = state.data.requests.filter((r) => r.status === 'open' && r.dueAt && r.dueAt <= horizon).sort((a, b) => a.dueAt.localeCompare(b.dueAt));

  if (!due.length && !upcoming.length && !deadlines.length) {
    root.append(h('div', { class: 'empty' },
      h('div', { class: 'title' }, 'Nothing to chase today'),
      h('div', { class: 'sub' }, state.data.requests.length ? 'All follow-ups are in the future.' : 'Open a client or tender email and create a request from the "This email" tab.')));
    return root;
  }
  root.append(h('div', { class: 'section-title' }, `Follow up now (${due.length})`));
  if (!due.length) root.append(h('div', { class: 'sub', style: 'margin-bottom:10px' }, 'Nothing due today.'));
  root.append(...groupByRequest(due).map(([r, rows]) => todoGroup(r, rows, true)));

  if (deadlines.length) {
    root.append(h('div', { class: 'section-title' }, 'Client deadlines (next 3 days)'));
    for (const r of deadlines) {
      root.append(h('div', { class: 'card link', onclick: () => go('request', r.id) },
        h('div', { class: 'row between' }, h('div', { class: 'title grow' }, r.title),
          h('span', { class: 'pill ' + (r.dueAt < today() ? 'danger' : 'pending') }, fmtDay(r.dueAt))),
        h('div', { class: 'sub' }, `${answeredCount(r)}/${r.suppliers.length} suppliers answered`)));
    }
  }
  if (upcoming.length) {
    root.append(h('div', { class: 'section-title' }, `Coming up (${upcoming.length})`));
    root.append(...groupByRequest(upcoming).map(([r, rows]) => todoGroup(r, rows, false)));
  }
  return root;
}

function groupByRequest(rows) {
  const map = new Map();
  for (const x of rows) { if (!map.has(x.r)) map.set(x.r, []); map.get(x.r).push(x.s); }
  return [...map.entries()];
}

function todoGroup(r, sups, isDue) {
  const card = h('div', { class: 'card' + (isDue ? ' alert' : '') },
    h('div', { class: 'row between' },
      h('button', { class: 'linklike title grow', onclick: () => go('request', r.id) }, r.title),
      r.dueAt ? h('span', { class: 'pill ' + (isReqOverdue(r) ? 'danger' : 'info') }, 'due ' + fmtDay(r.dueAt)) : null));
  for (const s of sups) {
    const row = h('div', { class: 'sup' },
      h('div', { class: 'row between' },
        h('div', { class: 'grow' },
          h('div', { style: 'font-weight:700' }, s.name),
          h('div', { class: 'sub' }, itemNames(r, s))),
        h('span', { class: 'pill ' + (isDue ? 'danger' : 'pending') }, relDay(s.followUpAt))),
      h('div', { class: 'sub' }, `sent ${fmtDay(s.sentAt)} · every ${s.followDays}d` + (s.followups ? ` · ${plural(s.followups, 'follow-up', 'follow-ups')}` : '')),
      h('div', { class: 'actions' },
        h('button', { class: 'btn sm primary', onclick: () => openMail(s.mailId, s.convKey) }, 'Open email'),
        h('button', { class: 'btn sm', onclick: () => draftReminder(r, s) }, 'Draft reminder'),
        isDue ? h('button', { class: 'btn sm', title: 'I chased them, check again after the interval', onclick: () => snooze(s, s.followDays) }, `Chased, +${s.followDays}d`) : null,
        h('button', { class: 'btn sm', onclick: () => setSupStatus(r, s, 'replied') }, 'Replied'),
        h('button', { class: 'btn sm', onclick: () => setSupStatus(r, s, 'quoted') }, 'Offer in')));
    card.append(row);
  }
  return card;
}

/* This email */
function emailHeader(info) {
  return h('div', { class: 'card flat' },
    h('div', { class: 'title' }, info.subject || '(no subject)'),
    h('div', { class: 'sub' }, info.fromMe ? 'Sent by you to ' + (info.to.map((t) => t.name || t.email).join(', ') || '-') : 'From ' + (info.from ? (info.from.name || info.from.email) : '-')),
    !state.inOutlook ? h('div', { class: 'row', style: 'margin-top:6px' },
      h('span', { class: 'sub grow' }, 'Demo mode (outside Outlook)'),
      h('button', { class: 'btn sm', onclick: () => { localSet('rt.mock', (localGet('rt.mock', 0) + 1) % MOCKS.length); state.info = readItem(); render(); } }, 'Next mock email')) : null);
}

function viewEmail() {
  const info = state.info;
  if (!info) return h('div', { class: 'empty' }, 'Select an email to start.');
  const root = h('div');
  root.append(emailHeader(info));
  const a = analyzeItem(info);

  for (const { r, s, by } of a.supplierHits) {
    if (!info.fromMe && s.status === 'sent') {
      root.append(h('div', { class: 'card alert' },
        h('div', { class: 'title' }, 'Reply from ' + s.name),
        h('div', { class: 'sub' }, r.title + ' · ' + itemNames(r, s) + (by === 'sender' ? ' (matched by sender)' : '')),
        h('div', { class: 'row wrap', style: 'margin-top:8px' },
          h('button', { class: 'btn primary', onclick: () => setSupStatus(r, s, 'replied', true) }, 'Mark replied'),
          h('button', { class: 'btn', onclick: () => setSupStatus(r, s, 'quoted', true) }, 'Offer received'),
          h('button', { class: 'btn ghost', onclick: () => go('request', r.id) }, 'Open request'))));
    } else {
      root.append(h('div', { class: 'card good link', onclick: () => go('request', r.id) },
        h('div', { class: 'row between' }, h('div', { class: 'title grow' }, s.name), h('span', { class: 'pill ' + SUP_STATUS[s.status].tone }, SUP_STATUS[s.status].label)),
        h('div', { class: 'sub' }, r.title + ' · ' + itemNames(r, s)),
        s.status === 'sent' ? h('div', { class: 'sub' }, 'Next check ' + fmtDay(s.followUpAt) + ' (' + relDay(s.followUpAt) + ')') : null));
    }
  }
  for (const r of a.clientOf) root.append(requestCard(r, 'Request from this email'));

  const tracked = a.clientOf.length || a.supplierHits.length;
  if (!tracked) {
    if (info.fromMe) root.append(formTrackSuppliers(info));
    else root.append(formCreateRequest(info));
  }
  const more = h('details');
  more.append(h('summary', null, tracked ? 'Other actions for this email' : (info.fromMe ? 'Or: create a new request from this email' : 'Or: this is an email to/from a supplier')));
  if (info.fromMe || tracked) more.append(formCreateRequest(info, true));
  if (!info.fromMe || tracked) more.append(formTrackSuppliers(info, true));
  root.append(h('div', { class: 'divider' }), more);
  return root;
}

function formCreateRequest(info, secondary) {
  const draft = {
    title: cleanSubject(info.subject),
    client: info.fromMe ? (info.to[0] ? info.to[0].name || info.to[0].email : '') : (info.from ? info.from.name || info.from.email : ''),
    dueAt: '',
    items: '',
    notes: '',
  };
  return h('div', { class: secondary ? 'stack-top' : 'card' },
    secondary ? null : h('div', { class: 'section-title' }, 'New request / tender'),
    field('Title', textInput(draft.title, (v) => (draft.title = v), 'e.g. LICITATIA SCN1180225')),
    field('Client', textInput(draft.client, (v) => (draft.client = v))),
    field('Equipment (one per line)', textArea('', (v) => (draft.items = v), 'Switch 48p PoE\nFirewall NGFW\nUPS 3kVA', 3)),
    field('Deadline to client', dateField(draft.dueAt, (v) => (draft.dueAt = v))),
    h('button', { class: 'btn primary block', onclick: () => {
      if (!draft.title.trim()) { toast('Title is required', true); return; }
      const r = {
        id: uid(), title: draft.title.trim(), client: draft.client.trim(), convKey: info.convKey, mailId: info.itemId,
        createdAt: today(), dueAt: draft.dueAt, status: 'open', notes: '',
        items: parseLines(draft.items).map((name) => ({ id: uid(), name })), suppliers: [],
      };
      state.data.requests.unshift(r);
      save().then((ok) => { if (ok) { toast('Request created'); go('request', r.id); } });
    } }, 'Create request'));
}

/* Link the open email (usually one I sent) to a request + equipment, with a follow-up interval */
function formTrackSuppliers(info, secondary) {
  const open = state.data.requests.filter((r) => r.status === 'open');
  const recipients = info.fromMe ? [...info.to, ...info.cc] : (info.from ? [info.from] : []);
  const pick = {
    reqId: state.reqId && reqById(state.reqId) && reqById(state.reqId).status === 'open' ? state.reqId : (open[0] ? open[0].id : null),
    emails: new Set(info.fromMe ? info.to.map((t) => t.email) : recipients.map((t) => t.email)),
    itemIds: new Set(),
    newItem: '',
    days: state.data.settings.followUpDays,
    date: '',
  };
  const root = h('div', { class: secondary ? 'stack-top' : 'card' });

  function paint() {
    root.replaceChildren();
    if (!secondary) root.append(h('div', { class: 'section-title' }, 'Track supplier request'));
    if (!open.length) { root.append(h('div', { class: 'sub' }, 'No open requests yet. Create one from the client / tender email first.')); return; }
    const r = reqById(pick.reqId);
    root.append(field('Request', customSelect({
      options: open.map((x) => ({ value: x.id, label: x.title, sub: [x.client, plural(x.items.length, 'item', 'items')].filter(Boolean).join(' · ') })),
      value: pick.reqId, searchable: open.length > 4, onChange: (v) => { pick.reqId = v; pick.itemIds.clear(); paint(); },
    })));

    const eq = h('div', { class: 'field' }, h('span', { class: 'label' }, 'For equipment'));
    for (const it of r.items) eq.append(checkbox(it.name, pick.itemIds.has(it.id), (on) => { on ? pick.itemIds.add(it.id) : pick.itemIds.delete(it.id); paint(); }));
    const ni = textInput(pick.newItem, (v) => (pick.newItem = v), r.items.length ? '+ new equipment (optional)' : 'Equipment name (optional)');
    eq.append(ni);
    root.append(eq);

    const sup = h('div', { class: 'field' }, h('span', { class: 'label' }, info.fromMe ? 'Supplier (recipients)' : 'Supplier (sender)'));
    for (const rc of recipients) sup.append(checkbox(rc.name || rc.email, pick.emails.has(rc.email), (on) => { on ? pick.emails.add(rc.email) : pick.emails.delete(rc.email); paint(); }, rc.name ? rc.email : null));
    root.append(sup);

    if (info.fromMe) {
      const st = state.data.settings;
      const preview = pick.date || addDays(today(), pick.days, st.workdaysOnly);
      root.append(field('Remind me after', daysChips(pick.date ? null : pick.days, (v) => { pick.days = v; pick.date = ''; paint(); })));
      root.append(field('...or deadline on a date', dateField(pick.date, (v) => { pick.date = v; paint(); })));
      root.append(h('div', { class: 'sub', style: 'margin:-4px 0 8px' }, 'Reminder on ' + fmtDay(preview) + ' (' + relDay(preview) + ')'));
    }

    root.append(h('button', { class: 'btn primary block', onclick: () => {
      const chosen = recipients.filter((rc) => pick.emails.has(rc.email));
      if (!chosen.length) { toast('Choose at least one supplier', true); return; }
      const itemIds = [...pick.itemIds];
      if (pick.newItem.trim()) { const it = { id: uid(), name: pick.newItem.trim() }; r.items.push(it); itemIds.push(it.id); }
      let added = 0;
      for (const rc of chosen) {
        // Same supplier + same thread already tracked: just extend the equipment list
        const existing = r.suppliers.find((s) => s.email === rc.email && s.convKey === info.convKey);
        if (existing) {
          existing.itemIds = [...new Set([...existing.itemIds, ...itemIds])];
          existing.mailId = info.itemId;
          continue;
        }
        const s = newSupplier({ name: rc.name, email: rc.email, convKey: info.convKey, mailId: info.itemId, sentAt: info.date, followDays: pick.days, itemIds });
        s.subject = info.subject;
        if (pick.date) s.followUpAt = pick.date;
        if (!info.fromMe) { s.status = 'replied'; s.answeredAt = today(); }
        r.suppliers.push(s);
        added++;
      }
      save().then((ok) => {
        if (!ok) return;
        toast(added ? `Tracking ${plural(added, 'supplier', 'suppliers')}` + (info.fromMe ? `, reminder ${fmtDay(pick.date || addDays(today(), pick.days, state.data.settings.workdaysOnly))}` : '') : 'Updated');
        render();
      });
    } }, info.fromMe ? 'Track & remind me' : 'Link supplier'));
  }
  paint();
  return root;
}

function requestCard(r, kicker) {
  const due = dueCount(r);
  const total = r.suppliers.length;
  const done = answeredCount(r);
  return h('div', { class: 'card link', onclick: () => go('request', r.id) },
    kicker ? h('div', { class: 'section-title' }, kicker) : null,
    h('div', { class: 'row between' }, h('div', { class: 'title grow' }, r.title), h('span', { class: 'pill ' + REQ_STATUS[r.status].tone }, REQ_STATUS[r.status].label)),
    h('div', { class: 'sub' }, [r.client, plural(r.items.length, 'item', 'items'), r.dueAt ? 'due ' + fmtDay(r.dueAt) : null].filter(Boolean).join(' · ')),
    total ? h('div', { class: 'progress' }, h('span', { style: `width:${Math.round((done / total) * 100)}%` })) : null,
    h('div', { class: 'row wrap', style: 'margin-top:4px' },
      h('span', { class: 'sub' }, total ? `${done}/${total} answered` : 'No suppliers yet'),
      due ? h('span', { class: 'pill danger' }, `${due} to follow up`) : null,
      isReqOverdue(r) ? h('span', { class: 'pill danger' }, 'deadline passed') : null));
}

function viewList() {
  const root = h('div');
  const search = h('input', { class: 'input', placeholder: 'Search title, client, equipment, supplier...', value: state.listQuery, autocomplete: 'off' });
  search.addEventListener('input', () => { state.listQuery = search.value; paintList(); });
  root.append(h('div', { class: 'stack', style: 'margin-bottom:10px' }, search,
    chipGroup([{ value: 'open', label: 'Open' }, { value: 'due', label: 'To follow up' }, { value: 'all', label: 'All' }], state.listFilter, (v) => { state.listFilter = v; render(); })));
  const list = h('div');
  root.append(list);
  function paintList() {
    const q = state.listQuery.toLowerCase();
    let items = state.data.requests.slice();
    if (state.listFilter === 'open') items = items.filter((r) => r.status === 'open');
    if (state.listFilter === 'due') items = items.filter((r) => dueCount(r) || isReqOverdue(r));
    if (q) items = items.filter((r) => [r.title, r.client, ...r.items.map((i) => i.name), ...r.suppliers.map((s) => s.name + ' ' + s.email)].join(' ').toLowerCase().includes(q));
    const urgency = (r) => (r.status !== 'open' ? 2 : (dueCount(r) || isReqOverdue(r) ? 0 : 1));
    items.sort((a, b) => urgency(a) - urgency(b) || (a.dueAt || '9999').localeCompare(b.dueAt || '9999'));
    list.replaceChildren(...(items.length ? items.map((r) => requestCard(r)) : [h('div', { class: 'empty' }, state.data.requests.length ? 'Nothing matches.' : 'No requests yet.')]));
  }
  paintList();
  return root;
}

/* Request detail: equipment > suppliers */
function viewRequest() {
  const r = reqById(state.reqId);
  if (!r) { state.view = 'list'; return viewList(); }
  const root = h('div');
  root.append(h('button', { class: 'btn ghost back', onclick: () => go('list') }, '← All requests'));

  root.append(h('div', { class: 'card' },
    h('div', { class: 'row between' }, h('div', { class: 'title grow' }, r.title), statusSelect(REQ_STATUS, r.status, (v) => { r.status = v; save().then(render); })),
    h('div', { class: 'sub' }, [r.client, 'created ' + fmtDay(r.createdAt)].filter(Boolean).join(' · ')),
    h('div', { class: 'row wrap', style: 'margin-top:6px' },
      r.dueAt ? h('span', { class: 'pill ' + (isReqOverdue(r) ? 'danger' : 'info') }, 'Deadline ' + fmtDay(r.dueAt) + ' (' + relDay(r.dueAt).replace('late', 'ago') + ')') : h('span', { class: 'sub' }, 'No deadline'),
      (r.mailId || r.convKey) ? h('button', { class: 'btn sm', onclick: () => openMail(r.mailId, r.convKey) }, 'Open request email') : null,
      r.dueAt ? h('button', { class: 'btn sm', onclick: () => addToCalendar('Deadline: ' + r.title, r.dueAt, 'Client: ' + (r.client || '-')) }, 'Calendar') : null)));

  root.append(h('div', { class: 'section-title' }, `Equipment & suppliers (${answeredCount(r)}/${r.suppliers.length} answered)`));
  const groups = r.items.map((it) => ({ it, sups: r.suppliers.filter((s) => s.itemIds.includes(it.id)) }));
  const general = r.suppliers.filter((s) => !s.itemIds.length || !s.itemIds.some((id) => r.items.some((i) => i.id === id)));
  if (!groups.length && !general.length) root.append(h('div', { class: 'sub', style: 'margin-bottom:8px' }, 'Add the equipment below, then open each email you sent to a supplier and press "Track & remind me".'));
  for (const g of groups) root.append(equipmentCard(r, g.it, g.sups));
  if (general.length) root.append(equipmentCard(r, null, general));

  const add = { text: '' };
  root.append(h('details', null, h('summary', null, '+ Add equipment'),
    h('div', { class: 'card flat', style: 'margin-top:6px' },
      field('One per line', textArea('', (v) => (add.text = v), 'Switch 48p PoE\nFirewall NGFW', 3)),
      h('button', { class: 'btn primary', onclick: () => {
        const names = parseLines(add.text);
        if (!names.length) return;
        r.items.push(...names.map((name) => ({ id: uid(), name })));
        save().then(render);
      } }, 'Add'))));

  root.append(h('div', { class: 'section-title' }, 'Notes'));
  const notes = textArea(r.notes, (v) => { r.notes = v; }, 'Notes...', 3);
  notes.addEventListener('blur', () => save());
  root.append(notes);

  root.append(h('div', { class: 'divider' }), h('button', { class: 'btn danger', onclick: () => {
    if (!confirmTwice(root, 'del')) return;
    state.data.requests = state.data.requests.filter((x) => x.id !== r.id);
    save().then(() => { toast('Request deleted'); go('list'); });
  } }, 'Delete request'));
  return root;
}

function equipmentCard(r, it, sups) {
  const answered = sups.filter((s) => s.status !== 'sent').length;
  const card = h('div', { class: 'card' });
  const head = h('div', { class: 'row between' },
    h('div', { class: 'grow' }, h('div', { class: 'title' }, it ? it.name : 'General (no equipment)'),
      h('div', { class: 'sub' }, sups.length ? `${answered}/${sups.length} answered` : 'No supplier asked yet')),
    it ? h('button', { class: 'btn sm ghost danger', onclick: () => {
      if (!confirmTwice(card, 'rmi')) return;
      r.items = r.items.filter((x) => x.id !== it.id);
      for (const s of r.suppliers) s.itemIds = s.itemIds.filter((id) => id !== it.id);
      save().then(render);
    } }, 'Remove') : null);
  card.append(head);
  const order = { sent: 0, replied: 1, quoted: 2, declined: 3 };
  for (const s of sups.slice().sort((a, b) => order[a.status] - order[b.status] || (a.followUpAt || '').localeCompare(b.followUpAt || ''))) {
    card.append(supplierRow(r, s));
  }
  return card;
}

function supplierRow(r, s) {
  const st = state.data.settings;
  const due = isSupDue(r, s);
  const row = h('div', { class: 'sup' });
  const meta = ['sent ' + fmtDay(s.sentAt)];
  if (s.followups) meta.push(plural(s.followups, 'follow-up', 'follow-ups'));
  if (s.answeredAt && s.status !== 'sent') meta.push('answered ' + fmtDay(s.answeredAt));
  row.append(
    h('div', { class: 'row between' },
      h('div', { class: 'grow' }, h('div', { style: 'font-weight:700' }, s.name), h('div', { class: 'sub' }, s.email || 'no email')),
      statusSelect(SUP_STATUS, s.status, (v) => {
        s.status = v;
        if (v !== 'sent') s.answeredAt = s.answeredAt || today();
        else { s.answeredAt = null; s.followUpAt = addDays(today(), s.followDays, st.workdaysOnly); }
        save().then(render);
      })),
    h('div', { class: 'sub', style: 'margin-top:2px' }, meta.join(' · ')));
  if (s.status === 'sent') {
    row.append(h('div', { class: 'row wrap', style: 'margin-top:4px' },
      h('span', { class: 'pill ' + (due ? 'danger' : 'pending') }, 'Check ' + relDay(s.followUpAt)),
      customSelect({
        compact: true, value: s.followDays,
        options: FOLLOW_UP_CHOICES.map((n) => ({ value: n, label: n === 1 ? 'every 1 day' : `every ${n} days` })),
        renderValue: () => `every ${s.followDays}d`,
        onChange: (v) => { s.followDays = v; s.followUpAt = addDays(s.sentAt > today() ? s.sentAt : today(), v, st.workdaysOnly); save().then(render); },
      })));
    row.append(h('details', { class: 'mini' }, h('summary', null, 'Change date'),
      dateField(s.followUpAt, (v) => { if (!v) return; s.followUpAt = v; save().then(() => { toast(`${s.name}: reminder ${fmtDay(v)}`); render(); }); })));
  }
  row.append(h('div', { class: 'actions' },
    (s.mailId || s.convKey) ? h('button', { class: 'btn sm' + (due ? ' primary' : ''), onclick: () => openMail(s.mailId, s.convKey) }, 'Open email') : null,
    s.status === 'sent' ? h('button', { class: 'btn sm', onclick: () => draftReminder(r, s) }, 'Draft reminder') : null,
    s.status === 'sent' && due ? h('button', { class: 'btn sm', onclick: () => snooze(s, s.followDays) }, `Chased, +${s.followDays}d`) : null,
    s.status === 'sent' ? h('button', { class: 'btn sm', onclick: () => addToCalendar('Follow up: ' + s.name + ' - ' + r.title, s.followUpAt, 'Supplier: ' + s.name + ' <' + s.email + '>\nEquipment: ' + itemNames(r, s) + '\nRequest: ' + r.title) }, 'Calendar') : null,
    h('button', { class: 'btn sm ghost danger', onclick: () => {
      if (!confirmTwice(row, 'rm')) return;
      r.suppliers = r.suppliers.filter((x) => x.id !== s.id);
      save().then(render);
    } }, 'Remove')));
  return row;
}

/* Settings */
function viewSettings() {
  const st = state.data.settings;
  const root = h('div');
  root.append(h('div', { class: 'section-title' }, 'Follow-up'));
  root.append(h('div', { class: 'card' },
    field('Default "remind me after"', daysChips(st.followUpDays, (v) => { st.followUpDays = v; save().then(render); })),
    toggleSwitch('Count working days only', st.workdaysOnly, (v) => { st.workdaysOnly = v; save().then(render); })));

  root.append(h('div', { class: 'section-title' }, 'Reminder draft (never sent automatically)'));
  const subj = textInput(st.subjectPrefix, (v) => (st.subjectPrefix = v));
  const tpl = textArea(st.template, (v) => (st.template = v), '', 7);
  subj.addEventListener('blur', () => save());
  tpl.addEventListener('blur', () => save());
  root.append(h('div', { class: 'card' },
    field('Subject prefix (new email only)', subj),
    field('Body', tpl),
    h('div', { class: 'sub' }, 'Placeholders: {supplier} {equipment} {subject} {sentDate} {client}'),
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
    h('div', { class: 'sub' }, `${plural(state.data.requests.length, 'request', 'requests')} · ${archive.length} archived on this device`),
    h('div', { class: 'row wrap', style: 'margin-top:8px' },
      h('button', { class: 'btn sm', disabled: !closed.length, onclick: archiveClosed }, `Archive closed (${closed.length})`),
      h('button', { class: 'btn sm', onclick: exportJson }, 'Export JSON'),
      h('button', { class: 'btn sm', onclick: () => fileIn.click() }, 'Import JSON'),
      fileIn)));
  root.append(h('div', { class: 'sub', style: 'margin-top:10px' }, 'Request Tracker v0.2.1 · reminders are drafts, never sent automatically · data stays in your mailbox'));
  return root;
}

function downloadJson(obj, name) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
  const a = h('a', { href: URL.createObjectURL(blob), download: name });
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}
function exportJson() { downloadJson({ exportedAt: new Date().toISOString(), ...state.data, archive: localGet(LOCAL_ARCHIVE, []) }, `request-tracker_${today()}.json`); }
function archiveClosed() {
  const closed = state.data.requests.filter((r) => r.status !== 'open');
  if (!closed.length) return;
  localSet(LOCAL_ARCHIVE, localGet(LOCAL_ARCHIVE, []).concat(closed));
  downloadJson({ archivedAt: new Date().toISOString(), requests: closed }, `request-tracker_archive_${today()}.json`);
  state.data.requests = state.data.requests.filter((r) => r.status === 'open');
  save().then(() => { toast(`${plural(closed.length, 'request', 'requests')} archived`); render(); });
}
function importJson(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const obj = JSON.parse(reader.result);
      if (!Array.isArray(obj.requests)) throw new Error('missing requests');
      const known = new Set(state.data.requests.map((r) => r.id));
      const fresh = obj.requests.filter((r) => !known.has(r.id));
      state.data.requests.push(...fresh);
      if (obj.settings) state.data.settings = { ...state.data.settings, ...obj.settings };
      migrate(state.data);
      save().then((ok) => { if (ok) { toast(`${plural(fresh.length, 'request', 'requests')} imported`); render(); } });
    } catch (e) { toast('Invalid file: ' + e.message, true); }
  };
  reader.readAsText(file);
}

/* ---------------- Boot ---------------- */

function pickInitialView() {
  if (!state.info) return 'todo';
  const a = analyzeItem(state.info);
  // Tracked email or a new email worth tracking -> email tab; otherwise the to-do list
  if (a.clientOf.length || a.supplierHits.length) return 'email';
  return totalDue() ? 'todo' : 'email';
}

function boot() {
  document.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => go(t.dataset.view)));
  loadData();
  state.info = readItem();
  state.view = pickInitialView();
  render();
  if (state.inOutlook && Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      state.info = readItem();
      if (state.view === 'email' || state.view === 'todo') state.view = pickInitialView();
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
