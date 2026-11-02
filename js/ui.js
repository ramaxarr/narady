// Дрібні помічники для інтерфейсу: створення елементів, тости, модальні вікна, формати.

export function h(tag, attrs, ...children) {
  const el = document.createElement(tag);
  if (attrs) {
    for (const [k, v] of Object.entries(attrs)) {
      if (v === null || v === undefined || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
      else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
      else if (k === 'dataset') Object.assign(el.dataset, v);
      else if (k === 'html') el.innerHTML = v;
      else if (k in el && typeof v !== 'string') el[k] = v;
      else el.setAttribute(k, v === true ? '' : v);
    }
  }
  append(el, children);
  return el;
}

function append(el, children) {
  for (const c of children) {
    if (c === null || c === undefined || c === false) continue;
    if (Array.isArray(c)) append(el, c);
    else if (c instanceof Node) el.appendChild(c);
    else el.appendChild(document.createTextNode(String(c)));
  }
}

// Безпечне додавання дітей: пропускає null/false (рідний append перетворив би їх на текст «null»).
export function put(el, ...children) { append(el, children); return el; }

export const $ = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

export function clear(el) { while (el.firstChild) el.removeChild(el.firstChild); return el; }

export function icon(name) {
  const paths = {
    calendar: 'M7 2v2H5a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2h-2V2h-2v2H9V2zM5 9h14v10H5z',
    list: 'M4 6h16v2H4zm0 5h16v2H4zm0 5h10v2H4z',
    people: 'M16 11a3 3 0 1 0-3-3 3 3 0 0 0 3 3m-8 0a3 3 0 1 0-3-3 3 3 0 0 0 3 3m0 2c-2.3 0-7 1.2-7 3.5V19h14v-2.5C15 14.2 10.3 13 8 13m8 0c-.3 0-.6 0-1 .1a4.4 4.4 0 0 1 2 3.4V19h6v-2.5c0-2.3-4.7-3.5-7-3.5',
    settings: 'M19.4 13a7.6 7.6 0 0 0 0-2l2.1-1.6-2-3.5-2.5 1a7.3 7.3 0 0 0-1.7-1L15 3h-4l-.4 2.9a7.3 7.3 0 0 0-1.7 1l-2.5-1-2 3.5L6.6 11a7.6 7.6 0 0 0 0 2l-2.1 1.6 2 3.5 2.5-1a7.3 7.3 0 0 0 1.7 1L11 21h4l.4-2.9a7.3 7.3 0 0 0 1.7-1l2.5 1 2-3.5zM13 15.5a3.5 3.5 0 1 1 3.5-3.5 3.5 3.5 0 0 1-3.5 3.5',
    mic: 'M12 14a3 3 0 0 0 3-3V5a3 3 0 0 0-6 0v6a3 3 0 0 0 3 3m5-3a5 5 0 0 1-10 0H5a7 7 0 0 0 6 6.9V21h2v-3.1a7 7 0 0 0 6-6.9z',
    upload: 'M5 20h14v-2H5zm7-16-5 5h3v6h4V9h3z',
    plus: 'M11 5h2v6h6v2h-6v6h-2v-6H5v-2h6z',
    back: 'M15.4 7.4 14 6l-6 6 6 6 1.4-1.4L10.8 12z',
    left: 'M15.4 7.4 14 6l-6 6 6 6 1.4-1.4L10.8 12z',
    right: 'M8.6 16.6 10 18l6-6-6-6-1.4 1.4 4.6 4.6z',
    edit: 'M3 17.2V21h3.8l11-11-3.8-3.8zM20.7 7a1 1 0 0 0 0-1.4l-2.3-2.3a1 1 0 0 0-1.4 0l-1.8 1.8 3.8 3.8z',
    trash: 'M6 19a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2V7H6zM19 4h-3.5l-1-1h-5l-1 1H5v2h14z',
    play: 'M8 5v14l11-7z',
    pause: 'M6 19h4V5H6zm8-14v14h4V5z',
    stop: 'M6 6h12v12H6z',
    doc: 'M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8zm-1 7V3.5L18.5 9zM8 13h8v2H8zm0 4h8v2H8z',
    copy: 'M16 1H4a2 2 0 0 0-2 2v14h2V3h12zm3 4H8a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h11a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2m0 16H8V7h11z',
    wave: 'M3 10h2v4H3zm4-3h2v10H7zm4-4h2v18h-2zm4 5h2v8h-2zm4 2h2v4h-2z',
    check: 'M9 16.2 4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z',
    x: 'M19 6.4 17.6 5 12 10.6 6.4 5 5 6.4 10.6 12 5 17.6 6.4 19 12 13.4 17.6 19 19 17.6 13.4 12z',
    open: 'M14 3v2h3.6l-9.8 9.8 1.4 1.4L19 6.4V10h2V3zm5 16H5V5h7V3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7h-2z'
  };
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('class', 'ic');
  svg.setAttribute('aria-hidden', 'true');
  const p = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  p.setAttribute('d', paths[name] || '');
  svg.appendChild(p);
  return svg;
}

let toastBox;
export function toast(msg, kind = 'info', ms = 4000) {
  if (!toastBox) { toastBox = h('div', { class: 'toasts', role: 'status', 'aria-live': 'polite' }); document.body.appendChild(toastBox); }
  const t = h('div', { class: 'toast ' + kind }, msg);
  toastBox.appendChild(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 300); }, ms);
}

export function modal({ title, body, actions = [], wide = false, onClose }) {
  const back = h('div', { class: 'modal-back' });
  const close = () => { back.remove(); document.removeEventListener('keydown', onKey); if (onClose) onClose(); };
  const onKey = (e) => { if (e.key === 'Escape') close(); };
  const box = h('div', { class: 'modal' + (wide ? ' wide' : ''), role: 'dialog', 'aria-modal': 'true' },
    h('div', { class: 'modal-head' }, h('h2', null, title),
      h('button', { class: 'icon-btn', 'aria-label': 'Закрити', onclick: close }, icon('x'))),
    h('div', { class: 'modal-body' }, body),
    actions.length ? h('div', { class: 'modal-actions' }, actions.map(a =>
      h('button', {
        class: 'btn ' + (a.kind || ''), disabled: a.disabled, onclick: async (e) => {
          if (a.onClick) { const r = await a.onClick(e); if (r === false) return; }
          close();
        }
      }, a.label))) : null);
  back.appendChild(box);
  back.addEventListener('click', (e) => { if (e.target === back) close(); });
  document.addEventListener('keydown', onKey);
  document.body.appendChild(back);
  const first = box.querySelector('input, select, textarea');
  if (first) setTimeout(() => first.focus(), 50);
  return { close, el: box };
}

export function confirmDialog(text, okLabel = 'Так', danger = false) {
  return new Promise((resolve) => {
    let done = false;
    modal({
      title: 'Підтвердження',
      body: h('p', null, text),
      actions: [
        { label: 'Скасувати', onClick: () => { done = true; resolve(false); } },
        { label: okLabel, kind: danger ? 'danger' : 'primary', onClick: () => { done = true; resolve(true); } }
      ],
      onClose: () => { if (!done) resolve(false); }
    });
  });
}

export function fmtDuration(sec) {
  sec = Math.max(0, Math.round(sec || 0));
  const hh = Math.floor(sec / 3600), mm = Math.floor((sec % 3600) / 60), ss = sec % 60;
  const p = (n) => String(n).padStart(2, '0');
  return hh ? `${hh}:${p(mm)}:${p(ss)}` : `${p(mm)}:${p(ss)}`;
}

export function fmtHuman(sec) {
  sec = Math.round(sec || 0);
  const hh = Math.floor(sec / 3600), mm = Math.round((sec % 3600) / 60);
  if (hh) return `${hh} год ${mm} хв`;
  if (sec < 60) return `${sec} с`;
  return `${mm} хв`;
}

const dtf = (o) => new Intl.DateTimeFormat('uk-UA', o);
export const fmtTime = (d) => dtf({ hour: '2-digit', minute: '2-digit' }).format(new Date(d));
export const fmtDay = (d) => dtf({ weekday: 'long', day: 'numeric', month: 'long' }).format(new Date(d));
export const fmtDateShort = (d) => dtf({ day: 'numeric', month: 'short' }).format(new Date(d));
export const fmtDateFull = (d) => dtf({ day: 'numeric', month: 'long', year: 'numeric' }).format(new Date(d));
export const fmtWeekday = (d) => dtf({ weekday: 'short' }).format(new Date(d));

export function fmtBytes(b) {
  if (b < 1024) return b + ' Б';
  if (b < 1024 * 1024) return (b / 1024).toFixed(0) + ' КБ';
  return (b / 1024 / 1024).toFixed(1) + ' МБ';
}

export function ymd(d) {
  d = new Date(d);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
export function hm(d) {
  d = new Date(d);
  return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
}
export function startOfWeek(d) {
  const x = new Date(d); x.setHours(0, 0, 0, 0);
  const wd = (x.getDay() + 6) % 7; x.setDate(x.getDate() - wd); return x;
}
export function addDays(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
export function sameDay(a, b) { return ymd(a) === ymd(b); }

export function initials(name) {
  return String(name || '?').split(/\s+/).filter(Boolean).slice(0, 2).map(s => s[0].toUpperCase()).join('');
}

const PALETTE = ['#2f7d6d', '#3b6fb6', '#a35a2a', '#7a4fa3', '#b0435e', '#4f7f2f', '#8a6d1f', '#2f6f8a'];
export function colorFor(key) {
  let x = 0; const s = String(key || '');
  for (let i = 0; i < s.length; i++) x = (x * 31 + s.charCodeAt(i)) >>> 0;
  return PALETTE[x % PALETTE.length];
}

export async function copyText(text) {
  try { await navigator.clipboard.writeText(text); toast('Скопійовано', 'ok', 2000); }
  catch { toast('Не вдалося скопіювати', 'err'); }
}

export function spinner(label) { return h('div', { class: 'loading' }, h('span', { class: 'spin' }), label || 'Завантаження…'); }
