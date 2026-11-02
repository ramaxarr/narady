// Задачі з нарад: картка задачі, редагування, сортування, текст для тижневого звіту.

import { h, toast, modal, confirmDialog, fmtDateShort, ymd, startOfWeek, addDays } from './ui.js';
import { call, peopleSync } from './api.js';

export const todayYmd = () => ymd(new Date());
export const isOverdue = (t) => t.status !== 'done' && !!t.dueDate && t.dueDate < todayYmd();

export function fmtDue(d) {
  const [y, m, dd] = String(d).split('-').map(Number);
  return new Intl.DateTimeFormat('uk-UA', { day: 'numeric', month: 'short' }).format(new Date(y, m - 1, dd));
}

export function sortTasks(list) {
  const rank = (t) => t.status === 'done' ? 3 : isOverdue(t) ? 0 : t.dueDate ? 1 : 2;
  return list.slice().sort((a, b) => {
    const r = rank(a) - rank(b);
    if (r) return r;
    if (rank(a) === 3) return String(b.doneAt || '').localeCompare(String(a.doneAt || ''));
    if (a.dueDate && b.dueDate && a.dueDate !== b.dueDate) return a.dueDate.localeCompare(b.dueDate);
    return String(b.meetingStart || '').localeCompare(String(a.meetingStart || ''));
  });
}

function dueLabel(t) {
  if (t.dueDate) return h('span', { class: 'due' + (isOverdue(t) ? ' late' : '') }, (isOverdue(t) ? 'прострочено, було до ' : 'до ') + fmtDue(t.dueDate));
  if (t.due) return h('span', { class: 'due muted' }, t.due);
  return h('span', { class: 'due muted' }, 'без терміну');
}

/** Картка задачі з галочкою «виконано». Натискання на текст відкриває редагування. */
export function taskCard(t, { onChange, showMeeting = true, suggestedDone = false } = {}) {
  let card;
  const cb = h('input', {
    type: 'checkbox', checked: t.status === 'done', 'aria-label': 'Виконано',
    onchange: async () => {
      const prev = t.status;
      t.status = cb.checked ? 'done' : 'open';
      card.classList.toggle('done', cb.checked);
      try {
        const r = await call('saveTask', { task: { id: t.id, status: t.status } });
        Object.assign(t, r);
        card.classList.toggle('overdue', isOverdue(t));
        if (onChange) onChange(t, 'status');
      } catch (e) {
        t.status = prev; cb.checked = prev === 'done'; card.classList.toggle('done', prev === 'done');
        toast(e.message, 'err');
      }
    }
  });
  card = h('div', { class: 'card task' + (t.status === 'done' ? ' done' : '') + (isOverdue(t) ? ' overdue' : '') },
    h('label', { class: 'task-check' }, cb),
    h('div', { class: 'grow task-body', role: 'button', tabindex: '0', onclick: () => editTask(t, onChange) },
      h('div', { class: 'task-text' }, t.task),
      h('div', { class: 'task-meta' },
        h('span', { class: 'chip small' }, t.owner || '—'),
        dueLabel(t),
        suggestedDone && t.status !== 'done' ? h('span', { class: 'badge ok' }, 'схоже, виконано — за розмовою') : null,
        showMeeting && t.meetingId ? h('a', {
          class: 'mlink', href: '#/m/' + encodeURIComponent(t.meetingId), onclick: (e) => e.stopPropagation()
        }, '⤷ ' + (t.meetingTitle || 'нарада') + (t.meetingStart ? ', ' + fmtDateShort(t.meetingStart) : '')) : null)));
  return card;
}

/** Створення (t = null) або редагування задачі. */
export function editTask(t, onChange) {
  const people = peopleSync();
  const listId = 'people-dl-' + Date.now();
  const text = h('textarea', { class: 'input', rows: 3, placeholder: 'Що зробити' }, t ? t.task : '');
  const owner = h('input', { class: 'input', list: listId, value: t ? (t.owner === '—' ? '' : t.owner || '') : '', placeholder: 'Хто відповідає' });
  const dl = h('datalist', { id: listId }, people.map(p => h('option', { value: p.name })));
  const due = h('input', { class: 'input', type: 'date', value: t && t.dueDate ? t.dueDate : '' });
  const done = h('input', { type: 'checkbox', checked: !!(t && t.status === 'done') });
  const actions = [];
  if (t) actions.push({
    label: 'Видалити', kind: 'danger ghost', onClick: async () => {
      if (!await confirmDialog('Видалити задачу?', 'Видалити', true)) return false;
      try { await call('deleteTask', { id: t.id }); if (onChange) onChange(t, 'deleted'); }
      catch (e) { toast(e.message, 'err'); return false; }
    }
  });
  actions.push({ label: 'Скасувати' });
  actions.push({
    label: 'Зберегти', kind: 'primary', onClick: async () => {
      if (!text.value.trim()) { text.focus(); return false; }
      const payload = { task: text.value.trim(), owner: owner.value.trim() || '—', dueDate: due.value || '', status: done.checked ? 'done' : 'open' };
      if (t) payload.id = t.id;
      try {
        const r = await call('saveTask', { task: payload });
        if (t) Object.assign(t, r);
        if (onChange) onChange(r, t ? 'edited' : 'created');
      } catch (e) { toast(e.message, 'err'); return false; }
    }
  });
  modal({
    title: t ? 'Задача' : 'Нова задача',
    body: h('div', { class: 'form' },
      h('label', { class: 'field' }, 'Що зробити', text),
      h('label', { class: 'field' }, 'Хто відповідає', owner, dl),
      h('label', { class: 'field' }, 'Термін', due),
      t && t.due && !t.dueDate ? h('p', { class: 'hint' }, `На нараді сказали: «${t.due}»`) : null,
      h('label', { class: 'check' }, done, ' Виконано'),
      t && t.meetingTitle ? h('p', { class: 'hint' }, `З наради: ${t.meetingTitle}, ${fmtDateShort(t.meetingStart)}`) : null),
    actions
  });
}

/** Текст для тижневого звіту: виконане за цей тиждень, прострочене, у роботі. */
export function weeklyReport(tasks) {
  const ws = startOfWeek(new Date()), we = addDays(ws, 7);
  const f = (d) => new Intl.DateTimeFormat('uk-UA', { day: 'numeric', month: 'long' }).format(d);
  const line = (t) => `• ${t.task} — ${t.owner || '—'}${t.dueDate ? ` (до ${fmtDue(t.dueDate)})` : ''}`;
  const done = tasks.filter(t => t.status === 'done' && t.doneAt && new Date(t.doneAt) >= ws && new Date(t.doneAt) < we);
  const late = tasks.filter(isOverdue);
  const open = tasks.filter(t => t.status !== 'done' && !isOverdue(t));
  const out = [`Задачі з нарад, тиждень ${f(ws)} – ${f(addDays(ws, 6))}`];
  out.push('', `Виконано за тиждень (${done.length}):`, ...(done.length ? done.map(line) : ['—']));
  out.push('', `Прострочено (${late.length}):`, ...(late.length ? sortTasks(late).map(line) : ['—']));
  out.push('', `У роботі (${open.length}):`, ...(open.length ? sortTasks(open).map(line) : ['—']));
  return out.join('\n');
}

export function tasksText(tasks) {
  return tasks.map((t, i) => `${i + 1}. ${t.task} — ${t.owner || '—'}${t.dueDate ? ' (до ' + fmtDue(t.dueDate) + ')' : t.due ? ' (' + t.due + ')' : ''}`).join('\n');
}
