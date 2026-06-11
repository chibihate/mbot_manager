// app.js — MBot Manager web frontend

const API = '';
let mbots   = [];
let selected = new Set();
let accounts = [];

// ---------------------------------------------------------------------------
// WebSocket
// ---------------------------------------------------------------------------
function connectWS() {
  const proto = location.protocol === 'https:' ? 'wss' : 'ws';
  const ws    = new WebSocket(`${proto}://${location.host}/ws`);

  ws.onopen = () => {
    document.getElementById('ws-status').textContent = '⬤ connected';
    document.getElementById('ws-status').style.color = '#6dc28a';
  };

  ws.onmessage = (e) => {
    const data = JSON.parse(e.data);
    if (data.mbots) renderMbots(data.mbots);
    if (data.logs)  renderLogs(data.logs);
  };

  ws.onclose = () => {
    document.getElementById('ws-status').textContent = '⬤ disconnected';
    document.getElementById('ws-status').style.color = '#d35d5d';
    setTimeout(connectWS, 3000);
  };
}

// ---------------------------------------------------------------------------
// Mbot cards
// ---------------------------------------------------------------------------
function renderMbots(data) {
  mbots = data;
  const grid = document.getElementById('mbot-grid');

  const online = data.filter(m => !m.is_dc).length;
  document.getElementById('online-pill').textContent = `● ${online} online`;

  // Update or create cards
  const existing = {};
  grid.querySelectorAll('.mbot-card').forEach(el => {
    existing[el.dataset.id] = el;
  });

  data.forEach(m => {
    let card = existing[m.id];
    if (!card) {
      card = document.createElement('div');
      card.className = 'mbot-card';
      card.dataset.id = m.id;
      card.innerHTML = cardHtml(m);
      card.addEventListener('click', () => toggleSelect(m.id, card));
      grid.appendChild(card);
    } else {
      card.innerHTML = cardHtml(m);
    }
    card.classList.toggle('selected', selected.has(m.id));
    card.classList.toggle('dc', !!m.is_dc);
  });

  // Remove stale cards
  const ids = new Set(data.map(m => m.id));
  Object.entries(existing).forEach(([id, el]) => {
    if (!ids.has(Number(id))) el.remove();
  });

  updateSelPill();
  updateChatMbotSelect();
}

function cardHtml(m) {
  const hp  = m.is_dc ? 0 : Math.round(m.hp);
  const mp  = m.is_dc ? 0 : Math.round(m.mp);
  const kph = m.kph || '–';
  const status = m.is_dc ? 'DC' : 'online';
  const statusColor = m.is_dc ? '#d35d5d' : '#6dc28a';
  return `
    <div class="flex items-center gap-2 mb-2">
      <span class="font-bold text-sm">${esc(m.char)}</span>
      <span class="ml-auto text-xs" style="color:${statusColor}">${status}</span>
      <span class="text-xs text-mute">${esc(m.kph)} K/h</span>
    </div>
    <div class="flex items-center gap-1 mb-1 text-xs text-mute">
      <span class="w-5">HP</span>
      <div class="flex-1 bg-deep rounded h-1.5 overflow-hidden border border-border">
        <div class="bar h-full rounded" style="width:${hp}%;background:#6dc28a"></div>
      </div>
      <span class="w-8 text-right">${hp}%</span>
    </div>
    <div class="flex items-center gap-1 text-xs text-mute">
      <span class="w-5">MP</span>
      <div class="flex-1 bg-deep rounded h-1.5 overflow-hidden border border-border">
        <div class="bar h-full rounded" style="width:${mp}%;background:#5a8fd6"></div>
      </div>
      <span class="w-8 text-right">${mp}%</span>
    </div>
  `;
}

function toggleSelect(id, card) {
  if (selected.has(id)) {
    selected.delete(id);
    card.classList.remove('selected');
  } else {
    selected.add(id);
    card.classList.add('selected');
  }
  updateSelPill();
}

function updateSelPill() {
  document.getElementById('sel-pill').textContent = `${selected.size} selected`;
}

// ---------------------------------------------------------------------------
// Actions
// ---------------------------------------------------------------------------
document.querySelectorAll('.action-btn[data-action]').forEach(btn => {
  btn.addEventListener('click', () => {
    const action = btn.dataset.action;
    if (action === 'refresh') {
      fetch('/api/mbots').then(r => r.json()).then(renderMbots);
      return;
    }
    const targets = selected.size > 0 ? [...selected] : [null];
    targets.forEach(id => sendCommand(action, id));
  });
});

function sendCommand(action, targetId = null, params = null) {
  return fetch(`${API}/api/command`, {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify({ action, target_id: targetId, params }),
  }).then(r => r.json());
}

// ---------------------------------------------------------------------------
// Accounts
// ---------------------------------------------------------------------------
async function loadAccounts() {
  const res  = await fetch('/api/accounts');
  accounts   = await res.json();
  const list = document.getElementById('account-list');
  list.innerHTML = '';
  accounts.forEach((acc, i) => {
    const row = document.createElement('div');
    row.className = 'flex items-center gap-2 bg-panel border border-border rounded px-3 py-2';
    row.innerHTML = `
      <input type="checkbox" class="acc-cb" data-idx="${i}" />
      <span class="font-bold">${esc(acc.username)}</span>
      <span class="text-mute">${esc(acc.character || '')}</span>
      <span class="ml-auto text-mute text-xs truncate max-w-48" title="${esc(acc.mbot_file_path || '')}">${esc(acc.mbot_file_path || '')}</span>
    `;
    list.appendChild(row);
  });
}

function selectedAccountIndices() {
  return [...document.querySelectorAll('.acc-cb:checked')].map(cb => Number(cb.dataset.idx));
}

function loginSelected() {
  const indices = selectedAccountIndices();
  sendCommand('login', null, indices.length ? { indices } : null);
}

function hideAllMbots() {
  const indices = selectedAccountIndices();
  sendCommand('hide_mbots', null, indices.length ? { indices } : null);
}

// ---------------------------------------------------------------------------
// Chat
// ---------------------------------------------------------------------------
function updateChatMbotSelect() {
  const sel = document.getElementById('chat-mbot');
  const cur = sel.value;
  sel.innerHTML = mbots.map(m =>
    `<option value="${m.id}">${esc(m.char)}</option>`
  ).join('');
  if (cur) sel.value = cur;
}

async function loadChat() {
  const id  = document.getElementById('chat-mbot').value;
  const ch  = document.getElementById('chat-channel').value;
  if (!id) return;
  const res = await fetch(`/api/chat/${id}/${encodeURIComponent(ch)}`);
  const el  = document.getElementById('chat-content');
  if (res.ok) {
    const data = await res.json();
    el.textContent = data.content;
  } else {
    el.textContent = '(no data)';
  }
}

// ---------------------------------------------------------------------------
// Logs
// ---------------------------------------------------------------------------
const KIND_COLOR = {
  ok:     '#6dc28a',
  err:    '#d35d5d',
  warn:   '#d6b35a',
  accent: '#7c6af7',
  info:   '#8888aa',
};

function renderLogs(logs) {
  const el = document.getElementById('log-content');
  const atBottom = el.scrollHeight - el.scrollTop - el.clientHeight < 40;
  el.innerHTML = [...logs].reverse().map(l => {
    const color = KIND_COLOR[l.kind] || '#8888aa';
    return `<div><span style="color:#44445a">${esc(l.ts)}</span> <span style="color:${color}">${esc(l.msg)}</span></div>`;
  }).join('');
  if (atBottom) el.scrollTop = el.scrollHeight;
}

// ---------------------------------------------------------------------------
// Tabs
// ---------------------------------------------------------------------------
document.querySelectorAll('.nav-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const tab = btn.dataset.tab;
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(t => t.classList.add('hidden'));
    btn.classList.add('active');
    document.getElementById(`tab-${tab}`).classList.remove('hidden');
    if (tab === 'accounts') loadAccounts();
  });
});

// ---------------------------------------------------------------------------
// Utils
// ---------------------------------------------------------------------------
function esc(s) {
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------
connectWS();
fetch('/api/mbots').then(r => r.json()).then(renderMbots);
fetch('/api/logs').then(r => r.json()).then(renderLogs);
