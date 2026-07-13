'use strict';

/* ═══════════════════════════════════════════════════════════
   CONSTANTS
═══════════════════════════════════════════════════════════ */
const EXPENSE_CATS = [
  { id: 'food',      label: 'Comida',         icon: '🍔' },
  { id: 'health',    label: 'Salud',          icon: '💊' },
  { id: 'fun',       label: 'Entretenimiento',icon: '🎬' },
  { id: 'shopping',  label: 'Shopping',       icon: '🛍️' },
  { id: 'transport', label: 'Transporte',     icon: '🚗' },
  { id: 'data',      label: 'Data',           icon: '📶' },
  { id: 'subs',      label: 'Membresías',     icon: '📱' },
  { id: 'travel',    label: 'Viajes',         icon: '✈️' },
  { id: 'invest',    label: 'Inversiones',    icon: '📈' },
  { id: 'bets',      label: 'Apuestas',       icon: '🎲' },
  { id: 'comision',  label: 'Comisiones',     icon: '💸' },
  { id: 'bankfee',   label: 'Intereses banco',icon: '🏦' },
  { id: 'other',     label: 'Otros',          icon: '📌' },
];

const INCOME_CATS = [
  { id: 'papa',      label: 'Depósito papá',  icon: '👨‍👦' },
  { id: 'salary',    label: 'Salario',        icon: '💼' },
  { id: 'freelance', label: 'Freelance',      icon: '💻' },
  { id: 'inv-in',    label: 'Inversión',      icon: '📈' },
  { id: 'bets-in',   label: 'Apuestas',       icon: '🎲' },
  { id: 'comision-in',label: 'Comisiones',    icon: '💸' },
  { id: 'interest',  label: 'Intereses banco',icon: '🏦' },
  { id: 'other-in',  label: 'Otros',          icon: '💰' },
];

const DEFAULT_ACCOUNTS = [
  { id: 'mercantil-bs',  name: 'Mercantil Bs',  type: 'debit',  currency: 'VES', emoji: '🏦' },
  { id: 'bofa-azul',     name: 'BofA Azul',     type: 'credit', currency: 'USD', emoji: '💳' },
  { id: 'bofa-gris',     name: 'BofA Gris',     type: 'credit', currency: 'USD', emoji: '💳' },
  { id: 'bofa-debit',    name: 'BofA Débito',   type: 'debit',  currency: 'USD', emoji: '💳' },
  { id: 'capital-one',   name: 'Capital One',   type: 'credit', currency: 'USD', emoji: '💳' },
];

const APPS_SCRIPT_CODE = `function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = {
    accounts: readSheet(ss, 'Accounts'),
    transactions: readSheet(ss, 'Transactions'),
  };
  return out(data);
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = JSON.parse(e.postData.contents);
  if (data.accounts)      writeSheet(ss, 'Accounts', data.accounts);
  if (data.transactions)  writeSheet(ss, 'Transactions', data.transactions);
  return out({ success: true });
}

function readSheet(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

function writeSheet(ss, name, data) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clearContents();
  if (!data || !data.length) return;
  const headers = Object.keys(data[0]);
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  const rows = data.map(o => headers.map(h => o[h] ?? ''));
  sheet.getRange(2,1,rows.length,headers.length).setValues(rows);
}

function out(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}`;

const PIE_COLORS = [
  '#E24B4A','#FF8C42','#F7C59F','#3B6D11','#6BAB5E',
  '#4A90D9','#9B59B6','#F39C12','#1ABC9C','#95A5A6',
];

/* ═══════════════════════════════════════════════════════════
   PIN AUTH
═══════════════════════════════════════════════════════════ */
const Auth = {
  _key: 'finanzas_pin_v1',

  async _hash(pin) {
    const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode('fp:' + pin));
    return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
  },

  hasPin()  { return !!localStorage.getItem(this._key); },
  clearPin(){ localStorage.removeItem(this._key); },

  async savePin(pin) {
    localStorage.setItem(this._key, await this._hash(pin));
  },

  async verify(pin) {
    const stored = localStorage.getItem(this._key);
    if (!stored) return true;
    return stored === await this._hash(pin);
  },

  showScreen(mode, onSuccess) {
    const existing = document.getElementById('pin-screen');
    if (existing) existing.remove();

    const el = document.createElement('div');
    el.id = 'pin-screen';
    el.innerHTML = `
      <div class="pin-container">
        <div class="pin-logo">💰</div>
        <div class="pin-subtitle">${mode === 'setup' ? 'Crea tu PIN de acceso' : mode === 'confirm' ? 'Confirma tu PIN' : 'Ingresa tu PIN'}</div>
        <div class="pin-dots" id="pin-dots">
          <span class="pin-dot"></span><span class="pin-dot"></span>
          <span class="pin-dot"></span><span class="pin-dot"></span>
        </div>
        <div class="pin-error" id="pin-error"></div>
        <div class="pin-keypad">
          ${[1,2,3,4,5,6,7,8,9,'',0,'⌫'].map(k =>
            k === '' ? '<div></div>' :
            `<button class="pin-key" data-k="${k}">${k}</button>`
          ).join('')}
        </div>
        ${mode === 'unlock' && this.hasPin() ? '<button class="pin-skip" id="pin-forgot">¿Olvidaste el PIN? Borrar datos</button>' : ''}
      </div>`;
    document.body.appendChild(el);

    let entered = '';
    let firstPin = '';

    const dots = el.querySelectorAll('.pin-dot');
    const errEl = el.querySelector('#pin-error');

    const updateDots = () => dots.forEach((d, i) => d.classList.toggle('filled', i < entered.length));

    const reset = (msg = '') => {
      entered = '';
      updateDots();
      errEl.textContent = msg;
    };

    const submit = async () => {
      if (mode === 'unlock') {
        if (await this.verify(entered)) {
          el.remove();
          onSuccess();
        } else {
          reset('PIN incorrecto');
        }
      } else if (mode === 'setup') {
        firstPin = entered;
        reset();
        // Switch to confirm mode
        el.querySelector('.pin-subtitle').textContent = 'Confirma tu PIN';
        mode = 'confirm';
      } else if (mode === 'confirm') {
        if (entered === firstPin) {
          await this.savePin(entered);
          el.remove();
          onSuccess();
        } else {
          reset('Los PINs no coinciden');
          mode = 'setup';
          firstPin = '';
          el.querySelector('.pin-subtitle').textContent = 'Crea tu PIN de acceso';
        }
      }
    };

    el.querySelectorAll('.pin-key').forEach(btn => {
      btn.addEventListener('click', async () => {
        const k = btn.dataset.k;
        if (k === '⌫') {
          entered = entered.slice(0, -1);
          updateDots();
        } else if (entered.length < 4) {
          entered += k;
          updateDots();
          if (entered.length === 4) await submit();
        }
      });
    });

    el.querySelector('#pin-forgot')?.addEventListener('click', () => {
      if (confirm('¿Borrar el PIN? Solo se eliminará el PIN, no tus datos.')) {
        this.clearPin();
        el.remove();
        onSuccess();
      }
    });
  },

  async gate(onSuccess) {
    if (!this.hasPin()) {
      this.showScreen('setup', onSuccess);
    } else {
      this.showScreen('unlock', onSuccess);
    }
  },
};

/* ═══════════════════════════════════════════════════════════
   STATE  (localStorage)
═══════════════════════════════════════════════════════════ */
const State = {
  _key: 'finanzas_data_v1',

  _defaults() {
    return {
      accounts: JSON.parse(JSON.stringify(DEFAULT_ACCOUNTS)),
      transactions: [],
      config: { sheetsUrl: '', lastSync: null },
      rates: { bcvUsd: null, bcvEur: null, paralelo: null, updatedAt: null },
    };
  },

  load() {
    try {
      const raw = localStorage.getItem(this._key);
      if (!raw) return this._defaults();
      return { ...this._defaults(), ...JSON.parse(raw) };
    } catch {
      return this._defaults();
    }
  },

  save(state) {
    localStorage.setItem(this._key, JSON.stringify(state));
  },
};

/* ═══════════════════════════════════════════════════════════
   APP
═══════════════════════════════════════════════════════════ */
const App = {
  state: null,
  charts: { pie: null, bar: null },
  currentTab: 'inicio',
  currentTxType: 'expense',
  selectedAccount: null,
  selectedCategory: null,
  txFilterAccount: null,
  txFilterDateFrom: null,
  txFilterDateTo: null,
  txFilterType: null,
  txFilterCategory: null,
  txFilterText: '',

  // ── Init ────────────────────────────────────────────────
  async init() {
    this.state = State.load();
    this.setupTabs();
    this.setupForm();
    this.setupConfig();
    this.renderAll();
    this.setDefaultDate();
    document.getElementById('btn-sync-header').addEventListener('click', () => this.syncAll());
    document.getElementById('btn-reload').addEventListener('click', () => window.location.reload());
    document.getElementById('btn-calc').addEventListener('click', () => this.showCalculator());
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('/sw.js').catch(() => {});
    }
    // Auto-pull from Sheets on app open
    await this.pullFromSheets({ silent: true });
    await this.fetchRates();
    // Auto-pull: events + polling every 60s (cubre iOS PWA que no dispara eventos)
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible') this.pullFromSheets({ silent: true });
    });
    window.addEventListener('pageshow', (e) => {
      if (e.persisted) this.pullFromSheets({ silent: true });
    });
    window.addEventListener('focus', () => this.pullFromSheets({ silent: true }));
    setInterval(() => this.pullFromSheets({ silent: true }), 60000);
  },

  renderAll() {
    this.renderRates();
    this.renderDashboard();
    this.renderAccountsTab();
    this.renderReports();
    this.renderConfig();
    this.renderFormAccounts();
  },

  // ── Tabs ─────────────────────────────────────────────────
  setupTabs() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => this.switchTab(btn.dataset.tab));
    });
  },

  switchTab(tab) {
    this.currentTab = tab;
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('active', p.id === `tab-${tab}`));
    if (tab === 'reportes') this.renderReports();
  },

  // ── Exchange Rates ────────────────────────────────────────
  async fetchRates() {
    const btn = document.getElementById('btn-sync-header');
    btn.querySelector('svg').classList.add('spinning');
    try {
      const [usd, eur, par] = await Promise.all([
        this.fetchRate('oficial'),
        this.fetchRate('euro'),
        this.fetchRate('paralelo'),
      ]);
      this.state.rates = {
        bcvUsd: usd?.promedio ?? null,
        bcvEur: eur?.promedio ?? null,
        paralelo: par?.promedio ?? null,
        updatedAt: new Date().toISOString(),
      };
      State.save(this.state);
      this.renderRates();
      this.updateUsdEquiv();
    } catch {
      this.toast('No se pudieron cargar las tasas');
    } finally {
      btn.querySelector('svg').classList.remove('spinning');
    }
  },

  async fetchRate(type) {
    const res = await fetch(`/api/tasas?type=${type}`);
    if (!res.ok) throw new Error(`${res.status}`);
    return res.json();
  },

  renderRates() {
    const r = this.state.rates;
    const fmt = v => v ? `Bs ${v.toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : '--';
    document.querySelector('#rate-bcv .rate-val').textContent = fmt(r.bcvUsd);
    document.querySelector('#rate-eur .rate-val').textContent = fmt(r.bcvEur);
    document.querySelector('#rate-par .rate-val').textContent = fmt(r.paralelo);
  },

  // ── Helpers ───────────────────────────────────────────────
  toUSD(amount, currency) {
    const r = this.state.rates;
    switch (currency) {
      case 'USD':     return { usd: amount, rate: 1, rateType: 'usd' };
      case 'EUR':     return r.bcvEur ? { usd: amount * r.bcvEur / (r.bcvUsd || 1), rate: r.bcvEur, rateType: 'bcvEur' } : null;
      case 'VES_BCV': return r.bcvUsd ? { usd: amount / r.bcvUsd, rate: r.bcvUsd, rateType: 'bcvUsd' } : null;
      case 'VES_EUR': return r.bcvEur ? { usd: amount / r.bcvEur, rate: r.bcvEur, rateType: 'bcvEur' } : null;
      default: return null;
    }
  },

  fmtUSD(v) {
    return `$${Math.abs(v).toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  },

  fmtAmt(amount, currency) {
    const abs = Math.abs(amount);
    const fmtNum = (n, dec = 2) => n.toLocaleString('es-VE', { minimumFractionDigits: dec, maximumFractionDigits: dec });
    if (currency === 'USD') return `$${fmtNum(abs)}`;
    if (currency === 'EUR') return `€${fmtNum(abs)}`;
    if (currency === 'VES_BCV' || currency === 'VES_EUR' || currency === 'VES')
      return `Bs ${fmtNum(abs)}`;
    return fmtNum(abs);
  },

  fmtVES(v) {
    return `Bs ${Math.abs(v).toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  },

  uid() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  },

  getAccountBalance(accountId) {
    const acc = this.getAccountInfo(accountId);
    const isVES = acc?.currency === 'VES';
    let balance = 0;
    for (const tx of this.state.transactions) {
      // VES accounts: use original amount in Bs; others: use USD equivalent
      const val = isVES ? tx.amount : tx.amountUSD;
      if (tx.type === 'expense' && tx.account === accountId) balance -= val;
      if (tx.type === 'income'  && tx.account === accountId) balance += val;
      if (tx.type === 'transfer') {
        if (tx.account   === accountId) balance -= val;
        if (tx.toAccount === accountId) balance += val;
      }
    }
    return balance;
  },

  getCategoryInfo(catId, type) {
    const list = type === 'expense' ? EXPENSE_CATS : type === 'income' ? INCOME_CATS : [];
    return list.find(c => c.id === catId) || { label: catId, icon: '💸' };
  },

  getAccountInfo(id) {
    return this.state.accounts.find(a => a.id === id);
  },

  // ── Dashboard (Inicio) ────────────────────────────────────
  renderDashboard() {
    this.renderAccountsGrid();
    this.renderTotals();
    this.renderMonthSummary();
    this.renderRecentTx();
  },

  renderAccountsGrid() {
    const grid = document.getElementById('accounts-grid');
    grid.innerHTML = '';
    for (const acc of this.state.accounts) {
      const bal    = this.getAccountBalance(acc.id);
      const isVES  = acc.currency === 'VES';
      const rate   = this.state.rates.bcvUsd;
      const usdVal = isVES && rate ? bal / rate : bal;
      const isNeg  = usdVal < 0;
      const div = document.createElement('div');
      div.className = 'account-card';
      div.innerHTML = `
        <div class="acc-name">${acc.emoji || '💳'} ${acc.name}</div>
        <div class="acc-balance ${isNeg ? 'negative' : ''}">${this.fmtUSD(usdVal)}</div>
        <div class="acc-type">${isVES ? `${this.fmtVES(bal)} · ` : ''}${acc.type === 'credit' ? 'Crédito' : 'Débito'}</div>
      `;
      div.addEventListener('click', () => {
        this.switchTab('cuentas');
        this.txFilterAccount = acc.id;
        this.renderAccountsTab();
      });
      grid.appendChild(div);
    }
  },

  renderTotals() {
    let totalUSD = 0;
    let totalVES = 0;
    const rate = this.state.rates.bcvUsd;
    for (const acc of this.state.accounts) {
      const bal = this.getAccountBalance(acc.id);
      if (acc.currency === 'VES') {
        totalVES += bal;
        if (rate) totalUSD += bal / rate;
      } else {
        totalUSD += bal;
      }
    }
    document.getElementById('total-usd').textContent = this.fmtUSD(totalUSD);
    document.getElementById('total-ves').textContent = this.fmtVES(totalVES);
  },

  renderMonthSummary() {
    const now = new Date();
    const ym = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    let inc = 0, exp = 0;
    for (const tx of this.state.transactions) {
      if (!tx.date.startsWith(ym)) continue;
      if (tx.type === 'income')  inc += tx.amountUSD;
      if (tx.type === 'expense') exp += tx.amountUSD;
    }
    document.getElementById('month-income').textContent  = this.fmtUSD(inc);
    document.getElementById('month-expense').textContent = this.fmtUSD(exp);
  },

  renderRecentTx() {
    const list = document.getElementById('recent-list');
    const txs  = [...this.state.transactions].sort((a, b) => b.createdAt - a.createdAt).slice(0, 8);
    if (!txs.length) {
      list.innerHTML = '<div class="empty-state"><div class="empty-icon">📭</div>Sin transacciones aún</div>';
      return;
    }
    list.innerHTML = txs.map(tx => this.txItemHTML(tx)).join('');
    list.querySelectorAll('.tx-item').forEach((el, i) => {
      el.addEventListener('click', () => this.showTxDetail(txs[i]));
    });
  },

  txItemHTML(tx) {
    const cat = this.getCategoryInfo(tx.category, tx.type);
    const acc = this.getAccountInfo(tx.account);
    const meta = tx.type === 'transfer'
      ? `${acc?.name || '?'} → ${this.getAccountInfo(tx.toAccount)?.name || '?'}`
      : `${acc?.name || '?'} · ${tx.date}`;
    const label = tx.type === 'transfer' ? 'Transferencia' : (tx.note || cat.label);
    const sign  = tx.type === 'expense' ? '-' : tx.type === 'income' ? '+' : '';
    return `
      <div class="tx-item">
        <div class="tx-icon ${tx.type}">${cat.icon}</div>
        <div class="tx-info">
          <div class="tx-cat">${label}</div>
          <div class="tx-meta">${meta}</div>
        </div>
        <div class="tx-amount ${tx.type}">
          ${sign}${this.fmtAmt(tx.amount, tx.currency)}
          <span class="tx-usd">${sign}${this.fmtUSD(tx.amountUSD)}</span>
        </div>
      </div>`;
  },

  showTxDetail(tx) {
    const cat = this.getCategoryInfo(tx.category, tx.type);
    const acc = this.getAccountInfo(tx.account);
    const typeLabel = tx.type === 'expense' ? 'Gasto' : tx.type === 'income' ? 'Ingreso' : 'Transferencia';
    document.getElementById('modal-content').innerHTML = `
      <div class="modal-title">${cat.icon} ${cat.label}</div>
      <div class="modal-row"><span class="label">Tipo</span><span class="value">${typeLabel}</span></div>
      <div class="modal-row"><span class="label">Monto</span><span class="value">${this.fmtAmt(tx.amount, tx.currency)}</span></div>
      <div class="modal-row"><span class="label">En USD</span><span class="value">${this.fmtUSD(tx.amountUSD)}</span></div>
      ${tx.rate ? `<div class="modal-row"><span class="label">Tasa usada</span><span class="value">Bs ${tx.rate.toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span></div>` : ''}
      <div class="modal-row"><span class="label">Cuenta</span><span class="value">${acc?.name || '?'}</span></div>
      ${tx.toAccount ? `<div class="modal-row"><span class="label">Destino</span><span class="value">${this.getAccountInfo(tx.toAccount)?.name || '?'}</span></div>` : ''}
      <div class="modal-row"><span class="label">Fecha</span><span class="value">${tx.date}</span></div>
      ${tx.note ? `<div class="modal-row"><span class="label">Nota</span><span class="value">${tx.note}</span></div>` : ''}
      <div class="modal-actions">
        <button class="btn-danger" id="modal-delete">Eliminar</button>
        <button class="btn-secondary" id="modal-close">Cerrar</button>
      </div>`;
    document.getElementById('modal-delete').addEventListener('click', () => {
      if (confirm('¿Eliminar esta transacción?')) {
        this.state.transactions = this.state.transactions.filter(t => t.id !== tx.id);
        State.save(this.state);
        this.closeModal();
        this.renderAll();
        this.toast('Transacción eliminada');
        // Sync immediately so Sheets reflects the deletion before next pull
        if (this.state.config.sheetsUrl) this.syncAll({ silent: true });
      }
    });
    document.getElementById('modal-close').addEventListener('click', () => this.closeModal());
    this.openModal();
  },

  openModal() {
    document.getElementById('modal-overlay').classList.remove('hidden');
    document.getElementById('modal-overlay').addEventListener('click', e => {
      if (e.target === document.getElementById('modal-overlay')) this.closeModal();
    }, { once: true });
  },

  closeModal() {
    document.getElementById('modal-overlay').classList.add('hidden');
  },

  // ── Form (Registrar) ──────────────────────────────────────
  setupForm() {
    // Type buttons
    document.querySelectorAll('.type-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        this.currentTxType = btn.dataset.type;
        document.querySelectorAll('.type-btn').forEach(b => b.classList.toggle('active', b === btn));
        this.renderFormCategories();
        this.renderFormAccounts();
        const toGroup = document.getElementById('to-account-group');
        const catGroup = document.getElementById('category-group');
        toGroup.style.display = this.currentTxType === 'transfer' ? '' : 'none';
        catGroup.style.display = this.currentTxType === 'transfer' ? 'none' : '';
        this.selectedCategory = null;
      });
    });

    // Currency change → update equiv
    document.getElementById('f-currency').addEventListener('change', () => this.updateUsdEquiv());
    document.getElementById('f-amount').addEventListener('input', () => this.updateUsdEquiv());

    // Account quick buttons
    document.getElementById('account-quick-btns').addEventListener('click', e => {
      const btn = e.target.closest('.quick-btn');
      if (!btn) return;
      this.selectedAccount = btn.dataset.id;
      document.querySelectorAll('#account-quick-btns .quick-btn').forEach(b => b.classList.toggle('selected', b === btn));
      document.getElementById('f-account').value = this.selectedAccount;
    });

    document.getElementById('f-account').addEventListener('change', e => {
      this.selectedAccount = e.target.value;
      document.querySelectorAll('#account-quick-btns .quick-btn').forEach(b => b.classList.toggle('selected', b.dataset.id === this.selectedAccount));
    });

    // Form submit
    document.getElementById('tx-form').addEventListener('submit', e => {
      e.preventDefault();
      this.submitTransaction();
    });

    this.renderFormCategories();
    this.renderFormAccounts();
  },

  renderFormAccounts() {
    const qBtns = document.getElementById('account-quick-btns');
    const sel   = document.getElementById('f-account');
    const toSel = document.getElementById('f-to-account');

    qBtns.innerHTML = this.state.accounts.map(a =>
      `<button type="button" class="quick-btn${this.selectedAccount === a.id ? ' selected' : ''}" data-id="${a.id}">${a.emoji || '💳'} ${a.name}</button>`
    ).join('');

    const opts = this.state.accounts.map(a => `<option value="${a.id}">${a.name} (${a.currency === 'VES' ? 'Bs' : a.currency})</option>`).join('');
    sel.innerHTML   = `<option value="">Seleccionar cuenta…</option>${opts}`;
    toSel.innerHTML = `<option value="">Seleccionar destino…</option>${opts}`;

    if (this.selectedAccount) sel.value = this.selectedAccount;
  },

  renderFormCategories() {
    const cats = this.currentTxType === 'expense' ? EXPENSE_CATS : INCOME_CATS;
    const container = document.getElementById('category-chips');
    container.innerHTML = cats.map(c =>
      `<button type="button" class="cat-chip${this.selectedCategory === c.id ? ' selected' : ''}" data-id="${c.id}">${c.icon} ${c.label}</button>`
    ).join('');
    container.querySelectorAll('.cat-chip').forEach(chip => {
      chip.addEventListener('click', () => {
        this.selectedCategory = chip.dataset.id;
        container.querySelectorAll('.cat-chip').forEach(c => c.classList.toggle('selected', c === chip));
        document.getElementById('f-category').value = chip.dataset.id;
      });
    });
  },

  updateUsdEquiv() {
    const amount   = parseFloat(document.getElementById('f-amount').value) || 0;
    const currency = document.getElementById('f-currency').value;
    const result   = this.toUSD(amount, currency);
    const el       = document.getElementById('usd-equiv');
    const r        = this.state.rates;

    if (!result) {
      el.textContent = 'Tasa no disponible — actualiza las tasas';
      return;
    }
    const rateLabel = {
      bcvUsd: `BCV: Bs ${r.bcvUsd?.toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
      bcvEur: `BCV EUR: Bs ${r.bcvEur?.toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
      usd:    'Sin conversión',
    }[result.rateType] || '';
    el.textContent = `≈ ${this.fmtUSD(result.usd)} USD (${rateLabel})`;
  },

  setDefaultDate() {
    const today = new Date().toISOString().slice(0, 10);
    document.getElementById('f-date').value = today;
  },

  submitTransaction() {
    const type     = this.currentTxType;
    const amount   = parseFloat(document.getElementById('f-amount').value);
    const currency = document.getElementById('f-currency').value;
    const account  = document.getElementById('f-account').value;
    const toAcc    = document.getElementById('f-to-account').value;
    const note     = document.getElementById('f-note').value.trim();
    const date     = document.getElementById('f-date').value;
    const category = document.getElementById('f-category').value;

    if (!account) { this.toast('Selecciona una cuenta'); return; }
    if (!amount || amount <= 0) { this.toast('Ingresa un monto válido'); return; }
    if (type === 'transfer' && !toAcc) { this.toast('Selecciona cuenta destino'); return; }
    if (type !== 'transfer' && !category) { this.toast('Selecciona una categoría'); return; }
    if (type === 'transfer' && account === toAcc) { this.toast('Las cuentas deben ser distintas'); return; }

    const conv = this.toUSD(amount, currency);
    if (!conv) { this.toast('Tasa no disponible. Actualiza las tasas primero.'); return; }

    const tx = {
      id:        this.uid(),
      type,
      amount,
      currency,
      amountUSD: parseFloat(conv.usd.toFixed(6)),
      rate:      conv.rate,
      rateType:  conv.rateType,
      category:  category || 'transfer',
      account,
      toAccount: type === 'transfer' ? toAcc : null,
      note,
      date,
      createdAt: Date.now(),
    };

    this.state.transactions.push(tx);
    State.save(this.state);

    // Auto-sync in background if configured
    if (this.state.config.sheetsUrl) {
      this.syncAll({ silent: true });
    }

    // Reset form
    document.getElementById('f-amount').value = '';
    document.getElementById('f-note').value   = '';
    document.getElementById('f-category').value = '';
    this.selectedCategory = null;
    this.renderFormCategories();
    document.getElementById('usd-equiv').textContent = '≈ $0.00 USD';
    this.setDefaultDate();

    this.renderDashboard();
    this.renderAccountsTab();

    const label = type === 'expense' ? 'Gasto' : type === 'income' ? 'Ingreso' : 'Transferencia';
    this.toast(`✓ ${label} registrado`);
  },

  // ── Accounts tab ──────────────────────────────────────────
  renderAccountsTab() {
    const list = document.getElementById('accounts-list');
    list.innerHTML = this.state.accounts.map(acc => {
      const bal   = this.getAccountBalance(acc.id);
      const isNeg = bal < 0;
      return `
        <div class="account-detail-card">
          <div class="acc-emoji">${acc.emoji || '💳'}</div>
          <div class="acc-detail-info">
            <div class="acc-detail-name">${acc.name}</div>
            <div class="acc-detail-meta">${acc.type === 'credit' ? 'Crédito' : 'Débito'} · ${acc.currency === 'VES' ? 'Bolívares' : acc.currency}</div>
          </div>
          <div class="acc-detail-balance ${isNeg ? 'negative' : ''}">
            ${acc.currency === 'VES' && this.state.rates.bcvUsd
              ? `${this.fmtUSD(bal / this.state.rates.bcvUsd)}<small style="display:block;font-size:.72rem;font-weight:400;color:var(--text3)">${this.fmtVES(bal)}</small>`
              : this.fmtUSD(bal)}
          </div>
          <div class="acc-actions">
            <button class="acc-action-btn" data-action="filter" data-id="${acc.id}">Ver txs</button>
            ${!DEFAULT_ACCOUNTS.find(d => d.id === acc.id) ? `<button class="acc-action-btn" data-action="delete" data-id="${acc.id}">Eliminar</button>` : ''}
          </div>
        </div>`;
    }).join('');

    list.querySelectorAll('.acc-action-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        if (btn.dataset.action === 'filter') {
          this.txFilterAccount = this.txFilterAccount === id ? null : id;
          this.renderAccountsTab();
        } else if (btn.dataset.action === 'delete') {
          if (confirm('¿Eliminar esta cuenta y sus transacciones?')) {
            this.state.accounts = this.state.accounts.filter(a => a.id !== id);
            this.state.transactions = this.state.transactions.filter(t => t.account !== id && t.toAccount !== id);
            State.save(this.state);
            this.renderAll();
          }
        }
      });
    });

    document.getElementById('btn-add-account').addEventListener('click', () => this.showAddAccountModal(), { once: true });
    this.renderTxFilters();
    this.renderTxHistory();
  },

  renderTxFilters() {
    const existing = document.getElementById('tx-filters');
    if (existing) existing.remove();

    const now   = new Date();
    const today = now.toISOString().slice(0, 10);
    const ym    = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    const ymd   = (d) => d.toISOString().slice(0, 10);

    const quickRanges = [
      { label: 'Todo',        from: null,  to: null },
      { label: 'Hoy',         from: today, to: today },
      { label: 'Esta semana', from: ymd(new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay())), to: today },
      { label: 'Este mes',    from: `${ym(now)}-01`, to: today },
      { label: 'Mes anterior',from: `${ym(new Date(now.getFullYear(), now.getMonth() - 1, 1))}-01`,
                               to: `${ym(new Date(now.getFullYear(), now.getMonth(), 0))}-${String(new Date(now.getFullYear(), now.getMonth(), 0).getDate()).padStart(2,'0')}` },
      { label: 'Este año',    from: `${now.getFullYear()}-01-01`, to: today },
    ];

    const isQuickActive = (q) => this.txFilterDateFrom === q.from && this.txFilterDateTo === q.to;

    const allCats = [...EXPENSE_CATS, ...INCOME_CATS.filter(c => !EXPENSE_CATS.find(e => e.id === c.id))];

    const wrap = document.createElement('div');
    wrap.id = 'tx-filters';
    wrap.className = 'tx-filters-wrap';
    wrap.innerHTML = `
      <input type="search" id="f-tx-search" class="form-input" placeholder="🔍 Buscar por nota o categoría…" value="${this.txFilterText}" />

      <div class="filter-row">
        <button class="filter-chip${!this.txFilterType ? ' active' : ''}" data-ftype="">Todos</button>
        <button class="filter-chip${this.txFilterType === 'expense' ? ' active' : ''}" data-ftype="expense">Gastos</button>
        <button class="filter-chip${this.txFilterType === 'income' ? ' active' : ''}" data-ftype="income">Ingresos</button>
        <button class="filter-chip${this.txFilterType === 'transfer' ? ' active' : ''}" data-ftype="transfer">Transf.</button>
      </div>

      <div class="filter-row" id="quick-range-chips">
        ${quickRanges.map((q, i) => `<button class="filter-chip${isQuickActive(q) ? ' active' : ''}" data-qi="${i}">${q.label}</button>`).join('')}
      </div>

      <div class="date-range-row">
        <div class="date-range-group">
          <label>Desde</label>
          <input type="date" id="f-date-from" class="form-input" value="${this.txFilterDateFrom || ''}" />
        </div>
        <div class="date-range-sep">—</div>
        <div class="date-range-group">
          <label>Hasta</label>
          <input type="date" id="f-date-to" class="form-input" value="${this.txFilterDateTo || ''}" />
        </div>
      </div>

      <div class="filter-row" id="cat-chips-filter">
        <button class="filter-chip${!this.txFilterCategory ? ' active' : ''}" data-cat="">Todas</button>
        ${allCats.map(c => `<button class="filter-chip${this.txFilterCategory === c.id ? ' active' : ''}" data-cat="${c.id}">${c.icon} ${c.label}</button>`).join('')}
      </div>
    `;

    document.getElementById('tx-history').prepend(wrap);

    document.getElementById('f-tx-search').addEventListener('input', e => {
      this.txFilterText = e.target.value;
      this.renderTxHistory();
    });

    wrap.querySelectorAll('[data-ftype]').forEach(btn => {
      btn.addEventListener('click', () => {
        this.txFilterType = btn.dataset.ftype || null;
        wrap.querySelectorAll('[data-ftype]').forEach(b => b.classList.toggle('active', b === btn));
        this.renderTxHistory();
      });
    });

    wrap.querySelectorAll('[data-qi]').forEach(btn => {
      btn.addEventListener('click', () => {
        const q = quickRanges[parseInt(btn.dataset.qi)];
        this.txFilterDateFrom = q.from;
        this.txFilterDateTo   = q.to;
        document.getElementById('f-date-from').value = q.from || '';
        document.getElementById('f-date-to').value   = q.to   || '';
        wrap.querySelectorAll('[data-qi]').forEach(b => b.classList.toggle('active', b === btn));
        this.renderTxHistory();
      });
    });

    document.getElementById('f-date-from').addEventListener('change', e => {
      this.txFilterDateFrom = e.target.value || null;
      wrap.querySelectorAll('[data-qi]').forEach(b => b.classList.remove('active'));
      this.renderTxHistory();
    });

    document.getElementById('f-date-to').addEventListener('change', e => {
      this.txFilterDateTo = e.target.value || null;
      wrap.querySelectorAll('[data-qi]').forEach(b => b.classList.remove('active'));
      this.renderTxHistory();
    });

    wrap.querySelectorAll('[data-cat]').forEach(btn => {
      btn.addEventListener('click', () => {
        this.txFilterCategory = btn.dataset.cat || null;
        wrap.querySelectorAll('[data-cat]').forEach(b => b.classList.toggle('active', b === btn));
        this.renderTxHistory();
      });
    });
  },

  getFilteredTxs() {
    return this.state.transactions
      .filter(tx => {
        if (this.txFilterAccount && tx.account !== this.txFilterAccount && tx.toAccount !== this.txFilterAccount) return false;
        if (this.txFilterDateFrom && tx.date < this.txFilterDateFrom) return false;
        if (this.txFilterDateTo   && tx.date > this.txFilterDateTo)   return false;
        if (this.txFilterType && tx.type !== this.txFilterType) return false;
        if (this.txFilterCategory && tx.category !== this.txFilterCategory) return false;
        if (this.txFilterText) {
          const q   = this.txFilterText.toLowerCase();
          const cat = this.getCategoryInfo(tx.category, tx.type);
          const matches = (tx.note || '').toLowerCase().includes(q) || cat.label.toLowerCase().includes(q);
          if (!matches) return false;
        }
        return true;
      })
      .sort((a, b) => b.createdAt - a.createdAt);
  },

  renderTxHistory() {
    const listEl = document.getElementById('tx-history-list');
    const txs    = this.getFilteredTxs();

    // Summary
    let existing = document.getElementById('tx-filter-total');
    if (!existing) {
      existing = document.createElement('div');
      existing.id = 'tx-filter-total';
      document.getElementById('tx-history-list').before(existing);
    }

    const totalInc = txs.filter(t => t.type === 'income').reduce((s, t) => s + t.amountUSD, 0);
    const totalExp = txs.filter(t => t.type === 'expense').reduce((s, t) => s + t.amountUSD, 0);
    const net      = totalInc - totalExp;

    existing.className = 'filter-total-card';
    existing.innerHTML = txs.length
      ? `<span class="ft-count">${txs.length} transacción${txs.length !== 1 ? 'es' : ''}</span>
         ${totalInc ? `<span class="ft-inc">+${this.fmtUSD(totalInc)}</span>` : ''}
         ${totalExp ? `<span class="ft-exp">-${this.fmtUSD(totalExp)}</span>` : ''}
         ${totalInc && totalExp ? `<span class="ft-net" style="color:${net >= 0 ? 'var(--positive)' : 'var(--accent)'}">= ${net >= 0 ? '+' : ''}${this.fmtUSD(net)}</span>` : ''}`
      : '';

    if (!txs.length) {
      listEl.innerHTML = '<div class="empty-state"><div class="empty-icon">📭</div>Sin transacciones con estos filtros</div>';
      return;
    }

    listEl.innerHTML = txs.map(tx => this.txItemHTML(tx)).join('');
    listEl.querySelectorAll('.tx-item').forEach((el, i) => {
      el.addEventListener('click', () => this.showTxDetail(txs[i]));
    });
  },

  // ── Calculator ────────────────────────────────────────────
  showCalculator() {
    const r = this.state.rates;
    const rates = [
      { key: 'bcvUsd',   label: 'BCV $',    val: r.bcvUsd   },
      { key: 'bcvEur',   label: 'BCV €',    val: r.bcvEur   },
      { key: 'paralelo', label: 'Paralelo', val: r.paralelo },
    ];
    let selRate  = rates.find(rt => rt.val) || rates[0];
    let people   = 1;
    let mode     = 'usd'; // 'usd' = entro USD, 'ves' = entro Bs

    const fmt = (n, dec = 2) => isNaN(n) ? '--' : n.toLocaleString('es-VE', { minimumFractionDigits: dec, maximumFractionDigits: dec });

    const compute = (amt) => {
      const rv = selRate.val;
      if (!rv || !amt) return null;
      const usdAmt = mode === 'usd' ? amt : amt / rv;
      const vesAmt = mode === 'usd' ? amt * rv : amt;
      return { usdAmt, vesAmt, perUSD: usdAmt / people, perVES: vesAmt / people };
    };

    const renderHTML = () => `
      <div class="modal-title">🧮 Calculadora</div>

      <div class="calc-mode-toggle">
        <button class="type-btn${mode === 'usd' ? ' active' : ''}" id="mode-usd">Entro $ USD</button>
        <button class="type-btn${mode === 'ves' ? ' active' : ''}" id="mode-ves">Entro Bs</button>
      </div>

      <div class="form-group">
        <label id="calc-label">${mode === 'usd' ? 'Monto en USD' : 'Monto en Bs'}</label>
        <input type="number" id="calc-amount" class="form-input"
          placeholder="${mode === 'usd' ? '0,00 $' : '0,00 Bs'}"
          min="0" step="0.01" inputmode="decimal" />
      </div>

      <div class="calc-rate-btns">
        ${rates.map(rt => `<button class="type-btn${selRate.key === rt.key ? ' active' : ''}" data-rkey="${rt.key}">${rt.label}${rt.val ? ` · Bs ${fmt(rt.val)}` : ' · --'}</button>`).join('')}
      </div>

      <div class="calc-result" id="calc-result">
        <div class="calc-row"><span id="cr-label-a">${mode === 'usd' ? 'Equivale en Bs' : 'Equivale en USD'}</span><span id="cr-main">--</span></div>
        <div class="calc-row" style="font-size:.75rem;color:var(--text3)"><span>Tasa</span><span id="cr-rate">Bs ${fmt(selRate.val)}</span></div>
      </div>

      <div class="form-group" style="margin-top:14px">
        <label>Dividir entre</label>
        <div class="calc-split-btns">
          ${[1,2,3,4,5,6,7,8].map(n => `<button class="quick-btn${people === n ? ' selected' : ''}" data-people="${n}">${n}</button>`).join('')}
        </div>
      </div>
      <div class="calc-result" id="calc-split-result" style="display:none">
        <div class="calc-row"><b>Por persona en $</b><span id="cr-per-usd">--</span></div>
        <div class="calc-row"><span>Por persona en Bs</span><span id="cr-per-ves">--</span></div>
      </div>

      <button class="btn-secondary" id="modal-close" style="margin-top:16px">Cerrar</button>
    `;

    const update = () => {
      const amt = parseFloat(document.getElementById('calc-amount').value) || 0;
      const res = compute(amt);
      const mainEl = document.getElementById('cr-main');
      if (!res || !amt) {
        mainEl.textContent = '--';
        document.getElementById('calc-split-result').style.display = 'none';
        return;
      }
      mainEl.textContent = mode === 'usd' ? `Bs ${fmt(res.vesAmt)}` : `$${fmt(res.usdAmt)}`;
      document.getElementById('cr-rate').textContent = `Bs ${fmt(selRate.val)}`;
      const splitEl = document.getElementById('calc-split-result');
      splitEl.style.display = people > 1 ? '' : 'none';
      document.getElementById('cr-per-usd').textContent = `$${fmt(res.perUSD)}`;
      document.getElementById('cr-per-ves').textContent = `Bs ${fmt(res.perVES)}`;
    };

    const mount = () => {
      document.getElementById('modal-content').innerHTML = renderHTML();

      document.getElementById('calc-amount').addEventListener('input', update);

      document.getElementById('mode-usd').addEventListener('click', () => { mode = 'usd'; mount(); });
      document.getElementById('mode-ves').addEventListener('click', () => { mode = 'ves'; mount(); });

      document.querySelectorAll('[data-rkey]').forEach(btn => {
        btn.addEventListener('click', () => {
          selRate = rates.find(rt => rt.key === btn.dataset.rkey);
          document.querySelectorAll('[data-rkey]').forEach(b => b.classList.toggle('active', b === btn));
          update();
        });
      });

      document.querySelectorAll('[data-people]').forEach(btn => {
        btn.addEventListener('click', () => {
          people = parseInt(btn.dataset.people);
          document.querySelectorAll('[data-people]').forEach(b => b.classList.toggle('selected', b === btn));
          update();
        });
      });

      document.getElementById('modal-close').addEventListener('click', () => this.closeModal());
      setTimeout(() => document.getElementById('calc-amount')?.focus(), 100);
    };

    mount();
    this.openModal();
  },

  showAddAccountModal() {
    const emojis = ['🏦','💳','💰','🏧','💵','🏪','📱','🌐'];
    document.getElementById('modal-content').innerHTML = `
      <div class="modal-title">Agregar cuenta</div>
      <div class="form-group">
        <label>Nombre</label>
        <input type="text" id="new-acc-name" class="form-input" placeholder="Mi cuenta…" maxlength="30" />
      </div>
      <div class="form-group">
        <label>Tipo</label>
        <select id="new-acc-type" class="form-select">
          <option value="debit">Débito</option>
          <option value="credit">Crédito</option>
        </select>
      </div>
      <div class="form-group">
        <label>Moneda</label>
        <select id="new-acc-currency" class="form-select">
          <option value="USD">USD (Dólares)</option>
          <option value="VES">VES (Bolívares)</option>
          <option value="EUR">EUR (Euros)</option>
        </select>
      </div>
      <div class="form-group">
        <label>Emoji</label>
        <div style="display:flex;gap:8px;flex-wrap:wrap" id="emoji-picker">
          ${emojis.map(e => `<button type="button" class="quick-btn" data-emoji="${e}">${e}</button>`).join('')}
        </div>
        <input type="hidden" id="new-acc-emoji" value="💳" />
      </div>
      <div class="modal-actions">
        <button class="btn-primary" id="modal-add-acc">Agregar</button>
        <button class="btn-secondary" id="modal-close">Cancelar</button>
      </div>`;

    document.getElementById('emoji-picker').addEventListener('click', e => {
      const btn = e.target.closest('.quick-btn');
      if (!btn) return;
      document.querySelectorAll('#emoji-picker .quick-btn').forEach(b => b.classList.toggle('selected', b === btn));
      document.getElementById('new-acc-emoji').value = btn.dataset.emoji;
    });

    document.getElementById('modal-add-acc').addEventListener('click', () => {
      const name     = document.getElementById('new-acc-name').value.trim();
      const type     = document.getElementById('new-acc-type').value;
      const currency = document.getElementById('new-acc-currency').value;
      const emoji    = document.getElementById('new-acc-emoji').value;
      if (!name) { this.toast('Ingresa un nombre'); return; }
      this.state.accounts.push({ id: this.uid(), name, type, currency, emoji });
      State.save(this.state);
      this.closeModal();
      this.renderAll();
      this.toast('Cuenta agregada');
    });

    document.getElementById('modal-close').addEventListener('click', () => this.closeModal());
    this.openModal();
  },

  // ── Reports ───────────────────────────────────────────────
  renderReports() {
    this.renderPieChart();
    this.renderBarChart();
  },

  renderPieChart() {
    const now = new Date();
    const ym  = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

    const byCategory = {};
    for (const tx of this.state.transactions) {
      if (tx.type !== 'expense') continue;
      if (!tx.date.startsWith(ym)) continue;
      byCategory[tx.category] = (byCategory[tx.category] || 0) + tx.amountUSD;
    }

    const labels  = [];
    const data    = [];
    const colors  = [];
    let   ci      = 0;

    for (const [catId, total] of Object.entries(byCategory).sort((a, b) => b[1] - a[1])) {
      const cat = this.getCategoryInfo(catId, 'expense');
      labels.push(`${cat.icon} ${cat.label}`);
      data.push(parseFloat(total.toFixed(2)));
      colors.push(PIE_COLORS[ci % PIE_COLORS.length]);
      ci++;
    }

    const canvas = document.getElementById('chart-pie');
    if (this.charts.pie) this.charts.pie.destroy();

    if (!data.length) {
      document.getElementById('pie-legend').innerHTML = '<div class="empty-state">Sin gastos este mes</div>';
      canvas.getContext('2d').clearRect(0, 0, canvas.width, canvas.height);
      return;
    }

    this.charts.pie = new Chart(canvas, {
      type: 'doughnut',
      data: { labels, datasets: [{ data, backgroundColor: colors, borderWidth: 0 }] },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${this.fmtUSD(ctx.raw)}` } },
        },
      },
    });

    document.getElementById('pie-legend').innerHTML = labels.map((l, i) =>
      `<div class="legend-item"><div class="legend-dot" style="background:${colors[i]}"></div>${l} — ${this.fmtUSD(data[i])}</div>`
    ).join('');
  },

  renderBarChart() {
    const months = [];
    const now    = new Date();
    for (let i = 5; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      months.push({
        ym:    `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`,
        label: d.toLocaleString('es', { month: 'short', year: '2-digit' }),
      });
    }

    const income  = new Array(6).fill(0);
    const expense = new Array(6).fill(0);

    for (const tx of this.state.transactions) {
      const idx = months.findIndex(m => tx.date.startsWith(m.ym));
      if (idx < 0) continue;
      if (tx.type === 'income')  income[idx]  += tx.amountUSD;
      if (tx.type === 'expense') expense[idx] += tx.amountUSD;
    }

    const canvas = document.getElementById('chart-bar');
    if (this.charts.bar) this.charts.bar.destroy();

    const isDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    const gridColor = isDark ? 'rgba(255,255,255,.08)' : 'rgba(0,0,0,.06)';
    const textColor = isDark ? '#aaa' : '#666';

    this.charts.bar = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: months.map(m => m.label),
        datasets: [
          { label: 'Ingresos', data: income.map(v => +v.toFixed(2)),  backgroundColor: '#3B6D11' },
          { label: 'Gastos',   data: expense.map(v => +v.toFixed(2)), backgroundColor: '#E24B4A' },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        plugins: {
          legend: { labels: { color: textColor, font: { size: 11 } } },
          tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: ${this.fmtUSD(ctx.raw)}` } },
        },
        scales: {
          x: { ticks: { color: textColor }, grid: { color: gridColor } },
          y: { ticks: { color: textColor, callback: v => `$${v}` }, grid: { color: gridColor } },
        },
      },
    });
  },

  // ── Config ────────────────────────────────────────────────
  renderConfig() {
    document.getElementById('c-sheets-url').value = this.state.config.sheetsUrl || '';
    document.getElementById('apps-script-code').textContent = APPS_SCRIPT_CODE;
    const ls = this.state.config.lastSync;
    document.getElementById('last-sync-info').textContent = ls
      ? `Última sincronización: ${new Date(ls).toLocaleString('es')}`
      : 'Sin sincronizaciones previas.';
  },

  setupConfig() {
    document.getElementById('btn-save-config').addEventListener('click', () => {
      const url = document.getElementById('c-sheets-url').value.trim();
      this.state.config.sheetsUrl = url;
      State.save(this.state);
      this.toast('Configuración guardada');
    });

    document.getElementById('btn-sync-all').addEventListener('click', () => this.syncAll());
    document.getElementById('btn-restore').addEventListener('click', () => this.restoreFromSheets());
    document.getElementById('btn-copy-script').addEventListener('click', () => {
      navigator.clipboard.writeText(APPS_SCRIPT_CODE).then(() => this.toast('Código copiado'));
    });
    document.getElementById('btn-export-json').addEventListener('click', () => this.exportJSON());
    document.getElementById('btn-import-json').addEventListener('click', () => {
      document.getElementById('import-file').click();
    });
    document.getElementById('import-file').addEventListener('change', e => this.importJSON(e));
    document.getElementById('btn-change-pin').addEventListener('click', () => {
      Auth.showScreen('setup', () => this.toast('PIN actualizado'));
    });
    document.getElementById('btn-remove-pin').addEventListener('click', () => {
      if (confirm('¿Quitar el PIN? La app quedará sin protección.')) {
        Auth.clearPin();
        this.toast('PIN eliminado');
      }
    });

    document.getElementById('btn-clear-data').addEventListener('click', () => {
      if (confirm('¿Borrar TODOS los datos? Esta acción no se puede deshacer.')) {
        localStorage.removeItem('finanzas_data_v1');
        location.reload();
      }
    });
  },

  // ── Sync ─────────────────────────────────────────────────
  async pullFromSheets({ silent = false } = {}) {
    const url = this.state.config.sheetsUrl;
    if (!url) return;
    try {
      const res = await fetch('/api/sheets', { headers: { 'X-Sheets-Url': url } });
      if (!res.ok) return;
      const data = await res.json();
      if (!data.transactions) return;

      // Merge: keep local transactions not in Sheets, add Sheets ones not local
      const localIds  = new Set(this.state.transactions.map(t => t.id));
      const sheetsIds = new Set(data.transactions.map(t => t.id));

      const onlyLocal  = this.state.transactions.filter(t => !sheetsIds.has(t.id));
      const fromSheets = data.transactions.filter(t => !localIds.has(t.id));

      if (fromSheets.length === 0 && onlyLocal.length === 0) return;

      this.state.transactions = [...data.transactions, ...onlyLocal];
      if (data.accounts && data.accounts.length) this.state.accounts = data.accounts;
      State.save(this.state);
      this.renderAll();
      if (!silent && fromSheets.length > 0) {
        this.toast(`↓ ${fromSheets.length} transacción(es) nueva(s) desde Sheets`);
      }
      // Push merged state back so Sheets stays complete
      if (onlyLocal.length > 0) this.syncAll({ silent: true });
    } catch { /* sin conexión, ignorar */ }
  },

  async syncAll({ silent = false } = {}) {
    const url = this.state.config.sheetsUrl;
    if (!url) {
      if (!silent) { this.toast('Configura la URL de Google Sheets primero'); this.switchTab('config'); }
      return;
    }

    const statusEl = document.getElementById('sync-status');
    if (!silent) statusEl.textContent = '⏳ Sincronizando…';

    try {
      const res = await fetch('/api/sheets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-Sheets-Url': url },
        body: JSON.stringify({
          accounts:     this.state.accounts,
          transactions: this.state.transactions,
        }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      this.state.config.lastSync = new Date().toISOString();
      State.save(this.state);
      this.renderConfig();
      if (!silent) {
        statusEl.textContent = '✓ Sincronizado correctamente';
        this.toast('✓ Datos enviados a Sheets');
      }
    } catch (err) {
      if (!silent) {
        statusEl.textContent = `✗ Error: ${err.message}`;
        this.toast('Error al sincronizar');
      }
    }
  },

  async restoreFromSheets() {
    const url = this.state.config.sheetsUrl;
    if (!url) { this.toast('Configura la URL de Google Sheets primero'); return; }
    if (!confirm('¿Restaurar datos desde Sheets? Se sobreescribirán los datos locales.')) return;

    const statusEl = document.getElementById('sync-status');
    statusEl.textContent = '⏳ Restaurando…';

    try {
      const res = await fetch('/api/sheets', {
        headers: { 'X-Sheets-Url': url },
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      if (data.accounts)     this.state.accounts     = data.accounts;
      if (data.transactions) this.state.transactions = data.transactions;
      this.state.config.lastSync = new Date().toISOString();
      State.save(this.state);
      this.renderAll();
      statusEl.textContent = '✓ Datos restaurados';
      this.toast('✓ Datos restaurados desde Sheets');
    } catch (err) {
      statusEl.textContent = `✗ Error: ${err.message}`;
      this.toast('Error al restaurar');
    }
  },

  // ── Import / Export ───────────────────────────────────────
  exportJSON() {
    const blob = new Blob([JSON.stringify(this.state, null, 2)], { type: 'application/json' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `finanzas-backup-${new Date().toISOString().slice(0,10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
  },

  importJSON(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const data = JSON.parse(ev.target.result);
        if (!data.accounts || !data.transactions) throw new Error('Formato inválido');
        if (!confirm('¿Importar datos? Se sobreescribirán los datos locales.')) return;
        this.state = { ...this.state, ...data };
        State.save(this.state);
        this.renderAll();
        this.toast('✓ Datos importados');
      } catch {
        this.toast('Archivo JSON inválido');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  },

  // ── Toast ─────────────────────────────────────────────────
  toast(msg) {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.classList.remove('hidden');
    clearTimeout(this._toastTimer);
    this._toastTimer = setTimeout(() => el.classList.add('hidden'), 2600);
  },
};

/* ═══════════════════════════════════════════════════════════
   BOOT
═══════════════════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', () => {
  Auth.gate(() => App.init());
});
