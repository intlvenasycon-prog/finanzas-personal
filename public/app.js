'use strict';

/* ═══════════════════════════════════════════════════════════
   CONSTANTS
═══════════════════════════════════════════════════════════ */
const EXPENSE_CATS = [
  { id: 'food',     label: 'Comida',        icon: '🍔' },
  { id: 'health',   label: 'Salud',         icon: '💊' },
  { id: 'fun',      label: 'Entretenimiento',icon: '🎬' },
  { id: 'shopping', label: 'Shopping',      icon: '🛍️' },
  { id: 'transport',label: 'Transporte',    icon: '🚗' },
  { id: 'data',     label: 'Data',          icon: '📶' },
  { id: 'subs',     label: 'Membresías',    icon: '📱' },
  { id: 'travel',   label: 'Viajes',        icon: '✈️' },
  { id: 'invest',   label: 'Inversiones',   icon: '📈' },
  { id: 'other',    label: 'Otros',         icon: '📌' },
];

const INCOME_CATS = [
  { id: 'papa',     label: 'Depósito papá', icon: '👨‍👦' },
  { id: 'salary',   label: 'Salario',       icon: '💼' },
  { id: 'freelance',label: 'Freelance',     icon: '💻' },
  { id: 'inv-in',   label: 'Inversión',     icon: '📈' },
  { id: 'other-in', label: 'Otros',         icon: '💰' },
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

  // ── Init ────────────────────────────────────────────────
  async init() {
    this.state = State.load();
    this.setupTabs();
    this.setupForm();
    this.setupConfig();
    this.renderAll();
    this.setDefaultDate();
    document.getElementById('btn-sync-header').addEventListener('click', () => this.syncAll());
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('/sw.js').catch(() => {});
    }
    // Auto-pull from Sheets on app open
    await this.pullFromSheets({ silent: true });
    await this.fetchRates();
    // Auto-pull when app comes back to foreground (works on iOS PWA)
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible') this.pullFromSheets({ silent: true });
    });
    window.addEventListener('pageshow', (e) => {
      if (e.persisted) this.pullFromSheets({ silent: true });
    });
    window.addEventListener('focus', () => {
      this.pullFromSheets({ silent: true });
    });
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
      const bal = this.getAccountBalance(acc.id);
      const isNeg = acc.type === 'credit' ? bal < 0 : bal < 0;
      const div = document.createElement('div');
      div.className = 'account-card';
      div.innerHTML = `
        <div class="acc-name">${acc.emoji || '💳'} ${acc.name}</div>
        <div class="acc-balance ${isNeg ? 'negative' : ''}">${acc.currency === 'VES' ? this.fmtVES(bal) : this.fmtUSD(bal)}</div>
        <div class="acc-type">${acc.currency === 'VES' && this.state.rates.bcvUsd ? `≈ ${this.fmtUSD(bal / this.state.rates.bcvUsd)} · ` : ''}${acc.type === 'credit' ? 'Crédito' : 'Débito'} · ${acc.currency === 'VES' ? 'VES' : acc.currency}</div>
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
    for (const acc of this.state.accounts) {
      const bal = this.getAccountBalance(acc.id);
      totalUSD += bal;
      if (acc.currency === 'VES' && this.state.rates.bcvUsd) {
        totalVES += bal * this.state.rates.bcvUsd;
      }
    }
    document.getElementById('total-usd').textContent = this.fmtUSD(totalUSD);
    document.getElementById('total-ves').textContent = `Bs ${totalVES.toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
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
      const bal = this.getAccountBalance(acc.id);
      const isNeg = bal < 0;
      return `
        <div class="account-detail-card">
          <div class="acc-emoji">${acc.emoji || '💳'}</div>
          <div class="acc-detail-info">
            <div class="acc-detail-name">${acc.name}</div>
            <div class="acc-detail-meta">${acc.type === 'credit' ? 'Crédito' : 'Débito'} · ${acc.currency === 'VES' ? 'Bolívares' : acc.currency}</div>
          </div>
          <div class="acc-detail-balance ${isNeg ? 'negative' : ''}">
            ${acc.currency === 'VES' ? this.fmtVES(bal) : this.fmtUSD(bal)}
            ${acc.currency === 'VES' && this.state.rates.bcvUsd ? `<small style="display:block;font-size:.72rem;font-weight:400;color:var(--text3)">≈ ${this.fmtUSD(bal / this.state.rates.bcvUsd)}</small>` : ''}
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

    this.renderTxHistory();
  },

  renderTxHistory() {
    const titleEl = document.getElementById('tx-history-title');
    const listEl  = document.getElementById('tx-history-list');

    let txs = [...this.state.transactions].sort((a, b) => b.createdAt - a.createdAt);
    if (this.txFilterAccount) {
      const acc = this.getAccountInfo(this.txFilterAccount);
      titleEl.textContent = acc ? `Transacciones — ${acc.name}` : 'Transacciones';
      txs = txs.filter(t => t.account === this.txFilterAccount || t.toAccount === this.txFilterAccount);
    } else {
      titleEl.textContent = 'Todas las transacciones';
    }

    if (!txs.length) {
      listEl.innerHTML = '<div class="empty-state"><div class="empty-icon">📭</div>Sin transacciones</div>';
      return;
    }

    listEl.innerHTML = txs.map(tx => this.txItemHTML(tx)).join('');
    listEl.querySelectorAll('.tx-item').forEach((el, i) => {
      el.addEventListener('click', () => this.showTxDetail(txs[i]));
    });
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
document.addEventListener('DOMContentLoaded', () => App.init());
