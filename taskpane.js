"use strict";

/*
  taskpane.js — Professional Xero P&L Refresher Add-in v2.0
  All logic: OAuth, Excel read/write, Xero API, step UI management.
*/

// ── Config — update APP_URL after deploying to Netlify ──────────────────────
const APP_URL = "https://YOUR-NETLIFY-APP.netlify.app";

const XERO_AUTH_URL = "https://login.xero.com/identity/connect/authorize";
const XERO_API_BASE = "https://api.xero.com/api.xro/2.0";
const XERO_TENANTS  = "https://api.xero.com/connections";
const SCOPES = "accounting.reports.profitandloss.read accounting.settings.read offline_access";

const LS = {
  CLIENT_ID:     "xpnl_client_id",
  CLIENT_SECRET: "xpnl_client_secret",
  TOKEN:         "xpnl_token",
  TENANT_ID:     "xpnl_tenant_id",
  TENANT_NAME:   "xpnl_tenant_name",
};

const MONTH_MAP = {
  jan:1,feb:2,mar:3,apr:4,may:5,jun:6,
  jul:7,aug:8,sep:9,oct:10,nov:11,dec:12,
  january:1,february:2,march:3,april:4,june:6,
  july:7,august:8,september:9,october:10,november:11,december:12,
};

let _pkceVerifier = null;
let _oauthState   = null;
let _authDialog   = null;
let _currentStep  = 1;

// ── Office ready ─────────────────────────────────────────────────────────────
Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) return;
  document.getElementById("clientId").value     = localStorage.getItem(LS.CLIENT_ID)     || "";
  document.getElementById("clientSecret").value = localStorage.getItem(LS.CLIENT_SECRET) || "";
  const hasCreds = !!localStorage.getItem(LS.CLIENT_ID);
  setStep(hasCreds ? 2 : 1);
  updateConnectionUI();
  log("Add-in loaded. " + (hasCreds ? "Credentials found." : "Enter credentials to begin."));
});

// ── Step management ──────────────────────────────────────────────────────────
function setStep(n) {
  _currentStep = n;
  [1,2,3].forEach(i => {
    const item = document.getElementById(`stepItem${i}`);
    const dot  = document.getElementById(`stepDot${i}`);
    item.className = "step-item";
    if (i < n)      { item.classList.add("done");   dot.textContent = "✓"; }
    else if (i === n){ item.classList.add("active"); dot.textContent = String(i); }
    else             { dot.textContent = String(i); }
  });
}

// ── Credentials ──────────────────────────────────────────────────────────────
function saveCreds() {
  const cid    = document.getElementById("clientId").value.trim();
  const secret = document.getElementById("clientSecret").value.trim();
  if (!cid || !secret) { log("⚠ Enter both Client ID and Client Secret.", "warn"); return; }
  localStorage.setItem(LS.CLIENT_ID,     cid);
  localStorage.setItem(LS.CLIENT_SECRET, secret);
  log("✓ Credentials saved.", "ok");
  setStep(2);
}

// ── Token management ─────────────────────────────────────────────────────────
function getToken() {
  try {
    const raw = localStorage.getItem(LS.TOKEN);
    if (!raw) return null;
    const tok = JSON.parse(raw);
    const exp = (tok.obtained_at || 0) + (tok.expires_in || 1800) - 60;
    return Date.now() / 1000 < exp ? tok : null;
  } catch { return null; }
}

function saveToken(tok) {
  tok.obtained_at = Date.now() / 1000;
  localStorage.setItem(LS.TOKEN, JSON.stringify(tok));
}

function isConnected() {
  return !!getToken() && !!localStorage.getItem(LS.TENANT_ID);
}

function disconnect() {
  [LS.TOKEN, LS.TENANT_ID, LS.TENANT_NAME].forEach(k => localStorage.removeItem(k));
  updateConnectionUI();
  log("Disconnected from Xero.");
}

// ── UI state ─────────────────────────────────────────────────────────────────
function updateConnectionUI() {
  const connected  = isConnected();
  const orgName    = localStorage.getItem(LS.TENANT_NAME) || "";
  const statusEl   = document.getElementById("connStatus");
  const connectBtn = document.getElementById("connectBtn");
  const disconnBtn = document.getElementById("disconnectBtn");
  const refreshBtn = document.getElementById("refreshBtn");
  const orgInfo    = document.getElementById("orgInfo");

  if (connected) {
    statusEl.textContent = `● Connected`;
    statusEl.className   = "badge badge-success";
    connectBtn.style.display  = "none";
    disconnBtn.style.display  = "inline-flex";
    refreshBtn.disabled       = false;
    orgInfo.style.display     = "block";
    document.getElementById("orgStats").innerHTML = `
      <div class="info-stat">
        <div class="info-stat-val">✓</div>
        <div class="info-stat-label">${orgName}</div>
      </div>`;
    setStep(3);
  } else {
    statusEl.textContent = "● Not connected";
    statusEl.className   = "badge badge-neutral";
    connectBtn.style.display  = "block";
    disconnBtn.style.display  = "none";
    refreshBtn.disabled       = true;
    orgInfo.style.display     = "none";
    if (_currentStep > 1) setStep(2);
  }
}

function setProgress(pct, label, sub) {
  const wrap = document.getElementById("progressWrap");
  const bar  = document.getElementById("progressBar");
  const lbl  = document.getElementById("progressLabel");
  const slbl = document.getElementById("progressSub");
  wrap.classList.add("visible");
  bar.style.width    = Math.min(pct, 100) + "%";
  if (label) lbl.textContent = label;
  if (sub !== undefined) slbl.textContent = sub || "";
  if (pct >= 100) setTimeout(() => wrap.classList.remove("visible"), 2000);
}

function clearLog() {
  const area = document.getElementById("logArea");
  area.innerHTML = '<div class="log-line">Log cleared.</div>';
}

function log(msg, type = "") {
  const area = document.getElementById("logArea");
  const line = document.createElement("div");
  const t    = new Date().toLocaleTimeString("en-NZ", { hour12: false });
  line.className   = "log-line" + (type ? " log-" + type : "");
  line.textContent = `${t}  ${msg}`;
  area.appendChild(line);
  area.scrollTop = area.scrollHeight;
}

// ── PKCE helpers ─────────────────────────────────────────────────────────────
function genVerifier() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~";
  const arr   = new Uint8Array(64);
  crypto.getRandomValues(arr);
  return Array.from(arr, n => chars[n % chars.length]).join("");
}

async function genChallenge(v) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(v));
  return btoa(String.fromCharCode(...new Uint8Array(buf)))
    .replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}

// ── OAuth ─────────────────────────────────────────────────────────────────────
async function connectXero() {
  const cid    = localStorage.getItem(LS.CLIENT_ID);
  const secret = localStorage.getItem(LS.CLIENT_SECRET);
  if (!cid || !secret) {
    saveCreds();
    if (!localStorage.getItem(LS.CLIENT_ID)) return;
  }

  _pkceVerifier = genVerifier();
  _oauthState   = crypto.randomUUID();
  const challenge = await genChallenge(_pkceVerifier);

  const params = new URLSearchParams({
    response_type: "code", client_id: cid,
    redirect_uri: `${APP_URL}/auth-dialog.html`,
    scope: SCOPES, state: _oauthState,
    code_challenge: challenge, code_challenge_method: "S256",
  });

  log("Opening Xero login…");
  document.getElementById("connectBtn").disabled = true;
  document.getElementById("connectBtn").textContent = "⏳ Waiting for login…";

  Office.context.ui.displayDialogAsync(
    `${XERO_AUTH_URL}?${params}`,
    { height: 60, width: 40, displayInIframe: false },
    result => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        log("✗ Could not open login window: " + result.error.message, "err");
        resetConnectBtn();
        return;
      }
      _authDialog = result.value;
      _authDialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
      _authDialog.addEventHandler(Office.EventType.DialogEventReceived, args => {
        if (args.error === 12006) { log("Login cancelled.", "warn"); resetConnectBtn(); }
      });
    }
  );
}

function resetConnectBtn() {
  const btn = document.getElementById("connectBtn");
  btn.disabled = false;
  btn.textContent = "🔗 Connect to Xero";
}

async function onDialogMessage(args) {
  if (_authDialog) _authDialog.close();
  resetConnectBtn();

  let msg;
  try { msg = JSON.parse(args.message); } catch { log("✗ Unexpected response.", "err"); return; }
  if (msg.type === "error")  { log("✗ Login failed: " + msg.message, "err"); return; }
  if (msg.type !== "code")   { log("✗ Unexpected message type.", "err"); return; }
  if (msg.state !== _oauthState) { log("✗ Security check failed. Please try again.", "err"); return; }

  log("Exchanging authorisation code for access token…");

  try {
    const resp = await fetch(`${APP_URL}/.netlify/functions/token`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        code: msg.code, verifier: _pkceVerifier,
        redirect_uri: `${APP_URL}/auth-dialog.html`,
      })
    });
    if (!resp.ok) throw new Error(`Token exchange failed (${resp.status})`);
    const tok = await resp.json();
    saveToken(tok);

    const tenResp = await fetch(XERO_TENANTS, {
      headers: { "Authorization": `Bearer ${tok.access_token}`, "Accept": "application/json" }
    });
    const tenants = await tenResp.json();
    if (!tenants?.length) throw new Error("No Xero organisations found.");
    localStorage.setItem(LS.TENANT_ID,   tenants[0].tenantId);
    localStorage.setItem(LS.TENANT_NAME, tenants[0].tenantName);
    log(`✓ Connected: ${tenants[0].tenantName}`, "ok");
    updateConnectionUI();
  } catch(e) {
    log("✗ " + e.message, "err");
  }
}

// ── Excel helpers ─────────────────────────────────────────────────────────────
function parseMonth(val) {
  if (!val) return null;
  if (val instanceof Date) return { year: val.getFullYear(), month: val.getMonth() + 1 };
  if (typeof val !== "string") return null;
  const [m, y] = val.trim().split(/\s+/);
  const mn = MONTH_MAP[m?.slice(0,3).toLowerCase()];
  const yr = parseInt(y, 10);
  return mn && yr >= 2000 && yr <= 2100 ? { year: yr, month: mn } : null;
}

async function readSheet() {
  return Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const hdr   = sheet.getRange("A5:Z5");
    hdr.load("values");
    await ctx.sync();

    const periods = [];
    hdr.values[0].forEach((v, i) => {
      if (i === 0) return;
      const p = parseMonth(v);
      if (p) periods.push({ col: i + 1, letter: String.fromCharCode(65 + i), ...p });
    });

    if (!periods.length) throw new Error("No month headers found in row 5. Expected format: 'May 2026'");

    const acctRange = sheet.getRange("A6:A80");
    acctRange.load("values, format/font/bold");
    await ctx.sync();

    const accounts = [];
    acctRange.values.forEach((row, i) => {
      const v    = row[0];
      const bold = acctRange.format.font.bold[i]?.[0];
      if (v && !bold) accounts.push({ row: i + 6, name: String(v).trim() });
    });

    return { periods, accounts };
  });
}

// ── Xero fetch ────────────────────────────────────────────────────────────────
function pad2(n) { return String(n).padStart(2, "0"); }
function lastDay(y, m) { return new Date(y, m, 0).getDate(); }

async function fetchPnL(year, month) {
  const tok    = getToken();
  const tenant = localStorage.getItem(LS.TENANT_ID);
  const from   = `${year}-${pad2(month)}-01`;
  const to     = `${year}-${pad2(month)}-${pad2(lastDay(year, month))}`;
  const url    = `${XERO_API_BASE}/Reports/ProfitAndLoss?fromDate=${from}&toDate=${to}&standardLayout=true&paymentsOnly=false`;

  const resp = await fetch(url, {
    headers: {
      "Authorization":  `Bearer ${tok.access_token}`,
      "Xero-Tenant-Id": tenant,
      "Accept":         "application/json"
    }
  });

  if (!resp.ok) throw new Error(`Xero API error ${resp.status}`);
  const data = await resp.json();
  const out  = {};

  function walk(rows) {
    for (const r of rows) {
      if (r.RowType === "Section" || r.RowType === "SummaryRow") walk(r.Rows || []);
      else if (r.RowType === "Row") {
        const cells = r.Cells || [];
        if (cells.length >= 2) {
          const name = cells[0].Value?.trim();
          const amt  = parseFloat((cells[1].Value || "0").replace(/,/g, "")) || 0;
          if (name) out[name.toLowerCase()] = amt;
        }
      }
    }
  }
  walk(data.Reports?.[0]?.Rows || []);
  return out;
}

// ── Main refresh ──────────────────────────────────────────────────────────────
async function runRefresh() {
  if (!isConnected()) { log("Not connected to Xero.", "warn"); return; }

  const btn = document.getElementById("refreshBtn");
  btn.disabled = true;
  btn.textContent = "⏳ Refreshing…";
  document.getElementById("resultStats").style.display = "none";

  try {
    log("Reading spreadsheet…");
    setProgress(5, "Reading spreadsheet…", "");
    const { periods, accounts } = await readSheet();
    log(`Found ${periods.length} month(s) and ${accounts.length} account(s).`);

    const allData = {};
    for (let i = 0; i < periods.length; i++) {
      const { year, month, letter } = periods[i];
      const label = new Date(year, month - 1).toLocaleString("en-NZ", { month: "long", year: "numeric" });
      const pct   = 10 + Math.round((i / periods.length) * 72);
      setProgress(pct, `Fetching ${label}…`, `Column ${letter}`);
      log(`Fetching ${label}…`);
      try {
        allData[`${year}-${month}`] = await fetchPnL(year, month);
        log(`  ✓ ${Object.keys(allData[`${year}-${month}`]).length} accounts`, "ok");
      } catch(e) {
        log(`  ✗ ${label}: ${e.message}`, "err");
        allData[`${year}-${month}`] = {};
      }
    }

    setProgress(86, "Writing to spreadsheet…", "");
    log("Writing to spreadsheet…");

    let written = 0;
    const notFound = [];

    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const org   = localStorage.getItem(LS.TENANT_NAME) || "Xero";
      sheet.getRange("A2").values = [[org]];
      sheet.getRange("A3").values = [[`Refreshed: ${new Date().toLocaleString("en-NZ")}`]];

      for (const { row, name } of accounts) {
        const key = name.toLowerCase();
        for (const { col, letter, year, month } of periods) {
          const pd   = allData[`${year}-${month}`] || {};
          const cell = sheet.getRange(`${letter}${row}`);
          let amt    = null;

          if (key in pd) {
            amt = pd[key];
          } else {
            const matches = Object.entries(pd).filter(([k]) => k.includes(key) || key.includes(k));
            if (matches.length === 1) amt = matches[0][1];
          }

          if (amt !== null) {
            cell.values         = [[amt]];
            cell.numberFormat   = [['#,##0;(#,##0);"-"']];
            cell.format.fill.color = "#FFFFFF";
            cell.format.font.color = "#242424";
            written++;
          } else {
            cell.values = [[null]];
            cell.format.fill.color = "#FFF4CE";  // subtle amber — not harsh yellow
            if (!notFound.includes(name)) notFound.push(name);
          }
        }
      }
      await ctx.sync();
    });

    setProgress(100, "Complete!", "");
    log(`✓ ${written} cells updated.`, "ok");
    if (notFound.length) {
      log(`⚠ ${notFound.length} unmatched: ${notFound.join(", ")}`, "warn");
    }

    // Show result stats
    document.getElementById("statCells").textContent     = written;
    document.getElementById("statMonths").textContent    = periods.length;
    document.getElementById("statUnmatched").textContent = notFound.length;
    document.getElementById("resultStats").style.display = "block";

  } catch(e) {
    log("✗ " + e.message, "err");
    setProgress(0, "", "");
    console.error(e);
  } finally {
    btn.disabled = false;
    btn.textContent = "🔄 Fetch from Xero & Update Sheet";
  }
}
