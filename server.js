const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const QRCode = require('qrcode');
const crypto = require('crypto');
const nodemailer = require('nodemailer');

const PORT = process.env.PORT || 8000;
const BASE_DIR = __dirname;
const isVercel = Boolean(process.env.VERCEL);
const DATA_FILE = isVercel ? path.join('/tmp', 'dashboard_data.json') : path.join(BASE_DIR, 'dashboard_data.json');
const EXCEL_PATH = path.join(BASE_DIR, 'Hexalevel_Financial_Management_Dashboard.xlsx');

const BLOB_URL = 'https://chyqewr2sbowut1m.public.blob.vercel-storage.com/dashboard_data.json';
const BLOB_TOKEN = process.env.PUBLIC_BLOB_READ_WRITE_TOKEN || process.env.BLOB_READ_WRITE_TOKEN;

let memoryData = null;
let lastBlobSync = 0;

async function syncFromBlob(force = false) {
  if (!isVercel || !BLOB_TOKEN) return;
  if (!force && memoryData && (Date.now() - lastBlobSync < 1500)) return;
  try {
    const res = await fetch(BLOB_URL + '?t=' + Date.now(), {
      cache: 'no-store',
      headers: { 'Cache-Control': 'no-cache, no-store' }
    });
    if (res.ok) {
      const remoteData = await res.json();
      if (remoteData && remoteData.users && Array.isArray(remoteData.users)) {
        memoryData = remoteData;
        lastBlobSync = Date.now();
        try {
          fs.writeFileSync(DATA_FILE, JSON.stringify(remoteData, null, 2), 'utf-8');
        } catch (e) {}
      }
    }
  } catch (err) {
    console.error('Error syncing from Vercel Blob:', err.message);
  }
}

async function persistToBlob(data) {
  if (!isVercel || !BLOB_TOKEN) return;
  try {
    const { put } = require('@vercel/blob');
    const res = await put('dashboard_data.json', JSON.stringify(data, null, 2), {
      access: 'public',
      addRandomSuffix: false,
      allowOverwrite: true,
      token: BLOB_TOKEN
    });
    lastBlobSync = Date.now();
    console.log('✅ Successfully persisted to Vercel Blob:', res.url);
  } catch (err) {
    console.error('❌ Error persisting to Vercel Blob:', err.message);
  }
}
function loadData() {
  if (memoryData && memoryData.users && memoryData.users.length > 0) {
    return memoryData;
  }
  if (isVercel && !fs.existsSync(DATA_FILE)) {
    try {
      const defaultRaw = fs.readFileSync(path.join(BASE_DIR, 'dashboard_data.json'), 'utf-8');
      fs.writeFileSync(DATA_FILE, defaultRaw, 'utf-8');
    } catch (e) {
      console.error("Vercel /tmp init error:", e);
    }
  }
  if (!fs.existsSync(DATA_FILE)) {
    throw new Error("Data store not found");
  }
  const raw = fs.readFileSync(DATA_FILE, 'utf-8');
  const data = JSON.parse(raw);
  if (!data.activity_logs) {
    data.activity_logs = [];
  }
  if (!data.users || data.users.length === 0) {
    data.users = [
      {
        id: "USR-001",
        name: "Admin",
        email: "admin@hexalevel.com",
        password: "admin",
        passwords: ["admin", "admin123"],
        role: "Administrator",
        created_at: "2026-10-31 12:00:00",
        last_login: "2026-10-31 19:00:00",
        login_count: 1,
        status: "Active"
      }
    ];
  }
  memoryData = data;
  return data;
}

async function saveData(data) {
  memoryData = data;
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf-8');
  } catch (e) {}
  if (isVercel) {
    await persistToBlob(data);
  }
}

function logActivity(data, user, action, tab, details) {
  let cleanUser = user ? user.replace(/\(Co-Founder\)/g, '').replace(/\(Partner\)/g, '').trim() : "Admin";
  if (!cleanUser) cleanUser = "Admin";
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const timestamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  const logCount = (data.activity_logs || []).length + 1;
  const logId = `ACT-${String(logCount).padStart(3, '0')}`;
  const entry = {
    id: logId,
    timestamp: timestamp,
    user: cleanUser,
    action: action,
    tab: tab,
    details: details
  };
  if (!data.activity_logs) data.activity_logs = [];
  data.activity_logs.unshift(entry);
}

function smartCapitalize(str) {
  if (!str || typeof str !== 'string') return '';
  const uppercaseWords = new Set(['CRM', 'GHL', 'WP', 'PKR', 'USD', 'INR', 'HBL', 'UBL', 'USA', 'UK', 'UAE', 'SEO', 'A2P', 'API', 'AI']);
  str = str.replace(/\s+/g, ' ').trim();
  return str.split(' ').map(word => {
    if (!word) return '';
    const cleanWord = word.trim();
    const upper = cleanWord.toUpperCase();
    if (uppercaseWords.has(upper)) return upper;
    if (cleanWord.includes('+')) {
      return cleanWord.split('+').map(part => smartCapitalize(part)).join(' + ');
    }
    if (cleanWord.includes('/')) {
      if (/^\d/.test(cleanWord)) return cleanWord;
      return cleanWord.split('/').map(part => smartCapitalize(part)).join('/');
    }
    return cleanWord.charAt(0).toUpperCase() + cleanWord.slice(1);
  }).join(' ').replace(/\s+\+\s+/g, ' + ').replace(/\s+/g, ' ');
}

const DATE_MONTH_MAP = {
  '01': 'Jan', '1': 'Jan', 'jan': 'Jan', 'january': 'Jan',
  '02': 'Feb', '2': 'Feb', 'feb': 'Feb', 'february': 'Feb',
  '03': 'Mar', '3': 'Mar', 'mar': 'Mar', 'march': 'Mar',
  '04': 'Apr', '4': 'Apr', 'apr': 'Apr', 'april': 'Apr',
  '05': 'May', '5': 'May', 'may': 'May',
  '06': 'Jun', '6': 'Jun', 'jun': 'Jun', 'june': 'Jun',
  '07': 'Jul', '7': 'Jul', 'jul': 'Jul', 'july': 'Jul',
  '08': 'Aug', '8': 'Aug', 'aug': 'Aug', 'august': 'Aug',
  '09': 'Sep', '9': 'Sep', 'sep': 'Sep', 'september': 'Sep',
  '10': 'Oct', 'oct': 'Oct', 'october': 'Oct',
  '11': 'Nov', 'nov': 'Nov', 'november': 'Nov',
  '12': 'Dec', 'dec': 'Dec', 'december': 'Dec'
};

function formatPaymentDate(str) {
  if (!str || !str.trim()) return '';
  str = str.trim();
  let mIso = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (mIso) {
    let yr = mIso[1];
    let mon = DATE_MONTH_MAP[mIso[2]] || mIso[2];
    let day = mIso[3].padStart(2, '0');
    return `${day}-${mon}-${yr}`;
  }
  let m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    let day = m[1].padStart(2, '0');
    let mon = DATE_MONTH_MAP[m[2]] || m[2];
    let yr = m[3].length === 2 ? '20' + m[3] : m[3];
    return `${day}-${mon}-${yr}`;
  }
  m = str.match(/^(\d{1,2})[\s\-]+([A-Za-z]+)[\s\-]+(\d{2,4})$/);
  if (m) {
    let day = m[1].padStart(2, '0');
    let mon = DATE_MONTH_MAP[m[2].toLowerCase()] || m[2];
    let yr = m[3].length === 2 ? '20' + m[3] : m[3];
    return `${day}-${mon}-${yr}`;
  }
  return str;
}

function formatProjectStart(str) {
  if (!str || !str.trim()) return '';
  str = str.trim();
  let mIso = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (mIso) {
    let yr = mIso[1];
    let mon = DATE_MONTH_MAP[mIso[2]] || mIso[2];
    return `${mon} ${yr}`;
  }
  let m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    let mon = DATE_MONTH_MAP[m[2]] || m[2];
    let yr = m[3].length === 2 ? '20' + m[3] : m[3];
    return `${mon} ${yr}`;
  }
  m = str.match(/^([A-Za-z]+)[\s\-]+(\d{2,4})$/);
  if (m) {
    let mon = DATE_MONTH_MAP[m[1].toLowerCase()] || m[1];
    let yr = m[2].length === 2 ? '20' + m[2] : m[2];
    return `${mon} ${yr}`;
  }
  return str;
}

function parseScenarioShares(scenarioStr, amt) {
  scenarioStr = scenarioStr || '';
  if (scenarioStr.includes("Scenario A")) {
    return {
      shaista: amt * 0.25,
      amara: amt * 0.25,
      urooj: 0.0,
      company: amt * 0.50,
      skip_member: null,
      skip_co: 0.0
    };
  } else if (scenarioStr.includes("Scenario B")) {
    return {
      shaista: amt * (7 / 30),
      amara: amt * (7 / 30),
      urooj: amt * (7 / 30),
      company: amt * 0.30,
      skip_member: null,
      skip_co: 0.0
    };
  } else {
    let skip_member, sh, am, ur;
    if (scenarioStr.includes("Urooj Skipped") || scenarioStr.includes("Shaista & Amara")) {
      skip_member = "Urooj";
      sh = amt * (7 / 30);
      am = amt * (7 / 30);
      ur = 0.0;
    } else if (scenarioStr.includes("Amara Skipped") || scenarioStr.includes("Shaista & Urooj")) {
      skip_member = "Amara";
      sh = amt * (7 / 30);
      am = 0.0;
      ur = amt * (7 / 30);
    } else if (scenarioStr.includes("Shaista Skipped") || scenarioStr.includes("Amara & Urooj")) {
      skip_member = "Shaista";
      sh = 0.0;
      am = amt * (7 / 30);
      ur = amt * (7 / 30);
    } else {
      skip_member = "Urooj";
      sh = amt * (7 / 30);
      am = amt * (7 / 30);
      ur = 0.0;
    }
    const skip_co = amt * (7 / 30);
    const tot_co = (amt * 0.30) + skip_co;
    return {
      shaista: sh,
      amara: am,
      urooj: ur,
      company: tot_co,
      skip_member: skip_member,
      skip_co: skip_co
    };
  }
}

function calculateDashboardData(filterUser) {
  const data = loadData();
  const projects = data.projects || [];
  const split_old = data.split_old || [];
  const split_new = data.split_new || [];
  const expenses_old = data.expenses_old || [];
  const expenses_new = data.expenses_new || [];
  const send_money_old = data.send_money_old || [];
  const send_money_new = data.send_money_new || [];
  const scenario_c = data.scenario_c || [];
  const all_logs = data.activity_logs || [];

  let filtered_logs = all_logs;
  if (filterUser && filterUser.trim().toLowerCase() !== 'all' && filterUser.trim() !== '') {
    filtered_logs = all_logs.filter(l => (l.user || '').toLowerCase() === filterUser.trim().toLowerCase());
  }

  // Old Split Totals
  const old_rev = split_old.filter(x => x.status === "Received").reduce((acc, x) => acc + (x.amount || 0), 0);
  const old_co = old_rev * 0.50;
  const old_team = old_rev * 0.50;
  const old_purchases = expenses_old.reduce((acc, e) => acc + (e.price || 0), 0);
  const old_send = send_money_old.reduce((acc, s) => acc + (s.amount || 0), 0);
  const old_net = old_co - old_purchases - old_send;

  const old_projects_list = projects.filter(p => p.period === 'old');
  const old_projects_count = old_projects_list.length;
  const old_received_count = old_projects_list.filter(x => (x.status || '').toLowerCase() === 'received').length;
  const old_pending_count = old_projects_list.filter(x => (x.status || '').toLowerCase() === 'pending').length;
  const old_pending_items = old_projects_list.filter(x => (x.status || '').toLowerCase() === 'pending');
  const old_pending_amount = old_pending_items.map(p => `${p.currency || 'USD'} ${Number(p.amount).toLocaleString()}`).join(', ') || '0';

  // New Split Totals
  const new_rev = split_new.filter(x => x.status === "Received").reduce((acc, x) => acc + (x.amount || 0), 0);
  const new_co = split_new.filter(x => x.status === "Received").reduce((acc, x) => acc + (x.company || 0), 0);
  const new_team = split_new.filter(x => x.status === "Received").reduce((acc, x) => acc + ((x.shaista || 0) + (x.amara || 0) + (x.urooj || 0)), 0);
  const new_purchases = expenses_new.reduce((acc, e) => acc + (e.price || 0), 0);
  const new_send = send_money_new.reduce((acc, s) => acc + (s.amount || 0), 0);
  const new_net = new_co - new_purchases - new_send;

  const new_projects_list = projects.filter(p => p.period === 'new');
  const new_projects_count = new_projects_list.length;
  const new_received_count = new_projects_list.filter(x => (x.status || '').toLowerCase() === 'received').length;
  const new_pending_count = new_projects_list.filter(x => (x.status || '').toLowerCase() === 'pending').length;
  const new_pending_items = new_projects_list.filter(x => (x.status || '').toLowerCase() === 'pending');
  const new_pending_amount = new_pending_items.map(p => `${p.currency || 'PKR'} ${Number(p.amount).toLocaleString()}`).join(', ') || '0';

  const all_rev = old_rev + new_rev;
  const all_net = old_net + new_net;
  const all_projects_count = projects.length;
  const all_received_count = projects.filter(x => (x.status || '').toLowerCase() === 'received').length;
  const all_pending_count = projects.filter(x => (x.status || '').toLowerCase() === 'pending').length;
  const all_pending_items = projects.filter(x => (x.status || '').toLowerCase() === 'pending');
  const all_pending_amount = all_pending_items.map(p => `${p.currency || 'USD'} ${Number(p.amount).toLocaleString()} (${p.client})`).join(', ') || '0';

  // Scenario C Aggregates
  const sc_received = scenario_c.filter(c => c.status === "Received");
  const sc_total_amt = sc_received.reduce((acc, c) => acc + (c.amount || 0), 0);
  const sc_shaista_dir = sc_received.reduce((acc, c) => acc + (c.shaista || 0), 0);
  const sc_amara_dir = sc_received.reduce((acc, c) => acc + (c.amara || 0), 0);
  const sc_urooj_dir = sc_received.reduce((acc, c) => acc + (c.urooj || 0), 0);
  const sc_skip_co = sc_received.reduce((acc, c) => acc + (c.skip_to_co || 0), 0);
  const sc_total_co = sc_received.reduce((acc, c) => acc + (c.total_co || 0), 0);

  // Partner Calculations
  // Direct project payouts
  const sh_old_direct = split_old.filter(x => x.status === "Received").reduce((acc, x) => acc + ((x.team || 0) / 2), 0);
  const sh_new_direct = split_new.filter(x => x.status === "Received").reduce((acc, x) => acc + (x.shaista || 0), 0);
  const sh_total_direct = sh_old_direct + sh_new_direct;
  const sh_worked_count = split_old.length + split_new.filter(x => (x.shaista || 0) > 0).length;
  const sh_skipped_count = scenario_c.filter(c => (c.skip_member || '').toLowerCase() === 'shaista').length;
  const sh_skipped_deduction = scenario_c.filter(c => (c.skip_member || '').toLowerCase() === 'shaista').reduce((acc, c) => acc + ((c.amount || 0) * (7 / 30)), 0);

  const am_old_direct = split_old.filter(x => x.status === "Received").reduce((acc, x) => acc + ((x.team || 0) / 2), 0);
  const am_new_direct = split_new.filter(x => x.status === "Received").reduce((acc, x) => acc + (x.amara || 0), 0);
  const am_total_direct = am_old_direct + am_new_direct;
  const am_worked_count = split_old.length + split_new.filter(x => (x.amara || 0) > 0).length;
  const am_skipped_count = scenario_c.filter(c => (c.skip_member || '').toLowerCase() === 'amara').length;
  const am_skipped_deduction = scenario_c.filter(c => (c.skip_member || '').toLowerCase() === 'amara').reduce((acc, c) => acc + ((c.amount || 0) * (7 / 30)), 0);

  const ur_new_direct = split_new.filter(x => x.status === "Received").reduce((acc, x) => acc + (x.urooj || 0), 0);
  const ur_total_direct = ur_new_direct;
  const ur_worked_count = split_new.filter(x => (x.urooj || 0) > 0).length;
  const ur_skipped_count = scenario_c.filter(c => (c.skip_member || '').toLowerCase() === 'urooj').length;
  const ur_skipped_deduction = scenario_c.filter(c => (c.skip_member || '').toLowerCase() === 'urooj').reduce((acc, c) => acc + ((c.amount || 0) * (7 / 30)), 0);

  // Dedicated Company Financial Statement & Savings Allocation
  const new_co_base = split_new.filter(x => x.status === "Received").reduce((acc, x) => acc + ((x.amount || 0) * 0.30), 0);
  const total_gross_co = old_co + new_co;
  const total_purchases = old_purchases + new_purchases;
  const total_send = old_send + new_send;
  const total_deductions = total_purchases + total_send;
  const company_net = total_gross_co - total_deductions;

  // Company Savings / Asal Net Profit Allocation:
  // Old Period Net (50/50 Model): Shaista 50%, Amara 50%, Urooj 0%
  const sh_old_savings = old_net * 0.50;
  const am_old_savings = old_net * 0.50;
  const ur_old_savings = 0.0;

  // New Period Net (30/70 Model): Shared equally among all 3 members (1/3 each)
  const sh_new_savings = new_net / 3;
  const am_new_savings = new_net / 3;
  const ur_new_savings = new_net / 3;

  // Total Company Savings (Asal Net Profit in Company Tijori)
  const sh_company_savings = sh_old_savings + sh_new_savings;
  const am_company_savings = am_old_savings + am_new_savings;
  const ur_company_savings = ur_old_savings + ur_new_savings;

  const partners = {
    shaista: {
      name: "Shaista",
      direct: sh_total_direct,
      old_savings: sh_old_savings,
      new_savings: sh_new_savings,
      company_savings: sh_company_savings,
      settlement_payable: sh_company_savings,
      total_financial_value: sh_total_direct + sh_company_savings,
      projects_worked: sh_worked_count,
      projects_skipped: sh_skipped_count,
      skipped_deduction: sh_skipped_deduction
    },
    amara: {
      name: "Amara",
      direct: am_total_direct,
      old_savings: am_old_savings,
      new_savings: am_new_savings,
      company_savings: am_company_savings,
      settlement_payable: am_company_savings,
      total_financial_value: am_total_direct + am_company_savings,
      projects_worked: am_worked_count,
      projects_skipped: am_skipped_count,
      skipped_deduction: am_skipped_deduction
    },
    urooj: {
      name: "Urooj",
      direct: ur_total_direct,
      old_savings: ur_old_savings,
      new_savings: ur_new_savings,
      company_savings: ur_company_savings,
      settlement_payable: ur_company_savings,
      total_financial_value: ur_total_direct + ur_company_savings,
      projects_worked: ur_worked_count,
      projects_skipped: ur_skipped_count,
      skipped_deduction: ur_skipped_deduction
    }
  };

  const company_statement = {
    old_co_share: old_co,
    new_co_base_share: new_co_base,
    sc_skip_additions: sc_skip_co,
    total_gross_inflow: total_gross_co,
    purchases_deducted: total_purchases,
    send_money_deducted: total_send,
    total_deductions: total_deductions,
    overall_net_profit: company_net,
    old_co_net: old_net,
    new_co_net: new_net,
    allocated_shaista: sh_company_savings,
    allocated_amara: am_company_savings,
    allocated_urooj: ur_company_savings,
    total_revenue: all_rev,
    net_profit_margin_pct: all_rev > 0 ? parseFloat(((company_net / all_rev) * 100).toFixed(2)) : 0
  };

  const chart_data = {
    scenario_comparison: {
      labels: ["Scenario A (2 Working)", "Scenario B (3 Working)", "Scenario C (2 Working, 1 Skipped)"],
      company_pct: [50.0, 30.0, 53.33],
      shaista_pct: [25.0, 23.33, 23.33],
      amara_pct: [25.0, 23.33, 23.33],
      urooj_pct: [0.0, 23.33, 0.0]
    },
    partner_earnings: {
      labels: ["Shaista", "Amara", "Urooj"],
      direct: [sh_total_direct, am_total_direct, ur_total_direct],
      savings: [sh_company_savings, am_company_savings, ur_company_savings],
      total_value: [sh_total_direct + sh_company_savings, am_total_direct + am_company_savings, ur_total_direct + ur_company_savings],
      deductions: [sh_skipped_deduction, am_skipped_deduction, ur_skipped_deduction]
    },
    revenue_breakdown: {
      labels: ["Company Net Profit", "Purchases", "Send Money", "Team Payouts"],
      values: [all_net, total_purchases, total_send, old_team + new_team]
    }
  };

  return {
    dashboard: {
      old_rev,
      old_co,
      old_team,
      old_purchases,
      old_net,
      old_projects: old_projects_count,
      old_received: old_received_count,
      old_pending: old_pending_count,
      old_pending_str: old_pending_amount,
      old_send,
      new_rev,
      new_co,
      new_team,
      new_purchases,
      new_net,
      new_projects: new_projects_count,
      new_received: new_received_count,
      new_pending: new_pending_count,
      new_pending_str: new_pending_amount,
      new_send,
      all_rev,
      all_net,
      all_projects: all_projects_count,
      all_received: all_received_count,
      all_pending: all_pending_count,
      all_pending_str: all_pending_amount
    },
    projects,
    split_old,
    split_new,
    expenses_old,
    expenses_new,
    send_money_old,
    send_money_new,
    scenario_c,
    scenario_c_summary: {
      total_amount: sc_total_amt,
      shaista: sc_shaista_dir,
      amara: sc_amara_dir,
      urooj: sc_urooj_dir,
      skip_to_co: sc_skip_co,
      total_co: sc_total_co
    },
    partners,
    company_statement,
    chart_data,
    users: data.users || [],
    credentials: data.credentials || [],
    activity_logs: filtered_logs,
    total_logs_count: all_logs.length
  };
}

function parseJsonBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => {
      body += chunk.toString();
    });
    req.on('end', () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (e) {
        reject(e);
      }
    });
    req.on('error', reject);
  });
}

function sendJson(res, statusCode, data) {
  res.writeHead(statusCode, {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  });
  res.end(JSON.stringify(data));
}

const server = http.createServer(async (req, res) => {
  if (isVercel) await syncFromBlob();
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;
  const method = req.method.toUpperCase();

  // CORS preflight
  if (method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
    return res.end();
  }

  try {
    // 1. Root / Index Page
    if (pathname === '/' || pathname === '/index.html') {
      const htmlPath = path.join(BASE_DIR, 'index.html');
      if (fs.existsSync(htmlPath)) {
        const content = fs.readFileSync(htmlPath);
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        return res.end(content);
      } else {
        res.writeHead(404, { 'Content-Type': 'text/plain' });
        return res.end('index.html not found');
      }
    }

    // 1b. Static Logo & Images & Favicon
    if (pathname === '/logo.png' || pathname === '/favicon.ico' || pathname === '/favicon.png' || pathname.endsWith('.png') || pathname.endsWith('.jpg') || pathname.endsWith('.jpeg') || pathname.endsWith('.svg') || pathname.endsWith('.ico')) {
      let filename = path.basename(pathname);
      if (pathname === '/favicon.ico' || pathname === '/favicon.png') {
        filename = 'favicon.png';
      }
      const filePath = path.join(BASE_DIR, filename);
      if (fs.existsSync(filePath)) {
        const ext = path.extname(filePath).toLowerCase();
        const mimeTypes = {
          '.png': 'image/png',
          '.jpg': 'image/jpeg',
          '.jpeg': 'image/jpeg',
          '.svg': 'image/svg+xml',
          '.ico': 'image/x-icon'
        };
        const img = fs.readFileSync(filePath);
        res.writeHead(200, { 
          'Content-Type': (pathname === '/favicon.ico') ? 'image/x-icon' : (mimeTypes[ext] || 'image/png'),
          'Cache-Control': (pathname.includes('favicon')) ? 'no-cache, no-store, must-revalidate' : 'public, max-age=86400'
        });
        return res.end(img);
      } else {
        res.writeHead(404, { 'Content-Type': 'text/plain' });
        return res.end('Image not found');
      }
    }

    // 2. Download Excel
    if (pathname === '/download-excel') {
      if (fs.existsSync(EXCEL_PATH)) {
        const fileStream = fs.createReadStream(EXCEL_PATH);
        res.writeHead(200, {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'Content-Disposition': 'attachment; filename="Hexalevel_Financial_Management_Dashboard.xlsx"'
        });
        return fileStream.pipe(res);
      } else {
        return sendJson(res, 404, { detail: 'Excel file not found' });
      }
    }

    // 3. Auth Suite: Direct Email & Password Login, Forgot Password, and Invitation System
    // 3a. Login
    if (pathname === '/api/auth/login' && method === 'POST') {
      const creds = await parseJsonBody(req);
      const query = (creds.email || '').trim().toLowerCase();
      const pass = (creds.password || '').trim();
      const data = loadData();
      const users = data.users || [];

      let user = users.find(u => (u.email || '').toLowerCase() === query || (u.name || '').toLowerCase() === query);

      // Master Admin fallback if not in list
      if (!user && USERS_DB[query]) {
        user = {
          id: 'USR-001',
          email: query,
          name: USERS_DB[query].name,
          password: USERS_DB[query].passwords[0],
          passwords: USERS_DB[query].passwords,
          role: 'Administrator',
          created_at: "2026-10-31 12:00:00",
          last_login: "2026-10-31 19:00:00",
          login_count: 1,
          status: "Active"
        };
        users.push(user);
        await saveData(data);
      }

      if (!user) {
        return sendJson(res, 401, { detail: "User account not found. Please contact Admin to receive an invitation." });
      }

      if (user.status === "Pending Password Setup" && !user.password) {
        return sendJson(res, 401, { detail: "Your account is pending password setup. Please use the invitation link sent to your email to create your password." });
      }

      const validPasswords = [user.password, ...(user.passwords || [])].filter(Boolean);
      if (user.role === 'Administrator' || user.email === 'admin@hexalevel.com') {
        validPasswords.push('admin', 'admin123');
      }
      if (!validPasswords.includes(pass)) {
        return sendJson(res, 401, { detail: "Incorrect password. Click 'Forgot Password?' to reset it." });
      }

      const now = new Date();
      const pad = (n) => String(n).padStart(2, '0');
      const timestamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

      user.last_login = timestamp;
      user.login_count = (user.login_count || 0) + 1;
      logActivity(data, user.name, "USER_LOGIN", "Authentication", `${user.name} logged in with email ${user.email}.`);
      await saveData(data);

      return sendJson(res, 200, {
        status: "success",
        user: user.name,
        email: user.email,
        role: user.role || "Member"
      });
    }

    // 3b. Forgot Password Reset
    if (pathname === '/api/auth/forgot-password' && method === 'POST') {
      const body = await parseJsonBody(req);
      const query = (body.email || '').trim().toLowerCase();
      const newPass = (body.new_password || '').trim();

      if (!query || !newPass) {
        return sendJson(res, 400, { detail: "Please provide your email address and a new password." });
      }

      const data = loadData();
      if (!data.users) data.users = [];

      let user = data.users.find(u => (u.email || '').toLowerCase() === query || (u.name || '').toLowerCase() === query);

      if (!user) {
        return sendJson(res, 404, { detail: "No account found with this email. Please contact Admin to receive an invitation." });
      }

      const now = new Date();
      const pad = (n) => String(n).padStart(2, '0');
      const timestamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

      user.password = newPass;
      if (!user.passwords) user.passwords = [];
      if (!user.passwords.includes(newPass)) user.passwords.unshift(newPass);
      user.status = "Active";
      user.last_login = timestamp;

      logActivity(data, user.name, "PASSWORD_RESET", "Authentication", `${user.name} set/reset their password for ${user.email}.`);
      await saveData(data);

      return sendJson(res, 200, {
        status: "success",
        message: "Password updated successfully! You can now sign in.",
        user: user.name,
        email: user.email,
        role: user.role || "Member"
      });
    }

    // 3c. Set Password via Invitation Email Link
    if (pathname === '/api/auth/set-password' && method === 'POST') {
      const body = await parseJsonBody(req);
      const email = (body.email || '').trim().toLowerCase();
      const newPass = (body.password || '').trim();

      if (!email || !newPass) {
        return sendJson(res, 400, { detail: "Email and password are required." });
      }

      const data = loadData();
      if (!data.users) data.users = [];

      let user = data.users.find(u => (u.email || '').toLowerCase() === email);
      if (!user) {
        return sendJson(res, 404, { detail: "Invitation record not found. Please contact Admin." });
      }

      const now = new Date();
      const pad = (n) => String(n).padStart(2, '0');
      const timestamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

      user.password = newPass;
      if (!user.passwords) user.passwords = [];
      if (!user.passwords.includes(newPass)) user.passwords.unshift(newPass);
      user.status = "Active";
      user.last_login = timestamp;
      user.login_count = (user.login_count || 0) + 1;

      logActivity(data, user.name, "USER_ACTIVATED", "Users Management", `${user.name} activated account and set custom password via email notification.`);
      await saveData(data);

      return sendJson(res, 200, {
        status: "success",
        message: "Account activated successfully!",
        user: user.name,
        email: user.email,
        role: user.role || "Member"
      });
    }

    // 3d. Change Password from inside Dashboard
    if (pathname === '/api/auth/change-password' && method === 'POST') {
      const body = await parseJsonBody(req);
      const query = (body.email || '').trim().toLowerCase();
      const newPass = (body.new_password || '').trim();

      if (!newPass) {
        return sendJson(res, 400, { detail: "New password cannot be empty." });
      }

      const data = loadData();
      if (!data.users) data.users = [];

      let user = data.users.find(u => (u.email || '').toLowerCase() === query || (u.name || '').toLowerCase() === query);
      if (!user) {
        return sendJson(res, 404, { detail: "User account not found." });
      }

      user.password = newPass;
      if (!user.passwords) user.passwords = [];
      if (!user.passwords.includes(newPass)) user.passwords.unshift(newPass);

      logActivity(data, user.name, "PASSWORD_CHANGE", "Authentication", `${user.name} changed their account password.`);
      await saveData(data);

      return sendJson(res, 200, {
        status: "success",
        message: "Password updated successfully!"
      });
    }

    // 3dd. Change Account Email Address
    if (pathname === '/api/auth/change-email' && method === 'POST') {
      const body = await parseJsonBody(req);
      const query = (body.current_email || body.user || '').trim().toLowerCase();
      const newEmail = (body.new_email || '').trim().toLowerCase();
      const password = (body.password || '').trim();

      if (!newEmail || !newEmail.includes('@')) {
        return sendJson(res, 400, { detail: "Please provide a valid new email address." });
      }

      const data = loadData();
      const users = data.users || [];

      let user = users.find(u => (u.email || '').toLowerCase() === query || (u.name || '').toLowerCase() === query);
      if (!user) {
        return sendJson(res, 404, { detail: "User account not found." });
      }

      const validPasswords = [user.password, ...(user.passwords || [])].filter(Boolean);
      if (user.role === 'Administrator' || user.email === 'admin@hexalevel.com') {
        validPasswords.push('admin', 'admin123');
      }
      if (!validPasswords.includes(password)) {
        return sendJson(res, 401, { detail: "Incorrect password. Please verify your current password to change email." });
      }

      const existing = users.find(u => u.id !== user.id && (u.email || '').toLowerCase() === newEmail);
      if (existing) {
        return sendJson(res, 400, { detail: "This email address is already in use by another member." });
      }

      const oldEmail = user.email;
      user.email = newEmail;
      logActivity(data, user.name, "CHANGE_EMAIL", "Profile", `${user.name} changed account email from ${oldEmail} to ${newEmail}.`);
      await saveData(data);

      return sendJson(res, 200, {
        status: "success",
        message: "Email address updated successfully!",
        email: newEmail
      });
    }

    // 3e. Users Directory (Admin View)
    if (pathname === '/api/users' && method === 'GET') {
      const data = loadData();
      return sendJson(res, 200, {
        users: data.users || []
      });
    }

    // 3f. Admin Invites / Adds User (Sends Email Notification)
    if (pathname === '/api/user' && method === 'POST') {
      const body = await parseJsonBody(req);
      const name = smartCapitalize((body.name || '').trim());
      const email = (body.email || '').trim().toLowerCase();
      const role = body.role || 'Member';

      if (!name || !email) {
        return sendJson(res, 400, { detail: "Please provide member name and email address." });
      }

      const data = loadData();
      if (!data.users) data.users = [];

      const existing = data.users.find(u => (u.email || '').toLowerCase() === email);
      if (existing) {
        return sendJson(res, 400, { detail: "A user with this email address already exists in the system." });
      }

      const now = new Date();
      const pad = (n) => String(n).padStart(2, '0');
      const timestamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
      const token = crypto.randomBytes(16).toString('hex');

      const newUser = {
        id: `USR-${String(data.users.length + 1).padStart(3, '0')}`,
        name,
        email,
        password: "", // pending setup by user
        passwords: [],
        created_at: timestamp,
        last_login: "—",
        login_count: 0,
        status: "Pending Password Setup",
        invite_token: token,
        role
      };

      data.users.push(newUser);
      
      const baseUrl = getEffectiveBaseUrl(req, data);
      const inviteLink = `${baseUrl}/?action=set-password&email=${encodeURIComponent(email)}&token=${token}`;

      // Dispatch real email via SMTP if configured
      const emailResult = await sendInvitationEmail(newUser, inviteLink, data.email_settings);
      
      logActivity(data, "Admin", "INVITE_USER", "Users Management", `Invitation notification for ${name} (${email}). Real Email Sent: ${emailResult.sent ? 'YES' : 'NO (' + (emailResult.error || emailResult.reason) + ')'}`);
      await saveData(data);

      return sendJson(res, 201, {
        status: "success",
        user: newUser,
        invite_link: inviteLink,
        email_sent: emailResult.sent,
        email_error: emailResult.error || null,
        needs_config: (!emailResult.sent && (!data.email_settings || !data.email_settings.sender_password)),
        message: emailResult.sent
          ? `Real invitation email delivered directly to ${email}!`
          : `User registered. Real email requires SMTP configuration.`
      });
    }

    // 3g. Admin Resends Invitation Email
    if (pathname.startsWith('/api/user/') && pathname.endsWith('/resend-invite') && method === 'POST') {
      const uid = pathname.replace('/api/user/', '').replace('/resend-invite', '');
      const data = loadData();
      const users = data.users || [];
      const user = users.find(u => u.id === uid || (u.email || '').toLowerCase() === uid.toLowerCase());

      if (!user) {
        return sendJson(res, 404, { detail: "User not found." });
      }

      const token = crypto.randomBytes(16).toString('hex');
      user.invite_token = token;
      const baseUrl = getEffectiveBaseUrl(req, data);
      const inviteLink = `${baseUrl}/?action=set-password&email=${encodeURIComponent(user.email)}&token=${token}`;

      const emailResult = await sendInvitationEmail(user, inviteLink, data.email_settings);
      logActivity(data, "Admin", "RESEND_INVITE", "Users Management", `Resent invitation to ${user.email}. Real Email Sent: ${emailResult.sent ? 'YES' : 'NO'}`);
      await saveData(data);

      return sendJson(res, 200, {
        status: "success",
        invite_link: inviteLink,
        email_sent: emailResult.sent,
        email_error: emailResult.error || null,
        needs_config: (!emailResult.sent && (!data.email_settings || !data.email_settings.sender_password)),
        message: emailResult.sent ? `Real invitation email resent to ${user.email}.` : `Invitation link refreshed.`
      });
    }

    // 3h. Admin Edits User
    if (pathname.startsWith('/api/user/') && !pathname.endsWith('/resend-invite') && method === 'PUT') {
      const uid = pathname.replace('/api/user/', '');
      const body = await parseJsonBody(req);
      const data = loadData();
      const users = data.users || [];

      const user = users.find(u => u.id === uid || (u.email || '').toLowerCase() === uid.toLowerCase());
      if (!user) {
        return sendJson(res, 404, { detail: "User not found." });
      }

      if (body.name) user.name = smartCapitalize(body.name.trim());
      if (body.email) user.email = body.email.trim().toLowerCase();
      if (body.role) user.role = body.role;
      if (body.password && body.password.trim()) {
        user.password = body.password.trim();
        if (!user.passwords) user.passwords = [];
        if (!user.passwords.includes(body.password.trim())) user.passwords.unshift(body.password.trim());
        user.status = "Active";
      }

      logActivity(data, "Admin", "EDIT_USER", "Users Management", `Admin updated user account for "${user.name}" (${user.email}).`);
      await saveData(data);

      return sendJson(res, 200, { status: "success", user });
    }

    // 3i. Admin Deletes User
    if (pathname.startsWith('/api/user/') && method === 'DELETE') {
      const uid = pathname.replace('/api/user/', '');
      const data = loadData();
      const users = data.users || [];

      const target = users.find(u => u.id === uid || (u.email || '').toLowerCase() === uid.toLowerCase());
      if (!target) {
        return sendJson(res, 404, { detail: "User not found." });
      }

      if (target.role === 'Administrator' || (target.email || '').toLowerCase() === 'admin@hexalevel.com') {
        return sendJson(res, 400, { detail: "Cannot delete the primary Administrator account." });
      }

      data.users = users.filter(u => u.id !== target.id);
      logActivity(data, "Admin", "DELETE_USER", "Users Management", `Admin deleted user account "${target.name}" (${target.email}).`);
      await saveData(data);

      return sendJson(res, 200, { status: "success", deleted: target.id });
    }

    // 3j. Email Settings (Get)
    if (pathname === '/api/email/settings' && method === 'GET') {
      const data = loadData();
      const s = data.email_settings || {};
      return sendJson(res, 200, {
        sender_email: s.sender_email || '',
        sender_name: s.sender_name || 'HEXALEVEL',
        smtp_host: s.smtp_host || 'smtp.gmail.com',
        smtp_port: s.smtp_port || 465,
        smtp_secure: (s.smtp_secure !== undefined) ? s.smtp_secure : true,
        base_url: s.base_url || '',
        is_configured: !!(s.sender_email && s.sender_password)
      });
    }

    // 3k. Email Settings (Save)
    if (pathname === '/api/email/settings' && method === 'POST') {
      const body = await parseJsonBody(req);
      const data = loadData();
      if (!data.email_settings) data.email_settings = {};

      if (body.sender_email !== undefined) data.email_settings.sender_email = body.sender_email.trim();
      if (body.sender_name !== undefined) data.email_settings.sender_name = body.sender_name.trim();
      if (body.sender_password !== undefined && body.sender_password.trim() !== '') {
        data.email_settings.sender_password = body.sender_password.trim().replace(/\s+/g, '');
      }
      if (body.smtp_host !== undefined) data.email_settings.smtp_host = body.smtp_host.trim();
      if (body.smtp_port !== undefined) data.email_settings.smtp_port = parseInt(body.smtp_port) || 465;
      if (body.smtp_secure !== undefined) data.email_settings.smtp_secure = Boolean(body.smtp_secure);
      if (body.base_url !== undefined) data.email_settings.base_url = body.base_url.trim();

      logActivity(data, "Admin", "UPDATE_EMAIL_SETTINGS", "Settings", `Admin updated SMTP email settings (${data.email_settings.sender_email}).`);
      await saveData(data);

      return sendJson(res, 200, {
        status: "success",
        message: "Email settings saved successfully!",
        is_configured: !!(data.email_settings.sender_email && data.email_settings.sender_password)
      });
    }

    // 3l. Test Email Connection
    if (pathname === '/api/email/test' && method === 'POST') {
      const body = await parseJsonBody(req);
      const data = loadData();
      const settings = {
        ...(data.email_settings || {}),
        ...(body || {})
      };
      if (body.sender_password) settings.sender_password = body.sender_password.trim().replace(/\s+/g, '');

      if (!settings.sender_email || !settings.sender_password) {
        return sendJson(res, 400, { detail: "Sender email and password (or App Password) are required." });
      }

      const transporter = getEmailTransporter(settings);
      if (!transporter) {
        return sendJson(res, 400, { detail: "Invalid SMTP configuration." });
      }

      try {
        await transporter.verify();
        const testTo = body.test_recipient || settings.sender_email;
        const info = await transporter.sendMail({
          from: `"${settings.sender_name || 'HEXALEVEL Dashboard'}" <${settings.sender_email.trim()}>`,
          to: testTo,
          subject: '✅ HEXALEVEL Dashboard: Email Connection Test Successful',
          text: 'Congratulations! Your SMTP email connection is configured correctly and ready to dispatch real invitations.',
          html: '<div style="font-family:sans-serif; padding:20px; color:#0f172a;"><h2 style="color:#0d5c52;">✅ Connection Successful!</h2><p>Your Hexalevel Dashboard email settings are working perfectly. Automated invitation emails will now be sent directly to new members.</p></div>'
        });
        return sendJson(res, 200, {
          status: "success",
          message: `Test email successfully delivered to ${testTo}!`,
          messageId: info.messageId
        });
      } catch (err) {
        return sendJson(res, 400, {
          status: "error",
          detail: `SMTP Connection failed: ${err.message}`
        });
      }
    }

    // 4. Get Dashboard Data
    if (pathname === '/api/data' && method === 'GET') {
      const filterUser = parsedUrl.query.filter_user;
      const dashboardData = calculateDashboardData(filterUser);
      return sendJson(res, 200, dashboardData);
    }

    // 5. Project CRUD
    if (pathname === '/api/project' && method === 'POST') {
      const p = await parseJsonBody(req);
      const data = loadData();
      const pid = `PRJ-${String((data.projects || []).length + 1).padStart(3, '0')}`;
      const userName = p.user ? p.user.replace(/\(Co-Founder\)/g, '').replace(/\(Partner\)/g, '').trim() : "Admin";

      const newProj = {
        id: pid,
        start: formatProjectStart(p.start) || smartCapitalize(p.start) || '',
        client: smartCapitalize(p.client) || '',
        project: smartCapitalize(p.project) || '',
        service: smartCapitalize(p.service) || '',
        amount: parseFloat(p.amount) || 0,
        currency: (p.currency || 'PKR').toUpperCase(),
        status: smartCapitalize(p.status) || 'Received',
        date: formatPaymentDate(p.date) || p.date || '',
        notes: smartCapitalize(p.notes) || '',
        bank: smartCapitalize(p.bank) || '',
        period: p.period || 'new',
        scenario: p.scenario || ''
      };
      if (!data.projects) data.projects = [];
      data.projects.push(newProj);

      const amt = newProj.amount;
      const monthName = newProj.start || (newProj.date ? newProj.date.substring(3) : '');
      const shares = parseScenarioShares(p.scenario, amt);

      if (p.period === 'old') {
        if (!data.split_old) data.split_old = [];
        data.split_old.push({
          id: pid,
          month: monthName,
          amount: amt,
          company: shares.company,
          team: shares.shaista + shares.amara,
          status: newProj.status,
          date: newProj.date,
          shaista: p.bank || "Nayapay",
          amara: "Askari",
          tax: ""
        });
        logActivity(data, userName, "ADD_PROJECT", "Projects & Revenue Split",
          `${userName} added project ${pid} '${p.client}' (PKR ${amt.toLocaleString()}) into Upper 50/50 Table.`);
      } else {
        if (!data.split_new) data.split_new = [];
        if (!data.scenario_c) data.scenario_c = [];

        if ((p.scenario || '').includes("Scenario C")) {
          const skipName = shares.skip_member;
          const shShare = shares.shaista;
          const amShare = shares.amara;
          const urShare = shares.urooj;
          const skipCo = shares.skip_co;
          const totCo = shares.company;

          data.scenario_c.push({
            id: pid,
            client: `${p.client} (${p.project})`,
            amount: amt,
            status: p.status,
            scenario: p.scenario,
            shaista: shShare,
            amara: amShare,
            urooj: urShare,
            skip_member: skipName,
            skip_to_co: skipCo,
            total_co: totCo,
            notes: `${skipName} skipped -> unallocated 23.33% (PKR ${skipCo.toLocaleString()}) to Company`
          });

          data.split_new.push({
            id: pid,
            month: monthName,
            amount: amt,
            company: totCo,
            shaista: shShare,
            amara: amShare,
            urooj: urShare,
            status: p.status,
            date: p.date,
            shaista_acc: shShare > 0 ? "Nayapay" : "SKIP (0.00)",
            amara_acc: amShare > 0 ? "Askari" : "SKIP (0.00)",
            urooj_acc: urShare > 0 ? "UBL" : "SKIP (0.00)",
            tax: "",
            notes: `Scenario C: ${skipName} Skipped -> Co 53.33%`,
            scenario: p.scenario
          });

          logActivity(data, userName, "ADD_PROJECT", "Projects, Revenue Split & Scenario C",
            `${userName} added Scenario C project ${pid} '${p.client}' (PKR ${amt.toLocaleString()}) where ${skipName} was SKIPPED.`);
        } else {
          data.split_new.push({
            id: pid,
            month: monthName,
            amount: amt,
            company: shares.company,
            shaista: shares.shaista,
            amara: shares.amara,
            urooj: shares.urooj,
            status: p.status,
            date: p.date,
            shaista_acc: "Nayapay",
            amara_acc: "Askari",
            urooj_acc: "UBL",
            tax: "",
            notes: p.notes || "",
            scenario: p.scenario
          });
          logActivity(data, userName, "ADD_PROJECT", "Projects & Revenue Split",
            `${userName} added project ${pid} '${p.client}' (PKR ${amt.toLocaleString()}) into Lower 30/70 Table.`);
        }
      }

      await saveData(data);
      return sendJson(res, 200, { status: "success", id: pid });
    }

    if (pathname.startsWith('/api/project/') && method === 'PUT') {
      const pid = pathname.replace('/api/project/', '');
      const p = await parseJsonBody(req);
      const data = loadData();
      const idx = (data.projects || []).findIndex(item => item.id === pid);
      if (idx === -1) {
        return sendJson(res, 404, { detail: "Project not found" });
      }

      const userName = p.user ? p.user.replace(/\(Co-Founder\)/g, '').replace(/\(Partner\)/g, '').trim() : "Admin";

      data.projects[idx] = {
        id: pid,
        start: formatProjectStart(p.start) || smartCapitalize(p.start) || '',
        client: smartCapitalize(p.client) || '',
        project: smartCapitalize(p.project) || '',
        service: smartCapitalize(p.service) || '',
        amount: parseFloat(p.amount) || 0,
        currency: (p.currency || 'PKR').toUpperCase(),
        status: smartCapitalize(p.status) || 'Received',
        date: formatPaymentDate(p.date) || p.date || '',
        notes: smartCapitalize(p.notes) || '',
        bank: smartCapitalize(p.bank) || '',
        period: p.period || 'new',
        scenario: p.scenario || ''
      };

      data.split_old = (data.split_old || []).filter(x => x.id !== pid);
      data.split_new = (data.split_new || []).filter(x => x.id !== pid);
      data.scenario_c = (data.scenario_c || []).filter(x => x.id !== pid);

      const amt = data.projects[idx].amount;
      const monthName = data.projects[idx].start || (data.projects[idx].date ? data.projects[idx].date.substring(3) : '');
      const shares = parseScenarioShares(p.scenario, amt);

      if (p.period === 'old') {
        data.split_old.push({
          id: pid,
          month: monthName,
          amount: amt,
          company: shares.company,
          team: shares.shaista + shares.amara,
          status: p.status,
          date: p.date,
          shaista: p.bank || "Nayapay",
          amara: "Askari",
          tax: ""
        });
      } else {
        if ((p.scenario || '').includes("Scenario C")) {
          const skipName = shares.skip_member;
          const shShare = shares.shaista;
          const amShare = shares.amara;
          const urShare = shares.urooj;
          const skipCo = shares.skip_co;
          const totCo = shares.company;

          data.scenario_c.push({
            id: pid,
            client: `${p.client} (${p.project})`,
            amount: amt,
            status: p.status,
            scenario: p.scenario,
            shaista: shShare,
            amara: amShare,
            urooj: urShare,
            skip_member: skipName,
            skip_to_co: skipCo,
            total_co: totCo,
            notes: `${skipName} skipped -> unallocated 23.33% to Co`
          });
          data.split_new.push({
            id: pid,
            month: monthName,
            amount: amt,
            company: totCo,
            shaista: shShare,
            amara: amShare,
            urooj: urShare,
            status: p.status,
            date: p.date,
            shaista_acc: shShare > 0 ? "Nayapay" : "SKIP (0.00)",
            amara_acc: amShare > 0 ? "Askari" : "SKIP (0.00)",
            urooj_acc: urShare > 0 ? "UBL" : "SKIP (0.00)",
            tax: "",
            notes: `Scenario C: ${skipName} Skipped -> Co 53.33%`,
            scenario: p.scenario
          });
        } else {
          data.split_new.push({
            id: pid,
            month: monthName,
            amount: amt,
            company: shares.company,
            shaista: shares.shaista,
            amara: shares.amara,
            urooj: shares.urooj,
            status: p.status,
            date: p.date,
            shaista_acc: "Nayapay",
            amara_acc: "Askari",
            urooj_acc: "UBL",
            tax: "",
            notes: p.notes || "",
            scenario: p.scenario
          });
        }
      }

      logActivity(data, userName, "EDIT_PROJECT", "Projects & Revenue Split",
        `${userName} edited project ${pid} '${p.client}' (PKR ${amt.toLocaleString()}).`);
      await saveData(data);
      return sendJson(res, 200, { status: "success" });
    }

    if (pathname.startsWith('/api/project/') && method === 'DELETE') {
      const pid = pathname.replace('/api/project/', '');
      const user = parsedUrl.query.user || "Admin";
      const userName = user ? user.replace(/\(Co-Founder\)/g, '').replace(/\(Partner\)/g, '').trim() : "Admin";
      const data = loadData();
      const projToDelete = (data.projects || []).find(p => p.id === pid);
      const clientName = projToDelete ? projToDelete.client : pid;

      data.projects = (data.projects || []).filter(p => p.id !== pid);
      data.split_old = (data.split_old || []).filter(x => x.id !== pid);
      data.split_new = (data.split_new || []).filter(x => x.id !== pid);
      data.scenario_c = (data.scenario_c || []).filter(c => c.id !== pid);

      logActivity(data, userName, "DELETE_PROJECT", "Projects & Revenue Split",
        `${userName} deleted project ${pid} '${clientName}' and synchronized records.`);
      await saveData(data);
      return sendJson(res, 200, { status: "success" });
    }

    // 6. Expense CRUD
    if (pathname === '/api/expense' && method === 'POST') {
      const e = await parseJsonBody(req);
      const data = loadData();
      const userName = e.user ? e.user.replace(/\(Co-Founder\)/g, '').replace(/\(Partner\)/g, '').trim() : "Admin";
      let eid, tabName;
      if (e.period === 'old') {
        if (!data.expenses_old) data.expenses_old = [];
        eid = `EXP-OLD-${String(data.expenses_old.length + 1).padStart(3, '0')}`;
        data.expenses_old.push({ id: eid, item: smartCapitalize(e.item), link: e.link || '', price: parseFloat(e.price) || 0 });
        tabName = "Purchases (Upper)";
      } else {
        if (!data.expenses_new) data.expenses_new = [];
        eid = `EXP-NEW-${String(data.expenses_new.length + 1).padStart(3, '0')}`;
        data.expenses_new.push({ id: eid, item: e.item, link: e.link || '', price: parseFloat(e.price) || 0 });
        tabName = "Purchases (Lower)";
      }
      logActivity(data, userName, "ADD_PURCHASE", tabName,
        `${userName} added purchase '${e.item}' (PKR ${(parseFloat(e.price) || 0).toLocaleString()}).`);
      await saveData(data);
      return sendJson(res, 200, { status: "success", id: eid });
    }

    if (pathname.startsWith('/api/expense/') && method === 'PUT') {
      const eid = pathname.replace('/api/expense/', '');
      const e = await parseJsonBody(req);
      const data = loadData();
      const userName = e.user ? e.user.replace(/\(Co-Founder\)/g, '').replace(/\(Partner\)/g, '').trim() : "Admin";

      let found = false;
      for (let i = 0; i < (data.expenses_old || []).length; i++) {
        if (data.expenses_old[i].id === eid) {
          data.expenses_old[i] = { id: eid, item: e.item, link: e.link || '', price: parseFloat(e.price) || 0 };
          logActivity(data, userName, "EDIT_PURCHASE", "Purchases (Upper)", `${userName} edited purchase ${eid} '${e.item}'.`);
          found = true;
          break;
        }
      }
      if (!found) {
        for (let i = 0; i < (data.expenses_new || []).length; i++) {
          if (data.expenses_new[i].id === eid) {
            data.expenses_new[i] = { id: eid, item: e.item, link: e.link || '', price: parseFloat(e.price) || 0 };
            logActivity(data, userName, "EDIT_PURCHASE", "Purchases (Lower)", `${userName} edited purchase ${eid} '${e.item}'.`);
            found = true;
            break;
          }
        }
      }

      if (!found) {
        return sendJson(res, 404, { detail: "Expense not found" });
      }

      await saveData(data);
      return sendJson(res, 200, { status: "success" });
    }

    if (pathname.startsWith('/api/expense/') && method === 'DELETE') {
      const eid = pathname.replace('/api/expense/', '');
      const user = parsedUrl.query.user || "Admin";
      const userName = user ? user.replace(/\(Co-Founder\)/g, '').replace(/\(Partner\)/g, '').trim() : "Admin";
      const data = loadData();

      data.expenses_old = (data.expenses_old || []).filter(x => x.id !== eid);
      data.expenses_new = (data.expenses_new || []).filter(x => x.id !== eid);

      logActivity(data, userName, "DELETE_PURCHASE", "Purchases", `${userName} deleted purchase ${eid}.`);
      await saveData(data);
      return sendJson(res, 200, { status: "success" });
    }

    // 7. Send Money CRUD
    if (pathname === '/api/sendmoney' && method === 'POST') {
      const s = await parseJsonBody(req);
      const data = loadData();
      const userName = s.user ? s.user.replace(/\(Co-Founder\)/g, '').replace(/\(Partner\)/g, '').trim() : "Admin";
      let sid, tabName;
      if (s.period === 'old') {
        if (!data.send_money_old) data.send_money_old = [];
        sid = `SEND-OLD-${String(data.send_money_old.length + 1).padStart(3, '0')}`;
        data.send_money_old.push({ id: sid, recipient: smartCapitalize(s.recipient), amount: parseFloat(s.amount) || 0 });
        tabName = "Send Money (Upper)";
      } else {
        if (!data.send_money_new) data.send_money_new = [];
        sid = `SEND-NEW-${String(data.send_money_new.length + 1).padStart(3, '0')}`;
        data.send_money_new.push({ id: sid, recipient: s.recipient, amount: parseFloat(s.amount) || 0 });
        tabName = "Send Money (Lower)";
      }
      logActivity(data, userName, "ADD_SEND_MONEY", tabName,
        `${userName} sent money to '${s.recipient}' (PKR ${(parseFloat(s.amount) || 0).toLocaleString()}).`);
      await saveData(data);
      return sendJson(res, 200, { status: "success", id: sid });
    }

    if (pathname.startsWith('/api/sendmoney/') && method === 'PUT') {
      const sid = pathname.replace('/api/sendmoney/', '');
      const s = await parseJsonBody(req);
      const data = loadData();
      const userName = s.user ? s.user.replace(/\(Co-Founder\)/g, '').replace(/\(Partner\)/g, '').trim() : "Admin";

      let found = false;
      for (let i = 0; i < (data.send_money_old || []).length; i++) {
        if (data.send_money_old[i].id === sid) {
          data.send_money_old[i] = { id: sid, recipient: s.recipient, amount: parseFloat(s.amount) || 0 };
          logActivity(data, userName, "EDIT_SEND_MONEY", "Send Money", `${userName} edited payment to '${s.recipient}'.`);
          found = true;
          break;
        }
      }
      if (!found) {
        for (let i = 0; i < (data.send_money_new || []).length; i++) {
          if (data.send_money_new[i].id === sid) {
            data.send_money_new[i] = { id: sid, recipient: s.recipient, amount: parseFloat(s.amount) || 0 };
            logActivity(data, userName, "EDIT_SEND_MONEY", "Send Money", `${userName} edited payment to '${s.recipient}'.`);
            found = true;
            break;
          }
        }
      }

      if (!found) {
        return sendJson(res, 404, { detail: "Record not found" });
      }

      await saveData(data);
      return sendJson(res, 200, { status: "success" });
    }

    if (pathname.startsWith('/api/sendmoney/') && method === 'DELETE') {
      const sid = pathname.replace('/api/sendmoney/', '');
      const user = parsedUrl.query.user || "Admin";
      const userName = user ? user.replace(/\(Co-Founder\)/g, '').replace(/\(Partner\)/g, '').trim() : "Admin";
      const data = loadData();

      data.send_money_old = (data.send_money_old || []).filter(x => x.id !== sid);
      data.send_money_new = (data.send_money_new || []).filter(x => x.id !== sid);

      logActivity(data, userName, "DELETE_SEND_MONEY", "Send Money", `${userName} deleted payment ${sid}.`);
      await saveData(data);
      return sendJson(res, 200, { status: "success" });
    }


    // ----------------------------------------------------------
    // 8. Website Credentials & Passwords CRUD Endpoints
    // ----------------------------------------------------------
    if (pathname === '/api/credential' && method === 'POST') {
      const item = await parseJsonBody(req);
      const data = loadData();
      if (!data.credentials) data.credentials = [];

      const newId = 'CRED-' + String(data.credentials.length + 1).padStart(3, '0');
      const now = new Date();
      const pad = (n) => String(n).padStart(2, '0');
      const timestamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

      const newCred = {
        id: newId,
        website_name: (item.website_name || '').trim(),
        login_url: (item.login_url || '').trim(),
        user_name: (item.user_name || '').trim(),
        password: (item.password || '').trim(),
        category: (item.category || 'General').trim(),
        notes: (item.notes || '').trim(),
        created_at: timestamp
      };

      data.credentials.push(newCred);
      logActivity(data, "Admin", "ADD_CREDENTIAL", "Website Credentials", `Added login credential for "${newCred.website_name}" (${newCred.user_name || 'No username'}).`);
      await saveData(data);

      return sendJson(res, 201, { status: "success", credential: newCred });
    }

    if (pathname.startsWith('/api/credential/') && method === 'PUT') {
      const credId = pathname.replace('/api/credential/', '');
      const updates = await parseJsonBody(req);
      const data = loadData();
      if (!data.credentials) data.credentials = [];

      const index = data.credentials.findIndex(c => c.id === credId);
      if (index === -1) {
        return sendJson(res, 404, { detail: "Credential not found." });
      }

      data.credentials[index] = {
        ...data.credentials[index],
        website_name: (updates.website_name !== undefined ? updates.website_name : data.credentials[index].website_name).trim(),
        login_url: (updates.login_url !== undefined ? updates.login_url : data.credentials[index].login_url).trim(),
        user_name: (updates.user_name !== undefined ? updates.user_name : data.credentials[index].user_name).trim(),
        password: (updates.password !== undefined ? updates.password : data.credentials[index].password).trim(),
        category: (updates.category !== undefined ? updates.category : data.credentials[index].category).trim(),
        notes: (updates.notes !== undefined ? updates.notes : data.credentials[index].notes).trim()
      };

      logActivity(data, "Admin", "UPDATE_CREDENTIAL", "Website Credentials", `Updated login credential for "${data.credentials[index].website_name}".`);
      await saveData(data);

      return sendJson(res, 200, { status: "success", credential: data.credentials[index] });
    }

    if (pathname.startsWith('/api/credential/') && method === 'DELETE') {
      const credId = pathname.replace('/api/credential/', '');
      const data = loadData();
      if (!data.credentials) data.credentials = [];

      const target = data.credentials.find(c => c.id === credId);
      if (!target) {
        return sendJson(res, 404, { detail: "Credential not found." });
      }

      data.credentials = data.credentials.filter(c => c.id !== credId);
      logActivity(data, "Admin", "DELETE_CREDENTIAL", "Website Credentials", `Deleted login credential for "${target.website_name}".`);
      await saveData(data);

      return sendJson(res, 200, { status: "success", deleted: credId });
    }


    // Default 404
    sendJson(res, 404, { detail: "Not found" });
  } catch (err) {
    console.error("Server error:", err);
    sendJson(res, 500, { detail: err.message || "Internal server error" });
  }
});

if (require.main === module) {
  server.listen(PORT, '0.0.0.0', () => {
    console.log(`HEXALEVEL Dashboard server running at http://localhost:${PORT}`);
  });
}

module.exports = server;
