/*************** FluxFilm Backend — Utils ***************/

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("FluxFilm")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ---------- Sheet helpers ----------


function sh_(name) {
  const ss = SpreadsheetApp.openById(CORE_SS_ID);
  const s = ss.getSheetByName(name);
  if (!s) throw new Error(`Missing sheet: ${name}`);
  return s;
}

function bankSh_() {
  if (!BANK_SS_ID || BANK_SS_ID.includes("PASTE_")) throw new Error("BANK_SS_ID not set in Config.gs");
  const ss = SpreadsheetApp.openById(BANK_SS_ID);
  const s = ss.getSheetByName(BANK_SHEET_NAME);
  if (!s) throw new Error(`Missing bank sheet: ${BANK_SHEET_NAME}`);
  return s;
}

function headerIndex_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h||"").trim());
  const idx = {};
  headers.forEach((h,i)=>{ if(h) idx[h]=i; });
  return { headers, idx };
}

function now_() { return new Date(); }

function getSetting_(key, defVal="") {
  const s = sh_(TAB_SETTINGS);
  const data = s.getDataRange().getValues();
  for (let i=1;i<data.length;i++){
    if (String(data[i][0]||"").trim() === key) return data[i][1];
  }
  return defVal;
}

function setCellByHeader_(sheet, rowNum, headerName, value) {
  const { idx } = headerIndex_(sheet);
  if (idx[headerName] == null) throw new Error(`Missing header "${headerName}" in ${sheet.getName()}`);
  sheet.getRange(rowNum, idx[headerName]+1).setValue(value);
}

function getRowByValue_(sheet, headerName, value) {
  const { idx } = headerIndex_(sheet);
  if (idx[headerName] == null) throw new Error(`Missing header "${headerName}" in ${sheet.getName()}`);
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const col = idx[headerName]+1;
  const vals = sheet.getRange(2, col, last-1, 1).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][0]||"").trim() === String(value||"").trim()) return (i+2);
  }
  return null;
}

function ymd_(d) {
  const dt = new Date(d);
  const y = dt.getFullYear();
  const m = ("0"+(dt.getMonth()+1)).slice(-2);
  const day = ("0"+dt.getDate()).slice(-2);
  return `${y}-${m}-${day}`;
}

function addDays_(d, n) {
  const dt = new Date(d);
  dt.setDate(dt.getDate() + Number(n||0));
  return dt;
}

function maskPhone_(p) {
  const s = String(p||"");
  if (s.length <= 4) return "****";
  return s.slice(0,2) + "****" + s.slice(-2);
}

function maskEmail_(e) {
  const s = String(e||"");
  const at = s.indexOf("@");
  if (at <= 1) return "***";
  return s.slice(0,1) + "***" + s.slice(at);
}

function asNumber_(x) {
  if (typeof x === "number") return x;
  const s = String(x||"").replace(/[, ]/g,"").trim();
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function genOrderId_() {
  // ✅ OrderID should be like: FF1234567 (no hyphen)
  // We still keep ORDER_PREFIX as a setting, but we sanitize it to alphanumeric only.
  const rawPrefix = String(getSetting_("ORDER_PREFIX", "FF"));
  const prefix = rawPrefix.replace(/[^a-z0-9]/ig, "");

  // 5 digits from timestamp + 2 random digits => 7 digits
  const stamp = Date.now().toString().slice(-5);
  const rand = Math.floor(Math.random() * 90 + 10);
  return `${prefix || "FF"}${stamp}${rand}`;
}


function genSubId_() {
  return "SUB-" + Date.now().toString().slice(-7) + Math.floor(Math.random()*90+10);
}

function genInvoiceNo_() {
  const prefix = String(getSetting_("INVOICE_PREFIX","INV-"));
  return prefix + Date.now().toString().slice(-6);
}

// Global “spinner-friendly” response helper
function ok_(data) { return Object.assign({ ok:true }, data||{}); }
function bad_(message, meta) { return Object.assign({ ok:false, message }, meta||{}); }



function getPaymentSettings(){
  return ok_({
    upiVpa: String(getSetting_("UPI_VPA","")).trim(),
    payee: String(getSetting_("UPI_PAYEE_NAME","FluxFilm")).trim(),
    currency: String(getSetting_("CURRENCY","INR")).trim()
  });
}


function getCore_() {
  if (!CORE_SS_ID || String(CORE_SS_ID).includes("PASTE_")) {
    throw new Error("CORE_SS_ID not set in Config.gs");
  }
  return SpreadsheetApp.openById(CORE_SS_ID);
}






/*************** FluxFilm Backend — Telegram Notify (robust + auto-enrich) ***************/

/**
 * Sends Telegram alert.
 * If payload is missing name/phone/email, it auto-fetches from ORDERS sheet via orderId.
 * Requires SETTINGS:
 *  - TELEGRAM_BOT_TOKEN
 *  - TELEGRAM_CHAT_ID
 */
function notifyTelegram_(eventType, payload) {
  try {
    const token = String(getSetting_("TELEGRAM_BOT_TOKEN", "")).trim();
    const chatId = String(getSetting_("TELEGRAM_CHAT_ID", "")).trim();
    if (!token || !chatId || token.includes("PASTE_") || chatId.includes("PASTE_")) return;

    // Enrich payload (so even partial payload gets name/phone/email)
    const p = tgEnrichPayloadFromOrders_(payload || {});
    const a = (p && p.access) ? p.access : {};

    const lines = [];
    lines.push("🍿 FluxFilm Alert");
    lines.push("Event: " + String(eventType || "-"));

    if (p.orderId) lines.push("Order: " + p.orderId);

    if (p.name)  lines.push("Name: " + String(p.name));
    if (p.phone) lines.push("Phone: " + tgFullPhone_(p.phone));
    if (p.email) lines.push("Email: " + tgMaskEmail_(p.email));

    if (p.service) lines.push("Service: " + String(p.service));
    if (p.plan)    lines.push("Plan: " + String(p.plan));
    if (p.amount != null && p.amount !== "") lines.push("Amount: " + String(p.amount));

    // optional: include extra fields if you want
    if (p.subId) lines.push("SubID: " + p.subId);

    // If you store anything in access object
    if (a && a.login)    lines.push("Login: " + String(a.login));
    if (a && a.profile)  lines.push("Profile: " + String(a.profile));

    const text = lines.join("\n").slice(0, 3500);

    const url = `https://api.telegram.org/bot${token}/sendMessage`;
    UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        chat_id: chatId,
        text: text,
        disable_web_page_preview: true
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    // never throw from notify
  }
}

/**
 * If name/phone/email missing, fetch it from ORDERS sheet using orderId.
 * Uses your existing sh_(), headerIndex_(), TAB_ORDERS constants if present.
 */
function tgEnrichPayloadFromOrders_(payload) {
  const p = payload || {};
  const orderId = String(p.orderId || p.orderID || p.OrderID || "").trim();

  // if already has details or no orderId, return as-is
  const hasAny = !!(p.name || p.phone || p.email);
  if (hasAny || !orderId) return p;

  try {
    // Prefer your constants/helpers
    const ordersSheetName = (typeof TAB_ORDERS !== "undefined" && TAB_ORDERS) ? TAB_ORDERS : "ORDERS";
    const sh = (typeof sh_ === "function")
      ? sh_(ordersSheetName)
      : SpreadsheetApp.getActive().getSheetByName(ordersSheetName);

    if (!sh) return p;

    const last = sh.getLastRow();
    if (last < 2) return p;

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idx = {};
    headers.forEach((h, i) => idx[String(h || "").trim()] = i);

    // find OrderID column
    const colOrder = (idx["OrderID"] != null) ? idx["OrderID"]
                  : (idx["Order Id"] != null) ? idx["Order Id"]
                  : (idx["Order"] != null) ? idx["Order"]
                  : null;
    if (colOrder == null) return p;

    const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const oid = String(r[colOrder] || "").trim();
      if (oid !== orderId) continue;

      // common columns (adjust if your headers differ)
      const nameCol  = idx["Name"];
      const phoneCol = idx["Phone"];
      const emailCol = idx["Email"];
      const svcCol   = idx["Service"];
      const planCol  = idx["Plan"];
      const amtCol   = (idx["FinalAmount"] != null) ? idx["FinalAmount"] : idx["Amount"];

      if (nameCol != null  && !p.name)  p.name  = r[nameCol];
      if (phoneCol != null && !p.phone) p.phone = r[phoneCol];
      if (emailCol != null && !p.email) p.email = r[emailCol];
      if (svcCol != null   && !p.service) p.service = r[svcCol];
      if (planCol != null  && !p.plan)    p.plan = r[planCol];
      if (amtCol != null   && (p.amount == null || p.amount === "")) p.amount = r[amtCol];

      break;
    }
  } catch (e) {}

  return p;
}

function tgMaskPhone_(phone) {
  const s = String(phone || "").replace(/\D/g, "");
  if (!s) return "";
  const last4 = s.slice(-4);
  return "******" + last4;
}

function tgFullPhone_(phone) {
  const s = String(phone || "").replace(/\D/g, "");
  if (!s) return "";
  return s.length > 10 ? s.slice(-10) : s;
}


function tgMaskEmail_(email) {
  const s = String(email || "").trim();
  if (!s || s.indexOf("@") === -1) return s;
  const parts = s.split("@");
  const user = parts[0] || "";
  const dom = parts[1] || "";
  const head = user.slice(0, 2);
  return head + "***@" + dom;
}



