/*************** FluxFilm Backend — Telegram Order Summaries (Daily/Weekly/Monthly) ***************
 * File: ReportsTelegram.gs
 *
 * Sends Daily + Weekly + Monthly order summaries to your Telegram.
 *
 * Requires SETTINGS:
 *  - TELEGRAM_BOT_TOKEN
 *  - TELEGRAM_CHAT_ID
 * Optional:
 *  - CURRENCY (default "INR")
 *
 * ORDERS sheet headers expected:
 *  - CreatedAt, OrderID, Service, Plan, Status, FinalAmount
 * Optional headers (if present will be used):
 *  - FulfillmentStatus, Discount
 *************************************************************************************************/


// ---------- Telegram sender (unique name so it doesn't collide with other tgSend_) ----------
function ffTgSend_(text) {
  const token = String(getSetting_("TELEGRAM_BOT_TOKEN", "")).trim();
  const chatId = String(getSetting_("TELEGRAM_CHAT_ID", "")).trim();
  if (!token || !chatId || token.includes("PASTE_") || chatId.includes("PASTE_")) return;

  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: String(text || ""),
    disable_web_page_preview: true
  };

  try {
    UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {}
}


// ---------- Date helpers ----------
function ffStartOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}
function ffEndOfDay_(d) {
  const x = new Date(d);
  x.setHours(23, 59, 59, 999);
  return x;
}
function ffStartOfMonth_(d) {
  const x = new Date(d);
  x.setDate(1);
  x.setHours(0, 0, 0, 0);
  return x;
}
function ffAddDays_(d, n) {
  const x = new Date(d);
  x.setDate(x.getDate() + Number(n || 0));
  return x;
}
function ffFmtDate_(d) {
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(d), tz, "dd MMM yyyy");
}
function ffFmtDateTime_(d) {
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(d), tz, "dd MMM yyyy, HH:mm");
}


// ---------- Core summary builder ----------
function ffBuildOrdersSummary_(fromDate, toDate) {
  const orders = sh_(TAB_ORDERS);
  const { idx } = headerIndex_(orders);

  // required headers
  const must = ["CreatedAt", "OrderID", "Service", "Plan", "Status", "FinalAmount"];
  for (const h of must) if (idx[h] == null) throw new Error(`ORDERS missing header: ${h}`);

  const last = orders.getLastRow();
  if (last < 2) {
    return {
      fromDate, toDate,
      count: 0,
      byStatus: {},
      byFulfill: {},
      topServices: [],
      topPlans: [],
      topRevenueServices: [],
      revenue: 0,
      discount: 0
    };
  }

  const data = orders.getRange(2, 1, last - 1, orders.getLastColumn()).getValues();
  const fromMs = new Date(fromDate).getTime();
  const toMs = new Date(toDate).getTime();

  let count = 0;
  let revenue = 0;
  let discount = 0;
  let failed = 0;

  const byStatus = {};
  const byFulfill = {};
  const byService = {};
  const byPlan = {};
  const byServicePaidRevenue = {};

  const hasFulfill = (idx["FulfillmentStatus"] != null);
  const hasDiscount = (idx["Discount"] != null);

  for (const r of data) {
    const createdAt = r[idx["CreatedAt"]];
    const dt = (createdAt instanceof Date) ? createdAt : new Date(createdAt);
    const t = dt.getTime();
    if (isNaN(t) || t < fromMs || t > toMs) continue;

    count++;

    const status = String(r[idx["Status"]] || "").trim().toUpperCase() || "UNKNOWN";
    byStatus[status] = (byStatus[status] || 0) + 1;

    if (status === "FAILED" || status === "CANCELLED") failed++;

    const fstatus = hasFulfill
      ? (String(r[idx["FulfillmentStatus"]] || "").trim().toUpperCase() || "UNKNOWN")
      : null;

    if (fstatus) byFulfill[fstatus] = (byFulfill[fstatus] || 0) + 1;

    const service = String(r[idx["Service"]] || "").trim() || "Unknown";
    const plan = String(r[idx["Plan"]] || "").trim() || "Unknown";

    byService[service] = (byService[service] || 0) + 1;
    byPlan[`${service} • ${plan}`] = (byPlan[`${service} • ${plan}`] || 0) + 1;

    const amt = asNumber_(r[idx["FinalAmount"]]);
    const disc = hasDiscount ? asNumber_(r[idx["Discount"]]) : 0;

    // Revenue: only PAID
    if (status === "PAID") {
      revenue += amt;
      byServicePaidRevenue[service] = (byServicePaidRevenue[service] || 0) + amt;
    }

    // Discounts: total of discount column (any status)
    if (disc > 0) discount += disc;
  }

  function topN_(obj, n) {
    const arr = Object.keys(obj).map(k => ({ k, v: obj[k] }));
    arr.sort((a, b) => (b.v - a.v));
    return arr.slice(0, n);
  }

  return {
    fromDate, toDate,
    count,
    failed,
    byStatus,
    byFulfill,
    topServices: topN_(byService, 5),
    topPlans: topN_(byPlan, 5),
    topRevenueServices: topN_(byServicePaidRevenue, 5),
    revenue: Math.round(revenue),
    discount: Math.round(discount)
  };
}


// ---------- Message formatter ----------
function ffFormatSummaryMsg_(title, s) {
  const currency = String(getSetting_("CURRENCY", "INR")).trim();
  const lines = [];

  lines.push(`🍿 FluxFilm — ${title}`);
  lines.push(`📅 ${ffFmtDate_(s.fromDate)} → ${ffFmtDate_(s.toDate)}`);
  lines.push(`🕒 Generated: ${ffFmtDateTime_(new Date())}`);
  lines.push(``);
  lines.push(`📦 Orders: ${s.count}`);
  lines.push(`💰 Revenue (PAID): ${currency} ${s.revenue}`);
  if (s.discount > 0) lines.push(`🏷️ Discounts: ${currency} ${s.discount}`);
  if (s.failed != null) lines.push(`❌ Failed/Canceled: ${s.failed}`);
  lines.push(``);

  if (s.byStatus && Object.keys(s.byStatus).length) {
    lines.push(`✅ Status breakdown:`);
    Object.keys(s.byStatus).sort().forEach(k => lines.push(`• ${k}: ${s.byStatus[k]}`));
    lines.push(``);
  }

  if (s.byFulfill && Object.keys(s.byFulfill).length) {
    lines.push(`⚙️ Fulfillment breakdown:`);
    Object.keys(s.byFulfill).sort().forEach(k => lines.push(`• ${k}: ${s.byFulfill[k]}`));
    lines.push(``);
  }

  if (s.topServices && s.topServices.length) {
    lines.push(`🔥 Top services (by orders):`);
    s.topServices.forEach(x => lines.push(`• ${x.k}: ${x.v}`));
    lines.push(``);
  }

  if (s.topRevenueServices && s.topRevenueServices.length) {
    lines.push(`🏆 Top services (by PAID revenue):`);
    s.topRevenueServices.forEach(x => lines.push(`• ${x.k}: ${currency} ${Math.round(x.v)}`));
    lines.push(``);
  }

  if (s.topPlans && s.topPlans.length) {
    lines.push(`⭐ Top plans:`);
    s.topPlans.forEach(x => lines.push(`• ${x.k}: ${x.v}`));
  }

  return lines.join("\n").slice(0, 3500);
}


// ---------- Public: Daily summary (today only) ----------
function sendDailyOrderSummary() {
  const now = new Date();
  const from = ffStartOfDay_(now);
  const to = ffEndOfDay_(now);
  const s = ffBuildOrdersSummary_(from, to);
  ffTgSend_(ffFormatSummaryMsg_("Daily Order Summary", s));
  return ok_({ sent: true, from, to, count: s.count });
}


// ---------- Public: Weekly summary (last 7 days incl today) ----------
function sendWeeklyOrderSummary() {
  const to = ffEndOfDay_(new Date());
  const from = ffStartOfDay_(ffAddDays_(to, -6));
  const s = ffBuildOrdersSummary_(from, to);
  ffTgSend_(ffFormatSummaryMsg_("Weekly Order Summary", s));
  return ok_({ sent: true, from, to, count: s.count });
}


// ---------- Public: Monthly summary (current calendar month to today) ----------
function sendMonthlyOrderSummary() {
  const now = new Date();
  const from = ffStartOfMonth_(now);
  const to = ffEndOfDay_(now);
  const s = ffBuildOrdersSummary_(from, to);
  ffTgSend_(ffFormatSummaryMsg_("Monthly Order Summary", s));
  return ok_({ sent: true, from, to, count: s.count });
}


// ---------- Optional: Previous calendar month summary ----------
function sendLastMonthOrderSummary() {
  const now = new Date();
  const thisMonthStart = ffStartOfMonth_(now);
  const lastMonthEnd = ffAddDays_(thisMonthStart, -1);
  const from = ffStartOfMonth_(lastMonthEnd);
  const to = ffEndOfDay_(lastMonthEnd);
  const s = ffBuildOrdersSummary_(from, to);
  ffTgSend_(ffFormatSummaryMsg_("Last Month Order Summary", s));
  return ok_({ sent: true, from, to, count: s.count });
}


// ---------- One-time trigger installer (Daily + Weekly + Monthly) ----------
function installOrderSummaryTriggers() {
  // cleanup existing triggers for these functions
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    const h = t.getHandlerFunction();
    if (
      h === "sendDailyOrderSummary" ||
      h === "sendWeeklyOrderSummary" ||
      h === "sendMonthlyOrderSummary" ||
      h === "sendLastMonthOrderSummary"
    ) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Daily: every day at 10:00
  ScriptApp.newTrigger("sendDailyOrderSummary")
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();

  // Weekly: every Monday at 10:00
  ScriptApp.newTrigger("sendWeeklyOrderSummary")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();

  // Monthly: 1st day of month at 10:00
  ScriptApp.newTrigger("sendLastMonthOrderSummary")
    .timeBased()
    .onMonthDay(1)
    .atHour(10)
    .create();

  ffTgSend_("✅ Order summary triggers installed:\n• Daily: 10:00\n• Weekly: Monday 10:00\n• Monthly: Previous month report (1st day 10:00)");
  return ok_({ installed: true });
}
