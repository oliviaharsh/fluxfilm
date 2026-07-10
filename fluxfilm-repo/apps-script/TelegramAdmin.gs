/*************** FluxFilm Backend — Telegram Admin Console (FULL FILE) ***************/

// SECURITY: only respond to your admin chat
function tg_isAllowed_(chatId){
  Logger.log("TG chatId:", chatId);
  return true; // TEMP: allow all
}


/**
 * WEBHOOK ENTRY
 * NOTE: You currently use tg_doPost(e). Keep this name unless you change webhook handler.
 */
function tg_doPost(e){
  try{
    const token = String(getSetting_("TELEGRAM_BOT_TOKEN","")).trim();
    if(!token) return ContentService.createTextOutput("OK");

    const upd = JSON.parse((e && e.postData && e.postData.contents) ? e.postData.contents : "{}");
    const msg = upd.message || upd.edited_message;
    if(!msg || !msg.text) return ContentService.createTextOutput("OK");



    const chatId = (msg.chat && msg.chat.id) ? String(msg.chat.id) : "";
    if(!tg_isAllowed_(chatId)) return ContentService.createTextOutput("OK");

    const text = String(msg.text||"").trim();
    const parts = text.split(/\s+/);
    const cmd = (parts[0]||"").toLowerCase();

    // ---------- HELP ----------
    if(cmd === "/start" || cmd === "/help"){
      tgSendLong_(token, chatId, tgHelp_());
      return ContentService.createTextOutput("OK");
    }

    if(cmd === "/ping"){
      tgSendLong_(token, chatId, "✅ FluxFilm Admin Bot is alive.");
      return ContentService.createTextOutput("OK");
    }

    // ---------- SUBS LOOKUP ----------
    // /subs <phone> [service:netflix]
    if(cmd === "/subs"){
      const phone = normalize10_(parts[1]||"");
      const opt = parseKeyVals_(parts.slice(2));
      tgSendLong_(token, chatId, adminFetchSubs_(phone, { serviceFilter: opt.service || "" }));
      return ContentService.createTextOutput("OK");
    }

    // /netflix <phone> (shortcut)
    if(cmd === "/netflix"){
      const phone = normalize10_(parts[1]||"");
      tgSendLong_(token, chatId, adminFetchSubs_(phone, { serviceFilter:"netflix" }));
      return ContentService.createTextOutput("OK");
    }

    // ---------- ORDER DETAILS ----------
    // /order <OrderID>
    if(cmd === "/order"){
      const orderId = String(parts[1]||"").trim();
      tgSendLong_(token, chatId, adminFetchOrder_(orderId));
      return ContentService.createTextOutput("OK");
    }

    // ---------- TODAY PAID ----------
    // /paidtoday
    if(cmd === "/paidtoday"){
      tgSendLong_(token, chatId, adminPaidToday_());
      return ContentService.createTextOutput("OK");
    }

    // ---------- TOTAL ORDERS ----------
    // /totalorders [days:7]
    if(cmd === "/totalorders"){
      const opt = parseKeyVals_(parts.slice(1));
      tgSendLong_(token, chatId, adminTotalOrders_(opt));
      return ContentService.createTextOutput("OK");
    }

    // ---------- EXPIRING SUBS ----------
    // /expiring days:7 [service:netflix]
    if(cmd === "/expiring"){
      const opt = parseKeyVals_(parts.slice(1));
      tgSendLong_(token, chatId, adminExpiringSubs_(opt));
      return ContentService.createTextOutput("OK");
    }

    // ---------- FORCE PAID (when customer didn't put order id in notes) ----------
    // /forcepaid <OrderID> [txn:UTR123]
    // This edits ORDERS -> Status=PAID, VerifiedAt, TxnRef(optional), then triggers fulfill pipeline.
    if(cmd === "/forcepaid"){
      const orderId = String(parts[1]||"").trim();
      const opt = parseKeyVals_(parts.slice(2));
      const res = adminForcePaid_(orderId, opt);
      tgSendLong_(token, chatId, res);
      return ContentService.createTextOutput("OK");
    }

    // ---------- MANUAL ORDER FULFILLED (2-step) ----------
    // /manualdone <OrderID> -> asks /inv ...
    // /manualdone <OrderID> inv:<login|phone|accountId|AccountID#P2|skip>
    if(cmd === "/manualdone"){
      const orderId = String(parts[1]||"").trim();
      const opt = parseKeyVals_(parts.slice(2));
      const res = adminManualDone_(chatId, orderId, opt);
      tgSendLong_(token, chatId, res);
      return ContentService.createTextOutput("OK");
    }

    // Step 2:
    // /inv harshwalia8888@gmail.com
    // /inv 8076332049
    // /inv NFLX-D1#P2
    // /inv skip
    if(cmd === "/inv"){
      const invInput = String(parts[1]||"").trim();
      const res = adminManualInvReply_(chatId, invInput);
      tgSendLong_(token, chatId, res);
      return ContentService.createTextOutput("OK");
    }

    // ---------- REASSIGN / SWITCH INVENTORY (your existing) ----------
    // /assign <SubID> <newLoginId> [pno:2]
    if(cmd === "/assign"){
      const subId = String(parts[1]||"").trim();
      const newLogin = String(parts[2]||"").trim();
      const opt = parseKeyVals_(parts.slice(3));
      const res = adminReassign_(subId, newLogin, opt);
      tgSendLong_(token, chatId, res);
      return ContentService.createTextOutput("OK");
    }

    // Fallback help
    tgSendLong_(token, chatId, tgHelp_());
    return ContentService.createTextOutput("OK");

  }catch(err){
    // Silent fail
    return ContentService.createTextOutput("OK");
  }
}

/* ----------------------------- HELP TEXT ----------------------------- */

function tgHelp_(){
  return [
    "🍿 FluxFilm Admin Commands",
    "",
    "📌 Lookup",
    "• /order <OrderID>",
    "• /subs <phone> [service:netflix]",
    "• /netflix <phone>",
    "",
    "⚡ Ops",
    "• /manualdone <OrderID>   (then /inv <login|phone|accountId|AccountID#P2|skip>)",
    "• /forcepaid <OrderID> txn:UTR123",
    "• /assign <SubID> <newLoginId> pno:2",
    "",
    "📊 Reports",
    "• /paidtoday",
    "• /totalorders [days:7]",
    "• /expiring days:7 [service:netflix]",
    "",
    "🔧 Misc",
    "• /ping",
    "• /help",
  ].join("\n");
}

/* ----------------------------- TELEGRAM SENDERS ----------------------------- */

function tgSend_(token, chatId, text){
  const url = "https://api.telegram.org/bot" + token + "/sendMessage";
  try{
    const res = UrlFetchApp.fetch(url, {
      method:"post",
      contentType:"application/json",
      payload: JSON.stringify({
        chat_id: chatId,
        text: String(text||""),
        disable_web_page_preview:true
      }),
      muteHttpExceptions:true
    });

    // TEMP DEBUG (remove later)
    Logger.log("TG send status=" + res.getResponseCode() + " body=" + res.getContentText());
  }catch(e){
    Logger.log("TG send crash: " + (e && e.stack ? e.stack : e));
  }
}


// chunk sender (Telegram message limit)
function tgSendLong_(token, chatId, text){
  const chunks = tgChunk_(String(text||""), 3500);
  for(let i=0;i<chunks.length;i++){
    tgSend_(token, chatId, chunks[i]);
  }
}

function tgChunk_(s, maxLen){
  const out = [];
  let cur = "";
  const lines = String(s||"").split("\n");
  for(const ln of lines){
    if((cur + "\n" + ln).length > maxLen){
      if(cur) out.push(cur);
      cur = ln;
    }else{
      cur = cur ? (cur + "\n" + ln) : ln;
    }
  }
  if(cur) out.push(cur);
  return out;
}

/* ----------------------------- BASIC HELPERS ----------------------------- */

function normalize10_(p){
  const s = String(p||"").replace(/\D/g,"");
  if(!s) return "";
  return s.length > 10 ? s.slice(-10) : s;
}

function parseKeyVals_(arr){
  const out = {};
  (arr||[]).forEach(t=>{
    const m = String(t||"").match(/^([a-zA-Z0-9_]+)\:(.+)$/);
    if(m) out[m[1].toLowerCase()] = m[2];
  });
  return out;
}

/* ----------------------------- COMMAND: SUBS DETAILS ----------------------------- */
/**
 * REQUIRED SUBSCRIPTIONS HEADERS:
 * Phone, SubID, Service, Plan, ExpiryDate, InventoryRef, LoginId, Password
 * Optional: Name, ProfileName, ProfileNumber, ProfilePIN, Status
 */
function adminFetchSubs_(phone, opts){
  const ph = normalize10_(phone);
  if(!ph) return "❌ Phone missing. Usage: /subs 9818196079";

  const filter = String((opts && opts.serviceFilter) || "").toLowerCase();

  const subs = sh_("SUBSCRIPTIONS");
  const { idx } = headerIndex_(subs);

  const last = subs.getLastRow();
  if(last < 2) return "No subscriptions in sheet.";

  const data = subs.getRange(2,1,last-1,subs.getLastColumn()).getValues();
  const nowMs = Date.now();

  const hits = [];
  for(let i=0;i<data.length;i++){
    const r = data[i];
    const rp = normalize10_(r[idx["Phone"]]);
    if(rp !== ph) continue;

    const svc = String(r[idx["Service"]]||"").trim();
    if(filter && !svc.toLowerCase().includes(filter)) continue;

    const exp = r[idx["ExpiryDate"]];
    const expMs = (exp instanceof Date) ? exp.getTime() : new Date(exp).getTime();
    const daysLeft = isNaN(expMs) ? null : Math.ceil((expMs - nowMs)/86400000);

    hits.push({
      name: (idx["Name"]!=null) ? String(r[idx["Name"]]||"").trim() : "",
      subId: String(r[idx["SubID"]]||"").trim(),
      service: svc,
      plan: String(r[idx["Plan"]]||"").trim(),
      expiry: exp,
      daysLeft: daysLeft,
      status: (idx["Status"]!=null) ? String(r[idx["Status"]]||"").trim() : "",
      inv: String(r[idx["InventoryRef"]]||"").trim(),
      login: String(r[idx["LoginId"]]||"").trim(),
      pass: String(r[idx["Password"]]||"").trim(),
      profileName: (idx["ProfileName"]!=null) ? String(r[idx["ProfileName"]]||"").trim() : "",
      profileNo: (idx["ProfileNumber"]!=null) ? String(r[idx["ProfileNumber"]]||"").trim() : "",
      pin: (idx["ProfilePIN"]!=null) ? String(r[idx["ProfilePIN"]]||"").trim() : ""
    });
  }

  if(!hits.length) return "❌ No subscriptions found for this phone.";

  // Active first (more days left first)
  hits.sort((a,b)=>{
    const ax = (a.daysLeft==null ? -999999 : a.daysLeft);
    const bx = (b.daysLeft==null ? -999999 : b.daysLeft);
    return bx - ax;
  });

  const headerName = hits[0].name ? `👤 ${hits[0].name}\n` : "";
  let out = `${headerName}📱 ${ph} — Found ${hits.length}\n`;
  hits.slice(0,10).forEach(s=>{
    out += `\n🎬 ${s.service} • ${s.plan}\n` +
           `SubID: ${s.subId}\n` +
           `Status: ${s.status || "-"}\n` +
           `Expiry: ${s.expiry} (${s.daysLeft==null ? "?" : s.daysLeft+"d"})\n` +
           `Inv: ${s.inv}\n` +
           `Login: ${s.login}\nPass: ${s.pass}\n` +
           (s.profileNo ? `Profile: P${s.profileNo} ${s.profileName}\nPIN: ${s.pin}\n` : "");
  });
  if(hits.length > 10) out += `\n(+${hits.length-10} more not shown)`;
  return out;
}

/* ----------------------------- COMMAND: ORDER DETAILS ----------------------------- */
/**
 * REQUIRED ORDERS HEADERS:
 * OrderID, Status, Service, Plan, Phone, Email, FinalAmount
 * Optional: Currency, TxnRef, VerifiedAt, FulfillmentStatus, FulfilledAt, InventoryRef, OrderType, RenewSubID, Error
 */
function adminFetchOrder_(orderId){
  const oid = String(orderId||"").trim();
  if(!oid) return "❌ Usage: /order <OrderID>";

  const orders = sh_("ORDERS");
  const rn = getRowByValue_(orders, "OrderID", oid);
  if(!rn) return "❌ Order not found: " + oid;

  const { idx } = headerIndex_(orders);
  function get(h){
    if(idx[h]==null) return "";
    return orders.getRange(rn, idx[h]+1).getValue();
  }

  return [
    `🧾 Order ${oid}`,
    `Status: ${String(get("Status")||"")}`,
    (idx["FulfillmentStatus"]!=null ? `Fulfillment: ${String(get("FulfillmentStatus")||"")}` : ""),
    `Service: ${String(get("Service")||"")}`,
    `Plan: ${String(get("Plan")||"")}`,
    `Phone: ${String(get("Phone")||"")}`,
    `Email: ${String(get("Email")||"")}`,
    `Amount: ${String(get("FinalAmount")||"")} ${String(get("Currency")||"INR")}`,
    (idx["TxnRef"]!=null ? `TxnRef: ${String(get("TxnRef")||"")}` : ""),
    (idx["VerifiedAt"]!=null ? `VerifiedAt: ${String(get("VerifiedAt")||"")}` : ""),
    (idx["FulfilledAt"]!=null ? `FulfilledAt: ${String(get("FulfilledAt")||"")}` : ""),
    (idx["InventoryRef"]!=null ? `InventoryRef: ${String(get("InventoryRef")||"")}` : ""),
    (idx["OrderType"]!=null ? `OrderType: ${String(get("OrderType")||"")}` : ""),
    (idx["RenewSubID"]!=null ? `RenewSubID: ${String(get("RenewSubID")||"")}` : ""),
    (idx["Error"]!=null && get("Error") ? `Error: ${String(get("Error")||"")}` : "")
  ].filter(Boolean).join("\n");
}

/* ----------------------------- COMMAND: PAID TODAY ----------------------------- */

function adminPaidToday_(){
  const orders = sh_("ORDERS");
  const { idx } = headerIndex_(orders);

  const must = ["OrderID","Status","Service","Plan","FinalAmount"];
  for(const h of must) if(idx[h]==null) return "❌ ORDERS missing header: " + h;

  const hasVerifiedAt = idx["VerifiedAt"] != null;
  const hasCreatedAt = idx["CreatedAt"] != null;
  if(!hasVerifiedAt && !hasCreatedAt) return "❌ ORDERS needs VerifiedAt or CreatedAt.";

  const last = orders.getLastRow();
  if(last < 2) return "No orders.";

  const data = orders.getRange(2,1,last-1,orders.getLastColumn()).getValues();

  const now = new Date();
  const start = new Date(now); start.setHours(0,0,0,0);
  const end = new Date(now); end.setHours(23,59,59,999);

  const hits = [];
  for(const r of data){
    const status = String(r[idx["Status"]]||"").trim().toUpperCase();
    if(status !== "PAID") continue;

    const dtRaw = hasVerifiedAt ? r[idx["VerifiedAt"]] : r[idx["CreatedAt"]];
    const dt = (dtRaw instanceof Date) ? dtRaw : new Date(dtRaw);
    if(isNaN(dt.getTime())) continue;
    if(dt < start || dt > end) continue;

    hits.push({
      oid: String(r[idx["OrderID"]]||"").trim(),
      svc: String(r[idx["Service"]]||"").trim(),
      plan: String(r[idx["Plan"]]||"").trim(),
      amt: r[idx["FinalAmount"]],
      f: (idx["FulfillmentStatus"]!=null) ? String(r[idx["FulfillmentStatus"]]||"").trim() : ""
    });
  }

  if(!hits.length) return "✅ No PAID orders today.";

  let out = `📌 PAID Today — ${hits.length}\n`;
  hits.slice(0,40).forEach(o=>{
    out += `\n• ${o.oid} — ${o.svc} • ${o.plan} — ₹${o.amt}` + (o.f ? ` — ${o.f}` : "");
  });
  if(hits.length > 40) out += `\n\n(+${hits.length-40} more not shown)`;
  return out;
}

/* ----------------------------- COMMAND: TOTAL ORDERS ----------------------------- */
// /totalorders [days:7]
function adminTotalOrders_(opt){
  const days = Math.max(0, parseInt(String((opt && opt.days) || "0"), 10) || 0);

  const orders = sh_("ORDERS");
  const { idx } = headerIndex_(orders);
  if(idx["OrderID"]==null || idx["Status"]==null) return "❌ ORDERS missing header: OrderID/Status";

  const last = orders.getLastRow();
  if(last < 2) return "No orders.";

  const data = orders.getRange(2,1,last-1,orders.getLastColumn()).getValues();

  let start = null;
  if(days > 0){
    start = new Date();
    start.setHours(0,0,0,0);
    start = new Date(start.getTime() - (days-1)*86400000); // include today as day 1
  }

  const createdCol = (idx["CreatedAt"]!=null) ? "CreatedAt" : (idx["VerifiedAt"]!=null ? "VerifiedAt" : "");
  const hasDate = !!createdCol;

  const counts = {};
  let total = 0;
  let filtered = 0;

  for(const r of data){
    total++;
    if(start && hasDate){
      const dRaw = r[idx[createdCol]];
      const d = (dRaw instanceof Date) ? dRaw : new Date(dRaw);
      if(isNaN(d.getTime()) || d < start) continue;
    }
    filtered++;

    const st = String(r[idx["Status"]]||"").trim().toUpperCase() || "UNKNOWN";
    counts[st] = (counts[st]||0) + 1;
  }

  const keys = Object.keys(counts).sort((a,b)=>counts[b]-counts[a]);

  let out = `📦 Total Orders: ${total}\n`;
  if(days>0 && hasDate) out += `🗓️ In last ${days} day(s): ${filtered}\n`;
  if(days>0 && !hasDate) out += `⚠️ No CreatedAt/VerifiedAt column found, showing overall counts.\n`;

  out += "\nStatus breakdown:\n";
  keys.forEach(k=>{
    out += `• ${k}: ${counts[k]}\n`;
  });

  return out.trim();
}

/* ----------------------------- COMMAND: EXPIRING SUBS ----------------------------- */
// /expiring days:7 [service:netflix]
function adminExpiringSubs_(opt){
  const days = Math.max(0, parseInt(String((opt && opt.days) || "7"), 10) || 7);
  const filter = String((opt && opt.service) || "").trim().toLowerCase();

  const subs = sh_("SUBSCRIPTIONS");
  const { idx } = headerIndex_(subs);

  const must = ["ExpiryDate","Service","Plan"];
  for(const h of must) if(idx[h]==null) return "❌ SUBSCRIPTIONS missing header: " + h;

  const last = subs.getLastRow();
  if(last < 2) return "No subscriptions in sheet.";

  const data = subs.getRange(2,1,last-1,subs.getLastColumn()).getValues();

  const now = new Date();
  now.setHours(0,0,0,0);
  const end = new Date(now.getTime() + days*86400000);

  const hits = [];
  for(const r of data){
    const svc = String(r[idx["Service"]]||"").trim();
    if(filter && !svc.toLowerCase().includes(filter)) continue;

    const expRaw = r[idx["ExpiryDate"]];
    const exp = (expRaw instanceof Date) ? expRaw : new Date(expRaw);
    if(isNaN(exp.getTime())) continue;

    // only in window [today..end]
    const exp0 = new Date(exp); exp0.setHours(0,0,0,0);
    if(exp0 < now || exp0 > end) continue;

    const name = (idx["Name"]!=null) ? String(r[idx["Name"]]||"").trim() : "";
    const phone = (idx["Phone"]!=null) ? normalize10_(r[idx["Phone"]]) : "";
    const subId = (idx["SubID"]!=null) ? String(r[idx["SubID"]]||"").trim() : "";
    const plan = String(r[idx["Plan"]]||"").trim();

    const daysLeft = Math.ceil((exp0.getTime() - now.getTime())/86400000);

    hits.push({
      daysLeft,
      exp: exp0,
      name,
      phone,
      svc,
      plan,
      subId
    });
  }

  if(!hits.length) return `✅ No subs expiring in next ${days} day(s).`;

  hits.sort((a,b)=>a.daysLeft - b.daysLeft);

  let out = `⏳ Expiring in next ${days} day(s) — ${hits.length}\n`;
  hits.slice(0,60).forEach(s=>{
    out += `\n• ${s.daysLeft}d — ${s.svc} • ${s.plan}\n` +
           `  Exp: ${s.exp}\n` +
           (s.name ? `  Name: ${s.name}\n` : "") +
           (s.phone ? `  Phone: ${s.phone}\n` : "") +
           (s.subId ? `  SubID: ${s.subId}\n` : "");
  });
  if(hits.length > 60) out += `\n(+${hits.length-60} more not shown)`;
  return out.trim();
}

/* ----------------------------- COMMAND: FORCE PAID ----------------------------- */
/**
 * This is for cases where payment came but verifyPayment can't match because customer didn't put OrderID in bank notes.
 * It updates ORDERS -> PAID, VerifiedAt, TxnRef(optional) and then triggers normal pipeline:
 *   safeFulfill_(orderId) OR fulfillOrder(orderId)
 * That pipeline should handle NEW and RENEW automatically.
 */
function adminForcePaid_(orderId, opt){
  const oid = String(orderId||"").trim();
  if(!oid) return "❌ Usage: /forcepaid <OrderID> txn:UTR123";

  const orders = sh_("ORDERS");
  const rn = getRowByValue_(orders, "OrderID", oid);
  if(!rn) return "❌ Order not found: " + oid;

  const { idx } = headerIndex_(orders);
  if(idx["Status"]==null) return "❌ ORDERS missing header: Status";

  const cur = String(orders.getRange(rn, idx["Status"]+1).getValue()||"").trim().toUpperCase();
  if(cur !== "PAID"){
    orders.getRange(rn, idx["Status"]+1).setValue("PAID");
    if(idx["VerifiedAt"]!=null) orders.getRange(rn, idx["VerifiedAt"]+1).setValue(now_());
    if(idx["TxnRef"]!=null && opt && opt.txn) orders.getRange(rn, idx["TxnRef"]+1).setValue(String(opt.txn||"").trim());
  }

  // Trigger normal fulfill path (NEW/RENEW handled there)
  let res = null;
  if(typeof safeFulfill_ === "function"){
    res = safeFulfill_(oid);
  }else if(typeof fulfillOrder === "function"){
    res = fulfillOrder(oid);
  }

  if(res){
    return "✅ Marked PAID + triggered normal flow.\n" + JSON.stringify(res, null, 2).slice(0, 2500);
  }
  return "✅ Marked PAID. (Fulfill function not found in project: safeFulfill_/fulfillOrder)";
}

/* ----------------------------- COMMAND: MANUAL DONE (2-step) ----------------------------- */
/**
 * When you complete a MANUAL_PENDING order, this command:
 * - asks for inventory ref (email/phone/accountId or AccountID#P2) OR "skip"
 * - resolves inv if possible (else blank)
 * - fills SUBSCRIPTIONS row (InventoryRef + creds/profile if found)
 * - marks SUB status ACTIVE
 * - updates ORDERS FulfillmentStatus=FULFILLED, FulfilledAt, InventoryRef
 * - triggers email invoice/access: sendOrderEmailsOnce_(orderId,"FULFILLED") (fallback sendOrderEmails_)
 */
function adminManualDone_(chatId, orderId, opt){
  const oid = String(orderId||"").trim();
  if(!oid) return "❌ Usage: /manualdone <OrderID>  (then /inv <login|phone|accountId|AccountID#P2|skip>)\nOr: /manualdone <OrderID> inv:<value>";

  const inv = opt && opt.inv ? String(opt.inv||"").trim() : "";
  if(inv){
    return adminManualDoneFinish_(oid, inv);
  }

  // store pending state (10 min)
  CacheService.getScriptCache().put("TG_WAIT_INV_" + String(chatId), oid, 600);

  return [
    "✅ Manual order noted.",
    "Now send inventory input:",
    "/inv <account login OR phone OR accountId OR AccountID#P2>",
    "Or to leave blank:",
    "/inv skip",
    "",
    "Examples:",
    "/inv harshwalia8888@gmail.com",
    "/inv 8076332049",
    "/inv NFLX-D1#P2",
  ].join("\n");
}

function adminManualInvReply_(chatId, invInput){
  const oid = CacheService.getScriptCache().get("TG_WAIT_INV_" + String(chatId));
  if(!oid) return "❌ No pending manual order. Start: /manualdone <OrderID>";

  CacheService.getScriptCache().remove("TG_WAIT_INV_" + String(chatId));

  const inv = String(invInput||"").trim();
  if(!inv) return "❌ Usage: /inv <login|phone|accountId|AccountID#P2|skip>";

  return adminManualDoneFinish_(oid, inv);
}

function adminManualDoneFinish_(orderId, invInput){
  const oid = String(orderId||"").trim();
  const raw = String(invInput||"").trim();
  if(!oid) return "❌ Missing OrderID.";

  const orders = sh_("ORDERS");
  const or = getRowByValue_(orders, "OrderID", oid);
  if(!or) return "❌ Order not found: " + oid;

  const { idx:oIdx } = headerIndex_(orders);

  // Find SUB row for this order
  const subs = sh_("SUBSCRIPTIONS");
  const sr = getRowByValue_(subs, "OrderID", oid);
  if(!sr) return "❌ SUB row not found for this order (OrderID not found in SUBSCRIPTIONS).";

  const { idx:sIdx } = headerIndex_(subs);

  // Resolve inventory (or skip)
  const isSkip = (raw.toLowerCase() === "skip" || raw === "-" || raw.toLowerCase() === "blank");
  const resolved = isSkip ? { inventoryRef:"" } : tgResolveInventory_(raw);

  const invRef = String(resolved.inventoryRef || "").trim();

  // ---- Update SUBSCRIPTIONS cells ----
  if(sIdx["InventoryRef"]!=null) subs.getRange(sr, sIdx["InventoryRef"]+1).setValue(invRef);

  if(resolved.loginId && sIdx["LoginId"]!=null) subs.getRange(sr, sIdx["LoginId"]+1).setValue(resolved.loginId);
  if(resolved.pass && sIdx["Password"]!=null) subs.getRange(sr, sIdx["Password"]+1).setValue(resolved.pass);

  // Netflix profile if present
  if(resolved.profileNumber && sIdx["ProfileNumber"]!=null) subs.getRange(sr, sIdx["ProfileNumber"]+1).setValue(resolved.profileNumber);
  if(resolved.profileName && sIdx["ProfileName"]!=null) subs.getRange(sr, sIdx["ProfileName"]+1).setValue(resolved.profileName);
  if(resolved.profilePin && sIdx["ProfilePIN"]!=null) subs.getRange(sr, sIdx["ProfilePIN"]+1).setValue(resolved.profilePin);

  // Make ACTIVE
  if(sIdx["Status"]!=null) subs.getRange(sr, sIdx["Status"]+1).setValue("ACTIVE");

  // ---- Update ORDERS cells ----
  if(oIdx["FulfillmentStatus"]!=null) orders.getRange(or, oIdx["FulfillmentStatus"]+1).setValue("FULFILLED");
  if(oIdx["FulfilledAt"]!=null) orders.getRange(or, oIdx["FulfilledAt"]+1).setValue(now_());
  if(oIdx["InventoryRef"]!=null) orders.getRange(or, oIdx["InventoryRef"]+1).setValue(invRef);

  // ---- Trigger Email Invoice / Access ----
  try{
    if(typeof sendOrderEmailsOnce_ === "function"){
      sendOrderEmailsOnce_(oid, "FULFILLED");
    }else if(typeof sendOrderEmails_ === "function"){
      sendOrderEmails_(oid, "FULFILLED");
    }
  }catch(e){}

  // Return summary
  const svc = (oIdx["Service"]!=null) ? String(orders.getRange(or, oIdx["Service"]+1).getValue()||"").trim() : "";
  const plan = (oIdx["Plan"]!=null) ? String(orders.getRange(or, oIdx["Plan"]+1).getValue()||"").trim() : "";

  return [
    "✅ Manual order marked FULFILLED + email triggered.",
    `Order: ${oid}`,
    `Service: ${svc} • ${plan}`,
    `InventoryRef: ${invRef || "(blank)"}`,
    (resolved.loginId ? `Login: ${resolved.loginId}` : ""),
    (resolved.pass ? `Pass: ${resolved.pass}` : ""),
    (resolved.profileNumber ? `Profile: P${resolved.profileNumber} ${resolved.profileName||""}` : "")
  ].filter(Boolean).join("\n");
}

/* ----------------------------- INVENTORY RESOLVE ----------------------------- */
/**
 * Accepts:
 * - email/login id
 * - phone digits (match inside loginId)
 * - AccountID
 * - AccountID#P2 (also fetch profile from INV_PROFILES if present)
 *
 * Sheets assumed:
 * INV_ACCOUNTS: AccountID, LoginId, Password
 * INV_PROFILES (optional): AccountID, ProfileNumber, ProfileDisplayName, ProfilePIN
 */
function tgResolveInventory_(input){
  const raw = String(input||"").trim();
  if(!raw) return {};

// AccountID#P2
if(raw.indexOf("#P") > -1){
  const parts = raw.split("#P");
  const accId = String(parts[0]||"").trim();
  const pNo = Number(parts[1]||0);

  const base = tgFindInvAccountByAny_(accId);
  if (!base) return { inventoryRef: "" }; // ✅ skips inactive accounts

  const prof = tgFindInvProfile_(accId, pNo) || {};

  return {
    inventoryRef: raw,
    loginId: base.loginId || "",
    pass: base.pass || "",
    profileNumber: pNo || "",
    profileName: prof.profileName || "",
    profilePin: prof.profilePin || ""
  };
}

  // else resolve to account only
  const acc = tgFindInvAccountByAny_(raw);
  if(!acc) return { inventoryRef:"" };

  return {
    inventoryRef: acc.accountId || "",
    loginId: acc.loginId || "",
    pass: acc.pass || ""
  };
}

function tgFindInvAccountByAny_(q){
  const invAcc = sh_("INV_ACCOUNTS");
  const { idx } = headerIndex_(invAcc);

  if(idx["AccountID"]==null || idx["LoginId"]==null || idx["Password"]==null) return null;

  const query = String(q||"").trim();
  const digits = query.replace(/\D/g,"");
  const qLower = query.toLowerCase();

  const last = invAcc.getLastRow();
  if(last < 2) return null;

  const rows = invAcc.getRange(2,1,last-1,invAcc.getLastColumn()).getValues();

  for(const r of rows){

    // ✅ ADD THIS: skip inactive accounts
    const isActive = (idx["IsActive"] == null)
      ? true
      : String(r[idx["IsActive"]] || "TRUE").toUpperCase().trim() === "TRUE";

    if(!isActive) continue;

    const accountId = String(r[idx["AccountID"]]||"").trim();
    const loginId = String(r[idx["LoginId"]]||"").trim();
    const pass = String(r[idx["Password"]]||"").trim();
    if(!accountId) continue;

    if(accountId === query) return { accountId, loginId, pass };

    if(loginId){
      const ll = loginId.toLowerCase();
      if(ll === qLower || ll.indexOf(qLower) >= 0) return { accountId, loginId, pass };

      if(digits && loginId.replace(/\D/g,"").indexOf(digits) >= 0) return { accountId, loginId, pass };
    }
  }

  return null;
}

function tgFindInvProfile_(accountId, profileNo){
  try{
    const invProf = sh_("INV_PROFILES");
    const { idx } = headerIndex_(invProf);

    if(idx["AccountID"]==null || idx["ProfileNumber"]==null) return null;

    const last = invProf.getLastRow();
    if(last < 2) return null;

    const rows = invProf.getRange(2,1,last-1,invProf.getLastColumn()).getValues();
    for(const r of rows){
      if(String(r[idx["AccountID"]]||"").trim() !== String(accountId||"").trim()) continue;
      if(Number(r[idx["ProfileNumber"]]||0) !== Number(profileNo||0)) continue;

      const name = (idx["ProfileDisplayName"]!=null) ? String(r[idx["ProfileDisplayName"]]||"").trim() : "";
      const pin  = (idx["ProfilePIN"]!=null) ? String(r[idx["ProfilePIN"]]||"").trim() : "";

      return { profileName:name, profilePin:pin };
    }
  }catch(e){}
  return null;
}






function tg_sendToAdmin_(text){
  const token = String(getSetting_("TELEGRAM_BOT_TOKEN","")).trim();
  const chatId = String(getSetting_("TELEGRAM_CHAT_ID","")).trim();
  if(!token || !chatId) return;
  tgSendLong_(token, chatId, String(text||""));
}




function zzz_sendDailyTelegramSummary(){
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const stamp = Utilities.formatDate(now, tz, "dd MMM yyyy, HH:mm");

  const parts = [];
  parts.push("📊 FluxFilm — Daily Summary");
  parts.push("🕒 " + stamp);
  parts.push("");

  // Paid today list
  parts.push(adminPaidToday_());

  parts.push("\n— — —\n");

  // Status breakdown last 7 days
  parts.push(adminTotalOrders_({ days:"7" }));

  parts.push("\n— — —\n");

  // Expiring in next 3 days
  parts.push(adminExpiringSubs_({ days:"3" }));

  tg_sendToAdmin_(parts.join("\n").slice(0, 12000));
}


function zzz_sendWeeklyTelegramSummary(){
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const stamp = Utilities.formatDate(now, tz, "dd MMM yyyy, HH:mm");

  const parts = [];
  parts.push("📈 FluxFilm — Weekly Summary");
  parts.push("🕒 " + stamp);
  parts.push("");

  // 7-day breakdown
  parts.push(adminTotalOrders_({ days:"7" }));

  parts.push("\n— — —\n");

  // Paid today (optional)
  parts.push(adminPaidToday_());

  parts.push("\n— — —\n");

  // Expiring next 7 days
  parts.push(adminExpiringSubs_({ days:"7" }));

  tg_sendToAdmin_(parts.join("\n").slice(0, 12000));
}



function zzz_sendMonthlyTelegramSummary(){
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const stamp = Utilities.formatDate(now, tz, "dd MMM yyyy, HH:mm");

  const parts = [];
  parts.push("🧾 FluxFilm — Monthly Summary");
  parts.push("🕒 " + stamp);
  parts.push("");

  // 30-day breakdown
  parts.push(adminTotalOrders_({ days:"30" }));

  parts.push("\n— — —\n");

  // Expiring next 14 days (monthly ops view)
  parts.push(adminExpiringSubs_({ days:"14" }));

  tg_sendToAdmin_(parts.join("\n").slice(0, 12000));
}








function tg_checkWebhook_NOW(){
  const token = getSetting_("TELEGRAM_BOT_TOKEN", "");
  const url = "https://api.telegram.org/bot" + token + "/getWebhookInfo";
  Logger.log(UrlFetchApp.fetch(url).getContentText());
}





function tg_setWebhook_u0(){
  const token = getSetting_("TELEGRAM_BOT_TOKEN", "");
  const urlU0 = "https://script.google.com/u/0/macros/s/AKfycbwCV7ech_uYORFuR_ZP8DA4Ahgy_hewK9GJEX7FWQH32GPTqOI8z9nf61avDVgC925Z4w/exec";

  const setUrl = "https://api.telegram.org/bot" + token + "/setWebhook?url=" + encodeURIComponent(urlU0);
  Logger.log(UrlFetchApp.fetch(setUrl).getContentText());
  Logger.log("WEBHOOK SET TO: " + urlU0);
}





function tg_doPost_(e){
  try{
    return tg_doPost(e);
  }catch(err){
    try{
      const token = String(getSetting_("TELEGRAM_BOT_TOKEN","")).trim();
      const chatId = String(getSetting_("TELEGRAM_CHAT_ID","")).trim();
      if(token && chatId){
        tgSendLong_(token, chatId, "❌ TG handler error:\n" + (err && err.stack ? err.stack : err));
      }
    }catch(_){}
    return ContentService.createTextOutput("OK");
  }
}


function tg_dropPending(){
  const token = getSetting_("TELEGRAM_BOT_TOKEN", "");
  const url = "https://api.telegram.org/bot" + token + "/setWebhook";
  const hook = "https://script.google.com/u/0/macros/s/AKfycbwCV7ech_uYORFuR_ZP8DA4Ahgy_hewK9GJEX7FWQH32GPTqOI8z9nf61avDVgC925Z4w/exec";

  const res = UrlFetchApp.fetch(url, {
    method:"post",
    contentType:"application/json",
    payload: JSON.stringify({ url: hook, drop_pending_updates: true })
  }).getContentText();

  Logger.log(res);
}

