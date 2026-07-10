/*************** FluxFilm Backend — Manage / Renew (Manage.gs) ***************/
/**
 * Phone-only Manage/Renew (NO RATE LIMIT)
 * Returns shape expected by index:
 *   { ok:true, phone, actionable:[], history:[], infoBanner:"" }
 * or { ok:false, message:"" }
 *
 * Uses sheet: "SUBSCRIPTIONS" from CORE sheet via sh_()
 */

// Public function called by frontend
function getMySubscriptions(phone){
  return mg_getMySubscriptions_(phone);
}

function getMySubscriptions_(phone){
  return mg_getMySubscriptions_(phone);
}


/**
 * Create a renewal order for an existing subscription (connects to existing payment flow).
 * Frontend will reuse scrPay (UPI QR + verify).
 *
 * NEW: returns extra preview fields under `renew`:
 *   renew: {
 *     subId, oldExpiry, newExpiry, durationDays,
 *     discount, discountTier, discountTierLabel
 *   }
 */
function createRenewOrder(subId, planOverride, couponCode){
  try {
    const sid = String(subId || "").trim();
    if(!sid) return { ok:false, message:"Missing SubID." };

    const subs = sh_("SUBSCRIPTIONS");
    const sIdx = mg_headerIndex_(subs);

    if(sIdx["SubID"] == null) return { ok:false, message:"SUBSCRIPTIONS missing header: SubID" };

    const last = subs.getLastRow();
    if(last < 2) return { ok:false, message:"No subscriptions found." };

    const data = subs.getRange(2, 1, last - 1, subs.getLastColumn()).getValues();

    var row = null;
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (String(r[sIdx["SubID"]] || "").trim() === sid) { row = r; break; }
    }
    if(!row) return { ok:false, message:"Subscription not found." };

    const service = String(row[sIdx["Service"]] || "").trim();
    let plan = String(row[sIdx["Plan"]] || "").trim();
    const name = (sIdx["Name"] != null) ? String(row[sIdx["Name"]] || "").trim() : "";
    const email = String(row[sIdx["Email"]] || "").trim();
    const phone = mg_normalizePhone_(row[sIdx["Phone"]]);

    planOverride = String(planOverride || "").trim();
    couponCode   = String(couponCode || "").trim().toUpperCase();

    // ✅ Backward compatibility:
    // If old frontend calls createRenewOrder(sid, couponCode) (2 args)
    // ✅ Backward compatibility (safe):
    // Only treat 2nd arg as coupon if it "looks like" a coupon code.
    // Plan names like "Sharing 1M" won't match this.
    if (!couponCode && planOverride) {
      var maybe = String(planOverride || "").trim().toUpperCase();
      var looksLikeCoupon = /^[A-Z0-9_-]{3,20}$/.test(maybe); // no spaces
      if (looksLikeCoupon) {
        couponCode = maybe;
        planOverride = "";
      }
    }


    // ✅ Option 1: duration extension only (same service) → override plan name
    if(planOverride){
      // this must exist in PLANS for the same service (mg_getPlanInfo_ will enforce)
      plan = planOverride;
    }

    if(!service || !plan) return { ok:false, message:"Subscription missing Service/Plan." };
    if(!name) return { ok:false, message:"Name missing for this subscription. Pls contact support " };
    if(!email) return { ok:false, message:"Subscription missing Email." };
    if(!phone) return { ok:false, message:"Subscription missing Phone." };

    // Plan + discount info
    const pInfo = mg_getPlanInfo_(service, plan);

    // Days left (from subscription expiry)
    var expRaw = (sIdx["ExpiryDate"] != null) ? row[sIdx["ExpiryDate"]] : "";
    var expDate = mg_parseDate_(expRaw);
    var daysLeft = null;
    if(expDate){
      daysLeft = Math.ceil((expDate.getTime() - Date.now()) / 86400000);
    }

    // Tiered early-renew discount
    var discObj = mg_calcEarlyDiscount_(daysLeft, pInfo);
    const discountOverride = asNumber_(discObj.amount || 0);

    // Preview: compute new expiry using the renewal base policy:
    // - Advance or late ≤10 days → extend from old expiry date
    // - Late >10 days → extend from today (payment date)
    var durationDays = asNumber_(pInfo.durationDays || 0);
    var LATE_FROM_TODAY_DAYS = asNumber_(getSetting_("RENEW_BASE_TODAY_AFTER_DAYS", 10));
    var baseDate = (expDate && (daysLeft == null || daysLeft >= -LATE_FROM_TODAY_DAYS))
      ? expDate
      : new Date();
    var newExpiry = new Date(baseDate.getTime() + (durationDays * 86400000));

    var tierLabel =
      (String(discObj.tier || "") === "8PLUS") ? "8+ days early" :
      (String(discObj.tier || "") === "7TO2") ? "2–7 days early" :
      "No discount";

    // Reuse existing createOrder() to generate OrderID + UPI link + write ORDERS
    var ord = createOrder({
      service: service,
      plan: plan,
      name: name,
      email: email,
      phone: phone,
      couponCode: couponCode, // optional coupon from frontend 
      action: "RENEW",// ✅ important for coupon scope
      notes: "RENEW:" + sid,
      extraFieldKey: "",
      extraFieldValue: "",
      discountOverride: discountOverride
    });

    if(!ord || !ord.ok) return ord;

    // ✅ Mark this order as a RENEW order linked to this SubID (so fulfillment edits same subscription row)
    try {
      const orders = sh_(TAB_ORDERS); // uses Config TAB_ORDERS
      const rn = getRowByValue_(orders, "OrderID", ord.orderId);
      if (rn) {
        setCellByHeader_(orders, rn, "OrderType", "RENEW");
        setCellByHeader_(orders, rn, "RenewSubID", sid);
      }
    } catch (e) {}

    // ✅ Optional log for audit trail
    try {
      const rsh = sh_("RENEWALS");
      const { idx: rIdx } = headerIndex_(rsh);

      const row = new Array(rsh.getLastColumn()).fill("");

      if (rIdx["RenewID"] != null) row[rIdx["RenewID"]] = "REN-" + Date.now().toString().slice(-8);
      if (rIdx["CreatedAt"] != null) row[rIdx["CreatedAt"]] = now_();
      if (rIdx["OrderID"] != null) row[rIdx["OrderID"]] = ord.orderId;
      if (rIdx["SubID"] != null) row[rIdx["SubID"]] = sid;
      if (rIdx["OldExpiry"] != null) row[rIdx["OldExpiry"]] = expDate ? expDate : "";
      if (rIdx["NewExpiry"] != null) row[rIdx["NewExpiry"]] = newExpiry;
      if (rIdx["DurationDays"] != null) row[rIdx["DurationDays"]] = durationDays;
      if (rIdx["Phone"] != null) row[rIdx["Phone"]] = phone;
      if (rIdx["Email"] != null) row[rIdx["Email"]] = email;
      if (rIdx["Service"] != null) row[rIdx["Service"]] = service;
      if (rIdx["Plan"] != null) row[rIdx["Plan"]] = plan;
      if (rIdx["Status"] != null) row[rIdx["Status"]] = "CREATED";
      if (rIdx["Notes"] != null) row[rIdx["Notes"]] = "";

      rsh.appendRow(row);
    } catch(e) {}

    // Attach renewal preview fields for frontend popup / summary
    ord.renew = {
      subId: sid,
      oldExpiry: expDate ? expDate.toISOString() : "",
      newExpiry: newExpiry.toISOString(),
      durationDays: durationDays,
      discount: discountOverride,
      discountTier: String(discObj.tier || "NONE"),
      discountTierLabel: tierLabel
    };

    if (typeof notifyTelegram_ === "function") {
      notifyTelegram_("RENEW_ORDER_CREATED", {
        orderId: ord.orderId,
        subId: sid,
        service, plan,
        name, phone, email,
        amount: ord.finalAmount,
        discount: discountOverride,
        expiry: (expDate ? expDate.toISOString() : ""),
        note: "Renew initiated"
      });
    }
    return ord;

  } catch (e) {
    return { ok:false, message:"Renew error: " + (e && e.message ? e.message : e) };
  }
}

// Internal (unique name to avoid global collisions)
function mg_getMySubscriptions_(phone) {
  try {
    const ph = mg_normalizePhone_(phone);
    if (!ph) return { ok: false, message: "Phone is required." };

    // Use CORE sheet helper (from Utils.gs)
    const subs = sh_("SUBSCRIPTIONS");
    const idx = mg_headerIndex_(subs);

    const required = ["SubID","OrderID","Phone","Email","Service","Plan","StartDate","ExpiryDate","InventoryRef","ProfileName","ProfilePIN","ProfileNumber"];
    for (var i = 0; i < required.length; i++) {
      var h = required[i];
      if (idx[h] == null) {
        return { ok: false, message: "SUBSCRIPTIONS missing header: " + h };
      }
    }

    const last = subs.getLastRow();
    if (last < 2) {
      return JSON.parse(JSON.stringify({
        ok: true,
        phone: ph,
        infoBanner: "📱 Please enter the same phone number you used to buy subscriptions.",
        actionable: [],
        history: []
      }));
    }

    const data = subs.getRange(2, 1, last - 1, subs.getLastColumn()).getValues();
    const nowMs = Date.now();

    const all = [];

    for (var r = 0; r < data.length; r++) {
      var row = data[r];

      var rowPhone = mg_normalizePhone_(row[idx["Phone"]]);
      if (rowPhone !== ph) continue;

      var svc = String(row[idx["Service"]] || "").trim();
      var plan = String(row[idx["Plan"]] || "").trim();
      var email = String(row[idx["Email"]] || "").trim();

      var startRaw = row[idx["StartDate"]];
      var startDate = mg_parseDate_(startRaw);

      var expRaw = row[idx["ExpiryDate"]];
      var expDate = mg_parseDate_(expRaw);
      var expMs = expDate ? expDate.getTime() : NaN;

      var daysLeft = isNaN(expMs) ? null : Math.ceil((expMs - nowMs) / 86400000);

      var elig = mg_renewEligibility_(daysLeft);
      var mood = mg_expiryMood_(daysLeft);

      // Plan info from PLANS sheet (DurationDays + Price + Discounts + LogoUrl)
      var pInfo = mg_getPlanInfo_(svc, plan);
      var durationDays = asNumber_(pInfo.durationDays || 0);

      var discObj = mg_calcEarlyDiscount_(daysLeft, pInfo);
      var eligibleDisc = asNumber_(discObj.amount || 0);
      var discTier = String(discObj.tier || "NONE");

      all.push({
        subId: String(row[idx["SubID"]] || "").trim(),
        orderId: String(row[idx["OrderID"]] || "").trim(),
        service: svc,
        plan: plan,
        maskedEmail: mg_maskEmailFirst4_(email),

        // Optional (for UI icon)
        logoUrl: String(pInfo.logoUrl || "").trim(),

        // JSON-safe dates
        startDate: (startRaw instanceof Date) ? startRaw.toISOString() : (startRaw ? String(startRaw) : ""),
        expiryDate: (expRaw instanceof Date) ? expRaw.toISOString() : (expRaw ? String(expRaw) : ""),

        durationDays: durationDays,

        // Tiered discount fields for UI (two-strip locked/unlocked)
        earlyRenewDiscountEligible: eligibleDisc,
        earlyRenewDiscountTier: discTier,
        earlyDiscount8Plus: asNumber_(pInfo.d8plus || 0),
        earlyDiscount7to2: asNumber_(pInfo.d7to2 || 0),

        profileNumber: String(row[idx["ProfileNumber"]] || "").trim(),
        profileName: String(row[idx["ProfileName"]] || "").trim(),
        profilePIN: String(row[idx["ProfilePIN"]] || "").trim(),

        daysLeft: daysLeft,
        moodEmoji: mood.emoji,
        moodText: mood.text,
        renewEligibility: elig,
        uiTone: (elig === "CAN_RENEW") ? "normal" : (elig === "LATE_RENEW" ? "faded_red" : "faded_grey"),
        showRenewButton: (elig !== "TOO_LATE"),
        inventoryRef: String(row[idx["InventoryRef"]] || "").trim()
      });
    }

    var actionable = [];
    var tooLate = [];

    for (var j = 0; j < all.length; j++) {
      if (all[j].renewEligibility === "TOO_LATE") tooLate.push(all[j]);
      else actionable.push(all[j]);
    }

    // Soonest expiry first
    actionable.sort(function(a, b){
      var ax = mg_parseDate_(a.expiryDate);
      var bx = mg_parseDate_(b.expiryDate);
      return (ax ? ax.getTime() : 9e15) - (bx ? bx.getTime() : 9e15);
    });

    // Most recent expired first
    tooLate.sort(function(a, b){
      var ax = mg_parseDate_(a.expiryDate);
      var bx = mg_parseDate_(b.expiryDate);
      return (bx ? bx.getTime() : 0) - (ax ? ax.getTime() : 0);
    });

    // Force JSON-safe response
    return JSON.parse(JSON.stringify({
      ok: true,
      phone: ph,
      infoBanner: "📱 Please enter the same phone number you used to buy subscriptions.",
      actionable: actionable,
      history: tooLate.slice(0, 3)
    }));

  } catch (e) {
    return {
      ok: false,
      message: "Manage error: " + (e && e.message ? e.message : e)
    };
  }
}


/* -------------------- Helpers (unique mg_* names) -------------------- */

function mg_getPlanInfo_(service, plan){
  try {
    const plans = sh_("PLANS");
    const idx = mg_headerIndex_(plans);

    if(idx["Service"] == null || idx["Plan"] == null) return { durationDays:0, price:0, d8plus:0, d7to2:0, logoUrl:"" };

    const last = plans.getLastRow();
    if(last < 2) return { durationDays:0, price:0, d8plus:0, d7to2:0, logoUrl:"" };

    const rows = plans.getRange(2, 1, last - 1, plans.getLastColumn()).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var s = String(r[idx["Service"]] || "").trim();
      var p = String(r[idx["Plan"]] || "").trim();
      if(s === String(service||"").trim() && p === String(plan||"").trim()){
        return {
          durationDays: (idx["DurationDays"] != null) ? asNumber_(r[idx["DurationDays"]]) : 0,
          price: (idx["Price"] != null) ? asNumber_(r[idx["Price"]]) : 0,
          d8plus: (idx["EarlyRenewDiscount"] != null) ? asNumber_(r[idx["EarlyRenewDiscount"]]) : 0,
          d7to2: (idx["EarlyRenewDiscount_7to2"] != null) ? asNumber_(r[idx["EarlyRenewDiscount_7to2"]]) : 0,
          logoUrl: (idx["LogoUrl"] != null) ? String(r[idx["LogoUrl"]] || "").trim() : ""
        };
      }
    }
    return { durationDays:0, price:0, d8plus:0, d7to2:0, logoUrl:"" };
  } catch(e){
    return { durationDays:0, price:0, d8plus:0, d7to2:0, logoUrl:"" };
  }
}

function mg_calcEarlyDiscount_(daysLeft, planInfo){
  var d8 = asNumber_(planInfo && planInfo.d8plus || 0);
  var d72 = asNumber_(planInfo && planInfo.d7to2 || 0);

  // SETTINGS (fallback defaults)
  // Put these in SETTINGS sheet if you want to tweak without code:
  // EARLY_RENEW_TIER1_MIN_DAYS = 8
  // EARLY_RENEW_TIER2_MIN_DAYS = 2
  // EARLY_RENEW_TIER2_MAX_DAYS = 7
  var TIER1_MIN = asNumber_(getSetting_("EARLY_RENEW_TIER1_MIN_DAYS", 8));  // 8+ days
  var TIER2_MIN = asNumber_(getSetting_("EARLY_RENEW_TIER2_MIN_DAYS", 2));  // 2+ days
  var TIER2_MAX = asNumber_(getSetting_("EARLY_RENEW_TIER2_MAX_DAYS", 7));  // up to 7

  if(daysLeft == null) return { amount:0, tier:"NONE" };
  if(daysLeft >= TIER1_MIN && d8 > 0) return { amount:d8, tier:"8PLUS" };
  if(daysLeft >= TIER2_MIN && daysLeft <= TIER2_MAX && d72 > 0) return { amount:d72, tier:"7TO2" };
  return { amount:0, tier:"NONE" };
}

function mg_headerIndex_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idx = {};
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || "").trim();
    if (h) idx[h] = c;
  }
  return idx;
}

function mg_renewEligibility_(daysLeft) {
  if (daysLeft == null) return "TOO_LATE";
  if (daysLeft >= 0) return "CAN_RENEW";
  if (daysLeft >= -5) return "LATE_RENEW";
  return "TOO_LATE";
}

function mg_expiryMood_(daysLeft) {
  if (daysLeft == null) return { emoji: "❓", text: "Expiry unknown" };
  if (daysLeft > 10) return { emoji: "😄", text: "Safe" };
  if (daysLeft >= 6) return { emoji: "🙂", text: "All good" };
  if (daysLeft >= 1) return { emoji: "😰", text: "Expiring soon" };
  if (daysLeft === 0) return { emoji: "⚠️", text: "Expires today" };
  if (daysLeft >= -5) return { emoji: "😵", text: "Expired (late renew allowed)" };
  return { emoji: "🟥", text: "Expired" };
}

function mg_normalizePhone_(p) {
  var s = String(p || "").replace(/\D/g, "");
  if (!s) return "";
  if (s.length > 10) s = s.slice(-10);
  return s;
}

function mg_maskEmailFirst4_(email) {
  var e = String(email || "").trim();
  var at = e.indexOf("@");
  if (at < 0) return e ? (e.slice(0, 4) + "****") : "";
  var local = e.slice(0, at);
  var domain = e.slice(at);
  var head = local.slice(0, 4);
  return head + "****" + domain;
}

function mg_parseDate_(x) {
  if (!x) return null;
  if (x instanceof Date && !isNaN(x.getTime())) return x;

  var s = String(x).trim();

  // DD/MM/YYYY HH:mm:ss
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    return new Date(
      Number(m[3]),
      Number(m[2]) - 1,
      Number(m[1]),
      Number(m[4] || 0),
      Number(m[5] || 0),
      Number(m[6] || 0)
    );
  }

  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}