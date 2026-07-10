/*************** FluxFilm Backend — Coupons ***************/
/**
 * Sheets used:
 *  - COUPONS
 *  - COUPON_USAGE   Columns:
 *      [Timestamp, CouponCode, Phone, Email, Discount, OrderID, Action]
 *
 * IMPORTANT RULE:
 *  - validateCoupon() NEVER consumes coupon.
 *  - PerUserLimit / GlobalLimit count ONLY rows where Action === "USED"
 *
 * Optional:
 *  - logCouponUse_(..., "HOLD") at order creation (does NOT count)
 *  - logCouponUse_(..., "USED") when payment becomes PAID (counts)
 */

/**
 * Validate coupon and compute discount.
 * Does NOT write to COUPON_USAGE.
 */
function validateCoupon(code, ctx) {
  code = String(code || "").trim().toUpperCase();
  if (!code) return { ok: false, message: "Enter coupon code" };

  const phone = String((ctx && ctx.phone) || "").replace(/\D/g, "");
  if (!phone) return { ok: false, message: "Phone required" };

  const amount = Number(ctx.amount || 0);
  const service = String(ctx.service || "").trim();
  const plan = String(ctx.plan || "").trim();
  const scope = String(ctx.scope || "NEW").toUpperCase(); // NEW / RENEW

  const row = findCouponRow_(code);
  if (!row) return { ok: false, message: "Invalid coupon" };

  if (String(row.Active).toUpperCase() !== "TRUE") {
    return { ok: false, message: "Coupon inactive" };
  }

  if (row.Expiry) {
    const exp = new Date(row.Expiry);
    if (!isNaN(exp) && new Date() > endOfDay_(exp)) {
      return { ok: false, message: "Coupon expired" };
    }
  }

  const allowedScope = String(row.Scope || "ANY").toUpperCase();
  if (allowedScope !== "ANY" && allowedScope !== scope) {
    return { ok: false, message: `Only valid for ${allowedScope.toLowerCase()}` };
  }

  const minAmt = Number(row.MinAmount || 0);
  if (amount < minAmt) {
    return { ok: false, message: `Min amount ₹${minAmt}` };
  }

  if (row.Services) {
    const allowed = String(row.Services).split(",").map(s => s.trim()).filter(Boolean);
    if (allowed.length && !allowed.includes(service)) {
      return { ok: false, message: "Not valid for this service" };
    }
  }

  if (row.Plans) {
    const allowed = String(row.Plans).split(",").map(s => s.trim()).filter(Boolean);
    if (allowed.length && !allowed.includes(plan)) {
      return { ok: false, message: "Not valid for this plan" };
    }
  }

  // First time only: block if phone already has PAID orders
  if (String(row.FirstTimeOnly).toUpperCase() === "TRUE") {
    if (phoneHasPaidOrder_(phone)) {
      return { ok: false, message: "Only for first-time customers" };
    }
  }

  // per user limit (counts ONLY USED)
  const perUser = Number(row.PerUserLimit || 0);
  if (perUser > 0) {
    const used = countCouponUsageByPhone_(code, phone); // ONLY USED
    if (used >= perUser) return { ok: false, message: "Coupon already used" };
  }

  // global limit (counts ONLY USED)
  const globalLim = Number(row.GlobalLimit || 0);
  if (globalLim > 0) {
    const used = countCouponUsageTotal_(code); // ONLY USED
    if (used >= globalLim) return { ok: false, message: "Coupon limit reached" };
  }

  // compute discount
  const type = String(row.Type || "").toUpperCase();
  let discount = 0;

  if (type === "FLAT") {
    discount = Number(row.Value || 0);
  } else if (type === "PERCENT") {
    discount = Math.round(amount * (Number(row.Value || 0) / 100));
  } else {
    return { ok: false, message: "Coupon misconfigured" };
  }

  const cap = Number(row.MaxDiscount || 0);
  if (cap > 0) discount = Math.min(discount, cap);
  discount = Math.max(0, Math.min(discount, amount));

  return {
    ok: true,
    code,
    discount,
    finalAmount: amount - discount,
    message: `Applied ✅ ₹${discount} off`
  };
}

/**
 * Append a coupon usage row.
 * Action:
 *  - "HOLD" = reserved/attempted (does NOT count in limits)
 *  - "USED" = consumed after payment (counts)
 *  - "CANCEL" = optional
 *
 * COUPON_USAGE columns:
 *   A Timestamp
 *   B CouponCode
 *   C Phone
 *   D Email
 *   E Discount
 *   F OrderID
 *   G Action
 */
function logCouponUse_(code, phone, email, discount, orderId, action) {
  const ss = getCore_();
  const sh = ss.getSheetByName("COUPON_USAGE");
  if (!sh) return;

  sh.appendRow([
    new Date(),
    String(code || "").toUpperCase(),
    String(phone || "").replace(/\D/g, ""),
    String(email || "").trim(),
    Number(discount || 0),
    String(orderId || "").trim(),
    String(action || "USED").toUpperCase()
  ]);
}

/**
 * Optional helper: log HOLD at order creation (safe).
 */
function holdCouponForOrder_(code, phone, email, discount, orderId) {
  if (!code) return;
  logCouponUse_(code, phone, email, discount, orderId, "HOLD");
}

/**
 * Recommended helper: mark coupon USED on payment success.
 * This should be called only after order is PAID.
 */
function markCouponUsedForOrder_(code, phone, email, discount, orderId) {
  if (!code) return;
  logCouponUse_(code, phone, email, discount, orderId, "USED");
}

/*** helpers ***/
function findCouponRow_(code) {
  const ss = getCore_();
  const sh = ss.getSheetByName("COUPONS");
  if (!sh) return null;

  const values = sh.getDataRange().getValues();
  const head = values[0].map(String);

  const idx = {};
  head.forEach((h, i) => idx[h] = i);

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[idx.Code] || "").trim().toUpperCase() === code) {
      return {
        Code: row[idx.Code],
        Type: row[idx.Type],
        Value: row[idx.Value],
        MinAmount: row[idx.MinAmount],
        MaxDiscount: row[idx.MaxDiscount],
        Active: row[idx.Active],
        Expiry: row[idx.Expiry],
        Scope: row[idx.Scope],
        FirstTimeOnly: row[idx.FirstTimeOnly],
        PerUserLimit: row[idx.PerUserLimit],
        GlobalLimit: row[idx.GlobalLimit],
        Services: row[idx.Services],
        Plans: row[idx.Plans]
      };
    }
  }
  return null;
}

function phoneHasPaidOrder_(phone) {
  const ss = getCore_();
  const sh = ss.getSheetByName("ORDERS");
  const vals = sh.getDataRange().getValues();
  const head = vals[0].map(String);

  const iPhone = head.indexOf("Phone");
  const iStatus = head.indexOf("Status");
  if (iPhone < 0 || iStatus < 0) return false;

  for (let r = 1; r < vals.length; r++) {
    if (
      String(vals[r][iPhone]).replace(/\D/g, "") === phone &&
      String(vals[r][iStatus]).toUpperCase() === "PAID"
    ) return true;
  }
  return false;
}

/**
 * Counts ONLY Action === "USED" for a specific phone.
 * Uses header-based column lookup to be robust against any column order in COUPON_USAGE.
 */
function countCouponUsageByPhone_(code, phone) {
  const ss = getCore_();
  const sh = ss.getSheetByName("COUPON_USAGE");
  if (!sh) return 0;

  code  = String(code  || "").trim().toUpperCase();
  phone = String(phone || "").replace(/\D/g, "");

  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return 0;

  // Build header index from row 0
  const head = vals[0].map(String);
  const iCode   = head.indexOf("CouponCode");
  const iPhone  = head.indexOf("Phone");
  const iAction = head.indexOf("Action");

  // Fall back to positional if headers are missing (legacy safety)
  const getCode   = iCode   >= 0 ? r => String(vals[r][iCode]   || "").toUpperCase().trim() : r => String(vals[r][1] || "").toUpperCase().trim();
  const getPhone  = iPhone  >= 0 ? r => String(vals[r][iPhone]  || "").replace(/\D/g, "")   : r => String(vals[r][2] || "").replace(/\D/g, "");
  const getAction = iAction >= 0 ? r => String(vals[r][iAction] || "").toUpperCase().trim() : r => String(vals[r][6] || "").toUpperCase().trim();

  let count = 0;
  for (let r = 1; r < vals.length; r++) {
    if (getCode(r) === code && getPhone(r) === phone && getAction(r) === "USED") count++;
  }
  return count;
}

/**
 * Counts ONLY Action === "USED" globally for a coupon code.
 * Uses header-based column lookup to be robust against any column order in COUPON_USAGE.
 */
function countCouponUsageTotal_(code) {
  const ss = getCore_();
  const sh = ss.getSheetByName("COUPON_USAGE");
  if (!sh) return 0;

  code = String(code || "").trim().toUpperCase();

  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return 0;

  // Build header index from row 0
  const head = vals[0].map(String);
  const iCode   = head.indexOf("CouponCode");
  const iAction = head.indexOf("Action");

  const getCode   = iCode   >= 0 ? r => String(vals[r][iCode]   || "").toUpperCase().trim() : r => String(vals[r][1] || "").toUpperCase().trim();
  const getAction = iAction >= 0 ? r => String(vals[r][iAction] || "").toUpperCase().trim() : r => String(vals[r][6] || "").toUpperCase().trim();

  let count = 0;
  for (let r = 1; r < vals.length; r++) {
    if (getCode(r) === code && getAction(r) === "USED") count++;
  }
  return count;
}

function endOfDay_(d) {
  const x = new Date(d);
  x.setHours(23, 59, 59, 999);
  return x;
}

/******************* ANALYTICS *******************/

/**
 * Returns high-level coupon analytics
 * - usedToday
 * - used7d
 * - used30d
 * - discountToday
 * - topCoupons (last 30d)
 */
function getCouponAnalyticsSummary() {
  const ss = getCore_();
  const sh = ss.getSheetByName("COUPON_USAGE");
  if (!sh) return ok_({ usedToday:0, used7d:0, used30d:0, discountToday:0, topCoupons:[] });

  const vals = sh.getDataRange().getValues();
  const now = new Date();

  const startOfToday = new Date(now); startOfToday.setHours(0,0,0,0);
  const d7 = new Date(now); d7.setDate(d7.getDate()-7);
  const d30 = new Date(now); d30.setDate(d30.getDate()-30);

  let usedToday=0, used7d=0, used30d=0, discountToday=0;
  const byCode30 = {}; // {CODE: {uses, discount}}

  for (let r=1; r<vals.length; r++){
    const ts = vals[r][0];
    const code = String(vals[r][1]||"").toUpperCase();
    const disc = Number(vals[r][4]||0);                 // Discount
    const action = String(vals[r][6]||"").toUpperCase(); // Action

    if (action !== "USED") continue;
    const dt = new Date(ts);
    if (isNaN(dt)) continue;

    if (dt >= d30) {
      used30d++;
      byCode30[code] = byCode30[code] || { code, uses:0, discount:0 };
      byCode30[code].uses++;
      byCode30[code].discount += disc;
    }
    if (dt >= d7) used7d++;
    if (dt >= startOfToday) {
      usedToday++;
      discountToday += disc;
    }
  }

  const topCoupons = Object.values(byCode30)
    .sort((a,b)=> (b.uses - a.uses) || (b.discount - a.discount))
    .slice(0,10);

  return ok_({
    usedToday,
    used7d,
    used30d,
    discountToday: Math.round(discountToday),
    topCoupons
  });
}

/**
 * Per-coupon stats (USED only):
 * - totalUses
 * - totalDiscount
 * - uniquePhones
 * - lastUsedAt
 */
function getCouponStats(code) {
  code = String(code||"").trim().toUpperCase();
  if(!code) return bad_("Code required");

  const ss = getCore_();
  const sh = ss.getSheetByName("COUPON_USAGE");
  if (!sh) return ok_({ code, totalUses:0, totalDiscount:0, uniquePhones:0, lastUsedAt:null });

  const vals = sh.getDataRange().getValues();
  let totalUses=0, totalDiscount=0;
  const phones = {};
  let lastUsedAt = null;

  for (let r=1; r<vals.length; r++){
    const cCode = String(vals[r][1]||"").toUpperCase();
    const action = String(vals[r][6]||"").toUpperCase();
    if(cCode !== code || action !== "USED") continue;

    totalUses++;
    totalDiscount += Number(vals[r][4]||0); // Discount

    const ph = String(vals[r][2]||"").replace(/\D/g,""); // Phone
    if(ph) phones[ph]=true;

    const dt = new Date(vals[r][0]);
    if(!isNaN(dt) && (!lastUsedAt || dt > lastUsedAt)) lastUsedAt = dt;
  }

  return ok_({
    code,
    totalUses,
    totalDiscount: Math.round(totalDiscount),
    uniquePhones: Object.keys(phones).length,
    lastUsedAt: lastUsedAt ? lastUsedAt.toISOString() : null
  });
}

/**
 * Recent usage feed (USED only)
 */
function getRecentCouponUses(limit) {
  limit = Math.max(1, Math.min(200, Number(limit||50)));

  const ss = getCore_();
  const sh = ss.getSheetByName("COUPON_USAGE");
  if (!sh) return ok_({ rows: [] });

  const vals = sh.getDataRange().getValues();
  const out = [];

  for (let r = vals.length - 1; r >= 1 && out.length < limit; r--) {
    const action = String(vals[r][6]||"").toUpperCase();
    if (action !== "USED") continue;

    out.push({
      ts: vals[r][0],
      code: vals[r][1],
      phone: vals[r][2],
      email: vals[r][3],
      discount: Number(vals[r][4]||0),
      orderId: vals[r][5],
      action: vals[r][6]
    });
  }

  return ok_({ rows: out });
}