/*************** FluxFilm Backend — Orders ***************/

/**
 * BOOTSTRAP: returns active plans (sheet-driven) + currency.
 * Now also returns:
 *  - benefits: array of strings (split by newlines)
 *  - badgeText: string
 *  - logoUrl: optional string (if column exists)
 */
function getBootstrap() {
  const plans = sh_(TAB_PLANS);
  const { idx } = headerIndex_(plans);

  const required = [
    "Service","Plan","DurationDays","Price","IsActive",
    "FulfillmentMode","AllocationPolicy",
    "NeedsExtraField","ExtraFieldKey","ExtraFieldLabel",
    "DeviceRuleText","PostPaymentMessage","EarlyRenewDiscount"
  ];
  required.forEach(h => { if (idx[h] == null) throw new Error(`PLANS missing header: ${h}`); });

  const last = plans.getLastRow();
  const rows = last < 2 ? [] : plans.getRange(2, 1, last - 1, plans.getLastColumn()).getValues();

  const out = [];
  rows.forEach(r => {
    const active = String(r[idx["IsActive"]] || "").trim().toUpperCase() === "TRUE";
    if (!active) return;

    const benefitsRaw = (idx["Benefits"] != null) ? String(r[idx["Benefits"]] || "") : "";
    const badgeText = (idx["BadgeText"] != null) ? String(r[idx["BadgeText"]] || "").trim() : "";
    const logoUrl  = (idx["LogoUrl"]  != null) ? String(r[idx["LogoUrl"]]  || "").trim() : "";


    const service_ = String(r[idx["Service"]] || "").trim();  
    const plan_    = String(r[idx["Plan"]]    || "").trim();

    out.push({
      service: service_,
      plan: plan_,
      durationDays: asNumber_(r[idx["DurationDays"]]),
      price: asNumber_(r[idx["Price"]]),
      fulfillmentMode: String(r[idx["FulfillmentMode"]] || "").trim().toUpperCase(),
      allocationPolicy: String(r[idx["AllocationPolicy"]] || "").trim().toUpperCase(),
      needsExtraField: String(r[idx["NeedsExtraField"]] || "").trim().toUpperCase() === "TRUE",
      extraFieldKey: String(r[idx["ExtraFieldKey"]] || "").trim(),
      extraFieldLabel: String(r[idx["ExtraFieldLabel"]] || "").trim(),
      deviceRuleText: String(r[idx["DeviceRuleText"]] || "").trim(),
      postPaymentMessage: String(r[idx["PostPaymentMessage"]] || "").trim(),
      earlyRenewDiscount: asNumber_(r[idx["EarlyRenewDiscount"]]),
      requiresGroupJoin: String(r[idx["RequiresGroupJoin"]] || "").trim().toUpperCase() === "TRUE",
      groupJoinLink: String(r[idx["GroupJoinLink"]] || "").trim(),
      groupSuspendWarningDays: asNumber_(r[idx["GroupSuspendWarningDays"]]),
      benefits: benefitsRaw ? benefitsRaw.split(/\r?\n/).map(x => x.trim()).filter(Boolean) : [],
      badgeText: badgeText,
      logoUrl: logoUrl,
    });
  });

  return ok_({
    plans: out,
    currency: String(getSetting_("CURRENCY", "INR"))
  });
}

/**
 * payload:
 * {
 *  service, plan, email?, phone, couponCode?, notes?, extraFieldKey?, extraFieldValue?,
 *  discountOverride? (optional)
 * }
 *
 * NOTE: You said "KEEP PHONE ONLY" for coupons. Email is still allowed for receipt,
 * but coupon logic is phone-based.
 */
function createOrder(payload) {
  const p = payload || {};
  const service = String(p.service || "").trim();
  const plan = String(p.plan || "").trim();
  const name = String(p.name || "").trim();
  const email = String(p.email || "").trim(); // optional later if you want, but kept
  const phoneRaw = String(p.phone || "").trim();
  const phone = normalizePhone_(phoneRaw);

  const couponCode = String(p.couponCode || "").trim().toUpperCase();
  const action = String(p.action || "NEW").trim().toUpperCase(); // NEW / RENEW
  const notes = String(p.notes || "").trim();
  const extraKey = String(p.extraFieldKey || "").trim();
  const extraVal = String(p.extraFieldValue || "").trim();

  if (!service || !plan) return bad_("Select service and plan.");
  if (!phone) return bad_("Phone number is required.");
  if(!name) return bad_("Full name is required.");
  // Keep email required if you want receipts; otherwise comment the next line:
  if (!email) return bad_("Email is required.");

  // Find plan row
  const plans = sh_(TAB_PLANS);
  const { idx: pIdx } = headerIndex_(plans);
  const last = plans.getLastRow();
  if (last < 2) return bad_("No plans configured.");
  const data = plans.getRange(2, 1, last - 1, plans.getLastColumn()).getValues();

  let planRow = null;
  for (const r of data) {
    if (String(r[pIdx["Service"]] || "").trim() === service &&
        String(r[pIdx["Plan"]] || "").trim() === plan) {
      const active = String(r[pIdx["IsActive"]] || "").trim().toUpperCase() === "TRUE";
      if (active) { planRow = r; break; }
    }
  }
  if (!planRow) return bad_("Plan not found or inactive.");

  const durationDays = asNumber_(planRow[pIdx["DurationDays"]]);
  const price = asNumber_(planRow[pIdx["Price"]]);
  const currency = String(getSetting_("CURRENCY", "INR"));

  // Group join fields (from plan)
  const groupJoinRequired = String(planRow[pIdx["RequiresGroupJoin"]] || "").trim().toUpperCase() === "TRUE";
  const groupJoinLink = String(planRow[pIdx["GroupJoinLink"]] || "").trim();
  const groupWarnDays = asNumber_(planRow[pIdx["GroupSuspendWarningDays"]] || 1);
  const suspendIfFalseOn = groupJoinRequired ? addDays_(now_(), groupWarnDays) : "";

  // --- COUPON ENGINE ---
  let discount = 0;

  if (couponCode) {
    const res = validateCoupon_({
      code: couponCode,
      phone,
      email,
      service,
      plan,
      baseAmount: price,
      action: action
    });
    if (!res.ok) return res;
    discount = asNumber_(res.discount);
  }

  // Optional manual override (kept)
  // Optional manual override ONLY when no coupon is used.
  // Prevents coupon discount from being doubled.
  const discountOverride = asNumber_(p.discountOverride || 0);

  if (!couponCode && discountOverride > 0) {
    discount += discountOverride;
  }

  // Safety: discount can never be more than plan price.
  discount = Math.max(0, Math.min(discount, price));

  const finalAmount = Math.max(0, price - discount);
  const orderId = genOrderId_();

  // Write order by headers (future-proof)
  const orders = sh_(TAB_ORDERS);
  const { idx: oIdx } = headerIndex_(orders);
  const row = new Array(orders.getLastColumn()).fill("");

  function put(h, v) {
    if (oIdx[h] == null) return; // safe if sheet changes
    row[oIdx[h]] = v;
  }

  put("OrderID", orderId);
  put("CreatedAt", now_());
  put("Service", service);
  put("Plan", plan);
  put("DurationDays", durationDays);

  put("Name", name);
  put("Email", email);
  put("Phone", phone);

  put("CouponCode", couponCode);
  put("Discount", discount);
  put("FinalAmount", finalAmount);
  put("Currency", currency);

  put("Notes", notes);
  put("ExtraFieldKey", extraKey);
  put("ExtraFieldValue", extraVal);

  put("Status", ORDER_STATUS.CREATED);
  put("FulfillmentStatus", FULFILL_STATUS.PENDING);

  // Group join columns (optional)
  put("GroupJoinRequired", groupJoinRequired ? "TRUE" : "FALSE");
  put("GroupJoinLink", groupJoinLink || "");
  put("GroupJoined", "FALSE");
  put("GroupJoinedAt", "");
  put("SuspendIfFalseOn", suspendIfFalseOn || "");

  orders.appendRow(row);

  // Coupon usage log (APPLIED) — matches COUPON_USAGE headers:
  // Timestamp, CouponCode, Phone, OrderID, Discount, Action
  if (couponCode && discount > 0 && action !== "RENEW") {
    recordCouponUsage_({
      phone,
      couponCode,
      orderId,
      discount,
      action: "APPLIED"
    });
  }

  if (typeof notifyTelegram_ === "function") {
    notifyTelegram_("ORDER_CREATED", {
      orderId, service, plan,
      name, phone, email,
      amount: finalAmount,
      discount: discount,
      note: notes
    });
  }


  // UPI deep link
  const upiVpa = String(getSetting_("UPI_VPA", "YOURUPI@BANK"));
  const payee = String(getSetting_("UPI_PAYEE_NAME", "FluxFilm"));
  const tn = encodeURIComponent(orderId);
  const upiLink = `upi://pay?pa=${encodeURIComponent(upiVpa)}&pn=${encodeURIComponent(payee)}&am=${encodeURIComponent(finalAmount)}&cu=${encodeURIComponent(currency)}&tn=${tn}`;


  return ok_({
    orderId,
    amount: finalAmount,
    baseAmount: price,
    discount,
    couponCode: couponCode || "",
    currency,
    upiVpa,
    payee,
    upiLink,
    paymentNote: orderId,
    groupJoinRequired,
    groupJoinLink
  });
}


// ===================== COUPONS =====================

/**
 * Internal coupon validator (sheet-driven).
 * Returns: ok_({discount, coupon})
 */
function validateCoupon_(ctx) {
  const code = String(ctx.code || "").trim().toUpperCase();
  const phone = normalizePhone_(String(ctx.phone || "").trim());
  const email = String(ctx.email || "").trim();
  const service = String(ctx.service || "").trim();
  const plan = String(ctx.plan || "").trim();
  const baseAmount = asNumber_(ctx.baseAmount);
  const action = String(ctx.action || "ANY").trim().toUpperCase(); // NEW / RENEW / ANY

  if (!code) return ok_({ discount: 0, coupon: null });

  const coupons = sh_(TAB_COUPONS);
  const { idx } = headerIndex_(coupons);

  const needed = [
    "Code","Type","Value","MinAmount","MaxDiscount","Active","Expiry","Scope",
    "FirstTimeOnly","PerUserLimit","GlobalLimit","Services","Plans"
  ];
  needed.forEach(h => { if (idx[h] == null) throw new Error(`COUPONS missing header: ${h}`); });

  const last = coupons.getLastRow();
  if (last < 2) return bad_("Invalid coupon.");

  const rows = coupons.getRange(2, 1, last - 1, coupons.getLastColumn()).getValues();

  let row = null;
  for (const r of rows) {
    if (String(r[idx["Code"]] || "").trim().toUpperCase() === code) { row = r; break; }
  }
  if (!row) return bad_("Invalid coupon.");

  const active = String(row[idx["Active"]] || "").trim().toUpperCase() === "TRUE";
  if (!active) return bad_("Coupon not active.");

  // Expiry
  const expiryCell = row[idx["Expiry"]];
  if (expiryCell) {
    const exp = new Date(expiryCell);
    if (!isNaN(exp.getTime()) && now_().getTime() > exp.getTime()) return bad_("Coupon expired.");
  }

  // Scope
  const scope = String(row[idx["Scope"]] || "ANY").trim().toUpperCase();
  if (scope !== "ANY" && scope !== action) return bad_("Coupon not valid for this action.");

  // Min amount
  const minAmount = asNumber_(row[idx["MinAmount"]]);
  if (minAmount && baseAmount < minAmount) return bad_(`Minimum amount for this coupon is ₹${Math.round(minAmount)}.`);

  // Service restriction allowlist
  const services = String(row[idx["Services"]] || "").trim();
  if (services) {
    const allow = services.split(",").map(x => x.trim()).filter(Boolean);
    if (allow.length && !allow.includes(service)) return bad_("Coupon not valid for this service.");
  }

  // Plan restriction allowlist
  const plans = String(row[idx["Plans"]] || "").trim();
  if (plans) {
    const allow = plans.split(",").map(x => x.trim()).filter(Boolean);
    if (allow.length && !allow.includes(plan)) return bad_("Coupon not valid for this plan.");
  }

  // First-time only (phone-based): no PAID order for this phone
  const firstOnly = String(row[idx["FirstTimeOnly"]] || "").trim().toUpperCase() === "TRUE";
  if (firstOnly) {
    const everPaid = hasPaidOrderForPhone_(phone);
    if (everPaid) return bad_("Coupon only for first-time users.");
  }

  // Limits
  const perUserLimit = asNumber_(row[idx["PerUserLimit"]]);
  const globalLimit = asNumber_(row[idx["GlobalLimit"]]);

  const usage = couponUsageCounts_(code, phone);
  if (perUserLimit && usage.userCount >= perUserLimit) return bad_("Coupon already used on your number.");
  if (globalLimit && usage.globalCount >= globalLimit) return bad_("Coupon limit reached.");

  // Discount calc
  const type = String(row[idx["Type"]] || "FLAT").trim().toUpperCase(); // FLAT / PERCENT
  const value = asNumber_(row[idx["Value"]]);
  const maxDiscount = asNumber_(row[idx["MaxDiscount"]]);

  let discount = 0;
  if (type === "PERCENT") discount = baseAmount * (value / 100);
  else discount = value;

  if (maxDiscount) discount = Math.min(discount, maxDiscount);
  discount = Math.max(0, Math.floor(discount));

  return ok_({
    discount,
    coupon: { code, type, value, scope, firstOnly }
  });
}

function hasPaidOrderForPhone_(phone) {
  const ph = normalizePhone_(phone);
  if (!ph) return false;

  const orders = sh_(TAB_ORDERS);
  const { idx } = headerIndex_(orders);
  const last = orders.getLastRow();
  if (last < 2) return false;

  const data = orders.getRange(2, 1, last - 1, orders.getLastColumn()).getValues();
  for (const r of data) {
    const rowPhone = normalizePhone_(String(r[idx["Phone"]] || ""));
    if (rowPhone !== ph) continue;
    const st = String(r[idx["Status"]] || "").trim().toUpperCase();
    if (st === ORDER_STATUS.PAID) return true;
  }
  return false;
}

function couponUsageCounts_(code, phone) {
  const usage = sh_(TAB_COUPON_USAGE);
  const { idx } = headerIndex_(usage);
  const last = usage.getLastRow();
  if (last < 2) return { globalCount: 0, userCount: 0 };

  const ph = normalizePhone_(phone);
  const data = usage.getRange(2, 1, last - 1, usage.getLastColumn()).getValues();
  let globalCount = 0, userCount = 0;

  for (const r of data) {
    if (String(r[idx["CouponCode"]] || "").trim().toUpperCase() !== code) continue;
    const action = String(r[idx["Action"]] || "").trim().toUpperCase();
    if (action !== "USED") continue; // ✅ only count paid usage
    globalCount++;
    if (normalizePhone_(String(r[idx["Phone"]] || "")) === ph) userCount++;
  }
  return { globalCount, userCount };
}

/**
 * COUPON_USAGE headers expected:
 * Timestamp | CouponCode | Phone | OrderID | Discount | Action
 */
function recordCouponUsage_(p) {
  try {
    const usage = sh_(TAB_COUPON_USAGE);
    const { idx } = headerIndex_(usage);

    // If headers exist, write in correct order by header names
    if (idx["Timestamp"] != null && idx["CouponCode"] != null) {
      const row = new Array(usage.getLastColumn()).fill("");
      row[idx["Timestamp"]]  = now_();
      row[idx["CouponCode"]] = String(p.couponCode || "").trim().toUpperCase();
      row[idx["Phone"]]      = normalizePhone_(String(p.phone || ""));
      row[idx["OrderID"]]    = String(p.orderId || "").trim();
      row[idx["Discount"]]   = asNumber_(p.discount || 0);
      row[idx["Action"]]     = String(p.action || "APPLIED").trim().toUpperCase();
      usage.appendRow(row);
      return;
    }

    // Fallback append (if sheet is simpler)
    usage.appendRow([
      now_(),
      String(p.couponCode || "").trim().toUpperCase(),
      normalizePhone_(String(p.phone || "")),
      String(p.orderId || "").trim(),
      asNumber_(p.discount || 0),
      String(p.action || "APPLIED").trim().toUpperCase()
    ]);
  } catch (_) {}
}

function normalizePhone_(p) {
  const digits = String(p || "").replace(/\D/g, "");
  if (!digits) return "";
  // keep last 10 for India numbers (handles +91)
  return digits.length > 10 ? digits.slice(-10) : digits;
}


// ===================== RESUME PAYMENT =====================

/**
 * Returns latest resumable order for this phone.
 * Resumable = Status=CREATED (unpaid) and not older than lookbackDays (default 3).
 */

function getLatestResumableOrder(phone, lookbackDays) {
  const ph = normalizePhone_(String(phone || "").trim());
  if (!ph) return bad_("Phone is required.");

  const days = asNumber_(lookbackDays || 3);
  const cutoff = now_().getTime() - days * 24 * 60 * 60 * 1000;

  const orders = sh_(TAB_ORDERS);
  const { idx } = headerIndex_(orders);
  const last = orders.getLastRow();
  if (last < 2) return ok_({ found: false });

  const data = orders.getRange(2, 1, last - 1, orders.getLastColumn()).getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];

    if (normalizePhone_(String(r[idx["Phone"]] || "")) !== ph) continue;

    const createdAt = r[idx["CreatedAt"]];
    const cMs = createdAt instanceof Date ? createdAt.getTime() : new Date(createdAt).getTime();
    if (!isNaN(cMs) && cMs < cutoff) continue;

    const st = String(r[idx["Status"]] || "").trim().toUpperCase();
    if (st !== ORDER_STATUS.CREATED) continue;

    return ok_({ found: true, order: buildPublicOrder_(r, idx) });
  }

  return ok_({ found: false });
}


function getCustomerOrders(phone, limit) {
  try {
    const ph = normalizePhone_(phone);
    const max = asNumber_(limit || 15);
    if (!ph) return bad_("Phone is required.");

    const orders = sh_(TAB_ORDERS);
    const { idx } = headerIndex_(orders);
    const last = orders.getLastRow();

    if (last < 2) return ok_({ orders: [] });

    const data = orders.getRange(2, 1, last - 1, orders.getLastColumn()).getValues();
    const out = [];

    for (let i = data.length - 1; i >= 0; i--) {
      const r = data[i];
      const rowPhone = normalizePhone_(String(r[idx["Phone"]] || ""));
      if (rowPhone !== ph) continue;

      out.push({
        orderId: String(r[idx["OrderID"]] || "").trim(),
        createdAt: r[idx["CreatedAt"]],
        service: String(r[idx["Service"]] || "").trim(),
        plan: String(r[idx["Plan"]] || "").trim(),
        status: String(r[idx["Status"]] || "").trim(),
        fulfillmentStatus: String(r[idx["FulfillmentStatus"]] || "").trim(),
        amount: asNumber_(r[idx["FinalAmount"]]),
        discount: asNumber_(r[idx["Discount"]]),
        currency: String(r[idx["Currency"]] || "INR").trim()
      });

      if (out.length >= max) break;
    }

    return ok_({ orders: JSON.parse(JSON.stringify(out)) });

  } catch(e) {
    return bad_("getCustomerOrders failed: " + e);
  }
}

/**
 * Resume by orderId + phone (prevents leaking someone else’s order).
 */
function getResumableOrderById(orderId, phone) {
  const oid = String(orderId || "").trim();
  const ph = normalizePhone_(String(phone || "").trim());
  if (!oid || !ph) return bad_("OrderID + phone required.");

  const orders = sh_(TAB_ORDERS);
  const { idx } = headerIndex_(orders);
  const rn = getRowByValue_(orders, "OrderID", oid);
  if (!rn) return ok_({ found: false });

  const row = orders.getRange(rn, 1, 1, orders.getLastColumn()).getValues()[0];
  if (normalizePhone_(String(row[idx["Phone"]] || "")) !== ph) return ok_({ found: false });

  const st = String(row[idx["Status"]] || "").trim().toUpperCase();
  if (st !== ORDER_STATUS.CREATED) return ok_({ found: false });

  return ok_({ found: true, order: buildPublicOrder_(row, idx) });
}

function buildPublicOrder_(row, idx) {
  const orderId = String(row[idx["OrderID"]] || "").trim();
  const amount = asNumber_(row[idx["FinalAmount"]]);
  const currency = String(row[idx["Currency"]] || getSetting_("CURRENCY", "INR")).trim();

  const service = String(row[idx["Service"]] || "").trim();
  const plan = String(row[idx["Plan"]] || "").trim();
  const phone = normalizePhone_(String(row[idx["Phone"]] || ""));
  const email = String(row[idx["Email"]] || "").trim();

  const upiVpa = String(getSetting_("UPI_VPA", "YOURUPI@BANK"));
  const payee = String(getSetting_("UPI_PAYEE_NAME", "FluxFilm"));
  const tn = encodeURIComponent(orderId);
  const upiLink = `upi://pay?pa=${encodeURIComponent(upiVpa)}&pn=${encodeURIComponent(payee)}&am=${encodeURIComponent(amount)}&cu=${encodeURIComponent(currency)}&tn=${tn}`;

  return {
    orderId, amount, currency,
    service, plan, phone, email,
    upiLink, paymentNote: orderId
  };
}

function orders_getMySubscriptions(phone) {
  // ✅ Delegate entirely to manage.gs implementation to avoid logic drift
  // manage.gs has the authoritative eligibility logic (LATE_RENEW/WARN_RENEW/TOO_LATE thresholds)
  if (typeof mg_getMySubscriptions_ === "function") {
    return mg_getMySubscriptions_(phone);
  }
  return bad_("getMySubscriptions not available.");
}

