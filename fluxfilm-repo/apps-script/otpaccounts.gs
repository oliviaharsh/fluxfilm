/*************** FluxFilm Backend — OTP-Service Account Allocation ***************
 *
 * Handles INSTANT delivery for OTP login services: JioHotstar, Zee5, SonyLIV
 *
 * HOW IT WORKS (updated):
 *  1. Read Stock column from PLANS sheet for this service+plan.
 *     Stock is auto-updated every 5 min by updateOtpStockCounts() trigger in stock.gs.
 *     Stock = remaining slots = (MaxTotal on that account) - (active subs using that account's LoginId)
 *  2. If Stock > 0  → find the matching account in INVENTORY_ACCOUNTS → deliver.
 *  3. If Stock <= 0 → out of stock. Simple.
 *
 * WHY THIS WORKS ACROSS PLANS:
 *  - The trigger counts active subs by LoginId (phone number), not by plan.
 *  - So 1 Month and 3 Month sharing the same phone number both draw from the same pool.
 *  - When that phone hits capacity, both plan rows get Stock = 0 simultaneously.
 *
 * INVENTORY_ACCOUNTS columns used:
 *   AccountID | Service | LoginId | Password | IsActive | Plan
 *   Plan column must match exactly (e.g. "1 Month", "1 Year").
 *
 * INVENTORY_CAPACITY columns used (for MaxTotal per account):
 *   Service | AccountID | MaxTotal | MaxTV | IsActive | Notes
 *
 **********************************************************************************/

// Canonical OTP service keys
const OTP_ACCOUNT_SERVICES = ["JioHotstar", "Zee5", "SonyLIV"];

/**
 * Is this service an OTP-login service?
 */
function otp_isOtpService_(service) {
  const s = String(service || "").toLowerCase();
  if (!s) return false;
  if (s.indexOf("jiohotstar") !== -1) return true;
  if (s.indexOf("hotstar") !== -1) return true;
  if (s.indexOf("zee5") !== -1 || s.indexOf("zee 5") !== -1) return true;
  if (s.indexOf("sonyliv") !== -1 || s.indexOf("sony liv") !== -1) return true;
  return false;
}

/**
 * Main allocation function — simple Stock-based gate.
 *
 * Stock in PLANS is kept up to date by the 5-min trigger.
 * If Stock > 0 → allocate. If Stock <= 0 → out of stock.
 * No active-sub counting here at all — trigger handles that separately.
 */
function allocateOtpAccount_(service, plan, durationDays) {

  // ── Step 1: Read Stock from PLANS ─────────────────────────────────────────
  const stock = otp_getStockFromPlans_(service, plan);

  // Stock column missing or blank → treat as in-stock, try to pick an account
  if (stock === null) {
    const acc = otp_pickAccount_(service, plan);
    if (!acc) {
      return bad_("No accounts available for this plan. Please contact support.");
    }
    return ok_({
      inventoryRef: acc.accountId,
      access: { user: acc.loginId, pass: acc.pass }
    });
  }

  // Stock = 0 → out of stock
  if (stock <= 0) {
    return bad_("This plan is currently out of stock. Please contact support.");
  }

  // Stock > 0 → slots available, find the account for this plan
  const acc = otp_pickAccount_(service, plan);
  if (!acc) {
    return bad_("No accounts configured for this plan. Please contact support.");
  }

  return ok_({
    inventoryRef: acc.accountId,
    access: { user: acc.loginId, pass: acc.pass }
  });
}

/**
 * Read the Stock value from PLANS for a service+plan.
 * Returns null if Stock column missing or cell is blank.
 * Returns the number (could be 0) if present.
 */
function otp_getStockFromPlans_(service, plan) {
  try {
    const plans = sh_(TAB_PLANS);
    const { idx } = headerIndex_(plans);
    if (idx["Stock"] == null) return null;

    const last = plans.getLastRow();
    if (last < 2) return null;

    const data = plans.getRange(2, 1, last - 1, plans.getLastColumn()).getValues();
    for (const r of data) {
      const s = String(r[idx["Service"]] || "").trim();
      const p = String(r[idx["Plan"]]    || "").trim();
      if (s === service && p === plan) {
        const raw = r[idx["Stock"]];
        const str = String(raw == null ? "" : raw).trim();
        if (str === "") return null; // blank = no limit set
        return Math.max(0, Math.floor(Number(str) || 0));
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Count active non-expired subscriptions for a given LoginId across ALL plans.
 * Used by the stock trigger in stock.gs — NOT used during allocation.
 *
 * Matches InventoryRef in SUBSCRIPTIONS against AccountIDs that share the same LoginId
 * in INVENTORY_ACCOUNTS.
 */
function otp_countActiveByLoginId_(service, loginId) {
  try {
    // First, get all AccountIDs that have this LoginId for this service
    const invAcc = sh_(TAB_INV_ACCOUNTS);
    const { idx: aIdx } = headerIndex_(invAcc);
    const aLast = invAcc.getLastRow();
    if (aLast < 2) return 0;

    const aData = invAcc.getRange(2, 1, aLast - 1, invAcc.getLastColumn()).getValues();
    const matchingAccountIds = new Set();

    for (const r of aData) {
      const svc = String(r[aIdx["Service"]] || "").trim();
      if (!otp_serviceMatch_(svc.toLowerCase(), service.toLowerCase())) continue;
      const lid = String(r[aIdx["LoginId"]] || "").trim();
      if (lid !== loginId) continue;
      const accId = String(r[aIdx["AccountID"]] || "").trim();
      if (accId) matchingAccountIds.add(accId);
    }

    if (matchingAccountIds.size === 0) return 0;

    // Now count active non-expired subs in SUBSCRIPTIONS whose InventoryRef is one of those AccountIDs
    const subs = sh_(TAB_SUBS);
    const { idx: sIdx } = headerIndex_(subs);
    const sLast = subs.getLastRow();
    if (sLast < 2) return 0;

    const sData = subs.getRange(2, 1, sLast - 1, subs.getLastColumn()).getValues();
    const nowMs = now_().getTime();
    let count = 0;

    for (const r of sData) {
      // Service match
      const rowSvc = String(r[sIdx["Service"]] || "").trim();
      if (!otp_serviceMatch_(rowSvc.toLowerCase(), service.toLowerCase())) continue;

      // InventoryRef must be one of the AccountIDs sharing this LoginId
      const ref = String(r[sIdx["InventoryRef"]] || "").trim();
      if (!matchingAccountIds.has(ref)) continue;

      // Must be ACTIVE
      const status = String(r[sIdx["Status"]] || "").trim().toUpperCase();
      if (status !== "ACTIVE") continue;

      // Must not be expired
      if (sIdx["ExpiryDate"] == null) continue;
      const exp = r[sIdx["ExpiryDate"]];
      const expMs = exp instanceof Date ? exp.getTime() : new Date(exp).getTime();
      if (isNaN(expMs) || nowMs >= expMs) continue;

      count++;
    }
    return count;
  } catch (e) {
    return 0;
  }
}

/**
 * Used by stock.gs fallback (otp_countFreeAccounts_) — kept for compatibility.
 * Returns remaining slots for a plan based on current Stock value in PLANS.
 */
function otp_countFreeAccounts_(service, plan) {
  try {
    const stock = otp_getStockFromPlans_(service, plan);
    if (stock !== null) return Math.max(0, stock);
    // No stock column — return 1 if any account exists for this plan, else 0
    const acc = otp_pickAccount_(service, plan);
    return acc ? 1 : 0;
  } catch (e) {
    return null;
  }
}

/**
 * Find the first active account in INVENTORY_ACCOUNTS for this service+plan.
 * Returns { accountId, loginId, pass } or null if none found.
 */
function otp_pickAccount_(service, plan) {
  try {
    const invAcc = sh_(TAB_INV_ACCOUNTS);
    const { idx } = headerIndex_(invAcc);
    const last = invAcc.getLastRow();
    if (last < 2) return null;

    const svcLower = String(service || "").toLowerCase();
    const planWanted = String(plan || "").trim();
    const hasPlanCol = idx["Plan"] != null;

    const data = invAcc.getRange(2, 1, last - 1, invAcc.getLastColumn()).getValues();

    for (const r of data) {
      // Service match
      const s = String(r[idx["Service"]] || "").trim().toLowerCase();
      if (!otp_serviceMatch_(s, svcLower)) continue;

      // Must be active
      const active = String(r[idx["IsActive"]] || "").trim().toUpperCase() === "TRUE";
      if (!active) continue;

      // Plan match — if Plan column exists, must match exactly
      if (hasPlanCol) {
        const accPlan = String(r[idx["Plan"]] || "").trim();
        if (accPlan !== planWanted) continue;
      }

      const accountId = String(r[idx["AccountID"]] || "").trim();
      const loginId   = String(r[idx["LoginId"]]   || "").trim();
      const pass      = String(r[idx["Password"]]  || "").trim();
      if (!accountId || !loginId) continue;

      return { accountId, loginId, pass };
    }
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Loose service name matching for OTP services.
 * "zee5 premium" → matches "zee5"
 * "sony liv"     → matches "sonyliv"
 */
function otp_serviceMatch_(a, b) {
  function canon(x) {
    x = String(x || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    if (x.indexOf("jiohotstar") !== -1 || x.indexOf("hotstar") !== -1) return "jiohotstar";
    if (x.indexOf("zee5") !== -1) return "zee5";
    if (x.indexOf("sonyliv") !== -1) return "sonyliv";
    return x;
  }
  return canon(a) === canon(b);
}