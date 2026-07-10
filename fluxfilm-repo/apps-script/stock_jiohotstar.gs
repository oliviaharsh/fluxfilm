/*************** FluxFilm — JioHotstar Stock Sync (stock_jiohotstar.gs) ***************
 *
 * NEW FILE. Handles JioHotstar stock ONLY.
 * Safe to add — does not touch YouTube or general OTP stock logic.
 *
 * ─── WHY THIS EXISTS ───────────────────────────────────────────────────────
 * The general OTP stock trigger caps by AccountID (via INVENTORY_CAPACITY).
 * But for JioHotstar, the REAL capacity limit lives on the PHONE NUMBER —
 * one phone can serve multiple plans (3M + 6M + 1Y all on 8076332049),
 * and the phone itself has a total login cap (e.g. 9).
 *
 * So we cap per phone, then split the remaining pool across every plan
 * that uses that phone.
 *
 * ─── HOW IT WORKS ──────────────────────────────────────────────────────────
 *  1. For each active JioHotstar row in INVENTORY_ACCOUNTS, collect its
 *     LoginId (phone) and its Plan.
 *  2. For each unique phone, count active subs (Status = "ACTIVE") from
 *     SUBSCRIPTIONS — we ignore ExpiryDate on purpose because the customer's
 *     device stays logged in until we officially flip Status.
 *  3. For each phone, read its max capacity from SETTINGS:
 *        Key = "JH_MAXCAP_<phone>"       (e.g. JH_MAXCAP_8076332049 = 9)
 *        Fallback = "JH_DEFAULT_MAX"     (e.g. 9)
 *  4. remaining[phone] = max(0, cap - active)
 *  5. For each JioHotstar plan in PLANS:
 *        Stock = sum of remaining[phone] for every phone that serves this plan
 *  6. Write Stock into PLANS.
 *
 * ─── EXAMPLE (matches the user's data) ─────────────────────────────────────
 *   Phone 7428576079 → cap 9, active 9 → remaining 0
 *     Serves: 1 Month  → Stock(1 Month)  = 0
 *
 *   Phone 8076332049 → cap 9, active 9 → remaining 0
 *     Serves: 3 Months → Stock(3 Months) = 0
 *              6 Months → Stock(6 Months) = 0
 *              1 Year   → contributes 0 to Stock(1 Year)
 *
 *   Phone 9818196079 → cap 9, active 7 → remaining 2
 *     Serves: 1 Year   → contributes 2 to Stock(1 Year)
 *
 *   Final Stock(1 Year) = 0 + 2 = 2  ✅
 *
 * ─── SETUP ─────────────────────────────────────────────────────────────────
 *   SETTINGS sheet — add rows:
 *      JH_MAXCAP_7428576079  = 9
 *      JH_MAXCAP_8076332049  = 9
 *      JH_MAXCAP_9818196079  = 9
 *      JH_DEFAULT_MAX        = 9     (fallback for any phone without a row)
 *
 * ─── REGISTER TRIGGER (run once in GAS editor) ─────────────────────────────
 *   registerJiohotstarStockTrigger()
 *
 *   Or run  updateJiohotstarStock()  manually anytime to test.
 ***************************************************************************/

const JH_SERVICE_NAME     = "JioHotstar";
const JH_MAXCAP_PREFIX    = "JH_MAXCAP_";      // + phone → per-phone cap key
const JH_DEFAULT_MAX_KEY  = "JH_DEFAULT_MAX";  // fallback cap
const JH_DEFAULT_MAX_VAL  = 9;                 // hard fallback if setting missing
// NOTE: TAB_INV_ACCOUNTS is already declared in code.gs (Config) — reusing it here, not redeclaring.

// ─────────────────────────────────────────────────────────────────────────────
// MAIN JOB — runs every 5 minutes via trigger
// ─────────────────────────────────────────────────────────────────────────────

function updateJiohotstarStock() {
  try {
    const plans        = sh_(TAB_PLANS);
    const { idx: pIdx } = headerIndex_(plans);

    if (pIdx["Stock"] == null || pIdx["Service"] == null || pIdx["Plan"] == null) {
      console.log("updateJiohotstarStock: PLANS missing Stock/Service/Plan column — skipping.");
      return;
    }

    // ── Step 1: Read INVENTORY_ACCOUNTS, build phone→plans map ────────────
    const inv         = sh_(TAB_INV_ACCOUNTS);
    const { idx: iIdx } = headerIndex_(inv);

    if (iIdx["Service"] == null || iIdx["LoginId"] == null || iIdx["Plan"] == null) {
      console.log("updateJiohotstarStock: INVENTORY_ACCOUNTS missing Service/LoginId/Plan — skipping.");
      return;
    }

    const invLast = inv.getLastRow();
    if (invLast < 2) return;

    const invRows = inv.getRange(2, 1, invLast - 1, inv.getLastColumn()).getValues();

    // phone → Set of plans it serves        e.g. "8076332049" → {"3 Months","6 Months","1 Year"}
    // plan  → Set of phones that serve it   e.g. "1 Year"     → {"8076332049","9818196079"}
    const phoneToPlans = {};
    const planToPhones = {};

    invRows.forEach(r => {
      const svc = String(r[iIdx["Service"]] || "").trim();
      if (svc.toLowerCase() !== JH_SERVICE_NAME.toLowerCase()) return;

      // Respect IsActive if column exists
      if (iIdx["IsActive"] != null) {
        const active = String(r[iIdx["IsActive"]] || "TRUE").trim().toUpperCase();
        if (active === "FALSE") return;
      }

      const phone = String(r[iIdx["LoginId"]] || "").trim();
      const plan  = String(r[iIdx["Plan"]]    || "").trim();
      if (!phone || !plan) return;

      if (!phoneToPlans[phone]) phoneToPlans[phone] = {};
      phoneToPlans[phone][plan] = true;

      if (!planToPhones[plan]) planToPhones[plan] = {};
      planToPhones[plan][phone] = true;
    });

    // ── Step 2: For each unique phone, count active + read cap ────────────
    const phoneRemaining = {};  // phone → remaining slots (max - active)

    Object.keys(phoneToPlans).forEach(phone => {
      const active    = jh_countActiveByPhone_(phone);
      const cap       = jh_readPhoneCap_(phone);
      const remaining = Math.max(0, cap - active);
      phoneRemaining[phone] = remaining;

      console.log("JH phone " + phone + " → cap=" + cap + ", active=" + active + ", remaining=" + remaining);
    });

    // ── Step 3: For each plan, sum remaining across its phones ────────────
    const planStock = {};
    Object.keys(planToPhones).forEach(plan => {
      let total = 0;
      Object.keys(planToPhones[plan]).forEach(phone => {
        total += (phoneRemaining[phone] || 0);
      });
      planStock[plan] = total;
    });

    // ── Step 4: Write into PLANS.Stock for every JioHotstar row ───────────
    const plansLast = plans.getLastRow();
    if (plansLast < 2) return;

    const plansRows = plans.getRange(2, 1, plansLast - 1, plans.getLastColumn()).getValues();
    let updated = 0;

    plansRows.forEach((r, i) => {
      const rowNum = i + 2;
      const svc    = String(r[pIdx["Service"]] || "").trim();
      const plan   = String(r[pIdx["Plan"]]    || "").trim();

      if (svc.toLowerCase() !== JH_SERVICE_NAME.toLowerCase()) return;
      if (!plan) return;

      if (pIdx["IsActive"] != null) {
        const active = String(r[pIdx["IsActive"]] || "TRUE").trim().toUpperCase();
        if (active === "FALSE") return;
      }

      // If plan has no inventory row at all, write 0 (out of stock)
      const stock = planStock[plan] != null ? planStock[plan] : 0;
      plans.getRange(rowNum, pIdx["Stock"] + 1).setValue(stock);
      updated++;
    });

    console.log("updateJiohotstarStock: wrote stock to " + updated + " plan(s) at " + new Date().toISOString());
  } catch (e) {
    console.log("updateJiohotstarStock ERROR: " + (e && e.message ? e.message : e));
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Count active JioHotstar subs sharing this phone number (LoginId).
 * IMPORTANT: uses Status = "ACTIVE" only. We do NOT filter on ExpiryDate
 * because the customer's device stays logged in until Status flips.
 */
function jh_countActiveByPhone_(phone) {
  try {
    const subs        = sh_(TAB_SUBS);
    const { idx }     = headerIndex_(subs);
    const last        = subs.getLastRow();
    if (last < 2) return 0;
    if (idx["Service"] == null || idx["LoginId"] == null || idx["Status"] == null) return 0;

    const data  = subs.getRange(2, 1, last - 1, subs.getLastColumn()).getValues();
    const phoneStr = String(phone).trim();
    let count   = 0;

    data.forEach(r => {
      const svc    = String(r[idx["Service"]] || "").trim().toLowerCase();
      const phn    = String(r[idx["LoginId"]] || "").trim();
      const status = String(r[idx["Status"]]  || "").trim().toUpperCase();

      if (svc === JH_SERVICE_NAME.toLowerCase() && phn === phoneStr && status === "ACTIVE") {
        count++;
      }
    });

    return count;
  } catch (e) {
    return 0;
  }
}

/**
 * Read the max login capacity for a phone number from SETTINGS.
 *   Primary key:  JH_MAXCAP_<phone>   (e.g. JH_MAXCAP_8076332049)
 *   Fallback key: JH_DEFAULT_MAX
 *   Hard fallback: JH_DEFAULT_MAX_VAL (9)
 */
function jh_readPhoneCap_(phone) {
  const perPhoneKey = JH_MAXCAP_PREFIX + String(phone).trim();
  const perPhoneVal = asNumber_(getSetting_(perPhoneKey, 0));
  if (perPhoneVal > 0) return perPhoneVal;

  const defVal = asNumber_(getSetting_(JH_DEFAULT_MAX_KEY, JH_DEFAULT_MAX_VAL));
  return defVal > 0 ? defVal : JH_DEFAULT_MAX_VAL;
}

// ─────────────────────────────────────────────────────────────────────────────
// TRIGGER REGISTRATION
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Run ONCE manually in the GAS editor to schedule the sync.
 * Safe to re-run — removes duplicates first.
 */
function registerJiohotstarStockTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "updateJiohotstarStock") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("updateJiohotstarStock")
    .timeBased()
    .everyMinutes(5)
    .create();
  console.log("Trigger registered: updateJiohotstarStock every 5 minutes.");
}
