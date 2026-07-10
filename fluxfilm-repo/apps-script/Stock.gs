/*************** FluxFilm Backend — Stock & Restock ***************
 *
 * Provides:
 *  - stock_getForPlan_(service, plan)   → number | null  (internal, reads PLANS Stock cell)
 *  - stock_levelLabel_(n)               → "OUT" | "LOW" | "OK"
 *  - submitRestockRequest(payload)      → records a "notify me" request
 *  - getStockLevels()                   → frontend call, returns all plan stock levels
 *  - updateOtpStockCounts()             → TIME-DRIVEN TRIGGER (every 5 min)
 *                                         Counts active subs per LoginId from INVENTORY_ACCOUNTS,
 *                                         reads MaxTotal from INVENTORY_CAPACITY,
 *                                         writes remaining slots into PLANS.Stock for each OTP plan.
 *
 * SHEET SETUP:
 *  PLANS sheet          → needs a "Stock" column (number, written by trigger)
 *  INVENTORY_ACCOUNTS   → Service | AccountID | LoginId | Password | IsActive | Plan
 *  INVENTORY_CAPACITY   → Service | AccountID | MaxTotal | MaxTV | IsActive | Notes
 *
 * HOW THE TRIGGER WORKS:
 *  1. Reads every active OTP plan row from PLANS.
 *  2. Finds the account for that plan in INVENTORY_ACCOUNTS (by Service + Plan).
 *  3. Gets that account's LoginId.
 *  4. Looks up MaxTotal for that AccountID in INVENTORY_CAPACITY.
 *  5. Counts all active non-expired subs in SUBSCRIPTIONS sharing that LoginId.
 *  6. Writes (MaxTotal - activeCount) into the Stock cell for that plan row.
 *
 * Because multiple plan rows can map to the same LoginId (e.g. 1 Month + 3 Month
 * both using phone 98765XXXXX), the trigger correctly reduces ALL those plan rows
 * together whenever any of them sells a slot.
 *
 * TO REGISTER THE TRIGGER (run once manually in GAS editor):
 *   registerOtpStockTrigger()
 *
 * "LOW" threshold default = 3, override via SETTINGS: STOCK_LOW_THRESHOLD = 3
 ***************************************************************************/

const STOCK_LOW_DEFAULT = 3;
const TAB_RESTOCK = "RESTOCK_REQUESTS";

// ─────────────────────────────────────────────────────────────────────────────
// TRIGGER — runs every 5 minutes
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Time-driven trigger. Register once with registerOtpStockTrigger().
 * Writes updated remaining stock into PLANS.Stock for every active OTP plan.
 */
function updateOtpStockCounts() {
  try {
    const plans    = sh_(TAB_PLANS);
    const { idx }  = headerIndex_(plans);

    // Stock column must exist in PLANS
    if (idx["Stock"] == null) {
      console.log("updateOtpStockCounts: PLANS sheet has no 'Stock' column — skipping.");
      return;
    }
    if (idx["Service"] == null || idx["Plan"] == null) return;

    const last = plans.getLastRow();
    if (last < 2) return;

    const rows = plans.getRange(2, 1, last - 1, plans.getLastColumn()).getValues();

    // Build a cache: LoginId → activeCount  so we don't re-scan SUBSCRIPTIONS
    // for every plan row that shares the same LoginId.
    const loginIdCountCache = {}; // "JioHotstar|98765XXXXX" → number

    rows.forEach((r, i) => {
      const rowNum  = i + 2;
      const service = String(r[idx["Service"]] || "").trim();
      const plan    = String(r[idx["Plan"]]    || "").trim();
      const active  = String(r[idx["IsActive"]] != null ? r[idx["IsActive"]] : "TRUE").trim().toUpperCase();

    if (active === "FALSE") return;
    if (!otp_isOtpService_(service)) return;
    if (!service || !plan) return;
    if (service.toLowerCase().indexOf("hotstar") !== -1) return;   // ← NEW: JioHotstar now handled by stock_jiohotstar.gs

      // Find which account handles this plan
      const acc = otp_pickAccount_(service, plan);
      if (!acc) {
        // No account configured — write 0 so it shows out of stock rather than
        // silently appearing available
        plans.getRange(rowNum, idx["Stock"] + 1).setValue(0);
        return;
      }

      const loginId  = acc.loginId;
      const cacheKey = service.toLowerCase() + "|" + loginId;

      // Count active subs for this LoginId (cached per LoginId)
      if (loginIdCountCache[cacheKey] == null) {
        loginIdCountCache[cacheKey] = otp_countActiveByLoginId_(service, loginId);
      }
      const activeCount = loginIdCountCache[cacheKey];

      // Get MaxTotal from INVENTORY_CAPACITY for this AccountID
      const maxTotal = otp_getMaxTotal_(service, acc.accountId);

      const remaining = Math.max(0, maxTotal - activeCount);
      plans.getRange(rowNum, idx["Stock"] + 1).setValue(remaining);
    });

    console.log("updateOtpStockCounts: completed at " + new Date().toISOString());
  } catch (e) {
    console.log("updateOtpStockCounts ERROR: " + (e && e.message ? e.message : e));
  }
}

/**
 * Read MaxTotal for an AccountID from INVENTORY_CAPACITY.
 * Falls back to SETTINGS: OTP_DEFAULT_MAX_TOTAL (default 10) if not found.
 */
function otp_getMaxTotal_(service, accountId) {
  try {
    const cap     = sh_(TAB_INV_CAPACITY);
    const { idx } = headerIndex_(cap);
    const last    = cap.getLastRow();
    if (last < 2 || idx["AccountID"] == null || idx["MaxTotal"] == null) {
      return asNumber_(getSetting_("OTP_DEFAULT_MAX_TOTAL", 10)) || 10;
    }

    const data = cap.getRange(2, 1, last - 1, cap.getLastColumn()).getValues();
    for (const r of data) {
      const rowSvc  = String(r[idx["Service"]]   || "").trim().toLowerCase();
      const rowAccId = String(r[idx["AccountID"]] || "").trim();
      const isActive = idx["IsActive"] != null
        ? String(r[idx["IsActive"]] || "").trim().toUpperCase()
        : "TRUE";

      if (isActive === "FALSE") continue;
      if (!otp_serviceMatch_(rowSvc, service.toLowerCase())) continue;
      if (rowAccId !== accountId) continue;

      const max = asNumber_(r[idx["MaxTotal"]]);
      return max > 0 ? max : (asNumber_(getSetting_("OTP_DEFAULT_MAX_TOTAL", 10)) || 10);
    }
  } catch (e) {}

  return asNumber_(getSetting_("OTP_DEFAULT_MAX_TOTAL", 10)) || 10;
}

/**
 * Run this once manually in the GAS editor to register the 5-min trigger.
 * Safe to run again — removes old duplicates first.
 */
function registerOtpStockTrigger() {
  // Remove existing triggers for this function to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "updateOtpStockCounts") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("updateOtpStockCounts")
    .timeBased()
    .everyMinutes(5)
    .create();
  console.log("Trigger registered: updateOtpStockCounts every 5 minutes.");
}

// ─────────────────────────────────────────────────────────────────────────────
// INTERNAL HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Returns the current Stock number for a plan (reads directly from PLANS cell).
 * Used internally and by getStockLevels().
 * Returns null if Stock column missing or cell blank.
 */
function stock_getForPlan_(service, plan) {
  service = String(service || "").trim();
  plan    = String(plan    || "").trim();
  if (!service || !plan) return null;

  try {
    const plans   = sh_(TAB_PLANS);
    const { idx } = headerIndex_(plans);
    if (idx["Stock"] == null) return null;

    const last = plans.getLastRow();
    if (last < 2) return null;

    const data = plans.getRange(2, 1, last - 1, plans.getLastColumn()).getValues();
    for (const r of data) {
      if (String(r[idx["Service"]] || "").trim() === service &&
          String(r[idx["Plan"]]    || "").trim() === plan) {
        const raw = r[idx["Stock"]];
        const s   = String(raw == null ? "" : raw).trim();
        if (s !== "") return Math.max(0, Math.floor(asNumber_(raw)));
        break;
      }
    }
  } catch (e) {}

  return null;
}

/**
 * Convert a stock number to a UI level label.
 *  null → "OK"   (unknown / engine-managed)
 *  0    → "OUT"
 *  <LOW → "LOW"
 *  else → "OK"
 */
function stock_levelLabel_(n) {
  if (n == null) return "OK";
  const low = asNumber_(getSetting_("STOCK_LOW_THRESHOLD", STOCK_LOW_DEFAULT)) || STOCK_LOW_DEFAULT;
  if (n <= 0)  return "OUT";
  if (n < low) return "LOW";
  return "OK";
}

// ─────────────────────────────────────────────────────────────────────────────
// FRONTEND CALLS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Returns stock levels for all active plans in one call.
 * Called by the frontend separately from getBootstrap (non-blocking).
 * Returns: { ok:true, levels:{ "Service|||Plan": { stock:N, stockLevel:"OUT"|"LOW"|"OK" } } }
 */
function getStockLevels() {
  try {
    const plans   = sh_(TAB_PLANS);
    const { idx } = headerIndex_(plans);
    const last    = plans.getLastRow();
    if (last < 2) return ok_({ levels: {} });

    const rows   = plans.getRange(2, 1, last - 1, plans.getLastColumn()).getValues();
    const levels = {};

    rows.forEach(r => {
      const active = String(r[idx["IsActive"]] || "").trim().toUpperCase() === "TRUE";
      if (!active) return;
      const service = String(r[idx["Service"]] || "").trim();
      const plan    = String(r[idx["Plan"]]    || "").trim();
      if (!service || !plan) return;

      let stockNum = null;
      if (idx["Stock"] != null) {
        const raw = r[idx["Stock"]];
        const s   = String(raw == null ? "" : raw).trim();
        if (s !== "") stockNum = Math.max(0, Math.floor(asNumber_(raw)));
      }

      const stockLevel = stock_levelLabel_(stockNum);
      levels[service + "|||" + plan] = { stock: stockNum, stockLevel };
    });

    return ok_({ levels });
  } catch (e) {
    return ok_({ levels: {} }); // never fail — stock is non-critical
  }
}

/**
 * Frontend: "Notify me when restocked".
 * payload: { name, phone, service, plan }
 */
function submitRestockRequest(payload) {
  try {
    const p       = payload || {};
    const name    = String(p.name    || "").trim();
    const phone   = normalizePhone_(p.phone);
    const service = String(p.service || "").trim();
    const plan    = String(p.plan    || "").trim();

    if (!phone)           return bad_("Phone number is required.");
    if (!service || !plan) return bad_("Service and plan are required.");

    const sh      = stock_restockSheet_();
    const { idx } = headerIndex_(sh);

    // Avoid duplicate pending requests for same phone+service+plan
    try {
      const last = sh.getLastRow();
      if (last >= 2 && idx["Phone"] != null && idx["Service"] != null && idx["Plan"] != null) {
        const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
        for (const r of data) {
          const rp  = normalizePhone_(r[idx["Phone"]]);
          const rs  = String(r[idx["Service"]] || "").trim();
          const rpl = String(r[idx["Plan"]]    || "").trim();
          const st  = idx["Status"] != null ? String(r[idx["Status"]] || "").trim().toUpperCase() : "";
          if (rp === phone && rs === service && rpl === plan && st !== "DONE") {
            return ok_({ message: "You're already on the notify list for this plan. We'll message you when it's back.", duplicate: true });
          }
        }
      }
    } catch (e) {}

    const row = new Array(sh.getLastColumn()).fill("");
    function put(h, v) { if (idx[h] != null) row[idx[h]] = v; }
    put("Timestamp", now_());
    put("Name",    name);
    put("Phone",   phone);
    put("Service", service);
    put("Plan",    plan);
    put("Status",  "PENDING");
    put("Notes",   "");
    sh.appendRow(row);

    if (typeof notifyTelegram_ === "function") {
      notifyTelegram_("RESTOCK_REQUEST", { name, phone, service, plan, note: "Customer wants restock alert" });
    }

    return ok_({ message: "You're on the list! We'll message you when this plan is back in stock." });
  } catch (e) {
    return bad_("submitRestockRequest error: " + (e && e.message ? e.message : e));
  }
}

function stock_restockSheet_() {
  const ss = getCore_();
  let sh   = ss.getSheetByName(TAB_RESTOCK);
  if (!sh) {
    sh = ss.insertSheet(TAB_RESTOCK);
    sh.appendRow(["Timestamp", "Name", "Phone", "Service", "Plan", "Status", "Notes"]);
    sh.setFrozenRows(1);
  }
  const headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(h => String(h || "").trim());
  if (!headers.some(Boolean)) {
    sh.getRange(1, 1, 1, 7).setValues([["Timestamp", "Name", "Phone", "Service", "Plan", "Status", "Notes"]]);
  }
  return sh;
}









/*************** FluxFilm — YouTube Premium Stock Sync (stock_youtube.gs) ***************
 *
 * NEW FILE. Safe to add — does not touch any existing feature.
 *
 * WHAT IT DOES:
 *  Your DASHBOARD sheet already has a little table:
 *
 *      YouTube Slots Occupied | 26 | 27
 *      YouTube Slots Empty    |  1 | 🟢 1 Slot Empty
 *
 *  The "YouTube Slots Empty" number IS your available stock.
 *  This trigger copies that number into the PLANS "Stock" column for
 *  every active YouTube Premium plan (1M / 3M / 6M / 1Year) — so all
 *  YT plans show OUT / LOW / OK correctly on the website, just like OTP.
 *
 * HOW IT FINDS THE NUMBER (no hardcoded cell — safe if you move the table):
 *  1. Scans the DASHBOARD sheet for the text "YouTube Slots Empty".
 *  2. Takes the number in the cell immediately to its right.
 *  3. Writes that number into PLANS.Stock for each YouTube Premium row.
 *
 * SETUP:
 *  - PLANS sheet must have a "Stock" column (you already added it for OTP).
 *  - Nothing else. The DASHBOARD table already exists.
 *
 * REGISTER THE TRIGGER (run once manually in the GAS editor):
 *      registerYoutubeStockTrigger()
 *
 * You can also just run  updateYoutubeStock()  once by hand to test it.
 ***************************************************************************/

const YT_DASHBOARD_TAB   = "DASHBOARD";
const YT_EMPTY_LABEL      = "YouTube Slots Empty"; // the label we search for
const YT_SERVICE_NAME     = "YouTube Premium";      // must match PLANS "Service"

/**
 * Main job. Reads empty-slot count from DASHBOARD and writes it into
 * PLANS.Stock for every active YouTube Premium plan row.
 * Register with registerYoutubeStockTrigger() to run automatically.
 */
function updateYoutubeStock() {
  try {
    const empty = yt_readEmptySlots_();
    if (empty == null) {
      console.log("updateYoutubeStock: couldn't find '" + YT_EMPTY_LABEL + "' on DASHBOARD — skipping.");
      return;
    }

    const plans   = sh_(TAB_PLANS);
    const { idx } = headerIndex_(plans);

    if (idx["Stock"] == null) {
      console.log("updateYoutubeStock: PLANS has no 'Stock' column — skipping.");
      return;
    }
    if (idx["Service"] == null) {
      console.log("updateYoutubeStock: PLANS has no 'Service' column — skipping.");
      return;
    }

    const last = plans.getLastRow();
    if (last < 2) return;

    const rows      = plans.getRange(2, 1, last - 1, plans.getLastColumn()).getValues();
    let   updated   = 0;

    rows.forEach((r, i) => {
      const rowNum  = i + 2;
      const service = String(r[idx["Service"]] || "").trim();

      // Only YouTube Premium rows
      if (service.toLowerCase() !== YT_SERVICE_NAME.toLowerCase()) return;

      // Optional: skip inactive plans if IsActive column exists
      if (idx["IsActive"] != null) {
        const active = String(r[idx["IsActive"]] || "TRUE").trim().toUpperCase();
        if (active === "FALSE") return;
      }

      plans.getRange(rowNum, idx["Stock"] + 1).setValue(empty);
      updated++;
    });

    console.log("updateYoutubeStock: wrote Stock=" + empty + " to " + updated +
                " YouTube Premium plan(s) at " + new Date().toISOString());
  } catch (e) {
    console.log("updateYoutubeStock ERROR: " + (e && e.message ? e.message : e));
  }
}

/**
 * Finds the "YouTube Slots Empty" label on the DASHBOARD sheet and returns
 * the number in the cell right next to it. Returns null if not found.
 */
function yt_readEmptySlots_() {
  try {
    const ss = getCore_();
    const sh = ss.getSheetByName(YT_DASHBOARD_TAB);
    if (!sh) return null;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return null;

    const data = sh.getRange(1, 1, lastRow, lastCol).getValues();

    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const cell = String(data[r][c] == null ? "" : data[r][c]).trim();
        if (cell.toLowerCase() === YT_EMPTY_LABEL.toLowerCase()) {
          // number is in the cell immediately to the right
          const raw = (c + 1 < data[r].length) ? data[r][c + 1] : null;
          const n   = asNumber_(raw);
          return Math.max(0, Math.floor(n)); // never negative, whole number
        }
      }
    }
  } catch (e) {}
  return null;
}

/**
 * Run this ONCE manually in the GAS editor to schedule the sync.
 * Safe to re-run — it removes old duplicates first.
 * Runs every 5 minutes (same rhythm as your OTP stock trigger).
 */
function registerYoutubeStockTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "updateYoutubeStock") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("updateYoutubeStock")
    .timeBased()
    .everyMinutes(5)
    .create();
  console.log("Trigger registered: updateYoutubeStock every 5 minutes.");
}