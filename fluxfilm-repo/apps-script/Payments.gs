/*************** FluxFilm Backend — Payments ***************/

/**
 * ARCHITECTURE (v2 — fast UI + safe allocation):
 *
 * verifyPayment(orderId)
 *   → Does the bank scan + marks PAID
 *   → Does NOT call fulfillOrder directly anymore
 *   → Returns { found:true, paid:true } as soon as payment is confirmed
 *   → Frontend shows "Payment Verified!" immediately
 *
 * fulfillAndGetAccess(orderId)
 *   → Called by frontend AFTER showing the celebration screen
 *   → Acquires lock, checks/sets PROCESSING, calls fulfillOrder
 *   → Returns credentials when done
 *   → Frontend polls this every 3s until it gets FULFILLED or MANUAL_PENDING
 *
 * This separation means:
 *   - UI is fast (payment confirmed in ~5s with no allocation delay)
 *   - Only ONE fulfillment call ever happens per order (no parallel racing)
 *   - Double-allocation is impossible because fulfillAndGetAccess is the
 *     only entry point to fulfillOrder, and it has a hard lock + PROCESSING guard
 */


/**
 * Step 1: Verify payment only. Fast. No allocation.
 * Returns { found:true, paid:true } as soon as bank match is found.
 * Frontend shows "Payment Verified!" and then calls fulfillAndGetAccess.
 */
function verifyPayment(orderId) {
  let orders, rowNum, idx, status, fulfillStatus;

  try {
    orders = sh_(TAB_ORDERS);
    rowNum = getRowByValue_(orders, "OrderID", orderId);
    if (!rowNum) return bad_("Order not found.");

    ({ idx } = headerIndex_(orders));

    const mustOrders = ["Status", "FulfillmentStatus", "FinalAmount", "Service", "Plan", "Phone", "CouponCode", "Discount"];
    for (const h of mustOrders) {
      if (idx[h] == null) return bad_(`ORDERS sheet missing header: ${h}`);
    }

    status        = String(orders.getRange(rowNum, idx["Status"]            + 1).getValue() || "").trim().toUpperCase();
    fulfillStatus = String(orders.getRange(rowNum, idx["FulfillmentStatus"] + 1).getValue() || "").trim().toUpperCase();
  } catch (e) {
    return bad_("Order lookup failed: " + e);
  }

  // Already fully done — return credentials directly
  if (fulfillStatus === FULFILL_STATUS.FULFILLED) {
    return _buildFulfilledResponse_(orderId, orders, rowNum, idx);
  }

  // Already processing or manual — pass through
  if (fulfillStatus === "PROCESSING") {
    return ok_({ found: true, paid: true, fulfillment: "PROCESSING", message: "Processing your order…" });
  }
  if (fulfillStatus === FULFILL_STATUS.MANUAL_PENDING) {
    return ok_({ found: true, paid: true, fulfillment: "MANUAL_PENDING", message: "✅ Payment received. Activation will be done manually.", orderId });
  }

  // Already marked PAID — payment confirmed, skip bank scan
  if (status === ORDER_STATUS.PAID) {
    return ok_({ found: true, paid: true, fulfillment: fulfillStatus || "PENDING", message: "✅ Payment confirmed." });
  }

  // ---- Bank Scan ----
  const finalAmount = asNumber_(orders.getRange(rowNum, idx["FinalAmount"] + 1).getValue());
  const service     = String(orders.getRange(rowNum, idx["Service"] + 1).getValue() || "").trim();
  const plan        = String(orders.getRange(rowNum, idx["Plan"]    + 1).getValue() || "").trim();
  const oid         = String(orderId || "").trim();
  const oidNoDash   = normalizeOrderId_(oid);

  let match = null;

  try {
    const bank = bankSh_();
    const { idx: bIdx } = headerIndex_(bank);

    const mustBank = ["Date", "Type", "Amount", "Reference", "RawLine"];
    for (const h of mustBank) {
      if (bIdx[h] == null) {
        return ok_({ found: false, retryAfterSec: 15, message: `Bank sheet missing header: ${h}. Auto-checking…` });
      }
    }

    const last = bank.getLastRow();
    if (last < 2) {
      return ok_({ found: false, retryAfterSec: 15, message: "Bank sheet has no transactions yet. Auto-checking…" });
    }

    const N    = Number(getSetting_("PAYMENT_SCAN_ROWS", 250)) || 250;
    const from = Math.max(2, last - N + 1);
    const data = bank.getRange(from, 1, last - from + 1, bank.getLastColumn()).getValues();

    for (let i = data.length - 1; i >= 0; i--) {
      const r    = data[i];
      const type = String(r[bIdx["Type"]] || "").trim().toUpperCase();
      if (type !== "CREDIT") continue;

      const amt = asNumber_(r[bIdx["Amount"]]);
      const raw = String(r[bIdx["RawLine"]]    || "");
      const ref = String(r[bIdx["Reference"]]  || "");

      const hit =
        (oid       && raw.includes(oid))       ||
        (oidNoDash && raw.includes(oidNoDash)) ||
        (oid       && ref.includes(oid))        ||
        (oidNoDash && ref.includes(oidNoDash));

      if (!hit) continue;
      if (finalAmount > 0 && Math.round(amt) !== Math.round(finalAmount)) continue;

      match = { date: r[bIdx["Date"]], txnRef: String(r[bIdx["Reference"]] || "").trim(), amount: amt, rawLine: raw.slice(0, 250) };
      break;
    }
  } catch (e) {
    return ok_({ found: false, retryAfterSec: 15, message: "Temporary bank scan issue. Auto-checking…" });
  }

  if (!match) {
    return ok_({ found: false, retryAfterSec: 10, message: "Payment not detected yet. Auto-checking…" });
  }

  // ---- Payment found: mark PAID immediately, return fast ----
  try {
    orders.getRange(rowNum, idx["Status"] + 1).setValue(ORDER_STATUS.PAID);
    if (idx["TxnRef"]     != null) orders.getRange(rowNum, idx["TxnRef"]     + 1).setValue(match.txnRef);
    if (idx["VerifiedAt"] != null) orders.getRange(rowNum, idx["VerifiedAt"] + 1).setValue(now_());
  } catch (e) {
    return ok_({ found: false, retryAfterSec: 10, message: "Payment found but write failed. Auto-checking…" });
  }

  // Coins + coupon (fire-and-forget, non-blocking)
  try {
    const phoneCoins = String(orders.getRange(rowNum, idx["Phone"] + 1).getValue() || "").replace(/\D/g, "");
    if (typeof awardCoinsForOrder_ === "function") {
      awardCoinsForOrder_("NEW_PURCHASE", orderId, phoneCoins, service, plan, finalAmount, "Paid & verified");
    }
  } catch (e) {}

  try {
    const couponCode = String(orders.getRange(rowNum, idx["CouponCode"] + 1).getValue() || "").trim().toUpperCase();
    const discount   = asNumber_(orders.getRange(rowNum, idx["Discount"]   + 1).getValue());
    const phone      = String(orders.getRange(rowNum, idx["Phone"]        + 1).getValue() || "").replace(/\D/g, "");
    const email      = String(orders.getRange(rowNum, idx["Email"]        + 1).getValue() || "").trim();
    if (couponCode && discount > 0) {
      if (!couponAlreadyUsedForOrder_(couponCode, phone, orderId)) {
        if      (typeof markCouponUsedForOrder_ === "function") markCouponUsedForOrder_(couponCode, phone, email, discount, orderId);
        else if (typeof logCouponUse_            === "function") logCouponUse_(couponCode, phone, orderId, discount, orderId, "USED");
      }
    }
  } catch (e) {}

  try {
    if (typeof notifyTelegram_ === "function") {
      notifyTelegram_("PAYMENT_VERIFIED", { orderId, service, plan, amount: finalAmount, txnRef: match.txnRef });
    }
  } catch (e) {}

  // ✅ Return immediately — frontend shows "Payment Verified!" right now.
  // Allocation happens when frontend calls fulfillAndGetAccess.
  return ok_({
    found: true,
    paid:  true,
    fulfillment: "PAID",
    message: "✅ Payment verified! Allocating your subscription…"
  });
}


/**
 * Step 2: Fulfill and return credentials.
 * Called by frontend AFTER showing the payment-verified celebration.
 * Has a hard lock + PROCESSING guard — only one execution per order ever runs.
 * Frontend polls this every 3s until fulfillment === "FULFILLED" or "MANUAL_PENDING".
 */
function fulfillAndGetAccess(orderId) {
  let orders, rowNum, idx, fulfillStatus;

  try {
    orders = sh_(TAB_ORDERS);
    rowNum = getRowByValue_(orders, "OrderID", String(orderId || "").trim());
    if (!rowNum) return bad_("Order not found.");
    ({ idx } = headerIndex_(orders));
    fulfillStatus = String(orders.getRange(rowNum, idx["FulfillmentStatus"] + 1).getValue() || "").trim().toUpperCase();
  } catch (e) {
    return bad_("Order lookup failed: " + e);
  }

  // Already done — return credentials right away without acquiring lock
  if (fulfillStatus === FULFILL_STATUS.FULFILLED) {
    return _buildFulfilledResponse_(orderId, orders, rowNum, idx);
  }
  if (fulfillStatus === FULFILL_STATUS.MANUAL_PENDING) {
    return ok_({ found: true, fulfillment: "MANUAL_PENDING", message: "✅ Payment received. Activation will be done manually.", orderId });
  }
  if (fulfillStatus === "FAILED") {
    return ok_({ found: true, fulfillment: "FAILED", message: "😔 Fulfillment failed. Please contact WhatsApp support.", orderId });
  }

  // PROCESSING = another request is already running fulfillOrder — just wait
  if (fulfillStatus === "PROCESSING") {
    return ok_({ found: false, fulfillment: "PROCESSING", retryAfterSec: 3, message: "Allocating your account…" });
  }

  // Acquire lock BEFORE writing PROCESSING — this is the single choke point
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (e) {
    return ok_({ found: false, fulfillment: "PROCESSING", retryAfterSec: 3, message: "Server busy. Retrying…" });
  }

  try {
    // Re-read inside lock — another request may have won the race while we waited
    const fresh = String(orders.getRange(rowNum, idx["FulfillmentStatus"] + 1).getValue() || "").trim().toUpperCase();
    if (fresh === FULFILL_STATUS.FULFILLED) {
      lock.releaseLock();
      return _buildFulfilledResponse_(orderId, orders, rowNum, idx);
    }
    if (fresh === "PROCESSING" || fresh === FULFILL_STATUS.MANUAL_PENDING) {
      lock.releaseLock();
      return ok_({ found: false, fulfillment: fresh, retryAfterSec: 3, message: "Allocating your account…" });
    }

    // ✅ Mark PROCESSING — any concurrent fulfillAndGetAccess call will now back off
    orders.getRange(rowNum, idx["FulfillmentStatus"] + 1).setValue("PROCESSING");
  } catch (e) {
    try { lock.releaseLock(); } catch (_) {}
    return ok_({ found: false, fulfillment: "PROCESSING", retryAfterSec: 3, message: "Retrying…" });
  }

  lock.releaseLock();

  // Run fulfillment — PROCESSING is set, so no concurrent call will re-allocate
  try {
    return fulfillOrder(orderId);
  } catch (e) {
    return ok_({ found: false, retryAfterSec: 5, message: "Fulfillment error. Retrying…" });
  }
}


/**
 * Build the standard FULFILLED response with credentials.
 * Used by both verifyPayment and fulfillAndGetAccess.
 */
function _buildFulfilledResponse_(orderId, orders, rowNum, idx) {
  try {
    const orderType  = (idx["OrderType"]  != null) ? String(orders.getRange(rowNum, idx["OrderType"]  + 1).getValue() || "").trim().toUpperCase() : "";
    const renewSubId = (idx["RenewSubID"] != null) ? String(orders.getRange(rowNum, idx["RenewSubID"] + 1).getValue() || "").trim() : "";
    const service    = String(orders.getRange(rowNum, idx["Service"] + 1).getValue() || "").trim();
    const plan       = String(orders.getRange(rowNum, idx["Plan"]    + 1).getValue() || "").trim();
    const planCfg    = getPlanCfg_(service, plan) || {};

    let access = {};
    if (orderType === "RENEW" && renewSubId) {
      const subs = sh_(TAB_SUBS);
      const srn  = getRowByValue_(subs, "SubID", renewSubId);
      if (srn) {
        const { idx: sIdx } = headerIndex_(subs);
        access = {
          user:          (sIdx["LoginId"]      != null) ? String(subs.getRange(srn, sIdx["LoginId"]      + 1).getValue() || "") : "",
          pass:          (sIdx["Password"]      != null) ? String(subs.getRange(srn, sIdx["Password"]      + 1).getValue() || "") : "",
          profileNumber: (sIdx["ProfileNumber"] != null) ? subs.getRange(srn, sIdx["ProfileNumber"] + 1).getValue() : "",
          profileName:   (sIdx["ProfileName"]   != null) ? String(subs.getRange(srn, sIdx["ProfileName"]   + 1).getValue() || "") : "",
          profilePin:    (sIdx["ProfilePIN"]    != null) ? String(subs.getRange(srn, sIdx["ProfilePIN"]    + 1).getValue() || "") : ""
        };
      }
    } else {
      access = getAccessFromSubsByOrderId_(orderId) || {};
    }

    const emailMode = (orderType === "RENEW") ? "RENEWED" : "FULFILLED";
    sendOrderEmailsOnce_(orderId, emailMode);

    return ok_({
      found:       true,
      orderId,
      fulfillment: "FULFILLED",
      message:     (orderType === "RENEW") ? "🔁 Renewed. Here is your access." : "✅ Access granted.",
      postPaymentMessage: planCfg.postPaymentMessage || "",
      access
    });
  } catch (e) {
    return ok_({ found: true, orderId, fulfillment: "FULFILLED", message: "✅ Done. Check your email for credentials.", access: {} });
  }
}


function normalizeOrderId_(orderId) {
  return String(orderId || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}


function couponAlreadyUsedForOrder_(code, phone, orderId) {
  try {
    const ss  = getCore_();
    const sh  = ss.getSheetByName("COUPON_USAGE");
    if (!sh) return false;
    const vals = sh.getDataRange().getValues();
    code    = String(code    || "").toUpperCase();
    phone   = String(phone   || "").replace(/\D/g, "");
    orderId = String(orderId || "").trim();
    for (let r = vals.length - 1; r >= 1; r--) {
      if (String(vals[r][1] || "").toUpperCase()         === code    &&
          String(vals[r][2] || "").replace(/\D/g, "")    === phone   &&
          String(vals[r][3] || "").trim()                === orderId &&
          String(vals[r][5] || "").toUpperCase()         === "USED") return true;
    }
    return false;
  } catch (e) { return false; }
}


/**
 * Fast status check — no bank scan, no fulfillment.
 */
function getOrderStatus(orderId) {
  try {
    const orders  = sh_(TAB_ORDERS);
    const rowNum  = getRowByValue_(orders, "OrderID", String(orderId || "").trim());
    if (!rowNum) return ok_({ paid: false, fulfilled: false, manual: false, found: false });

    const { idx } = headerIndex_(orders);

    const status        = idx["Status"]            != null ? String(orders.getRange(rowNum, idx["Status"]            + 1).getValue() || "").trim().toUpperCase() : "";
    const fulfillStatus = idx["FulfillmentStatus"] != null ? String(orders.getRange(rowNum, idx["FulfillmentStatus"] + 1).getValue() || "").trim().toUpperCase() : "";

    const paid       = status === ORDER_STATUS.PAID;
    const fulfilled  = fulfillStatus === FULFILL_STATUS.FULFILLED;
    const manual     = fulfillStatus === FULFILL_STATUS.MANUAL_PENDING;
    const processing = fulfillStatus === "PROCESSING";

    return ok_({ paid, fulfilled, manual, processing, status, fulfillStatus, found: true });
  } catch (e) {
    return ok_({ paid: false, fulfilled: false, manual: false, processing: false, error: String(e) });
  }
}


/**
 * safeFulfill_ kept for any internal callers that still reference it.
 * Routes through the new fulfillAndGetAccess guard.
 */
function safeFulfill_(orderId) {
  return fulfillAndGetAccess(orderId);
}