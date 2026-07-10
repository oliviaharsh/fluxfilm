/*************** FluxFilm Backend — Fulfillment (FIXED: PROCESSING guard + renewal expiry sync) ***************/

/**
 * Fulfillment rules:
 * - If already FULFILLED -> return credentials again (from SUBSCRIPTIONS)
 * - If PROCESSING -> return retry (another request is mid-fulfillment)
 * - If MANUAL_PENDING -> return postPaymentMessage (no credentials)
 * - Else require Status=PAID
 * - MANUAL plans -> create MANUAL task + create SUB row (no credentials yet)
 * - INSTANT plans:
 *    - Netflix:
 *        - Sharing plan -> allocate SAME reserved profile (P1) with CAPACITY (MaxTotal)
 *        - Private plan -> allocate rotating profile (P2+) 1-to-1 with cooldown
 *    - Prime:
 *        - Capacity allocation by device type (TV/NON_TV)
 *    - Account allocation:
 *        - whole account
 *
 * ✅ FIX 1: PROCESSING guard. verifyPayment writes PROCESSING before calling fulfillOrder.
 *           fulfillOrder checks for PROCESSING and treats it as "in flight — skip".
 *           This is the primary double-allocation fix. No RESERVE rows needed.
 *
 * ✅ FIX 2: fulfillRenew_ now syncs the actual computed NewExpiry back to the RENEWALS
 *           sheet row, so the preview written at order-creation time is corrected to
 *           match the real expiry written to SUBSCRIPTIONS.
 */

function fulfillOrder(orderId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const orders = sh_(TAB_ORDERS);
    const rowNum = getRowByValue_(orders, "OrderID", orderId);
    if (!rowNum) return bad_("Order not found.");

    const { idx } = headerIndex_(orders);

    // ---- Read current order state (inside lock — reads are fresh) ----
    const status = String(orders.getRange(rowNum, idx["Status"] + 1).getValue() || "")
      .trim()
      .toUpperCase();

    const fulfillStatus = String(orders.getRange(rowNum, idx["FulfillmentStatus"] + 1).getValue() || "")
      .trim()
      .toUpperCase();

    const service      = String(orders.getRange(rowNum, idx["Service"]      + 1).getValue() || "").trim();
    const plan         = String(orders.getRange(rowNum, idx["Plan"]         + 1).getValue() || "").trim();
    const durationDays = asNumber_(orders.getRange(rowNum, idx["DurationDays"] + 1).getValue());
    const email        = String(orders.getRange(rowNum, idx["Email"]        + 1).getValue() || "").trim();
    const phone        = String(orders.getRange(rowNum, idx["Phone"]        + 1).getValue() || "").trim();
    const notes        = String(orders.getRange(rowNum, idx["Notes"]        + 1).getValue() || "").trim();
    const extraKey     = String(orders.getRange(rowNum, idx["ExtraFieldKey"]   + 1).getValue() || "").trim();
    const extraVal     = String(orders.getRange(rowNum, idx["ExtraFieldValue"] + 1).getValue() || "").trim();
    const name         = (idx["Name"] != null)
      ? String(orders.getRange(rowNum, idx["Name"] + 1).getValue() || "").trim()
      : "";

    const orderType = (idx["OrderType"] != null)
      ? String(orders.getRange(rowNum, idx["OrderType"] + 1).getValue() || "").trim().toUpperCase()
      : "";

    const renewSubId = (idx["RenewSubID"] != null)
      ? String(orders.getRange(rowNum, idx["RenewSubID"] + 1).getValue() || "").trim()
      : "";

    // Read plan config (from PLANS)
    const planCfg = getPlanCfg_(service, plan);
    if (!planCfg) throw new Error("Plan config missing.");

    // ✅ Already fulfilled — return credentials, don't re-allocate
    if (fulfillStatus === FULFILL_STATUS.FULFILLED) {
      const access = getAccessFromSubsByOrderId_(orderId) || {};
      sendOrderEmailsOnce_(orderId, "FULFILLED");
      return ok_({
        found: true,
        orderId,
        fulfillment: "FULFILLED",
        message: "✅ Already fulfilled. Showing credentials again.",
        postPaymentMessage: planCfg.postPaymentMessage || "",
        access
      });
    }

    // ✅ PROCESSING = verifyPayment is mid-fulfillment right now in another execution.
    //    Return a retry signal. The frontend will re-poll and eventually get FULFILLED.
    if (fulfillStatus === "PROCESSING") {
      return ok_({
        found: false,
        retryAfterSec: 4,
        message: "Your order is being processed. Auto-checking…"
      });
    }

    if (fulfillStatus === FULFILL_STATUS.MANUAL_PENDING) {
      sendOrderEmailsOnce_(orderId, "MANUAL_PENDING");
      return ok_({
        found: true,
        orderId,
        fulfillment: "MANUAL_PENDING",
        message: planCfg.postPaymentMessage || "✅ Payment received. Activation will be done manually soon.",
        postPaymentMessage: planCfg.postPaymentMessage || ""
      });
    }

    // Enforce PAID (or PROCESSING treated as PAID — we set it in verifyPayment after finding match)
    if (status !== ORDER_STATUS.PAID) return bad_("Payment not verified yet.");

    // ✅ RENEW FLOW: extend existing subscription instead of creating new SubID
    if (orderType === "RENEW" && renewSubId) {
      const res = fulfillRenew_(orderId, renewSubId, durationDays, {
        name, phone, email, service, plan, notes
      }, planCfg);
      if (typeof notifyTelegram_ === "function") {
        notifyTelegram_("RENEWED", { orderId, subId: renewSubId, service, plan, phone, email });
      }
      return res;
    }

    // Manual fulfillment (YouTube etc.)
    if (planCfg.fulfillmentMode === "MANUAL" || planCfg.allocationPolicy === "NONE") {
      const taskId = "TASK-" + Date.now().toString().slice(-7);
      sh_(TAB_MANUAL).appendRow([
        taskId, orderId, now_(), service,
        "MANUAL_ACTIVATION", "NORMAL", "PENDING",
        email, phone, extraKey, extraVal, notes, "", "",
      ]);

      orders.getRange(rowNum, idx["FulfillmentStatus"] + 1).setValue(FULFILL_STATUS.MANUAL_PENDING);
      if (idx["FulfilledAt"] != null) orders.getRange(rowNum, idx["FulfilledAt"] + 1).setValue(now_());

      const sub = createSubscriptionRecord_(orderId, service, plan, name, email, phone, durationDays, planCfg, {
        inventoryRef: "", deviceType: "", deviceConcurrency: "", access: {}
      });

      if (typeof notifyTelegram_ === "function") {
        notifyTelegram_("FULFILLMENT_MANUAL", { orderId, service, plan, email, phone, extra: extraVal });
      }

      sendOrderEmailsOnce_(orderId, "MANUAL_PENDING");

      return ok_({
        found: true,
        orderId,
        fulfillment: "MANUAL_PENDING",
        message: planCfg.postPaymentMessage || "✅ Payment received. Activation will be done manually within a few hours.",
        postPaymentMessage: planCfg.postPaymentMessage || "",
        subscription: sub
      });
    }

    // -------- Instant fulfillment --------
    let invRes = null;

    const svcLower = String(service || "").toLowerCase();

    if (planCfg.allocationPolicy === "PROFILE" && svcLower.includes("netflix")) {
      invRes = allocateNetflixProfile_(plan, durationDays);
    } else if (planCfg.allocationPolicy === "CAPACITY" && svcLower.includes("prime")) {
      const deviceType = (extraVal || "").toUpperCase();
      if (deviceType !== "TV" && deviceType !== "NON_TV") {
        throw new Error("Prime requires device type: TV or NON_TV (ExtraFieldValue).");
      }
      invRes = allocatePrimeAccountCapacity_(deviceType, durationDays);
    } else if (planCfg.allocationPolicy === "OTP_ACCOUNT") {
      invRes = allocateOtpAccount_(service, plan, durationDays);
    } else if (planCfg.allocationPolicy === "ACCOUNT") {
      invRes = allocateWholeAccount_(service, durationDays);
    } else {
      throw new Error(`Unsupported allocationPolicy: ${planCfg.allocationPolicy}`);
    }

    if (!invRes || !invRes.ok) {
      orders.getRange(rowNum, idx["FulfillmentStatus"] + 1).setValue(FULFILL_STATUS.FAILED);
      if (idx["Error"] != null) orders.getRange(rowNum, idx["Error"] + 1).setValue(invRes?.message || "No accounts available");
      if (typeof notifyTelegram_ === "function") {
        notifyTelegram_("FULFILLMENT_FAILED_NO_STOCK", { orderId, service, plan, phone, email, reason: invRes?.message || "No accounts available" });
      }
      return ok_({
        found: true,
        orderId,
        fulfillment: "NO_STOCK",
        message: "😔 We ran out of accounts for this plan just as your payment came in. Please contact WhatsApp support — we'll sort it out immediately.",
        postPaymentMessage: ""
      });
    }

    // Mark order fulfilled
    orders.getRange(rowNum, idx["FulfillmentStatus"] + 1).setValue(FULFILL_STATUS.FULFILLED);
    if (idx["FulfilledAt"]  != null) orders.getRange(rowNum, idx["FulfilledAt"]  + 1).setValue(now_());
    if (idx["InventoryRef"] != null) orders.getRange(rowNum, idx["InventoryRef"] + 1).setValue(invRes.inventoryRef || "");

    // Create subscription record
    const sub = createSubscriptionRecord_(orderId, service, plan, name, email, phone, durationDays, planCfg, invRes);

    notifyTelegram_("FULFILLED", {
      orderId, service, plan,
      name, phone, email,
      inv: invRes.inventoryRef,
      access: invRes.access || {}
    });

    sendOrderEmailsOnce_(orderId, "FULFILLED");

    return ok_({
      found: true,
      orderId,
      fulfillment: "FULFILLED",
      postPaymentMessage: planCfg.postPaymentMessage || "",
      access: invRes.access || {},
      inventoryRef: invRes.inventoryRef || "",
      subscription: sub,
      message: "✅ Payment verified and access granted."
    });

  } catch (e) {
    // Record error on order — but do NOT leave it as PROCESSING (reset to FAILED)
    try {
      const orders = sh_(TAB_ORDERS);
      const rn = getRowByValue_(orders, "OrderID", orderId);
      if (rn) {
        const { idx } = headerIndex_(orders);
        if (idx["FulfillmentStatus"] != null) orders.getRange(rn, idx["FulfillmentStatus"] + 1).setValue(FULFILL_STATUS.FAILED);
        if (idx["Error"]             != null) orders.getRange(rn, idx["Error"]             + 1).setValue(String(e));
      }
    } catch (_) {}

    return bad_("Fulfillment failed: " + e);

  } finally {
    lock.releaseLock();
  }
}


function fulfillRenew_(orderId, subId, durationDays, meta, planCfg) {
  const subs = sh_(TAB_SUBS);
  const rn = getRowByValue_(subs, "SubID", subId);
  if (!rn) return bad_("Renew failed: SubID not found.");

  const { idx } = headerIndex_(subs);

  // Read old expiry
  const oldExpiry = (idx["ExpiryDate"] != null) ? subs.getRange(rn, idx["ExpiryDate"] + 1).getValue() : "";
  const oldExpiryDate = (oldExpiry instanceof Date && !isNaN(oldExpiry.getTime())) ? oldExpiry : null;

  // Renewal base policy:
  // - Renewing in advance OR late by ≤10 days → extend from old expiry (customer keeps continuity)
  // - Late by >10 days → extend from today (payment date), customer lost the gap days
  const LATE_FROM_TODAY_DAYS = asNumber_(getSetting_("RENEW_BASE_TODAY_AFTER_DAYS", 10));
  const nowDate = now_();
  let base;
  if (!oldExpiryDate) {
    base = nowDate;
  } else {
    const daysLate = Math.ceil((nowDate.getTime() - oldExpiryDate.getTime()) / 86400000);
    base = (daysLate <= LATE_FROM_TODAY_DAYS) ? oldExpiryDate : nowDate;
  }
  const newExpiry = addDays_(base, durationDays || 30);

  // ✅ Update SAME subscription row
  if (idx["ExpiryDate"] != null) subs.getRange(rn, idx["ExpiryDate"] + 1).setValue(newExpiry);
  if (idx["OrderID"]    != null) subs.getRange(rn, idx["OrderID"]    + 1).setValue(orderId);
  if (idx["Status"]     != null) subs.getRange(rn, idx["Status"]     + 1).setValue(SUB_STATUS.ACTIVE);

  // Extend release eligible
  const cooldownDays = asNumber_(getSetting_("REUSE_COOLDOWN_DAYS", 10));
  if (idx["ReleaseEligibleAt"] != null) subs.getRange(rn, idx["ReleaseEligibleAt"] + 1).setValue(addDays_(newExpiry, cooldownDays));

  // Mark order fulfilled
  const orders = sh_(TAB_ORDERS);
  const or = getRowByValue_(orders, "OrderID", orderId);
  const { idx: oIdx } = headerIndex_(orders);

  if (oIdx["FulfillmentStatus"] != null) orders.getRange(or, oIdx["FulfillmentStatus"] + 1).setValue(FULFILL_STATUS.FULFILLED);
  if (oIdx["FulfilledAt"]       != null) orders.getRange(or, oIdx["FulfilledAt"]       + 1).setValue(now_());

  // Keep InventoryRef same
  if (oIdx["InventoryRef"] != null && idx["InventoryRef"] != null) {
    const invRef = subs.getRange(rn, idx["InventoryRef"] + 1).getValue();
    orders.getRange(or, oIdx["InventoryRef"] + 1).setValue(invRef || "");
  }

  // ✅ Update RENEWALS sheet status + sync the actual NewExpiry back
  // (the preview written at createRenewOrder time may differ from the real computed expiry)
  updateRenewalStatusByOrder_(orderId, "RENEWED", newExpiry);

  // Send email
  try { sendOrderEmailsOnce_(orderId, "RENEWED"); } catch(e) {}

  // Return access from same Sub row
  const access = {
    user:          (idx["LoginId"]       != null) ? String(subs.getRange(rn, idx["LoginId"]       + 1).getValue() || "") : "",
    pass:          (idx["Password"]      != null) ? String(subs.getRange(rn, idx["Password"]      + 1).getValue() || "") : "",
    profileNumber: (idx["ProfileNumber"] != null) ? subs.getRange(rn, idx["ProfileNumber"] + 1).getValue() : "",
    profileName:   (idx["ProfileName"]   != null) ? String(subs.getRange(rn, idx["ProfileName"]   + 1).getValue() || "") : "",
    profilePin:    (idx["ProfilePIN"]    != null) ? String(subs.getRange(rn, idx["ProfilePIN"]    + 1).getValue() || "") : ""
  };

  return ok_({
    found: true,
    orderId,
    fulfillment: "RENEWED",
    message: "🔁 Renewed successfully (same SubID).",
    subscription: { subId, oldExpiry: oldExpiry || "", newExpiry: newExpiry },
    postPaymentMessage: planCfg?.postPaymentMessage || "",
    access
  });
}


// ---------- Plan config loader ----------
function getPlanCfg_(service, plan) {
  const plans = sh_(TAB_PLANS);
  const { idx } = headerIndex_(plans);
  const last = plans.getLastRow();
  if (last < 2) return null;
  const data = plans.getRange(2, 1, last - 1, plans.getLastColumn()).getValues();

  for (const r of data) {
    if (String(r[idx["Service"]] || "").trim() === service && String(r[idx["Plan"]] || "").trim() === plan) {
      const active = String(r[idx["IsActive"]] || "").trim().toUpperCase() === "TRUE";
      if (!active) return null;
      return {
        fulfillmentMode:    String(r[idx["FulfillmentMode"]]    || "").trim().toUpperCase(),
        allocationPolicy:   String(r[idx["AllocationPolicy"]]   || "").trim().toUpperCase(),
        postPaymentMessage: String(r[idx["PostPaymentMessage"]] || "").trim()
      };
    }
  }
  return null;
}


// ---------- Email sender (IDEMPOTENT) ----------
function sendOrderEmailsOnce_(orderId, mode) {
  try {
    if (typeof sendOrderEmails_ !== "function") return;
    if (wasEmailSent_(orderId, mode)) return;

    sendOrderEmails_(orderId, mode);

    if (mode === "FULFILLED") {
      try {
        const subs = sh_(TAB_SUBS);
        const rn = getRowByValue_(subs, "OrderID", orderId);
        if (rn) {
          const { idx } = headerIndex_(subs);
          const sentCol = idx["LastAccessSentAt"];
          if (sentCol != null && !subs.getRange(rn, sentCol + 1).getValue()) {
            subs.getRange(rn, sentCol + 1).setValue(now_());
          }
        }
      } catch (_) {}
    }
  } catch (e) {}
}


// ---------- Subscription writer ----------
function createSubscriptionRecord_(orderId, service, plan, name, email, phone, durationDays, planCfg, invRes) {
  const subs = sh_(TAB_SUBS);
  const subId = genSubId_();

  // ✅ If Netflix PRIVATE_ROTATING marked profile with placeholder, replace it with the real SubID
  try {
    if (invRes && invRes.inventoryRef && String(invRes.inventoryRef).indexOf("#P") > -1) {
      const invProf = sh_(TAB_INV_PROFILES);
      const { idx: pIdx } = headerIndex_(invProf);

      if (pIdx["AccountID"] != null && pIdx["ProfileNumber"] != null && pIdx["CurrentSubID"] != null) {
        const parts = String(invRes.inventoryRef).split("#P");
        const accId = parts[0];
        const pNo   = Number(parts[1]);

        const last = invProf.getLastRow();
        if (last >= 2) {
          const data = invProf.getRange(2, 1, last - 1, invProf.getLastColumn()).getValues();
          for (let i = 0; i < data.length; i++) {
            const r = data[i];
            if (String(r[pIdx["AccountID"]] || "").trim() === accId && Number(r[pIdx["ProfileNumber"]]) === pNo) {
              const cur = String(r[pIdx["CurrentSubID"]] || "").trim();
              if (cur === "PENDING_SUB_ID" || cur === "") {
                invProf.getRange(i + 2, pIdx["CurrentSubID"] + 1).setValue(subId);
              }
              break;
            }
          }
        }
      }
    }
  } catch (e) {}

  const start  = now_();
  const expiry = addDays_(start, durationDays || 30);
  const cooldownDays      = asNumber_(getSetting_("REUSE_COOLDOWN_DAYS", 10));
  const releaseEligibleAt = addDays_(expiry, cooldownDays);

  const access            = invRes.access || {};
  const deviceType        = invRes.deviceType || "";
  const deviceConcurrency = invRes.deviceConcurrency || "";

  subs.appendRow([
    subId, orderId, name, phone, email,
    service, plan, start, expiry,
    SUB_STATUS.ACTIVE,
    planCfg.fulfillmentMode,
    planCfg.allocationPolicy,
    invRes.inventoryRef || "",
    deviceType, deviceConcurrency,
    access.user || "",
    access.pass || "",
    access.profileNumber || "",
    access.profileName   || "",
    access.profilePin    || "",
    "",  // Notes
    "",  // LastAccessSentAt
    releaseEligibleAt,
    "", "", ""  // AdminStatusOverride, AdminNote, AdminUpdatedAt
  ]);

  return { subId, start, expiry, status: SUB_STATUS.ACTIVE };
}


// ---------- Helper: return access from SUBSCRIPTIONS by OrderID ----------
function getAccessFromSubsByOrderId_(orderId) {
  try {
    const subs = sh_(TAB_SUBS);
    const rn   = getRowByValue_(subs, "OrderID", orderId);
    if (!rn) return {};

    const { idx } = headerIndex_(subs);

    function gv(col) {
      if (idx[col] == null) return "";
      return subs.getRange(rn, idx[col] + 1).getValue();
    }

    return {
      user:          String(gv("LoginId")       || ""),
      pass:          String(gv("Password")       || ""),
      profileNumber: gv("ProfileNumber")          || "",
      profileName:   String(gv("ProfileName")    || ""),
      profilePin:    String(gv("ProfilePIN")     || "")
    };
  } catch (e) {
    return {};
  }
}


// ---------- Allocation engines ----------

function allocateNetflixProfile_(plan, durationDays) {
  const planLower = String(plan || "").toLowerCase();
  const isSharing = planLower.includes("sharing") || planLower.includes("group");

  if (isSharing) {
    return allocateNetflixSharingSlot_(durationDays);
  }
  return allocateNetflixPrivateProfile_(durationDays);
}

function allocateNetflixSharingSlot_(durationDays) {
  const invAcc  = sh_(TAB_INV_ACCOUNTS);
  const invProf = sh_(TAB_INV_PROFILES);
  const cap     = sh_(TAB_INV_CAPACITY);

  const { idx: aIdx } = headerIndex_(invAcc);
  const { idx: pIdx } = headerIndex_(invProf);
  const { idx: cIdx } = headerIndex_(cap);

  const sharingProfileNo = asNumber_(getSetting_("NETFLIX_SHARING_PROFILE_NO", 1));
  const maxTotalDefault  = asNumber_(getSetting_("NETFLIX_SHARING_MAX_TOTAL", 5));

  const aLast = invAcc.getLastRow();
  if (aLast < 2) return bad_("No inventory accounts.");
  const aData = invAcc.getRange(2, 1, aLast - 1, invAcc.getLastColumn()).getValues();

  const netflixAccounts = aData
    .map((r, i) => ({ r, rowNum: i + 2 }))
    .filter(o => String(o.r[aIdx["Service"]] || "").trim().toLowerCase().includes("netflix"))
    .filter(o => String(o.r[aIdx["IsActive"]] || "").trim().toUpperCase() === "TRUE");

  if (!netflixAccounts.length) return bad_("No active Netflix accounts.");

  // Load capacity map
  const capMap = new Map();
  const cLast = cap.getLastRow();
  if (cLast >= 2) {
    const cData = cap.getRange(2, 1, cLast - 1, cap.getLastColumn()).getValues();
    cData.forEach(r => {
      const svc = String(r[cIdx["Service"]] || "").trim().toLowerCase();
      if (!svc.includes("netflix")) return;
      const accountId = String(r[cIdx["AccountID"]] || "").trim();
      const isActive  = String(r[cIdx["IsActive"]]  || "").trim().toUpperCase() === "TRUE";
      const maxTotal  = asNumber_(r[cIdx["MaxTotal"]]) || maxTotalDefault;
      if (accountId) capMap.set(accountId, { maxTotal, isActive });
    });
  }

  // Load all profiles once
  const pLast = invProf.getLastRow();
  const pData = pLast < 2 ? [] : invProf.getRange(2, 1, pLast - 1, invProf.getLastColumn()).getValues();

  function findSharingProfile_(accountId) {
    let fallback = null;
    for (let i = 0; i < pData.length; i++) {
      const r     = pData[i];
      if (String(r[pIdx["AccountID"]] || "").trim() !== accountId) continue;
      const type  = String(r[pIdx["ProfileType"]] || "").trim().toUpperCase();
      const isRes = String(r[pIdx["IsReserved"]]  || "").trim().toUpperCase() === "TRUE";
      const pNo   = asNumber_(r[pIdx["ProfileNumber"]]);
      if (type === "SHARING_RESERVED" || type === "SHARING_RES" || isRes) {
        if (pNo === sharingProfileNo) return { r, rowNum: i + 2 };
        if (!fallback) fallback = { r, rowNum: i + 2 };
      }
    }
    return fallback;
  }

  // ✅ Always scan SUBSCRIPTIONS live for occupancy — never trust formula columns
  // for concurrent safety. The PROCESSING guard in verifyPayment means only one
  // fulfillOrder runs per order, but for separate simultaneous orders on the same
  // account we need a live count.
  const subsSheet = sh_(TAB_SUBS);
  const { idx: sIdx } = headerIndex_(subsSheet);
  const sLast = subsSheet.getLastRow();
  const sData = sLast < 2 ? [] : subsSheet.getRange(2, 1, sLast - 1, subsSheet.getLastColumn()).getValues();
  const nowMs = now_().getTime();

  function countSharingOccupancy_(inventoryRef) {
    let count = 0;
    for (const r of sData) {
      const svc = String(r[sIdx["Service"]] || "").trim().toLowerCase();
      if (!svc.includes("netflix")) continue;
      if (String(r[sIdx["InventoryRef"]] || "").trim() !== inventoryRef) continue;
      const st = String(r[sIdx["Status"]] || "").trim().toUpperCase();
      if (st !== "ACTIVE") continue;
      const exp   = r[sIdx["ExpiryDate"]];
      const expMs = exp instanceof Date ? exp.getTime() : new Date(exp).getTime();
      if (isNaN(expMs) || nowMs >= expMs) continue;
      count++;
    }
    return count;
  }

  // Build candidate list, sort least-used first
  const sharingCandidates = [];

  for (const accObj of netflixAccounts) {
    const accountId = String(accObj.r[aIdx["AccountID"]] || "").trim();
    const loginId   = String(accObj.r[aIdx["LoginId"]]   || "").trim();
    const pass      = String(accObj.r[aIdx["Password"]]   || "").trim();
    if (!accountId || !loginId || !pass) continue;

    const capCfg = capMap.get(accountId) || { maxTotal: maxTotalDefault, isActive: true };
    if (!capCfg.isActive) continue;

    const prof = findSharingProfile_(accountId);
    if (!prof) continue;

    const pNo = asNumber_(prof.r[pIdx["ProfileNumber"]]);
    if (!pNo) continue;

    const inventoryRef = `${accountId}#P${pNo}`;
    const used         = countSharingOccupancy_(inventoryRef);

    if (used >= capCfg.maxTotal) continue;

    sharingCandidates.push({ accObj, accountId, loginId, pass, capCfg, prof, pNo, inventoryRef, used });
  }

  if (!sharingCandidates.length) {
    return bad_("No Netflix sharing slots available right now. All accounts are at capacity.");
  }

  sharingCandidates.sort((a, b) => a.used - b.used);
  const picked = sharingCandidates[0];

  const profileName = String(picked.prof.r[pIdx["ProfileDisplayName"]] || "").trim() || "FluxFilm";
  const profilePin  = String(picked.prof.r[pIdx["ProfilePIN"]]         || "").trim();

  return ok_({
    inventoryRef: picked.inventoryRef,
    access: {
      user:          picked.loginId,
      pass:          picked.pass,
      profileNumber: picked.pNo,
      profileName:   profileName,
      profilePin:    profilePin
    }
  });
}

function allocateNetflixPrivateProfile_(durationDays) {
  const invAcc  = sh_(TAB_INV_ACCOUNTS);
  const invProf = sh_(TAB_INV_PROFILES);

  const { idx: aIdx } = headerIndex_(invAcc);
  const { idx: pIdx } = headerIndex_(invProf);

  const sharingProfileNo = asNumber_(getSetting_("NETFLIX_SHARING_PROFILE_NO", 1));
  const cooldownDays     = asNumber_(getSetting_("REUSE_COOLDOWN_DAYS", 10));

  const aLast = invAcc.getLastRow();
  if (aLast < 2) return bad_("No inventory accounts.");
  const aData = invAcc.getRange(2, 1, aLast - 1, invAcc.getLastColumn()).getValues();

  const netflixAccounts = aData
    .map((r, i) => ({ r, rowNum: i + 2 }))
    .filter(o => String(o.r[aIdx["Service"]] || "").trim().toLowerCase().includes("netflix"))
    .filter(o => String(o.r[aIdx["IsActive"]] || "").trim().toUpperCase() === "TRUE");

  if (!netflixAccounts.length) return bad_("No active Netflix accounts.");

  const pLast = invProf.getLastRow();
  const pData = pLast < 2 ? [] : invProf.getRange(2, 1, pLast - 1, invProf.getLastColumn()).getValues();
  const nowMs = now_().getTime();

  function pickPrivateProfile_(accountId) {
    let best = null;

    for (let i = 0; i < pData.length; i++) {
      const r = pData[i];
      if (String(r[pIdx["AccountID"]] || "").trim() !== accountId) continue;

      const pNo = asNumber_(r[pIdx["ProfileNumber"]]);
      if (pNo === sharingProfileNo) continue;

      const type = String(r[pIdx["ProfileType"]] || "").trim().toUpperCase();
      if (type !== "PRIVATE_ROTATING") continue;

      const statusRaw = String(r[pIdx["Status"]] || "").trim().toUpperCase();
      const status    = statusRaw || "FREE";
      const curSub    = String(r[pIdx["CurrentSubID"]] || "").trim();

      if (status === "ASSIGNED" && !curSub) {
        return { r, rowNum: i + 2 };
      }
      if (status === "FREE") {
        return { r, rowNum: i + 2 };
      }

      const rel   = r[pIdx["ReleaseEligibleAt"]];
      const relMs = rel instanceof Date ? rel.getTime() : new Date(rel).getTime();
      if ((status === "COOLDOWN" || status === "ASSIGNED") && !isNaN(relMs) && nowMs >= relMs) {
        best = { r, rowNum: i + 2 };
      }
    }
    return best;
  }

  const privateCandidates = [];

  for (const accObj of netflixAccounts) {
    const accountId = String(accObj.r[aIdx["AccountID"]] || "").trim();
    const loginId   = String(accObj.r[aIdx["LoginId"]]   || "").trim();
    const pass      = String(accObj.r[aIdx["Password"]]   || "").trim();
    if (!accountId || !loginId || !pass) continue;

    const prof = pickPrivateProfile_(accountId);
    if (!prof) continue;

    let assignedCount = 0;
    for (let i = 0; i < pData.length; i++) {
      const r    = pData[i];
      if (String(r[pIdx["AccountID"]] || "").trim() !== accountId) continue;
      const type = String(r[pIdx["ProfileType"]] || "").trim().toUpperCase();
      if (type !== "PRIVATE_ROTATING") continue;
      const pNo  = asNumber_(r[pIdx["ProfileNumber"]]);
      if (pNo === sharingProfileNo) continue;
      const st   = String(r[pIdx["Status"]] || "").trim().toUpperCase();
      if (st === "ASSIGNED") assignedCount++;
    }

    privateCandidates.push({ accountId, loginId, pass, prof, assignedCount });
  }

  if (!privateCandidates.length) {
    return bad_("No Netflix private profiles available right now.");
  }

  privateCandidates.sort((a, b) => a.assignedCount - b.assignedCount);
  const picked = privateCandidates[0];

  const prof        = picked.prof;
  const pNo         = asNumber_(prof.r[pIdx["ProfileNumber"]]);
  const profileName = String(prof.r[pIdx["ProfileDisplayName"]] || "").trim() || "Private";
  const profilePin  = String(prof.r[pIdx["ProfilePIN"]]         || "").trim();

  const expiry  = addDays_(now_(), durationDays || 30);
  const release = addDays_(expiry, cooldownDays);

  if (pIdx["Status"]           != null) invProf.getRange(prof.rowNum, pIdx["Status"]           + 1).setValue("ASSIGNED");
  if (pIdx["CurrentSubID"]     != null) invProf.getRange(prof.rowNum, pIdx["CurrentSubID"]     + 1).setValue("PENDING_SUB_ID");
  if (pIdx["AssignedAt"]       != null) invProf.getRange(prof.rowNum, pIdx["AssignedAt"]       + 1).setValue(now_());
  if (pIdx["ExpiryDate"]       != null) invProf.getRange(prof.rowNum, pIdx["ExpiryDate"]       + 1).setValue(expiry);
  if (pIdx["ReleaseEligibleAt"]!= null) invProf.getRange(prof.rowNum, pIdx["ReleaseEligibleAt"]+ 1).setValue(release);

  const inventoryRef = `${picked.accountId}#P${pNo}`;

  return ok_({
    inventoryRef,
    access: {
      user:          picked.loginId,
      pass:          picked.pass,
      profileNumber: pNo,
      profileName:   profileName,
      profilePin:    profilePin
    }
  });
}

// Prime: CAPACITY allocation (Total<=4, TV<=2)
// ✅ Always scans SUBSCRIPTIONS live — no formula column dependency.
//    PROCESSING guard in verifyPayment means only one fulfillOrder per order runs,
//    so concurrent separate-order collisions are the only remaining edge case,
//    handled by the live scan.
function allocatePrimeAccountCapacity_(deviceType, durationDays) {
  const invAcc = sh_(TAB_INV_ACCOUNTS);
  const cap    = sh_(TAB_INV_CAPACITY);

  const { idx: aIdx } = headerIndex_(invAcc);
  const { idx: cIdx } = headerIndex_(cap);

  const maxTotalDefault = asNumber_(getSetting_("PRIME_MAX_TOTAL", 4));
  const maxTVDefault    = asNumber_(getSetting_("PRIME_MAX_TV", 2));

  const aLast = invAcc.getLastRow();
  if (aLast < 2) return bad_("No inventory accounts.");
  const aData = invAcc.getRange(2, 1, aLast - 1, invAcc.getLastColumn()).getValues();

  const primeAccounts = aData
    .map((r, i) => ({ r, rowNum: i + 2 }))
    .filter(o => String(o.r[aIdx["Service"]] || "").trim().toLowerCase().includes("prime"))
    .filter(o => String(o.r[aIdx["IsActive"]] || "").trim().toUpperCase() === "TRUE");

  if (!primeAccounts.length) return bad_("No active Prime accounts.");

  const capMap = new Map();
  const cLast  = cap.getLastRow();
  if (cLast >= 2) {
    const cData = cap.getRange(2, 1, cLast - 1, cap.getLastColumn()).getValues();
    cData.forEach(r => {
      const service   = String(r[cIdx["Service"]]   || "").trim().toLowerCase();
      if (!service.includes("prime")) return;
      const accountId = String(r[cIdx["AccountID"]] || "").trim();
      const isActive  = String(r[cIdx["IsActive"]]  || "").trim().toUpperCase() === "TRUE";
      const maxTotal  = asNumber_(r[cIdx["MaxTotal"]] || maxTotalDefault);
      const maxTV     = asNumber_(r[cIdx["MaxTV"]]    || maxTVDefault);
      if (accountId) capMap.set(accountId, { maxTotal, maxTV, isActive });
    });
  }

  // Live scan of SUBSCRIPTIONS — accurate even under concurrent load
  const subsSheet = sh_(TAB_SUBS);
  const { idx: sIdx } = headerIndex_(subsSheet);
  const sLast = subsSheet.getLastRow();
  const sData = sLast < 2 ? [] : subsSheet.getRange(2, 1, sLast - 1, subsSheet.getLastColumn()).getValues();
  const nowMs = now_().getTime();

  function occupancy(accountId) {
    let total = 0, tv = 0;
    for (const r of sData) {
      const svc = String(r[sIdx["Service"]] || "").trim().toLowerCase();
      if (!svc.includes("prime")) continue;
      if (String(r[sIdx["InventoryRef"]] || "").trim() !== accountId) continue;
      const st = String(r[sIdx["Status"]] || "").trim().toUpperCase();
      if (st !== "ACTIVE") continue;
      const release = r[sIdx["ReleaseEligibleAt"]];
      const relMs   = release instanceof Date ? release.getTime() : new Date(release).getTime();
      if (isNaN(relMs) || nowMs >= relMs) continue;
      total++;
      if (String(r[sIdx["DeviceType"]] || "").trim().toUpperCase() === "TV") tv++;
    }
    return { total, tv };
  }

  const primeCandidates = [];

  for (const accObj of primeAccounts) {
    const accountId = String(accObj.r[aIdx["AccountID"]] || "").trim();
    const loginId   = String(accObj.r[aIdx["LoginId"]]   || "").trim();
    const pass      = String(accObj.r[aIdx["Password"]]   || "").trim();
    if (!accountId || !loginId || !pass) continue;

    const capCfg = capMap.get(accountId) || { maxTotal: maxTotalDefault, maxTV: maxTVDefault, isActive: true };
    if (!capCfg.isActive) continue;

    const occ = occupancy(accountId);
    if (occ.total >= capCfg.maxTotal) continue;
    if (deviceType === "TV" && occ.tv >= capCfg.maxTV) continue;

    primeCandidates.push({ accObj, accountId, loginId, pass, occ });
  }

  if (!primeCandidates.length) {
    return bad_("Prime slots are full right now (TV/non-TV capacity).");
  }

  primeCandidates.sort((a, b) => a.occ.total - b.occ.total);
  const picked = primeCandidates[0];

  return ok_({
    inventoryRef:      picked.accountId,
    deviceType,
    deviceConcurrency: "1",
    access: { user: picked.loginId, pass: picked.pass }
  });
}

// Account-based: allocate any active account (whole account)
function allocateWholeAccount_(service, durationDays) {
  const invAcc = sh_(TAB_INV_ACCOUNTS);
  const { idx } = headerIndex_(invAcc);

  const last = invAcc.getLastRow();
  if (last < 2) return bad_("No inventory accounts.");
  const data = invAcc.getRange(2, 1, last - 1, invAcc.getLastColumn()).getValues();

  const svcLower = String(service || "").trim().toLowerCase();

  for (let i = 0; i < data.length; i++) {
    const r      = data[i];
    const s      = String(r[idx["Service"]]  || "").trim().toLowerCase();
    const active = String(r[idx["IsActive"]] || "").trim().toUpperCase() === "TRUE";
    if (!active) continue;
    if (!s.includes(svcLower)) continue;

    const accountId = String(r[idx["AccountID"]] || "").trim();
    const loginId   = String(r[idx["LoginId"]]   || "").trim();
    const pass      = String(r[idx["Password"]]   || "").trim();
    if (!accountId || !loginId || !pass) continue;

    return ok_({
      inventoryRef: accountId,
      access: { user: loginId, pass: pass }
    });
  }

  return bad_("No accounts available for this service.");
}


// ✅ UPDATED: now accepts actualNewExpiry and syncs it back to RENEWALS sheet
function updateRenewalStatusByOrder_(orderId, status, actualNewExpiry) {
  try {
    const rw = sh_("RENEWALS");
    const { idx } = headerIndex_(rw);

    if (idx["OrderID"] == null || idx["Status"] == null) return;

    const last = rw.getLastRow();
    if (last < 2) return;

    const data = rw.getRange(2, 1, last - 1, rw.getLastColumn()).getValues();

    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][idx["OrderID"]] || "").trim() === String(orderId || "").trim()) {
        rw.getRange(i + 2, idx["Status"] + 1).setValue(String(status || "").trim().toUpperCase());

        // ✅ Sync actual computed NewExpiry back so RENEWALS matches SUBSCRIPTIONS
        if (actualNewExpiry instanceof Date && !isNaN(actualNewExpiry.getTime())) {
          if (idx["NewExpiry"] != null) {
            rw.getRange(i + 2, idx["NewExpiry"] + 1).setValue(actualNewExpiry);
          }
        }
        return;
      }
    }
  } catch (e) {}
}