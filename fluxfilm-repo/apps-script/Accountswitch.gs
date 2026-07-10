/*************** FluxFilm Backend — Account Switch (Admin) ***************
 *
 * Lets an admin move a customer's subscription onto a DIFFERENT inventory
 * account, for cases like:
 *   - Netflix Sharing (Group offer + normal) → customer wants another account
 *   - Prime Video → customer wants another account
 *
 * Flow (two steps, mirrors the Quick Actions cards):
 *   1) asLookupSubForSwitch(subId)
 *        → returns the sub's name/plan + the account/login/pass/invRef
 *          currently in use, PLUS a dropdown list of OTHER accounts that
 *          still have capacity left (not full), each with used/total shown.
 *
 *   2) asSwitchAccount(subId, newAccountId)
 *        → reallocates: writes new LoginId, Password and InventoryRef onto
 *          the SUBSCRIPTIONS row for this SubID.
 *
 * IMPORTANT — register these in HostingerBridge.gs ACTIONS map:
 *     asLookupSubForSwitch: asLookupSubForSwitch,
 *     asSwitchAccount:      asSwitchAccount,
 *
 * Shared constants (TAB_SUBS, TAB_INV_ACCOUNTS, TAB_INV_CAPACITY) live in
 * code.gs — do NOT redeclare them here.
 ************************************************************************/


/**
 * STEP 1 — Look up a subscription and list switchable accounts.
 *
 * @param {string} subId  e.g. "SUB-12345678"
 * @return ok_({
 *   sub: { subId, name, phone, service, plan, currentAccountId,
 *          currentLogin, currentPass, currentInvRef },
 *   accounts: [ { accountId, login, used, total, free }, ... ]   // only non-full
 * })
 */
function asLookupSubForSwitch(subId) {
  try {
    const sid = String(subId || "").trim();
    if (!sid) return bad_("Enter a Subscription ID (SubID).");

    const subs = sh_(TAB_SUBS);
    const { idx: sIdx } = headerIndex_(subs);

    const need = ["SubID", "Service", "Plan", "LoginId", "Password", "InventoryRef", "Status"];
    for (var i = 0; i < need.length; i++) {
      if (sIdx[need[i]] == null) return bad_("SUBSCRIPTIONS missing header: " + need[i]);
    }

    // Find the subscription row
    const rn = getRowByValue_(subs, "SubID", sid);
    if (!rn) return bad_("Subscription not found: " + sid);

    const row = subs.getRange(rn, 1, 1, subs.getLastColumn()).getValues()[0];

    const service = String(row[sIdx["Service"]] || "").trim();
    const plan = String(row[sIdx["Plan"]] || "").trim();
    const name = (sIdx["Name"] != null) ? String(row[sIdx["Name"]] || "").trim() : "";
    const phone = (sIdx["Phone"] != null) ? String(row[sIdx["Phone"]] || "").trim() : "";
    const currentInvRef = String(row[sIdx["InventoryRef"]] || "").trim();
    const currentLogin = String(row[sIdx["LoginId"]] || "").trim();
    const currentPass = String(row[sIdx["Password"]] || "").trim();

    // Which service family are we switching within?
    const svcLower = service.toLowerCase();
    const isNetflix = svcLower.indexOf("netflix") >= 0;
    const isPrime = svcLower.indexOf("prime") >= 0;

    if (!isNetflix && !isPrime) {
      return bad_("Account switch currently supported only for Netflix and Prime Video.");
    }

    // Netflix: only Sharing / Group plans (Private uses profile rotation, not account switch)
    if (isNetflix) {
      const planLower = plan.toLowerCase();
      const isSharing = planLower.indexOf("sharing") >= 0 || planLower.indexOf("group") >= 0;
      if (!isSharing) {
        return bad_("Netflix switch is only for Sharing / Group plans. This sub is: " + plan);
      }
    }

    // InventoryRef for Netflix Sharing looks like "NFLX-D1#P1".
    // The account id is the part before "#".
    const currentAccountId = currentInvRef.indexOf("#") >= 0
      ? currentInvRef.split("#")[0].trim()
      : currentInvRef;

    // Build the list of candidate accounts with capacity info
    const family = isNetflix ? "netflix" : "prime";
    const accounts = as_listAvailableAccounts_(family, currentAccountId);

    return ok_({
      sub: {
        subId: sid,
        name: name,
        phone: phone,
        service: service,
        plan: plan,
        currentAccountId: currentAccountId,
        currentLogin: currentLogin,
        currentPass: currentPass,
        currentInvRef: currentInvRef
      },
      accounts: accounts
    });

  } catch (e) {
    return bad_("asLookupSubForSwitch error: " + (e && e.message ? e.message : e));
  }
}


/**
 * STEP 2 — Switch the subscription onto the chosen account.
 *
 * @param {string} subId         SubID to move
 * @param {string} newAccountId  AccountID chosen from the dropdown (e.g. "NFLX-D3")
 * @return ok_({ subId, newAccountId, newLogin, newInvRef, oldAccountId })
 */
function asSwitchAccount(subId, newAccountId) {
  try {
    const sid = String(subId || "").trim();
    const accId = String(newAccountId || "").trim();
    if (!sid) return bad_("Missing SubID.");
    if (!accId) return bad_("Choose an account to switch to.");

    const subs = sh_(TAB_SUBS);
    const { idx: sIdx } = headerIndex_(subs);

    const need = ["SubID", "Service", "Plan", "LoginId", "Password", "InventoryRef", "Status"];
    for (var i = 0; i < need.length; i++) {
      if (sIdx[need[i]] == null) return bad_("SUBSCRIPTIONS missing header: " + need[i]);
    }

    const rn = getRowByValue_(subs, "SubID", sid);
    if (!rn) return bad_("Subscription not found: " + sid);

    const row = subs.getRange(rn, 1, 1, subs.getLastColumn()).getValues()[0];

    const service = String(row[sIdx["Service"]] || "").trim();
    const plan = String(row[sIdx["Plan"]] || "").trim();
    const oldInvRef = String(row[sIdx["InventoryRef"]] || "").trim();
    const oldAccountId = oldInvRef.indexOf("#") >= 0 ? oldInvRef.split("#")[0].trim() : oldInvRef;

    const svcLower = service.toLowerCase();
    const isNetflix = svcLower.indexOf("netflix") >= 0;
    const isPrime = svcLower.indexOf("prime") >= 0;
    if (!isNetflix && !isPrime) {
      return bad_("Account switch supported only for Netflix / Prime Video.");
    }
    if (isNetflix) {
      const pl = plan.toLowerCase();
      if (pl.indexOf("sharing") < 0 && pl.indexOf("group") < 0) {
        return bad_("Netflix switch is only for Sharing / Group plans.");
      }
    }

    // Prevent a no-op switch to the same account
    if (accId === oldAccountId) {
      return bad_("That is the account already in use. Pick a different one.");
    }

    // Re-check the target account is valid, active, correct service and NOT full
    const family = isNetflix ? "netflix" : "prime";
    const available = as_listAvailableAccounts_(family, oldAccountId);
    var target = null;
    for (var j = 0; j < available.length; j++) {
      if (available[j].accountId === accId) { target = available[j]; break; }
    }
    if (!target) {
      return bad_("That account is not available (full, inactive, or wrong service). Refresh and try again.");
    }

    const newLogin = target.login;
    const newPass = target.pass;
    if (!newLogin || !newPass) {
      return bad_("Chosen account is missing login/password in INVENTORY_ACCOUNTS.");
    }

    // Build the new InventoryRef.
    //  - Netflix Sharing keeps the same shared profile number (#P<no>) as before,
    //    so the customer stays on the reserved sharing profile of the new account.
    //  - Prime uses the bare accountId as InventoryRef (matches allocator).
    var newInvRef = accId;
    if (isNetflix) {
      var profSuffix = "";
      if (oldInvRef.indexOf("#") >= 0) {
        profSuffix = oldInvRef.substring(oldInvRef.indexOf("#")); // e.g. "#P1"
      } else {
        var shareNo = asNumber_(getSetting_("NETFLIX_SHARING_PROFILE_NO", 1)) || 1;
        profSuffix = "#P" + shareNo;
      }
      newInvRef = accId + profSuffix;
    }

    // ---- Write the change onto the SUBSCRIPTIONS row ----
    subs.getRange(rn, sIdx["LoginId"] + 1).setValue(newLogin);
    subs.getRange(rn, sIdx["Password"] + 1).setValue(newPass);
    subs.getRange(rn, sIdx["InventoryRef"] + 1).setValue(newInvRef);

    if (sIdx["AdminNote"] != null) {
      subs.getRange(rn, sIdx["AdminNote"] + 1)
        .setValue("Account switched " + (oldAccountId || "?") + " → " + accId + " on " + new Date().toISOString());
    }
    if (sIdx["AdminUpdatedAt"] != null) {
      subs.getRange(rn, sIdx["AdminUpdatedAt"] + 1).setValue(now_());
    }

    // Telegram notify (best-effort)
    try {
      if (typeof notifyTelegram_ === "function") {
        notifyTelegram_("ADMIN_ACCOUNT_SWITCH", {
          subId: sid, service: service, plan: plan,
          from: oldAccountId, to: accId, login: newLogin
        });
      }
    } catch (e) {}

    return ok_({
      message: "Switched to " + accId + " successfully.",
      subId: sid,
      oldAccountId: oldAccountId,
      newAccountId: accId,
      newLogin: newLogin,
      newPass: newPass,
      newInvRef: newInvRef
    });

  } catch (e) {
    return bad_("asSwitchAccount error: " + (e && e.message ? e.message : e));
  }
}


/* ==================== internal helper ==================== */

/**
 * Lists active accounts for a service family with capacity that is NOT full.
 * Uses the pre-computed "Total" column in INVENTORY_ACCOUNTS for occupancy,
 * and INVENTORY_CAPACITY.MaxTotal for the cap. Falls back to MaxProfiles.
 *
 * @param {string} family        "netflix" | "prime"
 * @param {string} excludeAccId   account currently in use (still listed if not full,
 *                                 but excluded here so admin only sees alternatives)
 * @return Array<{ accountId, login, pass, used, total, free }>
 */
function as_listAvailableAccounts_(family, excludeAccId) {
  const invAcc = sh_(TAB_INV_ACCOUNTS);
  const cap = sh_(TAB_INV_CAPACITY);

  const { idx: aIdx } = headerIndex_(invAcc);
  const { idx: cIdx } = headerIndex_(cap);

  const fam = String(family || "").toLowerCase();

  // Default caps per family (used only if INVENTORY_CAPACITY has no row)
  const defaultCap = (fam === "netflix")
    ? asNumber_(getSetting_("NETFLIX_SHARING_MAX_TOTAL", 5))
    : asNumber_(getSetting_("PRIME_MAX_TOTAL", 4));

  // ---- Load capacity map: AccountID → { maxTotal, isActive } ----
  const capMap = {};
  const cLast = cap.getLastRow();
  if (cLast >= 2) {
    const cData = cap.getRange(2, 1, cLast - 1, cap.getLastColumn()).getValues();
    cData.forEach(function (r) {
      const svc = String(r[cIdx["Service"]] || "").trim().toLowerCase();
      if (svc.indexOf(fam) < 0) return;
      const accountId = String(r[cIdx["AccountID"]] || "").trim();
      if (!accountId) return;
      capMap[accountId] = {
        maxTotal: asNumber_(r[cIdx["MaxTotal"]]) || defaultCap,
        isActive: String(r[cIdx["IsActive"]] || "").trim().toUpperCase() === "TRUE"
      };
    });
  }

  // Occupancy comes from the pre-computed "Total" column (col L in INVENTORY_ACCOUNTS)
  const totalCol = aIdx["Total"];

  const aLast = invAcc.getLastRow();
  if (aLast < 2) return [];
  const aData = invAcc.getRange(2, 1, aLast - 1, invAcc.getLastColumn()).getValues();

  const out = [];
  for (var i = 0; i < aData.length; i++) {
    const r = aData[i];

    const svc = String(r[aIdx["Service"]] || "").trim().toLowerCase();
    if (svc.indexOf(fam) < 0) continue;

    const isActive = String(r[aIdx["IsActive"]] || "").trim().toUpperCase() === "TRUE";
    if (!isActive) continue;

    const accountId = String(r[aIdx["AccountID"]] || "").trim();
    if (!accountId) continue;
    if (excludeAccId && accountId === excludeAccId) continue; // show only alternatives

    const login = String(r[aIdx["LoginId"]] || "").trim();
    const pass = String(r[aIdx["Password"]] || "").trim();
    if (!login || !pass) continue;

    // capacity: prefer INVENTORY_CAPACITY, else MaxProfiles column, else default
    var capCfg = capMap[accountId];
    if (capCfg && capCfg.isActive === false) continue;
    var total = capCfg ? capCfg.maxTotal : (asNumber_(r[aIdx["MaxProfiles"]]) || defaultCap);

    // used: pre-computed "Total" column
    var used = (totalCol != null) ? (asNumber_(r[totalCol]) || 0) : 0;
    var free = total - used;

    // only accounts that are NOT full
    if (free <= 0) continue;

    out.push({
      accountId: accountId,
      login: login,
      pass: pass,
      used: used,
      total: total,
      free: free
    });
  }

  // Least-used first, so admin naturally picks the emptiest account
  out.sort(function (a, b) { return a.used - b.used; });

  return out;
}