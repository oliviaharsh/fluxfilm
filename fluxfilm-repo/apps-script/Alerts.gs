function debugSettings() {
  const out = {
    sheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
    sheetName: SpreadsheetApp.getActiveSpreadsheet().getName(),
    effectiveUser: Session.getEffectiveUser().getEmail(),
    UPI_VPA: getSetting_("UPI_VPA","(missing)"),
    UPI_PAYEE_NAME: getSetting_("UPI_PAYEE_NAME","(missing)"),
  };

  Logger.log(JSON.stringify(out, null, 2)); // also logs it
  return out;
}


function tg_debugGetUpdates(){
  const token = getSetting_("TELEGRAM_BOT_TOKEN", "");
  const url = "https://api.telegram.org/bot" + token + "/getUpdates";
  const res = UrlFetchApp.fetch(url).getContentText();
  Logger.log(res);
}



// ---- main entry ----
function zzz_lowInventoryAlert_() {
  var now = (typeof now_ === "function") ? now_() : new Date();
  var nowMs = now.getTime();

  var parts = [];
  parts.push("🕒 " + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd MMM yyyy, HH:mm"));

  parts.push("");
  parts.push(inv_dbgNetflixSharing_(nowMs));
  parts.push("");
  parts.push(inv_dbgNetflixPrivate_(nowMs));
  parts.push("");
  parts.push(inv_dbgPrime_(nowMs));

  // Telegram message cap safe
  var msg = parts.join("\n").slice(0, 3500);
  inv_dbgSend_("Inventory Debug", msg);
}


function sendDailyInventoryAlert() {
  var now = (typeof now_ === "function") ? now_() : new Date();
  var nowMs = now.getTime();

  var lines = [];
  lines.push("📦 FluxFilm Daily Inventory");
  lines.push("🕒 " + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd MMM yyyy, HH:mm"));
  lines.push("");

  // ===== NETFLIX SHARING =====
  lines.push("🎬 Netflix Sharing");
  try {
    var invAcc = sh_(TAB_INV_ACCOUNTS);
    var cap    = sh_(TAB_INV_CAPACITY);
    var subs   = sh_(TAB_SUBS);

    var aIdx = headerIndex_(invAcc).idx;
    var cIdx = headerIndex_(cap).idx;
    var sIdx = headerIndex_(subs).idx;

    var maxTotalDefault = Number(getSetting_("NETFLIX_SHARING_MAX_TOTAL", 5)) || 5;

    // Build cap map
    var capMap = {};
    var cLast = cap.getLastRow();
    if (cLast >= 2) {
      cap.getRange(2, 1, cLast - 1, cap.getLastColumn()).getValues().forEach(function(r) {
        var svc = String(r[cIdx["Service"]] || "").toLowerCase();
        if (svc.indexOf("netflix") < 0) return;
        var aid = String(r[cIdx["AccountID"]] || "").trim();
        if (!aid) return;
        capMap[aid] = {
          isActive: String(r[cIdx["IsActive"]] || "TRUE").toUpperCase().trim() === "TRUE",
          maxTotal: Number(r[cIdx["MaxTotal"]] || maxTotalDefault) || maxTotalDefault
        };
      });
    }

    // Build subs data
    var sLast = subs.getLastRow();
    var sData = sLast < 2 ? [] : subs.getRange(2, 1, sLast - 1, subs.getLastColumn()).getValues();

    function countSharingUsed(invRef) {
      var count = 0;
      sData.forEach(function(r) {
        var svc = String(r[sIdx["Service"]] || "").toLowerCase();
        if (svc.indexOf("netflix") < 0) return;
        if (String(r[sIdx["InventoryRef"]] || "").trim() !== invRef) return;
        var st = String(r[sIdx["Status"]] || "").toUpperCase().trim();
        if (st !== "ACTIVE") return;
        var exp = r[sIdx["ExpiryDate"]];
        var expMs = exp instanceof Date ? exp.getTime() : new Date(exp).getTime();
        if (!isNaN(expMs) && nowMs < expMs) count++;
      });
      return count;
    }

    var aLast = invAcc.getLastRow();
    var aData = aLast < 2 ? [] : invAcc.getRange(2, 1, aLast - 1, invAcc.getLastColumn()).getValues();

    var nsGrandMax = 0, nsGrandUsed = 0;

    aData.forEach(function(ar) {
      var svc = String(ar[aIdx["Service"]] || "").toLowerCase();
      if (svc.indexOf("netflix") < 0) return;
      if (aIdx["IsActive"] == null) return;

      var accActive = String(ar[aIdx["IsActive"]] || "").toUpperCase().trim() === "TRUE";
      if (!accActive) return;
      var accountId = String(ar[aIdx["AccountID"]] || "").trim();
      if (!accountId) return;
      var capCfg = capMap[accountId] || { isActive: true, maxTotal: maxTotalDefault };
      if (!capCfg.isActive) return;

      var invRef = accountId + "#P1";
      var used = countSharingUsed(invRef);
      var avail = Math.max(0, capCfg.maxTotal - used);
      nsGrandMax += capCfg.maxTotal;
      nsGrandUsed += used;

      var status = used >= capCfg.maxTotal ? "🔴 FULL" : avail === capCfg.maxTotal ? "🟢 EMPTY" : "🟡 PARTIAL";
      lines.push("  " + accountId + ": " + used + "/" + capCfg.maxTotal + " " + status);
    });

    lines.push("  TOTAL: " + nsGrandUsed + "/" + nsGrandMax + " used, " + (nsGrandMax - nsGrandUsed) + " free");
  } catch(e) {
    lines.push("  ⚠️ Error: " + e.message);
  }

  lines.push("");

// ===== NETFLIX PRIVATE =====
lines.push("🔒 Netflix Private");
try {
  var invProf = sh_(TAB_INV_PROFILES);
  var invAccP = sh_(TAB_INV_ACCOUNTS);

  var pIdx = headerIndex_(invProf).idx;
  var aIdxP = headerIndex_(invAccP).idx;

  var sharingNo = Number(getSetting_("NETFLIX_SHARING_PROFILE_NO", 1)) || 1;

  // ✅ Build active Netflix account map
  var activeAccMap = {};
  var aLastP = invAccP.getLastRow();
  if (aLastP >= 2) {
    invAccP.getRange(2, 1, aLastP - 1, invAccP.getLastColumn()).getValues().forEach(function(ar) {
      var svc = String(ar[aIdxP["Service"]] || "").toLowerCase();
      if (svc.indexOf("netflix") < 0) return;

      var accountId = String(ar[aIdxP["AccountID"]] || "").trim();
      if (!accountId) return;

      var isActive = String(ar[aIdxP["IsActive"]] || "TRUE").toUpperCase().trim() === "TRUE";
      if (!isActive) return;

      activeAccMap[accountId] = true;
    });
  }

  var pLast = invProf.getLastRow();
  var pData = pLast < 2 ? [] : invProf.getRange(2, 1, pLast - 1, invProf.getLastColumn()).getValues();

  var privMap = {}; // accountId -> { used, free, total }

  pData.forEach(function(r) {
    var svc = String(r[pIdx["Service"]] || "").toLowerCase();
    if (svc.indexOf("netflix") < 0) return;

    var type = String(r[pIdx["ProfileType"]] || "").toUpperCase().trim();
    if (type !== "PRIVATE_ROTATING") return;

    var pNo = Number(r[pIdx["ProfileNumber"]] || 0);
    if (pNo === sharingNo) return;

    var accountId = String(r[pIdx["AccountID"]] || "").trim();
    if (!accountId) return;
    if (!activeAccMap[accountId]) return;

    // ✅ Skip profiles if parent account is inactive
    if (!activeAccMap[accountId]) return;

    if (!privMap[accountId]) privMap[accountId] = { used: 0, free: 0, total: 0 };

    privMap[accountId].total++;

    var status = String(r[pIdx["Status"]] || "FREE").toUpperCase().trim();

    if (status === "FREE") {
      privMap[accountId].free++;
    } else if (status === "ASSIGNED") {
      privMap[accountId].used++;
    } else if (status === "COOLDOWN") {
      var relMs2 = inv_toMs_(r[pIdx["ReleaseEligibleAt"]]);
      if (!isNaN(relMs2) && nowMs >= relMs2) privMap[accountId].free++;
      else privMap[accountId].used++;
    }
  });

  var npGrandTotal = 0, npGrandUsed = 0;

  Object.keys(privMap).sort().forEach(function(acc) {
    var o = privMap[acc];
    var status = o.free === 0 ? "🔴 FULL" : o.used === 0 ? "🟢 EMPTY" : "🟡 PARTIAL";
    lines.push("  " + acc + ": " + o.used + "/" + o.total + " used, " + o.free + " free " + status);
    npGrandTotal += o.total;
    npGrandUsed += o.used;
  });

  lines.push("  TOTAL: " + npGrandUsed + "/" + npGrandTotal + " used, " + (npGrandTotal - npGrandUsed) + " free");
} catch(e) {
  lines.push("  ⚠️ Error: " + e.message);
}

  lines.push("");

  // ===== PRIME =====
  lines.push("🛒 Prime");
  try {
    var invAcc2  = sh_(TAB_INV_ACCOUNTS);
    var cap2     = sh_(TAB_INV_CAPACITY);
    var subs2    = sh_(TAB_SUBS);

    var aIdx2 = headerIndex_(invAcc2).idx;
    var cIdx2 = headerIndex_(cap2).idx;
    var sIdx2 = headerIndex_(subs2).idx;

    var prMaxTotalDefault = Number(getSetting_("PRIME_MAX_TOTAL", 4)) || 4;
    var prMaxTVDefault    = Number(getSetting_("PRIME_MAX_TV", 2)) || 2;

    var prCapMap = {};
    var cLast2 = cap2.getLastRow();
    if (cLast2 >= 2) {
      cap2.getRange(2, 1, cLast2 - 1, cap2.getLastColumn()).getValues().forEach(function(r) {
        var svc = String(r[cIdx2["Service"]] || "").toLowerCase();
        if (svc.indexOf("prime") < 0) return;
        var aid = String(r[cIdx2["AccountID"]] || "").trim();
        if (!aid) return;
        prCapMap[aid] = {
          isActive: String(r[cIdx2["IsActive"]] || "TRUE").toUpperCase().trim() === "TRUE",
          maxTotal: Number(r[cIdx2["MaxTotal"]] || prMaxTotalDefault) || prMaxTotalDefault,
          maxTV:    Number(r[cIdx2["MaxTV"]]    || prMaxTVDefault)    || prMaxTVDefault
        };
      });
    }

    var sLast2 = subs2.getLastRow();
    var sData2 = sLast2 < 2 ? [] : subs2.getRange(2, 1, sLast2 - 1, subs2.getLastColumn()).getValues();

    function countPrimeOcc(accountId) {
      var total = 0, tv = 0;
      sData2.forEach(function(r) {
        var svc = String(r[sIdx2["Service"]] || "").toLowerCase();
        if (svc.indexOf("prime") < 0) return;
        if (String(r[sIdx2["InventoryRef"]] || "").trim() !== accountId) return;
        var st = String(r[sIdx2["Status"]] || "").toUpperCase().trim();
        if (st !== "ACTIVE") return;
        var exp = r[sIdx2["ExpiryDate"]];
        var expMs = exp instanceof Date ? exp.getTime() : new Date(exp).getTime();
        if (isNaN(expMs) || nowMs >= expMs) return;
        total++;
        if (String(r[sIdx2["DeviceType"]] || "").toUpperCase().trim() === "TV") tv++;
      });
      return { total: total, tv: tv };
    }

    var aLast2 = invAcc2.getLastRow();
    var aData2 = aLast2 < 2 ? [] : invAcc2.getRange(2, 1, aLast2 - 1, invAcc2.getLastColumn()).getValues();

    var prGrandMax = 0, prGrandUsed = 0;

    aData2.forEach(function(ar) {
      var svc = String(ar[aIdx2["Service"]] || "").toLowerCase();
      if (svc.indexOf("prime") < 0) return;
      if (aIdx2["IsActive"] == null) return;

      var accActive2 = String(ar[aIdx2["IsActive"]] || "").toUpperCase().trim() === "TRUE";
      if (!accActive2) return;
      var accountId = String(ar[aIdx2["AccountID"]] || "").trim();
      if (!accountId) return;
      var capCfg = prCapMap[accountId] || { isActive: true, maxTotal: prMaxTotalDefault, maxTV: prMaxTVDefault };
      if (!capCfg.isActive) return;

      var o = countPrimeOcc(accountId);
      var avail = Math.max(0, capCfg.maxTotal - o.total);
      prGrandMax += capCfg.maxTotal;
      prGrandUsed += o.total;

      var statusIcon = o.total >= capCfg.maxTotal ? "🔴 FULL" : o.total === 0 ? "🟢 EMPTY" : "🟡 PARTIAL";
      lines.push("  " + accountId + ": " + o.total + "/" + capCfg.maxTotal + " (" + o.tv + " TV) " + statusIcon);
    });

    lines.push("  TOTAL: " + prGrandUsed + "/" + prGrandMax + " used, " + (prGrandMax - prGrandUsed) + " free");
  } catch(e) {
    lines.push("  ⚠️ Error: " + e.message);
  }

  // Send
  var msg = lines.join("\n").slice(0, 4000);
  try {
    if (typeof ffTgSend_ === "function") {
      ffTgSend_(msg);
    } else {
      var token = getSetting_("TELEGRAM_BOT_TOKEN", "");
      var chatId = getSetting_("TELEGRAM_CHAT_ID", "");
      UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendMessage", {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({ chat_id: chatId, text: msg }),
        muteHttpExceptions: true
      });
    }
  } catch(e) {}
}

function setupDailyInventoryAlertTrigger() {
  // Remove existing
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "sendDailyInventoryAlert") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Daily at 9 AM
  ScriptApp.newTrigger("sendDailyInventoryAlert")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  return "✅ Daily inventory alert trigger set for 9 AM";
}












function TEST_D6(){
  var sh = sh_(TAB_INV_ACCOUNTS);
  var idx = headerIndex_(sh).idx;

  var data = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();

  data.forEach(function(r){
    if(String(r[idx["AccountID"]] || "").trim() === "NFLX-D6"){
      Logger.log("AccountID = " + r[idx["AccountID"]]);
      Logger.log("IsActive RAW = " + r[idx["IsActive"]]);
      Logger.log("Parsed = " + (String(r[idx["IsActive"]] || "").toUpperCase().trim() === "TRUE"));
    }
  });
}



