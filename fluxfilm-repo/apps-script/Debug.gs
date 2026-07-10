/*************** FluxFilm Backend — Inventory Debug + Breakdown (Telegram) ***************/

/**
 * Sends a detailed inventory breakdown to Telegram so you can verify:
 * - Netflix Sharing slots per account (max/used/avail + which subs counted)
 * - Netflix Private profiles per account (free / cooldown done / total)
 * - Prime capacity per account (total + TV)
 *
 * Depends on:
 *  - sh_(), headerIndex_(), getSetting_()
 *  - notifyTelegram_() OR ffTgSend_() (it will use notifyTelegram_ if present)
 *  - TAB_INV_ACCOUNTS, TAB_INV_PROFILES, TAB_INV_CAPACITY, TAB_SUBS
 *  - Optional now_()
 */

// ---- send helper (uses your existing telegram notify flow if available) ----
function inv_dbgSend_(title, msg) {
  try {
    if (typeof notifyTelegram_ === "function") {
      notifyTelegram_("INV_DEBUG", { note: "📦 " + title + "\n\n" + msg });
      return;
    }
  } catch (e) {}
  try {
    if (typeof ffTgSend_ === "function") {
      ffTgSend_("📦 " + title + "\n\n" + msg);
      return;
    }
  } catch (e2) {}
}


// ===================== Netflix Sharing Debug =====================

function inv_dbgNetflixSharing_(nowMs) {
  var invAcc  = sh_(TAB_INV_ACCOUNTS);
  var invProf = sh_(TAB_INV_PROFILES);
  var cap     = sh_(TAB_INV_CAPACITY);
  var subs    = sh_(TAB_SUBS);

  var aIdx = headerIndex_(invAcc).idx;
  var pIdx = headerIndex_(invProf).idx;
  var cIdx = headerIndex_(cap).idx;
  var sIdx = headerIndex_(subs).idx;

  var sharingNo = Number(getSetting_("NETFLIX_SHARING_PROFILE_NO", 1)) || 1;
  var maxTotalDefault = Number(getSetting_("NETFLIX_SHARING_MAX_TOTAL", 4)) || 4;

  // cap map (Netflix)
  var capMap = {};
  var cLast = cap.getLastRow();
  if (cLast >= 2) {
    var cData = cap.getRange(2, 1, cLast - 1, cap.getLastColumn()).getValues();
    for (var i=0; i<cData.length; i++){
      var r = cData[i];
      var svc = String(r[cIdx["Service"]] || "").toLowerCase();
      if (svc.indexOf("netflix") < 0) continue;

      var accountId = String(r[cIdx["AccountID"]] || "").trim();
      if (!accountId) continue;

      capMap[accountId] = {
        isActive: String(r[cIdx["IsActive"]] || "TRUE").toUpperCase().trim() === "TRUE",
        maxTotal: Number(r[cIdx["MaxTotal"]] || maxTotalDefault) || maxTotalDefault
      };
    }
  }

  var pLast = invProf.getLastRow();
  var pData = (pLast < 2) ? [] : invProf.getRange(2, 1, pLast - 1, invProf.getLastColumn()).getValues();

  var sLast = subs.getLastRow();
  var sData = (sLast < 2) ? [] : subs.getRange(2, 1, sLast - 1, subs.getLastColumn()).getValues();

  function reservedProfileNo(accountId){
    var reservedPNo = null;

    for (var k=0; k<pData.length; k++){
      var pr = pData[k];
      if (String(pr[pIdx["AccountID"]] || "").trim() !== accountId) continue;

      var type = String(pr[pIdx["ProfileType"]] || "").toUpperCase().trim();
      var isRes = String(pr[pIdx["IsReserved"]] || "").toUpperCase().trim() === "TRUE";
      var pNo = Number(pr[pIdx["ProfileNumber"]] || 0);

      if (type === "SHARING_RESERVED" || type === "SHARING_RES" || isRes) {
        if (pNo === sharingNo) { reservedPNo = pNo; break; }
        if (reservedPNo == null) reservedPNo = pNo;
      }

      if (reservedPNo == null && pNo === sharingNo) reservedPNo = pNo;
    }

    return reservedPNo;
  }

  function activeSubsForInvRef(invRef){
    var list = [];
    for (var j=0; j<sData.length; j++){
      var sr = sData[j];

      var svc = String(sr[sIdx["Service"]] || "").toLowerCase();
      if (svc.indexOf("netflix") < 0) continue;

      if (String(sr[sIdx["InventoryRef"]] || "").trim() !== invRef) continue;

      var st = String(sr[sIdx["Status"]] || "").toUpperCase().trim();
      if (st === "FAILED" || st === "CANCELLED" || st === "REFUNDED") continue;

      var relMs = inv_toMs_(sr[sIdx["ReleaseEligibleAt"]]);
      if (!isNaN(relMs) && nowMs < relMs) {
        list.push({
          subId: String(sr[sIdx["SubID"]] || ""),
          status: st,
          rel: sr[sIdx["ReleaseEligibleAt"]]
        });
      }
    }
    return list;
  }

  var aLast = invAcc.getLastRow();
  if (aLast < 2) return "🎬 Netflix Sharing: (no accounts)";

  var aData = invAcc.getRange(2, 1, aLast - 1, invAcc.getLastColumn()).getValues();

  var lines = [];
  lines.push("🎬 Netflix Sharing (per account)");

  var grandAvail = 0, grandUsed = 0, grandMax = 0;
  for (var i=0; i<aData.length; i++){
    var ar = aData[i];
    var svc = String(ar[aIdx["Service"]] || "").toLowerCase();
    if (svc.indexOf("netflix") < 0) continue;

    var active = String(ar[aIdx["IsActive"]] || "TRUE").toUpperCase().trim() === "TRUE";
    if (!active) continue;

    var accountId = String(ar[aIdx["AccountID"]] || "").trim();
    if (!accountId) continue;

    var capCfg = capMap[accountId] || { isActive:true, maxTotal:maxTotalDefault };
    if (!capCfg.isActive) continue;

    var rNo = reservedProfileNo(accountId);
    if (!rNo) {
      lines.push("• " + accountId + ": ❌ reserved profile not found");
      continue;
    }

    var invRef = accountId + "#P" + rNo;
    var act = activeSubsForInvRef(invRef);

    var used = act.length;
    var maxT = capCfg.maxTotal;
    var avail = Math.max(0, maxT - used);

    grandAvail += avail; grandUsed += used; grandMax += maxT;

    lines.push("• " + accountId + " (" + invRef + "): max " + maxT + ", used " + used + ", avail " + avail);

    // show sub IDs (small)
    if (act.length) {
      var subLine = act.slice(0, 8).map(function(x){ return x.subId || "SUB?"; }).join(", ");
      if (act.length > 8) subLine += " …";
      lines.push("   ↳ subs: " + subLine);
    }
  }

  lines.push("—");
  lines.push("TOTAL: max " + grandMax + ", used " + grandUsed + ", avail " + grandAvail);
  return lines.join("\n");
}


// ===================== Netflix Private Debug =====================

function inv_dbgNetflixPrivate_(nowMs) {
  var invProf = sh_(TAB_INV_PROFILES);
  var idx = headerIndex_(invProf).idx;

  var last = invProf.getLastRow();
  if (last < 2) return "🔒 Netflix Private: (no profiles)";

  var sharingNo = Number(getSetting_("NETFLIX_SHARING_PROFILE_NO", 1)) || 1;

  var data = invProf.getRange(2, 1, last - 1, invProf.getLastColumn()).getValues();

  // group by AccountID
  var map = {}; // accountId -> {free, cooldown, total}
  for (var i=0; i<data.length; i++){
    var r = data[i];

    var accountId = String(r[idx["AccountID"]] || "").trim();
    if (!accountId) continue;

    var pNo = Number(r[idx["ProfileNumber"]] || 0);
    if (pNo === sharingNo) continue;

    var type = String(r[idx["ProfileType"]] || "").toUpperCase().trim();
    if (type !== "PRIVATE_ROTATING") continue;

    if (!map[accountId]) map[accountId] = { free:0, cooldown:0, total:0 };

    map[accountId].total++;

    var status = String(r[idx["Status"]] || "FREE").toUpperCase().trim();
    if (status === "FREE") { map[accountId].free++; continue; }

    var relMs = inv_toMs_(r[idx["ReleaseEligibleAt"]]);
    if (!isNaN(relMs) && nowMs >= relMs) map[accountId].cooldown++;
  }

  var lines = [];
  lines.push("🔒 Netflix Private (per account)");
  var totalAvail = 0;

  var keys = Object.keys(map).sort();
  for (var k=0; k<keys.length; k++){
    var acc = keys[k];
    var o = map[acc];
    var avail = o.free + o.cooldown;
    totalAvail += avail;
    lines.push("• " + acc + ": total " + o.total + ", free " + o.free + ", cooldownDone " + o.cooldown + ", avail " + avail);
  }

  lines.push("—");
  lines.push("TOTAL private avail: " + totalAvail);
  return lines.join("\n");
}


// ===================== Prime Debug =====================

function inv_dbgPrime_(nowMs) {
  var invAcc = sh_(TAB_INV_ACCOUNTS);
  var cap    = sh_(TAB_INV_CAPACITY);
  var subs   = sh_(TAB_SUBS);

  var aIdx = headerIndex_(invAcc).idx;
  var cIdx = headerIndex_(cap).idx;
  var sIdx = headerIndex_(subs).idx;

  var maxTotalDefault = Number(getSetting_("PRIME_MAX_TOTAL", 4)) || 4;
  var maxTVDefault    = Number(getSetting_("PRIME_MAX_TV", 2)) || 2;

  var capMap = {};
  var cLast = cap.getLastRow();
  if (cLast >= 2) {
    var cData = cap.getRange(2, 1, cLast - 1, cap.getLastColumn()).getValues();
    for (var i=0; i<cData.length; i++){
      var r = cData[i];
      var svc = String(r[cIdx["Service"]] || "").toLowerCase();
      if (svc.indexOf("prime") < 0) continue;

      var accountId = String(r[cIdx["AccountID"]] || "").trim();
      if (!accountId) continue;

      capMap[accountId] = {
        isActive: String(r[cIdx["IsActive"]] || "TRUE").toUpperCase().trim() === "TRUE",
        maxTotal: Number(r[cIdx["MaxTotal"]] || maxTotalDefault) || maxTotalDefault,
        maxTV:    Number(r[cIdx["MaxTV"]] || maxTVDefault) || maxTVDefault
      };
    }
  }

  var sLast = subs.getLastRow();
  var sData = (sLast < 2) ? [] : subs.getRange(2, 1, sLast - 1, subs.getLastColumn()).getValues();

  function occ(accountId){
    var total = 0, tv = 0;
    for (var j=0; j<sData.length; j++){
      var sr = sData[j];
      var svc = String(sr[sIdx["Service"]] || "").toLowerCase();
      if (svc.indexOf("prime") < 0) continue;

      if (String(sr[sIdx["InventoryRef"]] || "").trim() !== accountId) continue;

      var st = String(sr[sIdx["Status"]] || "").toUpperCase().trim();
      if (st === "FAILED" || st === "CANCELLED" || st === "REFUNDED") continue;

      var relMs = inv_toMs_(sr[sIdx["ReleaseEligibleAt"]]);
      if (!isNaN(relMs) && nowMs < relMs) {
        total++;
        if (String(sr[sIdx["DeviceType"]] || "").toUpperCase().trim() === "TV") tv++;
      }
    }
    return { total:total, tv:tv };
  }

  var aLast = invAcc.getLastRow();
  if (aLast < 2) return "🛒 Prime: (no accounts)";

  var aData = invAcc.getRange(2, 1, aLast - 1, invAcc.getLastColumn()).getValues();

  var lines = [];
  lines.push("🛒 Prime Capacity (per account)");

  var gMaxT=0, gUsedT=0, gAvailT=0;
  var gMaxTV=0, gUsedTV=0, gAvailTV=0;

  for (var i=0; i<aData.length; i++){
    var ar = aData[i];
    var svc = String(ar[aIdx["Service"]] || "").toLowerCase();
    if (svc.indexOf("prime") < 0) continue;

    var active = String(ar[aIdx["IsActive"]] || "TRUE").toUpperCase().trim() === "TRUE";
    if (!active) continue;

    var accountId = String(ar[aIdx["AccountID"]] || "").trim();
    if (!accountId) continue;

    var capCfg = capMap[accountId] || { isActive:true, maxTotal:maxTotalDefault, maxTV:maxTVDefault };
    if (!capCfg.isActive) continue;

    var o = occ(accountId);

    var availT = Math.max(0, capCfg.maxTotal - o.total);
    var availTV = Math.max(0, capCfg.maxTV - o.tv);

    gMaxT += capCfg.maxTotal; gUsedT += o.total; gAvailT += availT;
    gMaxTV += capCfg.maxTV; gUsedTV += o.tv; gAvailTV += availTV;

    lines.push("• " + accountId + ": max " + capCfg.maxTotal + ", used " + o.total + ", avail " + availT +
               " | TV max " + capCfg.maxTV + ", used " + o.tv + ", avail " + availTV);
  }

  lines.push("—");
  lines.push("TOTAL: max " + gMaxT + ", used " + gUsedT + ", avail " + gAvailT);
  lines.push("TV TOTAL: max " + gMaxTV + ", used " + gUsedTV + ", avail " + gAvailTV);
  return lines.join("\n");
}


// ===================== Shared helper =====================

function inv_toMs_(v){
  if (!v) return NaN;
  if (v instanceof Date) return v.getTime();
  var d = new Date(v);
  var t = d.getTime();
  return isNaN(t) ? NaN : t;
}

function runLowInventoryAlert() {
  zzz_lowInventoryAlert_();
}



// ---- Trigger-safe wrapper (DO NOT DELETE) ----
function runLowInventoryAlert() {
  zzz_lowInventoryAlert_();
}


function debugNetflixSharingCapacity() {

  const subs = sh_("SUBSCRIPTIONS");
  const acc = sh_("INVENTORY_ACCOUNTS");
  const prof = sh_("INVENTORY_PROFILES");
  const cap = sh_("INVENTORY_CAPACITY");

  const { idx: sIdx } = headerIndex_(subs);
  const { idx: aIdx } = headerIndex_(acc);
  const { idx: pIdx } = headerIndex_(prof);
  const { idx: cIdx } = headerIndex_(cap);

  const subsData = subs.getRange(2,1,subs.getLastRow()-1,subs.getLastColumn()).getValues();
  const accData = acc.getRange(2,1,acc.getLastRow()-1,acc.getLastColumn()).getValues();
  const profData = prof.getRange(2,1,prof.getLastRow()-1,prof.getLastColumn()).getValues();
  const capData = cap.getRange(2,1,cap.getLastRow()-1,cap.getLastColumn()).getValues();

  Logger.log("===== NETFLIX SHARING DEBUG =====");

  accData.forEach(a => {

    const service = String(a[aIdx["Service"]] || "");
    if (!service.includes("Netflix")) return;

    const accountId = a[aIdx["AccountID"]];
    const login = a[aIdx["LoginId"]];

    let maxTotal = 4;

    capData.forEach(c=>{
      if(c[cIdx["AccountID"]] == accountId){
        maxTotal = Number(c[cIdx["MaxTotal"]] || maxTotal);
      }
    });

    let used = 0;
    let subList = [];

    subsData.forEach(s=>{

      const inv = String(s[sIdx["InventoryRef"]] || "");
      const status = String(s[sIdx["Status"]] || "");

      if(inv === accountId+"#P1" && status === "ACTIVE"){
        used++;
        subList.push(s[sIdx["SubID"]]);
      }

    });

    Logger.log(
      accountId +
      " ("+login+") | max:"+maxTotal+
      " used:"+used+
      " avail:"+(maxTotal-used)+
      " | subs:"+subList.join(", ")
    );

  });

}














function debugOtpAllocation(){
  const r = allocateOtpAccount_("JioHotstar", "1 Month", 30);
  Logger.log(JSON.stringify(r));
  
  // Also check what's occupying it
  const occupied = otp_occupiedAccountIds_("JioHotstar");
  Logger.log("Occupied IDs: " + JSON.stringify([...occupied]));
}






