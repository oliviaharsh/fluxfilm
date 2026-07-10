function autoReleaseInventory() {
  const now = new Date();

  // 1) Release Netflix PRIVATE_ROTATING profiles
  const invProf = sh_(TAB_INV_PROFILES);
  const hi = headerIndex_(invProf);
  const idx = hi.idx;

  const last = invProf.getLastRow();
  if (last >= 2) {
    const data = invProf.getRange(2, 1, last - 1, invProf.getLastColumn()).getValues();

    for (let i = 0; i < data.length; i++) {
      const r = data[i];

      const svc = String(r[idx["Service"]] || "").trim().toLowerCase();
      if (!svc.includes("netflix")) continue;

      const type = String(r[idx["ProfileType"]] || "").trim().toUpperCase();
      if (type !== "PRIVATE_ROTATING") continue;

      const status = String(r[idx["Status"]] || "").trim().toUpperCase();
      if (status !== "ASSIGNED" && status !== "COOLDOWN") continue;

      const rel = r[idx["ReleaseEligibleAt"]];
      const relMs = rel instanceof Date ? rel.getTime() : new Date(rel).getTime();
      if (isNaN(relMs)) continue;

      if (relMs <= now.getTime()) {
        const rowNum = i + 2;
        invProf.getRange(rowNum, idx["Status"] + 1).setValue("FREE");
        invProf.getRange(rowNum, idx["CurrentSubID"] + 1).setValue("");

        if (idx["AssignedAt"] != null) invProf.getRange(rowNum, idx["AssignedAt"] + 1).setValue("");
        if (idx["ExpiryDate"] != null) invProf.getRange(rowNum, idx["ExpiryDate"] + 1).setValue("");
        if (idx["ReleaseEligibleAt"] != null) invProf.getRange(rowNum, idx["ReleaseEligibleAt"] + 1).setValue("");
      }
    }
  }

  // 2) Optional: mark SUBSCRIPTIONS expired
  const subs = sh_(TAB_SUBS);
  const hs = headerIndex_(subs);
  const sIdx = hs.idx;

  if (sIdx["ExpiryDate"] != null && sIdx["Status"] != null) {
    const sLast = subs.getLastRow();
    if (sLast >= 2) {
      const sData = subs.getRange(2, 1, sLast - 1, subs.getLastColumn()).getValues();

      for (let i = 0; i < sData.length; i++) {
        const exp = sData[i][sIdx["ExpiryDate"]];
        const expMs = exp instanceof Date ? exp.getTime() : new Date(exp).getTime();
        if (isNaN(expMs)) continue;

        if (expMs < now.getTime()) {
          subs.getRange(i + 2, sIdx["Status"] + 1).setValue("EXPIRED");
        }
      }
    }
  }

  return "autoReleaseInventory done ✅";
}

function setupAutoReleaseTrigger() {
  // delete existing triggers for same function
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "autoReleaseInventory") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // daily at 3 AM
  ScriptApp.newTrigger("autoReleaseInventory")
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  return "Daily trigger created ✅";
}
