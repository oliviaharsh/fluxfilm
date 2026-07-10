/*************** FluxFilm Backend — Customers + Profile + Coupons + Trending ***************
 * File: customers.gs
 *
 * Includes:
 *  - createOrUpdateCustomerProfile(payload)
 *  - createCustomerProfile(payload)
 *  - getCustomerProfile(phone)
 *  - updateCustomerProfilePic(phone, profilePicUrl)
 *  - getActiveCouponsForCustomer(phone)
 *  - getTrendingItems()
 ******************************************************************************************/

/**
 * Frontend account popup calls this:
 * createOrUpdateCustomerProfile({ name, phone, email })
 */
function createOrUpdateCustomerProfile(payload) {
  try {
    const p = payload || {};

    const name = String(p.name || "").trim();
    const phone = cus_normPhone_(p.phone);
    const email = String(p.email || "").trim().toLowerCase();

    if (!name) return bad_("Name is required.");
    if (!phone) return bad_("Phone number is required.");
    if (!email || email.indexOf("@") === -1) return bad_("Valid email is required.");

    const sh = cus_sheet_();
    cus_ensureCustomerHeaders_(sh);

    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h || "").trim());
    const idx = {};
    headers.forEach((h, i) => { if (h) idx[h] = i; });

    const now = cus_now_();
    let rowNum = cus_findCustomerRowByPhone_(sh, phone);

    // Existing customer → update
    if (rowNum) {
      cus_setByHeader_(sh, rowNum, idx, "Name", name);
      cus_setByHeader_(sh, rowNum, idx, "Phone", phone);
      cus_setByHeader_(sh, rowNum, idx, "Email", email);
      cus_setByHeader_(sh, rowNum, idx, "UpdatedAt", now);
      cus_setByHeader_(sh, rowNum, idx, "lastActivity", now);

      const oldStatus = cus_getByHeader_(sh, rowNum, idx, "Status");
      if (!oldStatus) cus_setByHeader_(sh, rowNum, idx, "Status", "ACTIVE");

      return ok_({
        message: "Account updated",
        profile: cus_readProfileFromRow_(sh, rowNum)
      });
    }

    // New customer → append row by headers
    const freshHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
    const freshIdx = {};
    freshHeaders.forEach((h, i) => { if (h) freshIdx[h] = i; });

    const row = new Array(sh.getLastColumn()).fill("");

    function put(h, v) {
      if (freshIdx[h] != null) row[freshIdx[h]] = v;
    }

    put("CustomerID", "CUS-" + Date.now().toString().slice(-8));
    put("MemberSince", now);
    put("UpdatedAt", now);
    put("Name", name);
    put("Phone", phone);
    put("Email", email);
    put("LastOrderID", "");
    put("TotalOrders", 0);
    put("TotalSpent", 0);
    put("lastActivity", now);
    put("Notes", "");
    put("Status", "ACTIVE");
    put("ProfilePicUrl", "");

    sh.appendRow(row);
    rowNum = sh.getLastRow();

    return ok_({
      message: "Account created",
      profile: cus_readProfileFromRow_(sh, rowNum)
    });

  } catch (e) {
    return bad_("createOrUpdateCustomerProfile error: " + cus_err_(e));
  }
}

/**
 * Alias, in case frontend / Hostinger calls this name.
 */
function createCustomerProfile(payload) {
  return createOrUpdateCustomerProfile(payload);
}

/**
 * Customer Profile — called by Account screen / login checks.
 */
function getCustomerProfile(phone) {
  try {
    const sheet = cus_sheet_();
    cus_ensureCustomerHeaders_(sheet);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok:false, message:"No customers found" };

    const headers = data[0].map(h => String(h || "").trim());
    const col = (names) => {
      for (const n of names) {
        const i = headers.indexOf(n);
        if (i !== -1) return i;
      }
      return -1;
    };

    const cPhone   = col(["Phone", "phone"]);
    const cName    = col(["Name", "CustomerName", "FullName"]);
    const cEmail   = col(["Email", "email"]);
    const cMember  = col(["MemberSince", "CreatedAt"]);
    const cUpdated = col(["UpdatedAt", "LastActivity", "lastActivity"]);
    const cOrders  = col(["TotalOrders"]);
    const cSpent   = col(["TotalSpent"]);
    const cPic     = col(["ProfilePicUrl", "AvatarUrl", "PhotoUrl"]);
    const cStatus  = col(["Status"]);

    const normPhone = cus_normPhone_(phone);
    if (!normPhone) return { ok:false, message:"Phone required" };
    if (cPhone === -1) return { ok:false, message:"CUSTOMERS missing Phone header" };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowPhone = cus_normPhone_(row[cPhone]);

      if (rowPhone === normPhone) {
        return {
          ok: true,
          customerId: col(["CustomerID"]) >= 0 ? String(row[col(["CustomerID"])] || "") : "",
          name: cName >= 0 ? String(row[cName] || "") : "",
          email: cEmail >= 0 ? String(row[cEmail] || "") : "",
          phone: row[cPhone] || normPhone,
          memberSince: cMember >= 0 ? cus_dateIso_(row[cMember]) : "",
          updatedAt: cUpdated >= 0 ? cus_dateIso_(row[cUpdated]) : "",
          lastActivity: cUpdated >= 0 ? cus_dateIso_(row[cUpdated]) : "",
          totalOrders: cOrders >= 0 ? Number(row[cOrders] || 0) : 0,
          totalSpent: cSpent >= 0 ? Number(row[cSpent] || 0) : 0,
          status: cStatus >= 0 ? String(row[cStatus] || "") : "",
          profilePicUrl: cPic >= 0 ? String(row[cPic] || "") : ""
        };
      }
    }

    return { ok:false, message:"Customer not found" };

  } catch (e) {
    return { ok:false, message:"getCustomerProfile error: " + cus_err_(e) };
  }
}

/**
 * Update customer profile image/avatar.
 */
function updateCustomerProfilePic(phone, profilePicUrl) {
  try {
    const ph = cus_normPhone_(phone);
    const url = String(profilePicUrl || "").trim();

    if (!ph) return { ok:false, message:"Phone required" };

    const sh = cus_sheet_();
    cus_ensureCustomerHeaders_(sh);

    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => String(h || "").trim());

    let cPhone = headers.indexOf("Phone");
    let cPic = headers.indexOf("ProfilePicUrl");
    let cUpdated = headers.indexOf("UpdatedAt");
    let cActivity = headers.indexOf("lastActivity");

    if (cPhone < 0) return { ok:false, message:"CUSTOMERS missing Phone header" };

    if (cPic < 0) {
      cPic = headers.length;
      sh.getRange(1, cPic + 1).setValue("ProfilePicUrl");
    }

    for (let i = 1; i < data.length; i++) {
      const rowPhone = cus_normPhone_(data[i][cPhone]);

      if (rowPhone === ph) {
        sh.getRange(i + 1, cPic + 1).setValue(url);

        if (cUpdated >= 0) sh.getRange(i + 1, cUpdated + 1).setValue(cus_now_());
        if (cActivity >= 0) sh.getRange(i + 1, cActivity + 1).setValue(cus_now_());

        return {
          ok:true,
          profilePicUrl:url,
          profile: cus_readProfileFromRow_(sh, i + 1)
        };
      }
    }

    return { ok:false, message:"Customer not found" };

  } catch(e) {
    return { ok:false, message: cus_err_(e) };
  }
}

/**
 * Coupons shown inside customer/account section.
 */
function getActiveCouponsForCustomer(phone) {
  try {
    const ph = cus_normPhone_(phone);
    if (!ph) return { ok:false, message:"Phone required" };

    const sh = sh_("COUPONS");
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok:true, coupons:[] };

    const headers = data[0].map(h => String(h || "").trim());
    const idx = {};
    headers.forEach((h,i) => idx[h] = i);

    const out = [];

    for (let i = 1; i < data.length; i++) {
      const r = data[i];

      if (idx.Active == null) continue;
      if (String(r[idx.Active] || "").toUpperCase() !== "TRUE") continue;

      const showInProfile = idx.ShowInProfile != null
        ? String(r[idx.ShowInProfile] || "").toUpperCase()
        : "TRUE";

      if (showInProfile !== "TRUE") continue;

      const code = String(
        (idx.CouponCode != null ? r[idx.CouponCode] : "") ||
        (idx.Code != null ? r[idx.Code] : "") ||
        ""
      ).trim().toUpperCase();

      if (!code) continue;

      const allowedRaw = idx.AllowedPhones != null
        ? String(r[idx.AllowedPhones] || "").trim()
        : "ALL";

      if (allowedRaw && allowedRaw.toUpperCase() !== "ALL") {
        const allowedPhones = allowedRaw
          .split(",")
          .map(x => cus_normPhone_(x))
          .filter(Boolean);

        if (!allowedPhones.includes(ph)) continue;
      }

      const perLimit = idx.PerUserLimit != null ? Number(r[idx.PerUserLimit] || 0) : 0;

      let used = 0;
      try {
        if (typeof countCouponUsageByPhone_ === "function") {
          used = countCouponUsageByPhone_(code, ph);
        }
      } catch(e) {
        used = 0;
      }

      if (perLimit > 0 && used >= perLimit) continue;

      out.push({
        code,
        description: idx.Description != null ? String(r[idx.Description] || "") : "",
        scope: idx.Scope != null ? String(r[idx.Scope] || "ANY") : "ANY",
        type: idx.Type != null ? String(r[idx.Type] || "") : "",
        value: idx.Value != null ? Number(r[idx.Value] || 0) : 0,
        minAmount: idx.MinAmount != null ? Number(r[idx.MinAmount] || 0) : 0,
        maxDiscount: idx.MaxDiscount != null ? Number(r[idx.MaxDiscount] || 0) : 0,
        expiry: idx.Expiry != null ? cus_dateIso_(r[idx.Expiry]) : "",
        perUserLimit: perLimit,
        usedByUser: used,
        remaining: perLimit > 0 ? Math.max(0, perLimit - used) : "Unlimited"
      });
    }

    return { ok:true, coupons: out };

  } catch(e) {
    return { ok:false, message: cus_err_(e) };
  }
}

/**
 * Trending OTT ticker items.
 */
function getTrendingItems() {
  try {
    const ss = getCore_();
    const sh = ss.getSheetByName("TRENDING");
    if (!sh) return ok_({ items: [] });

    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return ok_({ items: [] });

    const headers = vals[0].map(h => String(h || "").trim());
    const col = name => headers.indexOf(name);

    const cActive = col("Active");
    const cTitle = col("Title");
    const cPlatform = col("Platform");
    const cLine = col("Line");

    const out = [];

    for (let i = 1; i < vals.length; i++) {
      const r = vals[i];

      const active = cActive >= 0
        ? String(r[cActive] || "").toUpperCase() === "TRUE"
        : true;

      if (!active) continue;

      const title = cTitle >= 0 ? String(r[cTitle] || "").trim() : "";
      const platform = cPlatform >= 0 ? String(r[cPlatform] || "").trim() : "";
      const line = cLine >= 0 ? String(r[cLine] || "").trim() : "";

      if (!title && !line) continue;

      out.push(line || `🔥 ${title}${platform ? " on " + platform : ""}`);
    }

    return ok_({ items: out.slice(0, 8) });

  } catch(e) {
    return ok_({ items: [] });
  }
}

/**
 * Debug helpers.
 */
function testProfile(){
  Logger.log(JSON.stringify(getCustomerProfile("9818196079"), null, 2));
}

function DEBUG_getCustomerOrders() {
  const r = getCustomerOrders("9818196079", 15);
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}

function DEBUG_createCustomerProfile() {
  const r = createOrUpdateCustomerProfile({
    name: "Test User",
    phone: "9999999999",
    email: "test@example.com"
  });
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}

// ---------- Internal helpers ----------

function cus_sheet_() {
  const ss = typeof getCore_ === "function"
    ? getCore_()
    : SpreadsheetApp.getActiveSpreadsheet();

  let sh = ss.getSheetByName("CUSTOMERS");

  if (!sh) {
    sh = ss.insertSheet("CUSTOMERS");
    sh.appendRow([
      "CustomerID",
      "MemberSince",
      "UpdatedAt",
      "Name",
      "Phone",
      "Email",
      "LastOrderID",
      "TotalOrders",
      "TotalSpent",
      "lastActivity",
      "Notes",
      "Status",
      "ProfilePicUrl"
    ]);
  }

  return sh;
}

function cus_ensureCustomerHeaders_(sh) {
  const required = [
    "CustomerID",
    "MemberSince",
    "UpdatedAt",
    "Name",
    "Phone",
    "Email",
    "LastOrderID",
    "TotalOrders",
    "TotalSpent",
    "lastActivity",
    "Notes",
    "Status",
    "ProfilePicUrl"
  ];

  const lastCol = Math.max(1, sh.getLastColumn());
  let headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());

  // If first row is blank, set all headers
  const hasAnyHeader = headers.some(Boolean);
  if (!hasAnyHeader) {
    sh.getRange(1, 1, 1, required.length).setValues([required]);
    return;
  }

  required.forEach(h => {
    headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x || "").trim());
    if (!headers.includes(h)) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
    }
  });
}

function cus_normPhone_(phone) {
  let s = String(phone || "").replace(/\D/g, "");
  if (s.length > 10) s = s.slice(-10);
  return s;
}

function cus_now_() {
  try {
    if (typeof now_ === "function") return now_();
  } catch(e) {}
  return new Date();
}

function cus_findCustomerRowByPhone_(sh, phone) {
  const ph = cus_normPhone_(phone);
  if (!ph) return null;

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return null;

  const headers = data[0].map(h => String(h || "").trim());
  const cPhone = headers.indexOf("Phone");

  if (cPhone < 0) throw new Error("CUSTOMERS missing Phone header.");

  for (let i = 1; i < data.length; i++) {
    const rowPhone = cus_normPhone_(data[i][cPhone]);
    if (rowPhone === ph) return i + 1;
  }

  return null;
}

function cus_setByHeader_(sh, rowNum, idx, header, value) {
  if (idx[header] == null) return;
  sh.getRange(rowNum, idx[header] + 1).setValue(value);
}

function cus_getByHeader_(sh, rowNum, idx, header) {
  if (idx[header] == null) return "";
  return sh.getRange(rowNum, idx[header] + 1).getValue();
}

function cus_readProfileFromRow_(sh, rowNum) {
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h || "").trim());

  const col = (names) => {
    for (const n of names) {
      const i = headers.indexOf(n);
      if (i !== -1) return i;
    }
    return -1;
  };

  const row = data[rowNum - 1] || [];

  const cCustomer = col(["CustomerID"]);
  const cPhone    = col(["Phone", "phone"]);
  const cName     = col(["Name", "CustomerName", "FullName"]);
  const cEmail    = col(["Email", "email"]);
  const cMember   = col(["MemberSince", "CreatedAt"]);
  const cUpdated  = col(["UpdatedAt", "LastActivity", "lastActivity"]);
  const cOrders   = col(["TotalOrders"]);
  const cSpent    = col(["TotalSpent"]);
  const cNotes    = col(["Notes"]);
  const cStatus   = col(["Status"]);
  const cPic      = col(["ProfilePicUrl", "AvatarUrl", "PhotoUrl"]);
  const cLastOrd  = col(["LastOrderID"]);

  return {
    customerId: cCustomer >= 0 ? String(row[cCustomer] || "") : "",
    name: cName >= 0 ? String(row[cName] || "") : "",
    email: cEmail >= 0 ? String(row[cEmail] || "") : "",
    phone: cPhone >= 0 ? cus_normPhone_(row[cPhone]) : "",
    memberSince: cMember >= 0 ? cus_dateIso_(row[cMember]) : "",
    updatedAt: cUpdated >= 0 ? cus_dateIso_(row[cUpdated]) : "",
    lastActivity: cUpdated >= 0 ? cus_dateIso_(row[cUpdated]) : "",
    totalOrders: cOrders >= 0 ? Number(row[cOrders] || 0) : 0,
    totalSpent: cSpent >= 0 ? Number(row[cSpent] || 0) : 0,
    lastOrderId: cLastOrd >= 0 ? String(row[cLastOrd] || "") : "",
    notes: cNotes >= 0 ? String(row[cNotes] || "") : "",
    status: cStatus >= 0 ? String(row[cStatus] || "") : "",
    profilePicUrl: cPic >= 0 ? String(row[cPic] || "") : ""
  };
}

function cus_dateIso_(value) {
  try {
    if (!value) return "";
    const d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return "";
    return d.toISOString();
  } catch(e) {
    return "";
  }
}

function cus_err_(e) {
  return e && e.message ? e.message : String(e);
}










function testProfile(){
  Logger.log(getCustomerProfile("9818196079"));
}




function DEBUG_getCustomerOrders() {
  const r = getCustomerOrders("9818196079", 15);
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}