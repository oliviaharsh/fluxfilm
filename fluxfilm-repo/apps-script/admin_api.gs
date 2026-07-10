/*************** FluxFilm Backend — Admin Dashboard API (admin_api.gs) **************
 *
 * Add this file to your Apps Script project.
 * Then add these to the ACTIONS object in hostingerbridge.gs:
 *
 *   adminMarkPaidAndFulfill:  adminMarkPaidAndFulfill,
 *   getOrdersSummaryAdmin:    getOrdersSummaryAdmin,
 *   getSubsSummaryAdmin:      getSubsSummaryAdmin,
 *   getExpiringSubs:          getExpiringSubs,
 *   getOrdersAdmin:           getOrdersAdmin,
 *   getRecentPaidOrdersAdmin: getRecentPaidOrdersAdmin,
 *   adminGetSubs:             adminGetSubs,
 *   getCouponAnalyticsSummary: getCouponAnalyticsSummary,  ← already in coupons.gs
 *
 * getCouponAnalyticsSummary is already in coupons.gs — just add to ACTIONS.
 *****************************************************************************/


// ──────────────────────────────────────────────────────────────────────────────
// 1. adminMarkPaidAndFulfill(orderId, txnRef)
//    Called by Mark Paid button & Quick Pay in the admin dashboard.
//    Wraps your existing ADMIN_markPaidAndFulfill from debugmanage.gs.
// ──────────────────────────────────────────────────────────────────────────────
function adminMarkPaidAndFulfill(orderId, txnRef) {
  // ADMIN_markPaidAndFulfill already exists in debugmanage.gs
  return ADMIN_markPaidAndFulfill(orderId, txnRef);
}


// ──────────────────────────────────────────────────────────────────────────────
// 2. getOrdersSummaryAdmin(days)
//    Dashboard stats: revenue today, last 7d, daily breakdown, by-service.
//    days defaults to 14 for chart; pass 30/90 for revenue page.
// ──────────────────────────────────────────────────────────────────────────────
function getOrdersSummaryAdmin(days) {
  try {
    days = Math.max(1, parseInt(String(days || "14"), 10) || 14);

    const orders = sh_(TAB_ORDERS);
    const { idx } = headerIndex_(orders);
    const last = orders.getLastRow();
    const data = last < 2 ? [] : orders.getRange(2, 1, last - 1, orders.getLastColumn()).getValues();

    const now   = new Date();
    const today = new Date(now); today.setHours(0, 0, 0, 0);
    const cutoff = new Date(today.getTime() - (days - 1) * 86400000);

    let revenueToday = 0, paidToday = 0, pendingToday = 0, manualPending = 0, failed = 0;
    let revenue7d = 0;
    const d7cut = new Date(today.getTime() - 6 * 86400000);

    // daily map: "YYYY-MM-DD" → amount
    const dailyMap = {};
    // Initialize all days in range
    for (let i = 0; i < days; i++) {
      const d = new Date(cutoff.getTime() + i * 86400000);
      const key = ymd_(d);
      dailyMap[key] = 0;
    }

    const byService = {}; // service → { count, revenue }
    let totalRevenue = 0, paidCount = 0, totalDiscount = 0;

    for (const r of data) {
      const status   = String(r[idx["Status"]] || "").toUpperCase().trim();
      const fStatus  = idx["FulfillmentStatus"] != null ? String(r[idx["FulfillmentStatus"]] || "").toUpperCase().trim() : "";
      const amount   = asNumber_(r[idx["FinalAmount"]] || 0);
      const discount = idx["Discount"] != null ? asNumber_(r[idx["Discount"]] || 0) : 0;
      const svc      = String(r[idx["Service"]] || "").trim();

      // Manual pending count (all time)
      if (fStatus === "MANUAL_PENDING") manualPending++;
      if (fStatus === "FAILED")         failed++;

      if (status !== "PAID") {
        // Count pending today
        const dateRaw = idx["CreatedAt"] != null ? r[idx["CreatedAt"]] : null;
        if (dateRaw) {
          const d = dateRaw instanceof Date ? dateRaw : new Date(dateRaw);
          const d0 = new Date(d); d0.setHours(0, 0, 0, 0);
          if (d0.getTime() === today.getTime() && status === "CREATED") pendingToday++;
        }
        continue;
      }

      // PAID orders
      const verifiedRaw = idx["VerifiedAt"] != null ? r[idx["VerifiedAt"]] : (idx["CreatedAt"] != null ? r[idx["CreatedAt"]] : null);
      const dt = verifiedRaw instanceof Date ? verifiedRaw : (verifiedRaw ? new Date(verifiedRaw) : null);
      if (!dt || isNaN(dt.getTime())) continue;

      const dt0 = new Date(dt); dt0.setHours(0, 0, 0, 0);

      // Today's stats
      if (dt0.getTime() === today.getTime()) {
        revenueToday += amount;
        paidToday++;
      }

      // 7-day revenue
      if (dt >= d7cut) revenue7d += amount;

      // Range stats (for chart + totals)
      if (dt >= cutoff) {
        const key = ymd_(dt);
        if (dailyMap[key] !== undefined) dailyMap[key] += amount;
        totalRevenue += amount;
        totalDiscount += discount;
        paidCount++;

        if (svc) {
          if (!byService[svc]) byService[svc] = { count: 0, revenue: 0 };
          byService[svc].count++;
          byService[svc].revenue += amount;
        }
      }
    }

    // Build daily array for chart
    const dailyRevenue = Object.entries(dailyMap).sort().map(([date, amount]) => ({
      label: date.slice(5), // "MM-DD"
      date,
      amount: Math.round(amount)
    }));

    return ok_({
      revenueToday:  Math.round(revenueToday),
      revenue7d:     Math.round(revenue7d),
      paidToday,
      pendingToday,
      manualPending,
      failed,
      totalRevenue:  Math.round(totalRevenue),
      totalDiscount: Math.round(totalDiscount),
      paidCount,
      avgOrderValue: paidCount > 0 ? Math.round(totalRevenue / paidCount) : 0,
      dailyRevenue,
      byService
    });

  } catch (e) {
    return bad_("getOrdersSummaryAdmin error: " + (e && e.message ? e.message : e));
  }
}


// ──────────────────────────────────────────────────────────────────────────────
// 3. getSubsSummaryAdmin()
//    Active sub count, expiring counts, expiring list for dashboard widget.
// ──────────────────────────────────────────────────────────────────────────────
function getSubsSummaryAdmin() {
  try {
    const subs = sh_(TAB_SUBS);
    const { idx } = headerIndex_(subs);
    const last = subs.getLastRow();
    if (last < 2) return ok_({ activeCount: 0, totalCount: 0, expiringIn3: 0, expiringIn7: 0, customerCount: 0, newLast30: 0, expiringList: [] });

    const data = subs.getRange(2, 1, last - 1, subs.getLastColumn()).getValues();
    const now  = new Date(); now.setHours(0, 0, 0, 0);
    const in3  = new Date(now.getTime() + 3 * 86400000);
    const in7  = new Date(now.getTime() + 7 * 86400000);

    let activeCount = 0, totalCount = 0, expiringIn3 = 0, expiringIn7 = 0;
    const phones = new Set();
    const expiringList = [];

    for (const r of data) {
      totalCount++;
      const status = idx["Status"] != null ? String(r[idx["Status"]] || "").toUpperCase().trim() : "";
      if (status === "EXPIRED" || status === "CANCELLED") continue;

      const expRaw = idx["ExpiryDate"] != null ? r[idx["ExpiryDate"]] : null;
      const exp = expRaw instanceof Date ? expRaw : (expRaw ? new Date(expRaw) : null);
      if (!exp || isNaN(exp.getTime())) continue;

      const exp0 = new Date(exp); exp0.setHours(0, 0, 0, 0);
      if (exp0 >= now) activeCount++;

      const daysLeft = Math.ceil((exp0.getTime() - now.getTime()) / 86400000);
      if (daysLeft >= 0 && exp0 <= in3) expiringIn3++;
      if (daysLeft >= 0 && exp0 <= in7) expiringIn7++;

      if (daysLeft >= 0 && daysLeft <= 7) {
        const phone = idx["Phone"] != null ? String(r[idx["Phone"]] || "").replace(/\D/g, "").slice(-10) : "";
        if (phone) phones.add(phone);
        expiringList.push({
          subId:      idx["SubID"] != null ? String(r[idx["SubID"]] || "").trim() : "",
          name:       idx["Name"]  != null ? String(r[idx["Name"]]  || "").trim() : "",
          phone,
          service:    idx["Service"] != null ? String(r[idx["Service"]] || "").trim() : "",
          plan:       idx["Plan"]    != null ? String(r[idx["Plan"]]    || "").trim() : "",
          expiryDate: exp.toISOString(),
          daysLeft
        });
      }
    }

    // Customer count from CUSTOMERS sheet
    let customerCount = 0, newLast30 = 0;
    try {
      const csh  = sh_(TAB_CUS);
      const clast = csh.getLastRow();
      if (clast >= 2) {
        customerCount = clast - 1;
        const { idx: cIdx } = headerIndex_(csh);
        if (cIdx["MemberSince"] != null) {
          const cdata = csh.getRange(2, 1, clast - 1, csh.getLastColumn()).getValues();
          const d30 = new Date(now.getTime() - 30 * 86400000);
          for (const r of cdata) {
            const ms = r[cIdx["MemberSince"]];
            const d  = ms instanceof Date ? ms : new Date(ms);
            if (!isNaN(d.getTime()) && d >= d30) newLast30++;
          }
        }
      }
    } catch (e) {}

    expiringList.sort((a, b) => a.daysLeft - b.daysLeft);

    return ok_({ activeCount, totalCount, expiringIn3, expiringIn7, customerCount, newLast30, expiringList });

  } catch (e) {
    return bad_("getSubsSummaryAdmin error: " + (e && e.message ? e.message : e));
  }
}


// ──────────────────────────────────────────────────────────────────────────────
// 4. getExpiringSubs(days)
//    Full list for the Expiring Soon page.
// ──────────────────────────────────────────────────────────────────────────────
function getExpiringSubs(days) {
  try {
    days = Math.max(1, parseInt(String(days || "7"), 10) || 7);
    const subs = sh_(TAB_SUBS);
    const { idx } = headerIndex_(subs);
    const last = subs.getLastRow();
    if (last < 2) return ok_({ subscriptions: [] });

    const data  = subs.getRange(2, 1, last - 1, subs.getLastColumn()).getValues();
    const now   = new Date(); now.setHours(0, 0, 0, 0);
    const cutoff = new Date(now.getTime() + days * 86400000);
    const out = [];

    for (const r of data) {
      const expRaw = idx["ExpiryDate"] != null ? r[idx["ExpiryDate"]] : null;
      const exp = expRaw instanceof Date ? expRaw : (expRaw ? new Date(expRaw) : null);
      if (!exp || isNaN(exp.getTime())) continue;

      const exp0 = new Date(exp); exp0.setHours(0, 0, 0, 0);
      const daysLeft = Math.ceil((exp0.getTime() - now.getTime()) / 86400000);
      if (daysLeft < 0 || exp0 > cutoff) continue;

      out.push({
        subId:      idx["SubID"] != null     ? String(r[idx["SubID"]]     || "").trim() : "",
        name:       idx["Name"] != null      ? String(r[idx["Name"]]      || "").trim() : "",
        phone:      idx["Phone"] != null     ? String(r[idx["Phone"]]     || "").replace(/\D/g,"").slice(-10) : "",
        service:    idx["Service"] != null   ? String(r[idx["Service"]]   || "").trim() : "",
        plan:       idx["Plan"] != null      ? String(r[idx["Plan"]]      || "").trim() : "",
        expiryDate: exp.toISOString(),
        daysLeft
      });
    }

    out.sort((a, b) => a.daysLeft - b.daysLeft);
    return ok_({ subscriptions: out });

  } catch (e) {
    return bad_("getExpiringSubs error: " + (e && e.message ? e.message : e));
  }
}


// ──────────────────────────────────────────────────────────────────────────────
// 5. getOrdersAdmin(days, statusFilter)
//    Orders list page with optional status filter.
// ──────────────────────────────────────────────────────────────────────────────
function getOrdersAdmin(days, statusFilter) {
  try {
    days = Math.max(1, parseInt(String(days || "7"), 10) || 7);
    statusFilter = String(statusFilter || "").toUpperCase().trim();

    const orders = sh_(TAB_ORDERS);
    const { idx } = headerIndex_(orders);
    const last = orders.getLastRow();
    if (last < 2) return ok_({ orders: [] });

    const data   = orders.getRange(2, 1, last - 1, orders.getLastColumn()).getValues();
    const now    = new Date();
    const cutoff = new Date(now); cutoff.setHours(0, 0, 0, 0);
    cutoff.setDate(cutoff.getDate() - (days - 1));

    const out = [];
    for (let i = data.length - 1; i >= 0; i--) {
      const r = data[i];
      const status = String(r[idx["Status"]] || "").toUpperCase().trim();
      if (statusFilter && status !== statusFilter) continue;

      const dateRaw = idx["CreatedAt"] != null ? r[idx["CreatedAt"]] : null;
      const dt = dateRaw instanceof Date ? dateRaw : (dateRaw ? new Date(dateRaw) : null);
      if (dt && !isNaN(dt.getTime()) && dt < cutoff) continue;

      out.push({
        orderId:           String(r[idx["OrderID"]]          || "").trim(),
        name:              idx["Name"] != null        ? String(r[idx["Name"]]        || "").trim() : "",
        phone:             idx["Phone"] != null       ? String(r[idx["Phone"]]       || "").replace(/\D/g,"").slice(-10) : "",
        service:           String(r[idx["Service"]]          || "").trim(),
        plan:              String(r[idx["Plan"]]              || "").trim(),
        status,
        fulfillmentStatus: idx["FulfillmentStatus"] != null  ? String(r[idx["FulfillmentStatus"]] || "").trim() : "",
        amount:            asNumber_(r[idx["FinalAmount"]]   || 0),
        createdAt:         dt ? dt.toISOString() : ""
      });

      if (out.length >= 200) break; // cap
    }

    return ok_({ orders: JSON.parse(JSON.stringify(out)) });

  } catch (e) {
    return bad_("getOrdersAdmin error: " + (e && e.message ? e.message : e));
  }
}


// ──────────────────────────────────────────────────────────────────────────────
// 6. getRecentPaidOrdersAdmin(limit)
//    Recent PAID orders for dashboard widget.
// ──────────────────────────────────────────────────────────────────────────────
function getRecentPaidOrdersAdmin(limit) {
  try {
    limit = Math.max(1, Math.min(50, parseInt(String(limit || "8"), 10) || 8));

    const orders = sh_(TAB_ORDERS);
    const { idx } = headerIndex_(orders);
    const last = orders.getLastRow();
    if (last < 2) return ok_({ orders: [] });

    const data = orders.getRange(2, 1, last - 1, orders.getLastColumn()).getValues();
    const out = [];

    for (let i = data.length - 1; i >= 0 && out.length < limit; i--) {
      const r = data[i];
      if (String(r[idx["Status"]] || "").toUpperCase().trim() !== "PAID") continue;
      const dateRaw = idx["VerifiedAt"] != null ? r[idx["VerifiedAt"]] : (idx["CreatedAt"] != null ? r[idx["CreatedAt"]] : null);
      const dt = dateRaw instanceof Date ? dateRaw : (dateRaw ? new Date(dateRaw) : null);
      out.push({
        orderId:           String(r[idx["OrderID"]]          || "").trim(),
        service:           String(r[idx["Service"]]          || "").trim(),
        plan:              String(r[idx["Plan"]]              || "").trim(),
        amount:            asNumber_(r[idx["FinalAmount"]]   || 0),
        fulfillmentStatus: idx["FulfillmentStatus"] != null  ? String(r[idx["FulfillmentStatus"]] || "").trim() : "",
        createdAt:         dt ? dt.toISOString() : ""
      });
    }

    return ok_({ orders: JSON.parse(JSON.stringify(out)) });

  } catch (e) {
    return bad_("getRecentPaidOrdersAdmin error: " + (e && e.message ? e.message : e));
  }
}


// ──────────────────────────────────────────────────────────────────────────────
// 7. adminGetSubs(phone)
//    Returns full subscription data for a customer including credentials.
//    Used by the Search Customer page.
// ──────────────────────────────────────────────────────────────────────────────
function adminGetSubs(phone) {
  try {
    const ph = String(phone || "").replace(/\D/g, "").slice(-10);
    if (!ph) return bad_("Phone required");

    const subs = sh_(TAB_SUBS);
    const { idx } = headerIndex_(subs);
    const last = subs.getLastRow();
    if (last < 2) return ok_({ subscriptions: [] });

    const data  = subs.getRange(2, 1, last - 1, subs.getLastColumn()).getValues();
    const nowMs = Date.now();
    const out   = [];

    for (const r of data) {
      const rowPhone = idx["Phone"] != null
        ? String(r[idx["Phone"]] || "").replace(/\D/g, "").slice(-10)
        : "";
      if (rowPhone !== ph) continue;

      const expRaw  = idx["ExpiryDate"] != null ? r[idx["ExpiryDate"]] : null;
      const exp     = expRaw instanceof Date ? expRaw : (expRaw ? new Date(expRaw) : null);
      const daysLeft = exp && !isNaN(exp.getTime())
        ? Math.ceil((exp.getTime() - nowMs) / 86400000)
        : null;

      const startRaw = idx["StartDate"] != null ? r[idx["StartDate"]] : null;
      const start    = startRaw instanceof Date ? startRaw : (startRaw ? new Date(startRaw) : null);

      out.push({
        subId:        idx["SubID"]         != null ? String(r[idx["SubID"]]         || "").trim() : "",
        orderId:      idx["OrderID"]       != null ? String(r[idx["OrderID"]]       || "").trim() : "",
        service:      idx["Service"]       != null ? String(r[idx["Service"]]       || "").trim() : "",
        plan:         idx["Plan"]          != null ? String(r[idx["Plan"]]          || "").trim() : "",
        status:       idx["Status"]        != null ? String(r[idx["Status"]]        || "").trim() : "",
        startDate:    start && !isNaN(start.getTime()) ? start.toISOString() : "",
        expiryDate:   exp   && !isNaN(exp.getTime())   ? exp.toISOString()   : "",
        daysLeft,
        durationDays: idx["DurationDays"]  != null ? asNumber_(r[idx["DurationDays"]]  || 0) : 0,
        inventoryRef: idx["InventoryRef"]  != null ? String(r[idx["InventoryRef"]]  || "").trim() : "",
        loginId:      idx["LoginId"]       != null ? String(r[idx["LoginId"]]       || "").trim() : "",
        pass:         idx["Password"]      != null ? String(r[idx["Password"]]      || "").trim() : "",
        profileNumber:idx["ProfileNumber"] != null ? String(r[idx["ProfileNumber"]] || "").trim() : "",
        profileName:  idx["ProfileName"]   != null ? String(r[idx["ProfileName"]]   || "").trim() : "",
        profilePIN:   idx["ProfilePIN"]    != null ? String(r[idx["ProfilePIN"]]    || "").trim() : ""
      });
    }

    // active/expiring first
    out.sort((a, b) => {
      const ax = a.daysLeft ?? -9999;
      const bx = b.daysLeft ?? -9999;
      return bx - ax;
    });

    return ok_({ subscriptions: JSON.parse(JSON.stringify(out)) });

  } catch (e) {
    return bad_("adminGetSubs error: " + (e && e.message ? e.message : e));
  }
}


/* ─────────────────────────────────────────────────────────────────────────
 * INSTRUCTIONS — Add these lines to the ACTIONS object in hostingerbridge.gs
 *
 *   adminMarkPaidAndFulfill:   adminMarkPaidAndFulfill,
 *   getOrdersSummaryAdmin:     getOrdersSummaryAdmin,
 *   getSubsSummaryAdmin:       getSubsSummaryAdmin,
 *   getExpiringSubs:           getExpiringSubs,
 *   getOrdersAdmin:            getOrdersAdmin,
 *   getRecentPaidOrdersAdmin:  getRecentPaidOrdersAdmin,
 *   adminGetSubs:              adminGetSubs,
 *   getCouponAnalyticsSummary: getCouponAnalyticsSummary,   ← already in coupons.gs
 *
 * HOSTING OPTIONS:
 *   Option A — Apps Script (fastest, no cost):
 *     1. Paste admin.html into your project as a new file named "AdminIndex"
 *     2. In utils.gs, add a doGet route for /admin or a separate web app deploy
 *     3. Change API_URL in admin.html to the same exec URL as your main deploy
 *
 *   Option B — Hostinger:
 *     1. Upload admin.html to your Hostinger file manager
 *     2. Set API_URL to your api.php proxy URL
 *     3. Password-protect the file via .htaccess or Hostinger panel
 * ───────────────────────────────────────────────────────────────────────── */