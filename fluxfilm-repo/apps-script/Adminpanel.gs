/*************** FluxFilm Backend — Admin Panel API ***************
 *
 * Powers the FluxFilm Admin web panel (go.fluxfilm.in).
 * These functions were referenced by the panel but not yet defined —
 * this file implements them so the Dashboard, Revenue, Orders and
 * Expiring pages show REAL, correct numbers.
 *
 * REVENUE RULE (important):
 *   Revenue is counted ONLY from ORDERS where Status = "PAID",
 *   summed on the FinalAmount column, bucketed by CreatedAt date.
 *
 * ACTIVE SUBS RULE:
 *   activeCount = SUBSCRIPTIONS rows with Status = "ACTIVE"
 *   totalCount  = ACTIVE + EXPIRED rows (all real subscription rows)
 *
 * IMPORTANT — register these in HostingerBridge.gs ACTIONS map:
 *     getOrdersSummaryAdmin:     getOrdersSummaryAdmin,
 *     getSubsSummaryAdmin:       getSubsSummaryAdmin,
 *     getRecentPaidOrdersAdmin:  getRecentPaidOrdersAdmin,
 *     getOrdersAdmin:            getOrdersAdmin,
 *     getExpiringSubs:           getExpiringSubs,
 *     adminMarkPaidAndFulfill:   adminMarkPaidAndFulfill,
 *
 * Shared constants (TAB_ORDERS, TAB_SUBS, TAB_CUS, ORDER_STATUS,
 * SUB_STATUS) live in code.gs — do NOT redeclare them here.
 *****************************************************************/


/* ---------- small local date helpers ---------- */

function ap_toDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (v == null || v === "") return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

// "YYYY-MM-DD" in the script's timezone
function ap_ymd_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// short label like "6 Jul"
function ap_shortLabel_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "d MMM");
}

function ap_startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}


/**
 * ORDERS SUMMARY (revenue analytics + dashboard order stats).
 * Revenue = PAID orders only, on FinalAmount.
 *
 * @param {number} days  window length (default 14)
 * @return ok_({
 *   revenueToday, revenue7d, paidToday, pendingToday, manualPending, failed,
 *   totalRevenue, paidCount, avgOrderValue, totalDiscount,
 *   dailyRevenue: [{date,label,amount}],
 *   byService: { svc: {count,revenue} }
 * })
 */
function getOrdersSummaryAdmin(days) {
  try {
    const win = asNumber_(days || 14);
    const orders = sh_(TAB_ORDERS);
    const { idx } = headerIndex_(orders);

    const last = orders.getLastRow();
    if (last < 2) {
      return ok_({
        revenueToday: 0, revenue7d: 0, paidToday: 0, pendingToday: 0,
        manualPending: 0, failed: 0, totalRevenue: 0, paidCount: 0,
        avgOrderValue: 0, totalDiscount: 0, dailyRevenue: [], byService: {}
      });
    }

    const data = orders.getRange(2, 1, last - 1, orders.getLastColumn()).getValues();

    const now = now_();
    const todayYmd = ap_ymd_(now);
    const winStart = ap_startOfDay_(new Date(now.getTime() - (win - 1) * 86400000));
    const sevenStart = ap_startOfDay_(new Date(now.getTime() - 6 * 86400000));

    // Pre-build the daily buckets for the whole window (so empty days show 0)
    const dayBuckets = {}; // ymd -> {date,label,amount}
    const orderedYmd = [];
    for (let i = 0; i < win; i++) {
      const d = new Date(winStart.getTime() + i * 86400000);
      const ymd = ap_ymd_(d);
      dayBuckets[ymd] = { date: ymd, label: ap_shortLabel_(d), amount: 0 };
      orderedYmd.push(ymd);
    }

    let revenueToday = 0, revenue7d = 0, paidToday = 0, pendingToday = 0;
    let manualPending = 0, failed = 0;
    let totalRevenue = 0, paidCount = 0, totalDiscount = 0;
    const byService = {};

    for (const r of data) {
      const status = String(r[idx["Status"]] || "").trim().toUpperCase();
      const fstatus = String(r[idx["FulfillmentStatus"]] || "").trim().toUpperCase();
      const created = ap_toDate_(r[idx["CreatedAt"]]);
      const amount = asNumber_(r[idx["FinalAmount"]]);
      const discount = asNumber_(r[idx["Discount"]]);
      const service = String(r[idx["Service"]] || "").trim() || "Other";
      const createdYmd = created ? ap_ymd_(created) : "";

      // Operational counters (today)
      if (createdYmd === todayYmd) {
        if (status === ORDER_STATUS.PAID) paidToday++;
        else if (status === ORDER_STATUS.CREATED) pendingToday++;
      }
      if (fstatus === "MANUAL_PENDING") manualPending++;
      if (fstatus === "FAILED") failed++;

      // Revenue counters — PAID only
      if (status !== ORDER_STATUS.PAID) continue;

      // window-scoped revenue totals
      if (created && created >= winStart) {
        totalRevenue += amount;
        paidCount++;
        totalDiscount += discount;

        if (!byService[service]) byService[service] = { count: 0, revenue: 0 };
        byService[service].count++;
        byService[service].revenue += amount;

        if (dayBuckets[createdYmd]) dayBuckets[createdYmd].amount += amount;
      }

      if (createdYmd === todayYmd) revenueToday += amount;
      if (created && created >= sevenStart) revenue7d += amount;
    }

    const dailyRevenue = orderedYmd.map(y => dayBuckets[y]);
    const avgOrderValue = paidCount ? Math.round(totalRevenue / paidCount) : 0;

    return ok_({
      revenueToday: Math.round(revenueToday),
      revenue7d: Math.round(revenue7d),
      paidToday, pendingToday, manualPending, failed,
      totalRevenue: Math.round(totalRevenue),
      paidCount,
      avgOrderValue,
      totalDiscount: Math.round(totalDiscount),
      dailyRevenue,
      byService
    });

  } catch (e) {
    return bad_("getOrdersSummaryAdmin error: " + (e && e.message ? e.message : e));
  }
}


/**
 * SUBSCRIPTIONS + CUSTOMERS summary for dashboard cards.
 *
 * @return ok_({
 *   activeCount, totalCount, expiringIn3, expiringIn7,
 *   customerCount, customersWithActiveSubs, newLast30,
 *   expiringList: [{name,phone,service,daysLeft,subId}]
 * })
 */
function getSubsSummaryAdmin() {
  try {
    const subs = sh_(TAB_SUBS);
    const { idx } = headerIndex_(subs);

    const now = now_();
    const nowMs = now.getTime();
    const DAY = 86400000;

    let activeCount = 0, totalCount = 0, expiringIn3 = 0, expiringIn7 = 0;
    const expiringList = [];
    const activePhones = {}; // distinct phones that have >=1 ACTIVE sub

    const sLast = subs.getLastRow();
    if (sLast >= 2) {
      const data = subs.getRange(2, 1, sLast - 1, subs.getLastColumn()).getValues();

      for (const r of data) {
        const status = String(r[idx["Status"]] || "").trim().toUpperCase();
        if (!status) continue;

        // total = every real sub row (ACTIVE or EXPIRED); skip reservation rows
        const subId = String(r[idx["SubID"]] || "").trim();
        if (subId.indexOf("RESERVE-") === 0) continue;

        totalCount++;

        if (status !== SUB_STATUS.ACTIVE) continue;
        activeCount++;

        // distinct customer (by phone) with an active sub
        const phone = normalizePhone_(String(r[idx["Phone"]] || ""));
        if (phone) activePhones[phone] = true;

        // expiring buckets — only for ACTIVE subs
        const exp = ap_toDate_(r[idx["ExpiryDate"]]);
        if (!exp) continue;
        const daysLeft = Math.ceil((exp.getTime() - nowMs) / DAY);
        if (daysLeft >= 0 && daysLeft <= 7) {
          expiringIn7++;
          if (daysLeft <= 3) expiringIn3++;
          expiringList.push({
            name: (idx["Name"] != null) ? String(r[idx["Name"]] || "").trim() : "",
            phone: phone,
            service: String(r[idx["Service"]] || "").trim(),
            plan: String(r[idx["Plan"]] || "").trim(),
            daysLeft: daysLeft,
            subId: subId
          });
        }
      }
    }

    expiringList.sort((a, b) => a.daysLeft - b.daysLeft);

    // Customers sheet: total + new in last 30 days
    let customerCount = 0, newLast30 = 0;
    const cus = sh_(TAB_CUS);
    const { idx: cIdx } = headerIndex_(cus);
    const cLast = cus.getLastRow();
    if (cLast >= 2) {
      const cData = cus.getRange(2, 1, cLast - 1, cus.getLastColumn()).getValues();
      const cutoff = nowMs - 30 * DAY;
      for (const r of cData) {
        const phone = normalizePhone_(String(r[cIdx["Phone"]] || ""));
        if (!phone) continue;
        customerCount++;
        const since = ap_toDate_(r[cIdx["MemberSince"]]);
        if (since && since.getTime() >= cutoff) newLast30++;
      }
    }

    const customersWithActiveSubs = Object.keys(activePhones).length;

    return ok_({
      activeCount,
      totalCount,
      expiringIn3,
      expiringIn7,
      customerCount,
      customersWithActiveSubs,
      newLast30,
      expiringList: expiringList.slice(0, 20)
    });

  } catch (e) {
    return bad_("getSubsSummaryAdmin error: " + (e && e.message ? e.message : e));
  }
}


/**
 * Recent PAID orders for the dashboard list.
 * @param {number} limit default 8
 */
function getRecentPaidOrdersAdmin(limit) {
  try {
    const max = asNumber_(limit || 8);
    const orders = sh_(TAB_ORDERS);
    const { idx } = headerIndex_(orders);
    const last = orders.getLastRow();
    if (last < 2) return ok_({ orders: [] });

    const data = orders.getRange(2, 1, last - 1, orders.getLastColumn()).getValues();
    const out = [];

    for (let i = data.length - 1; i >= 0; i--) {
      const r = data[i];
      if (String(r[idx["Status"]] || "").trim().toUpperCase() !== ORDER_STATUS.PAID) continue;
      out.push(ap_publicOrderRow_(r, idx));
      if (out.length >= max) break;
    }

    return ok_({ orders: out });
  } catch (e) {
    return bad_("getRecentPaidOrdersAdmin error: " + (e && e.message ? e.message : e));
  }
}


/**
 * Orders list with status + date-window filter (Orders page).
 * @param {number} days   default 7
 * @param {string} status optional "PAID"|"CREATED"|"CANCELLED" ("" = all)
 */
function getOrdersAdmin(days, status) {
  try {
    const win = asNumber_(days || 7);
    const wantStatus = String(status || "").trim().toUpperCase();
    const orders = sh_(TAB_ORDERS);
    const { idx } = headerIndex_(orders);
    const last = orders.getLastRow();
    if (last < 2) return ok_({ orders: [] });

    const data = orders.getRange(2, 1, last - 1, orders.getLastColumn()).getValues();
    const cutoff = ap_startOfDay_(new Date(now_().getTime() - (win - 1) * 86400000));

    const out = [];
    for (let i = data.length - 1; i >= 0; i--) {
      const r = data[i];
      const created = ap_toDate_(r[idx["CreatedAt"]]);
      if (created && created < cutoff) continue;

      const st = String(r[idx["Status"]] || "").trim().toUpperCase();
      if (wantStatus && st !== wantStatus) continue;

      out.push(ap_publicOrderRow_(r, idx));
    }

    return ok_({ orders: out });
  } catch (e) {
    return bad_("getOrdersAdmin error: " + (e && e.message ? e.message : e));
  }
}


/**
 * Subscriptions expiring within N days (Expiring page).
 * ACTIVE subs only.
 * @param {number} days default 7
 */
function getExpiringSubs(days) {
  try {
    const win = asNumber_(days || 7);
    const subs = sh_(TAB_SUBS);
    const { idx } = headerIndex_(subs);
    const last = subs.getLastRow();
    if (last < 2) return ok_({ subscriptions: [] });

    const data = subs.getRange(2, 1, last - 1, subs.getLastColumn()).getValues();
    const nowMs = now_().getTime();
    const DAY = 86400000;

    const out = [];
    for (const r of data) {
      if (String(r[idx["Status"]] || "").trim().toUpperCase() !== SUB_STATUS.ACTIVE) continue;
      const exp = ap_toDate_(r[idx["ExpiryDate"]]);
      if (!exp) continue;
      const daysLeft = Math.ceil((exp.getTime() - nowMs) / DAY);
      if (daysLeft < 0 || daysLeft > win) continue;

      out.push({
        name: (idx["Name"] != null) ? String(r[idx["Name"]] || "").trim() : "",
        phone: normalizePhone_(String(r[idx["Phone"]] || "")),
        service: String(r[idx["Service"]] || "").trim(),
        plan: String(r[idx["Plan"]] || "").trim(),
        expiryDate: exp.toISOString(),
        daysLeft: daysLeft,
        subId: String(r[idx["SubID"]] || "").trim()
      });
    }

    out.sort((a, b) => a.daysLeft - b.daysLeft);
    return ok_({ subscriptions: out });
  } catch (e) {
    return bad_("getExpiringSubs error: " + (e && e.message ? e.message : e));
  }
}


/**
 * Mark an order PAID and fulfill it (Quick Actions + row buttons).
 * Delegates to your existing ADMIN_markPaidAndFulfill(orderId, txnRef).
 */
function adminMarkPaidAndFulfill(orderId, txnRef) {
  try {
    const oid = String(orderId || "").trim().toUpperCase();
    if (!oid) return bad_("Order ID is required.");

    if (typeof ADMIN_markPaidAndFulfill === "function") {
      const res = ADMIN_markPaidAndFulfill(oid, String(txnRef || "").trim());
      // Normalise to { ok, message } shape the panel expects
      if (res && typeof res === "object" && "ok" in res) return res;
      return ok_({ message: "Order " + oid + " processed." });
    }

    return bad_("Fulfillment function ADMIN_markPaidAndFulfill() not found in project.");
  } catch (e) {
    return bad_("adminMarkPaidAndFulfill error: " + (e && e.message ? e.message : e));
  }
}


/* ---------- internal ---------- */

function ap_publicOrderRow_(r, idx) {
  const created = ap_toDate_(r[idx["CreatedAt"]]);
  return {
    orderId: String(r[idx["OrderID"]] || "").trim(),
    createdAt: created ? created.toISOString() : "",
    service: String(r[idx["Service"]] || "").trim(),
    plan: String(r[idx["Plan"]] || "").trim(),
    name: (idx["Name"] != null) ? String(r[idx["Name"]] || "").trim() : "",
    phone: normalizePhone_(String(r[idx["Phone"]] || "")),
    amount: asNumber_(r[idx["FinalAmount"]]),
    status: String(r[idx["Status"]] || "").trim().toUpperCase(),
    fulfillmentStatus: String(r[idx["FulfillmentStatus"]] || "").trim().toUpperCase()
  };
}