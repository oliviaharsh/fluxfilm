/**************** FluxFilm Backend — HostingerBridge (FIXED) ****************
 * Purpose:
 *  - Single doPost router for Hostinger via /api.php proxy
 *  - Telegram updates can also POST here (no apiKey for Telegram)
 *
 * IMPORTANT:
 *  - ContentService TextOutput does NOT support setHeader()
 *  - So we do NOT set CORS headers here.
 ******************************************************************************/


// ====== CONFIG ======
const FF_API_KEY_EXPECTED = "FF_GO_2026_SUPER_SECRET_987654";


// ====== RESPONSE HELPERS ======
function json_(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj == null ? {} : obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// Preflight (not required when calling via api.php, but safe)
function doOptions(e){
  return json_({ ok:true });
}


// Simple test
function ping(){
  return { ok:true, pong:true, at:new Date().toISOString() };
}


// Detect Telegram update payload
function isTelegramUpdate_(body){
  return !!(body && (body.update_id || body.message || body.edited_message || body.callback_query));
}


/**
 * ✅ Hostinger API Router
 * Expected body:
 *   { apiKey:"...", action:"ping", args:[...] }
 *
 * Also supports Telegram bot updates posted to doPost (no apiKey).
 */
function doPost(e){
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    let body = {};
    try { body = JSON.parse(raw); } catch(_){ body = {}; }

    // 1) TELEGRAM ROUTE
    if (isTelegramUpdate_(body)) {
      try {
        if (typeof tg_doPost_ === "function") {
          return tg_doPost_(e);
        }
      } catch (err) {
        try { console.log("[TG] handler crash:", err); } catch(_){}
      }
      return ContentService.createTextOutput("OK");
    }


    // 2) HOSTINGER AUTH
    if (String(body.apiKey || "") !== String(FF_API_KEY_EXPECTED)) {
      return json_({ ok:false, message:"Unauthorized" });
    }

    // 3) ACTION ROUTER
    const action = String(body.action || "").trim();
    const args   = Array.isArray(body.args) ? body.args : [];

    if (!action) return json_({ ok:false, message:"Missing action" });

    const ACTIONS = {
      ping: ping,
      getBootstrap: getBootstrap,
      getFaqs: getFaqs,
      createOrder: createOrder,
      validateCoupon: validateCoupon,
      verifyPayment: verifyPayment,
      fulfillAndGetAccess: fulfillAndGetAccess,
      getOrderStatus: getOrderStatus,
      getResumePaymentByPhone: getResumePaymentByPhone,
      getMySubscriptions: getMySubscriptions,
      createRenewOrder: createRenewOrder,
      getWalletByPhone: getWalletByPhone,
      recoverSendOtp: recoverSendOtp,
      recoverVerifyOtp: recoverVerifyOtp,
      recoverListSubscriptionsSafe: recoverListSubscriptionsSafe,
      recoverGetAccess: recoverGetAccess,
      adminGetSubs: adminGetSubs,
      recoverReassignAccount: recoverReassignAccount,
      getNetflixHouseholdLink: getNetflixHouseholdLink,
      getCustomerOrders: getCustomerOrders,
      getCustomerProfile: getCustomerProfile,
      createOrUpdateCustomerProfile: createOrUpdateCustomerProfile,
      updateCustomerProfilePic: updateCustomerProfilePic,
      getActiveCouponsForCustomer: getActiveCouponsForCustomer,
      recoverReassignAccount: recoverReassignAccount,
      getLatestOtp: getLatestOtp,
      getOtpQuota: getOtpQuota,
      submitRestockRequest: submitRestockRequest,
      getStockLevels: getStockLevels,
      adminMarkPaidAndFulfill:   adminMarkPaidAndFulfill,
      getOrdersSummaryAdmin:     getOrdersSummaryAdmin,
      getSubsSummaryAdmin:       getSubsSummaryAdmin,
      getExpiringSubs:           getExpiringSubs,
      getOrdersAdmin:            getOrdersAdmin,
      getRecentPaidOrdersAdmin:  getRecentPaidOrdersAdmin,
      adminGetSubs:              adminGetSubs,
      getCouponAnalyticsSummary: getCouponAnalyticsSummary,
      asLookupSubForSwitch: asLookupSubForSwitch,
      asSwitchAccount: asSwitchAccount,
      getTrendingItems: getTrendingItems,
    };

    const fn = ACTIONS[action];
    if (typeof fn !== "function") {
      return json_({ ok:false, message:"Unknown action: " + action });
    }

    const out = fn.apply(null, args);

    if (typeof out === "undefined") {
      return json_({ ok:true, result:null });
    }

    return json_(out);

  } catch (err) {
    return json_({
      ok:false,
      message:"Server error",
      detail:String(err && err.message ? err.message : err)
    });
  }
}