/**************** FluxFilm Backend — OTP Service (otp.gs) ****************
 *
 * PURPOSE
 * - Reads OTP emails forwarded by SMS Forwarder app to harshwalia8888@gmail.com
 * - getLatestOtp(service, phone) → returns latest fresh OTP for JioHotstar/Zee5/SonyLIV
 * - getOtpQuota(phone, service)  → returns monthly usage quota for a customer
 *
 * NO WEBHOOK NEEDED — reads Gmail directly via GmailApp
 *
 * Add to ACTIONS in hostingerbridge.gs:
 *   getLatestOtp: getLatestOtp,
 *   getOtpQuota:  getOtpQuota,
 **************************************************************************/

const OTP_EXPIRY_MS  = 10 * 60 * 1000; // 10 minutes
const OTP_EMAIL_FROM = 'harshwalia8888@gmail.com';

// Service keyword map — matches email subject/body
const OTP_SERVICE_KEYWORDS = {
  'JioHotstar': ['jiohotstar', 'jiohtr', 'VM-JIOHTR', 'JD-JIOHTR', 'VA-JIOHTR'],
  'Zee5':       ['zee5', 'zeeott', 'VM-ZEEOTT', 'VA-ZEEOTT'],
  'SonyLIV':    ['sonyliv', 'sony liv', 'livotp', 'VM-LIVOTP'],
};


/**
 * Called by frontend
 * args[0] = service: "JioHotstar" | "Zee5" | "SonyLIV"
 * args[1] = phone:   customer's 10-digit phone number
 *
 * Returns latest fresh OTP from Gmail inbox
 */
function getLatestOtp(service, phone) {
  try {
    const svc = String(service || '').trim();
    // Normalize: "Zee5 Premium" → "Zee5", "JioHotstar Premium" → "JioHotstar" etc.
    const svcKey = Object.keys(OTP_SERVICE_KEYWORDS).find(function(k){
      return svc.toLowerCase().indexOf(k.toLowerCase()) !== -1;
    }) || svc;
    const ph  = String(phone  || '').replace(/\D/g,'').slice(-10);
    if (!svc) return bad_('Service is required.');

    // Search Gmail for recent OTP emails from the forwarder
    const query = 'from:' + OTP_EMAIL_FROM + ' subject:SMSForwarder newer_than:1d';
    const threads = GmailApp.search(query, 0, 20);

    if (!threads || threads.length === 0) {
      return ok_({
        found:   false,
        message: 'No OTP emails received yet. Try logging in to ' + svc + ' first to trigger an OTP.',
      });
    }


    // Scan threads newest first
    // Collect ALL messages from all threads first, then sort by date descending
    const allMsgs = [];
    for (var t = 0; t < threads.length; t++) {
      const msgs = threads[t].getMessages();
      for (var m = 0; m < msgs.length; m++) {
        allMsgs.push(msgs[m]);
      }
    }

    // Sort newest first
    allMsgs.sort(function(a, b) {
      return b.getDate().getTime() - a.getDate().getTime();
    });

    const now = Date.now();
    const keywords = (OTP_SERVICE_KEYWORDS[svcKey] || [svc.toLowerCase()]).map(k => k.toLowerCase());

    for (var i = 0; i < allMsgs.length; i++) {
      const msg    = allMsgs[i];
      const subject = msg.getSubject().toLowerCase();
      const body    = msg.getPlainBody();
      const bodyLC  = body.toLowerCase();
      const dateMs  = msg.getDate().getTime();

      // Must be recent
      if (now - dateMs > OTP_EXPIRY_MS) continue;

      // Skip already-read messages (already served)
      if (!msg.isUnread()) continue;

      // Must match service keywords
      const matchesService = keywords.some(function(kw) {
        return subject.indexOf(kw) !== -1 || bodyLC.indexOf(kw) !== -1;
      });
      if (!matchesService) continue;

      // Extract OTP
      const otp = otp_extractFromBody_(body);
      if (!otp) continue;

      try { msg.markRead(); } catch(e) {}

      const ageSec       = Math.round((now - dateMs) / 1000);
      const remainingSec = Math.max(0, Math.round((OTP_EXPIRY_MS - (now - dateMs)) / 1000));

      otp_logToSheet_(svcKey, otp, body, dateMs, ph);

      return ok_({
        found:        true,
        otp:          otp,
        service:      svc,
        receivedAt:   new Date(dateMs).toISOString(),
        ageSec:       ageSec,
        remainingSec: remainingSec,
      });
    }

    return ok_({
      found:   false,
      message: 'No fresh OTP found for ' + svc + '. OTPs expire in 5 minutes. Try logging in to ' + svc + ' again to get a new one.',
    });

  } catch(e) {
    return bad_('getLatestOtp error: ' + (e.message || String(e)));
  }
}


/**
 * Returns OTP usage quota for a customer + service (monthly).
 * Reads limit from SETTINGS: OTP_QUOTA_Zee5, OTP_QUOTA_JioHotstar, OTP_QUOTA_SonyLIV
 * Reads used count from SMS_OTP_LOG filtered by phone + service + current month.
 */
function getOtpQuota(phone, service) {
  try {
    const ph  = String(phone  || '').replace(/\D/g,'').slice(-10);
    const svc = String(service||'').trim();
    const svcKey = Object.keys(OTP_SERVICE_KEYWORDS).find(function(k){
      return svc.toLowerCase().indexOf(k.toLowerCase()) !== -1;
    }) || svc;
    if (!ph || !svc) return bad_('Phone and service required.');

    // Global limit per service from SETTINGS sheet
    const limitKey = 'OTP_QUOTA_' + svcKey.replace(/\s+/g, '');
    const limit    = asNumber_(getSetting_(limitKey, 5)); // default 5 if not set

    // Count this month's usage from SMS_OTP_LOG
    const ss = getCore_();
    const sh = ss.getSheetByName('SMS_OTP_LOG');
    if (!sh) return ok_({ used:0, limit:limit, remaining:limit });

    const last = sh.getLastRow();
    if (last < 2) return ok_({ used:0, limit:limit, remaining:limit });

    const data    = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h||'').trim());

    const iSvc    = headers.indexOf('Service');
    const iPhone  = headers.indexOf('Phone');
    // Column 0 (index 0) is always Timestamp

    const now       = new Date();
    const thisMonth = now.getMonth();
    const thisYear  = now.getFullYear();

    let used = 0;
    for (var i = 0; i < data.length; i++) {
      const rowSvc   = iSvc   >= 0 ? String(data[i][iSvc]  || '').trim()                   : '';
      const rowPhone = iPhone >= 0 ? String(data[i][iPhone] || '').replace(/\D/g,'').slice(-10) : '';

      if (rowSvc !== svc && rowSvc !== svcKey) continue; // wrong service
      if (rowPhone !== ph)         continue; // wrong customer
      
      const rowDate = new Date(data[i][0]); // Timestamp column
      if (isNaN(rowDate.getTime())) continue;
      if (rowDate.getMonth()    !== thisMonth) continue;
      if (rowDate.getFullYear() !== thisYear)  continue;

      used++;
    }

    return ok_({ used:used, limit:limit, remaining:Math.max(0, limit - used) });

  } catch(e) {
    return bad_('getOtpQuota error: ' + (e.message || String(e)));
  }
}


// ── OTP extractor ─────────────────────────────────────────
function otp_extractFromBody_(body) {
  if (!body) return '';

  // Zee5: "Your OTP is: 6150.@www.zee5.com"
  var m = body.match(/Your OTP is:\s*(\d{4,6})/i);
  if (m) return m[1];

  // JioHotstar: "<#> 7920 is your JioHotstar verification code"
  m = body.match(/(\d{4,6})\s+is your JioHotstar/i);
  if (m) return m[1];

  // JioHotstar: "Hi, your JioHotstar verification code is 7216"
  m = body.match(/verification code is\s+(\d{4,6})/i);
  if (m) return m[1];

  // SonyLIV: "Sony LIV OTP is 1148"
  m = body.match(/OTP is\s+(\d{4,6})/i);
  if (m) return m[1];

  // Generic fallback: first 4-6 digit number
  m = body.match(/\b(\d{4,6})\b/);
  if (m) return m[1];

  return '';
}


// ── Log to SMS_OTP_LOG sheet ──────────────────────────────
function otp_logToSheet_(service, otp, message, dateMs, phone) {
  try {
    const ss = getCore_();
    var sh = ss.getSheetByName('SMS_OTP_LOG');
    if (!sh) {
      sh = ss.insertSheet('SMS_OTP_LOG');
      sh.appendRow(['Timestamp', 'Service', 'OTP', 'Message', 'ServedAt', 'Phone']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([
      new Date(dateMs).toISOString(),
      service,
      otp,
      String(message || '').substring(0, 200),
      new Date().toISOString(),
      String(phone || ''),
    ]);
  } catch(e) {
    // never crash
  }
}


/**
 * Test function — run manually in GAS editor to verify Gmail reading works
 */
function testOtpGmail() {
  var result = getLatestOtp('Zee5', '9999999999');
  Logger.log(JSON.stringify(result));
  result = getLatestOtp('JioHotstar', '9999999999');
  Logger.log(JSON.stringify(result));
  result = getLatestOtp('SonyLIV', '9999999999');
  Logger.log(JSON.stringify(result));
}


function testQuotaDirect() {
  var r = getOtpQuota('9818196079', 'Zee5');
  Logger.log(JSON.stringify(r));
}


function testQuotaSettings() {
  Logger.log(getSetting_('OTP_QUOTA_Zee5', 'NOT_FOUND'));
  Logger.log(getSetting_('OTP_QUOTA_JioHotstar', 'NOT_FOUND'));
  Logger.log(getSetting_('OTP_QUOTA_SonyLIV', 'NOT_FOUND'));
}







function debugOtp() {
  // Step 1: check if function exists
  Logger.log("getLatestOtp exists: " + (typeof getLatestOtp === "function"));
  Logger.log("getOtpQuota exists: " + (typeof getOtpQuota === "function"));
  
  // Step 2: test quota
  var q = getOtpQuota("9818196079", "Zee5");
  Logger.log("quota result: " + JSON.stringify(q));
  
  // Step 3: test OTP fetch
  var r = getLatestOtp("Zee5", "9818196079");
  Logger.log("otp result: " + JSON.stringify(r));
}




function debugEmailTimes() {
  const query = 'from:harshwalia8888@gmail.com subject:SMSForwarder newer_than:1d';
  const threads = GmailApp.search(query, 0, 10);
  threads.forEach(function(t) {
    const msgs = t.getMessages();
    const msg = msgs[msgs.length - 1];
    Logger.log({
      subject: msg.getSubject(),
      date: msg.getDate(),
      ageMinutes: Math.round((Date.now() - msg.getDate().getTime()) / 60000),
      body: msg.getPlainBody().substring(0, 100)
    });
  });
}


