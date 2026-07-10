/*************** FluxFilm Backend — Email + Invoice (UPDATED with Renewal) ***************/

/**
 * Modes supported:
 * - "FULFILLED"      -> Access granted (includes creds if available)
 * - "MANUAL_PENDING" -> Payment received, manual activation
 * - "RENEWED"        -> Renewal confirmed + new expiry (includes creds if available)
 */
function sendOrderEmails_(orderId, mode) {
  const orders = sh_(TAB_ORDERS);
  const rn = getRowByValue_(orders, "OrderID", orderId);
  if (!rn) return;

  const { idx } = headerIndex_(orders);

  const email = String(orders.getRange(rn, idx["Email"] + 1).getValue() || "").trim();
  if (!email) return;

  const service = String(orders.getRange(rn, idx["Service"] + 1).getValue() || "").trim();
  const plan = String(orders.getRange(rn, idx["Plan"] + 1).getValue() || "").trim();
  const amount = asNumber_(orders.getRange(rn, idx["FinalAmount"] + 1).getValue());
  const currency = String(orders.getRange(rn, idx["Currency"] + 1).getValue() || "INR").trim();

  const extraKey = (idx["ExtraFieldKey"] != null) ? String(orders.getRange(rn, idx["ExtraFieldKey"] + 1).getValue() || "").trim() : "";
  const extraVal = (idx["ExtraFieldValue"] != null) ? String(orders.getRange(rn, idx["ExtraFieldValue"] + 1).getValue() || "").trim() : "";

  // Optional renew markers (if columns exist)
  const orderType = (idx["OrderType"] != null)
    ? String(orders.getRange(rn, idx["OrderType"] + 1).getValue() || "").trim().toUpperCase()
    : "";

  const renewSubId = (idx["RenewSubID"] != null)
    ? String(orders.getRange(rn, idx["RenewSubID"] + 1).getValue() || "").trim()
    : "";

  const isRenew = (mode === "RENEWED") || (orderType === "RENEW" && !!renewSubId);

  const fromName = String(getSetting_("FROM_NAME", "FluxFilm Support"));
  const supportEmail = String(getSetting_("SUPPORT_EMAIL", "support@fluxfilm.in"));
  const brand = String(getSetting_("BRAND_NAME", "FluxFilm"));
  const brandTagline = String(getSetting_("BRAND_TAGLINE", "Premium subscriptions • Instant support")).trim();

  // Post-payment message from PLANS
  let postPaymentMessage = "";
  try {
    const plansSh = sh_(TAB_PLANS);
    const { idx: pIdx } = headerIndex_(plansSh);
    const last = plansSh.getLastRow();
    if (last >= 2 && pIdx["Service"] != null && pIdx["Plan"] != null && pIdx["PostPaymentMessage"] != null) {
      const data = plansSh.getRange(2, 1, last - 1, plansSh.getLastColumn()).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][pIdx["Service"]] || "").trim() === service &&
            String(data[i][pIdx["Plan"]] || "").trim() === plan) {
          postPaymentMessage = String(data[i][pIdx["PostPaymentMessage"]] || "").trim();
          break;
        }
      }
    }
  } catch (e) {}

  // PostPaymentMessage (highlighted)
  let postPayBlock = "";
  if (postPaymentMessage) {
    const lines = postPaymentMessage
      .split(/\r?\n/)
      .map(s => String(s || "").trim())
      .filter(Boolean);

    const bullets = lines.length > 1
      ? `<ul style="margin:10px 0 0; padding-left:18px; color:#7c2d12; font-size:13px;">
           ${lines.map(x => `<li style="margin:6px 0;">${escapeHtml_(x)}</li>`).join("")}
         </ul>`
      : `<div style="margin-top:8px; color:#7c2d12; font-size:13px;"><b>${escapeHtml_(lines[0] || postPaymentMessage)}</b></div>`;

    postPayBlock = `
      <div style="margin-top:14px; padding:14px 14px; border-radius:14px; background:#fff7ed; border:1px solid #fed7aa;">
        <div style="font-weight:900; color:#9a3412;">IMPORTANT</div>
        ${bullets}
      </div>`;
  }

  // Pull access (for FULFILLED/RENEWED)
  const access = (mode === "FULFILLED" || mode === "RENEWED") ? getAccessFromSubsByOrderId_(orderId) : {};

  // Renewal info block: new expiry + SubID (if available)
  let renewalBlock = "";
  let renewalSubId = renewSubId;
  let renewalNewExpiry = "";
  let renewalOldExpiry = "";

  if (isRenew) {
    const info = getRenewInfoByOrder_(orderId, renewSubId);
    renewalSubId = info.subId || renewalSubId || "";
    renewalNewExpiry = info.newExpiry || "";
    renewalOldExpiry = info.oldExpiry || "";
    if (renewalSubId || renewalNewExpiry) {
      renewalBlock = `
        <div style="margin-top:14px; padding:14px 14px; border-radius:14px; border:1px solid #e5e7eb; background:#f8fafc;">
          <div style="font-weight:900; color:#0f172a;">🔁 Renewal confirmed</div>
          ${renewalSubId ? `<div style="margin-top:8px; font-size:14px;"><b>SubID:</b> <span style="font-family:monospace;">${escapeHtml_(renewalSubId)}</span></div>` : ""}
          ${renewalOldExpiry ? `<div style="margin-top:6px; font-size:13px; color:#334155;">Old expiry: <b>${escapeHtml_(renewalOldExpiry)}</b></div>` : ""}
          ${renewalNewExpiry ? `<div style="margin-top:6px; font-size:13px; color:#334155;">New expiry: <b>${escapeHtml_(renewalNewExpiry)}</b></div>` : ""}
          <div style="margin-top:10px; font-size:12px; color:#475569;">Your access remains the same. No need to change password.</div>
        </div>`;
    }
  }

  // Subject
  const subject =
    (mode === "RENEWED")
      ? `🔁 ${brand} — Renewal Confirmed (${service} • ${plan})`
      : (mode === "FULFILLED")
        ? `✅ ${brand} — Access Granted (${service} • ${plan})`
        : `✅ ${brand} — Payment Received (${service} • ${plan})`;

  // Invoice PDF
  let invNo = "";
  let pdfBlob = null;
  let pdfFileUrl = "";

  try {
    invNo = genInvoiceNo_();
    pdfBlob = generateInvoicePdf_(invNo, orderId, service, plan, amount, currency, email, extraKey, extraVal, {
      mode: mode,
      isRenew: isRenew,
      subId: renewalSubId || "",
      oldExpiry: renewalOldExpiry || "",
      newExpiry: renewalNewExpiry || ""
    });
    const saved = saveInvoicePdfToDrive_(invNo, pdfBlob);
    pdfFileUrl = saved && saved.url ? saved.url : "";
  } catch (e) {
    pdfBlob = null;
    pdfFileUrl = "";
  }

  // Badge
  const badgeText =
    (mode === "RENEWED") ? "RENEWED" :
    (mode === "FULFILLED") ? "ACCESS GRANTED" :
    "PAYMENT RECEIVED";

  const badgeBg =
    (mode === "RENEWED") ? "#7c3aed" :        // purple
    (mode === "FULFILLED") ? "#16a34a" :      // green
    "#0ea5e9";                                // blue

  const extraLine = (extraKey && extraVal)
    ? `<div style="margin-top:10px; padding:10px 12px; border:1px dashed #e5e7eb; border-radius:12px; background:#fafafa;">
         <div style="font-size:12px; color:#6b7280;">Extra info</div>
         <div style="margin-top:4px; font-size:14px; color:#111827;"><b>${escapeHtml_(extraKey)}:</b> ${escapeHtml_(extraVal)}</div>
       </div>`
    : "";

  // Access block
  let accessBlock = "";
  if ((mode === "FULFILLED" || mode === "RENEWED") && access && (access.user || access.pass || access.profileNumber || access.profileName || access.profilePin)) {
    const isNetflix = String(service).toLowerCase().includes("netflix");

    const profileLine =
      isNetflix && (access.profileName || access.profileNumber)
        ? `<div style="margin-top:8px; font-size:14px;"><b>Profile:</b> ${escapeHtml_(access.profileName || "Netflix")} ${access.profileNumber ? `(P${escapeHtml_(access.profileNumber)})` : ""}</div>`
        : "";

    const pinLine =
      isNetflix && access.profilePin
        ? `<div style="margin-top:8px; font-size:14px;"><b>Profile PIN:</b> <span style="font-family:monospace; letter-spacing:0.5px;">${escapeHtml_(access.profilePin)}</span></div>`
        : "";

    accessBlock = `
      <div style="margin-top:14px; padding:14px 14px; border-radius:14px; background:#0b1220; color:#ffffff;">
        <div style="font-size:13px; opacity:0.85;">Your access details</div>
        ${access.user ? `<div style="margin-top:10px; font-size:14px;"><b>Login ID:</b> <span style="font-family:monospace;">${escapeHtml_(access.user)}</span></div>` : ""}
        ${access.pass ? `<div style="margin-top:8px; font-size:14px;"><b>Password:</b> <span style="font-family:monospace;">${escapeHtml_(access.pass)}</span></div>` : ""}
        ${profileLine}
        ${pinLine}
        <div style="margin-top:10px; font-size:12px; opacity:0.85;">Don’t change password / email. If you face household/device issues, contact support.</div>
      </div>`;
  }

  // Main message
  const mainMessage =
    (mode === "RENEWED")
      ? `🔁 Your renewal payment is confirmed.`
      : (mode === "MANUAL_PENDING")
        ? `✅ Payment received! Your subscription will be activated soon.`
        : `✅ Payment verified! Your subscription is now active.`;

  const nextStep =
    (mode === "RENEWED")
      ? `✅ Your subscription has been extended successfully.`
      : (mode === "MANUAL_PENDING")
        ? `⌛ Activation is in progress. We’ll notify you once it’s done.`
        : `✅ You can start watching now.`;

  const htmlBody = `
  <div style="background:#f3f4f6; padding:24px 0; font-family:Arial, sans-serif;">
    <div style="max-width:620px; margin:0 auto; background:#ffffff; border-radius:18px; overflow:hidden; box-shadow:0 10px 28px rgba(0,0,0,0.08);">

      <div style="padding:18px 20px; background:#111827; color:#ffffff;">
        <div style="display:flex; justify-content:space-between; align-items:center;">
          <div>
            <div style="font-size:18px; font-weight:800;">${escapeHtml_(brand)}</div>
            <div style="margin-top:4px; font-size:12px; opacity:0.85;">${escapeHtml_(brandTagline || "")}</div>
          </div>
          <div style="padding:6px 10px; border-radius:999px; background:${badgeBg}; font-size:11px; font-weight:700; letter-spacing:0.6px;">
            ${badgeText}
          </div>
        </div>
      </div>

      <div style="padding:18px 20px;">
        <div style="font-size:16px; font-weight:800; color:#111827;">${mainMessage}</div>
        <div style="margin-top:6px; font-size:14px; color:#374151;">${nextStep}</div>

        <div style="margin-top:14px; padding:14px 14px; border:1px solid #e5e7eb; border-radius:14px;">
          <div style="display:flex; justify-content:space-between; gap:12px; flex-wrap:wrap;">
            <div style="min-width:220px;">
              <div style="font-size:12px; color:#6b7280;">Order ID</div>
              <div style="margin-top:4px; font-family:monospace; font-size:14px; color:#111827;"><b>${escapeHtml_(orderId)}</b></div>
            </div>
            <div style="text-align:right; min-width:160px;">
              <div style="font-size:12px; color:#6b7280;">Amount</div>
              <div style="margin-top:4px; font-size:14px; color:#111827;"><b>${escapeHtml_(currency)} ${Math.round(amount)}</b></div>
            </div>
          </div>

          <div style="margin-top:12px; border-top:1px dashed #e5e7eb; padding-top:12px;">
            <div style="font-size:12px; color:#6b7280;">Subscription</div>
            <div style="margin-top:4px; font-size:14px; color:#111827;"><b>${escapeHtml_(service)}</b> • ${escapeHtml_(plan)}</div>
          </div>

          ${extraLine}
        </div>

        ${renewalBlock}

        ${accessBlock}

        ${postPayBlock}

        <div style="margin-top:16px; padding:14px 14px; border-radius:14px; background:#fff7ed; border:1px solid #fed7aa;">
          <div style="font-weight:800; color:#9a3412;">✨ Helpful tips</div>
          <ul style="margin:8px 0 0; padding-left:18px; color:#7c2d12; font-size:13px;">
            <li>Keep this email safe (it includes your invoice).</li>
            <li>If you lose access later, use <b>Recover Access (OTP)</b> on the website.</li>
            <li>Need help? Reply to this email or contact support.</li>
          </ul>
        </div>

        <div style="margin-top:16px; font-size:12px; color:#6b7280;">
          Support: <b>${escapeHtml_(supportEmail)}</b><br/>
          — ${escapeHtml_(fromName)}
        </div>
      </div>

      <div style="padding:14px 20px; background:#f9fafb; font-size:11px; color:#6b7280;">
        This is an automated message from ${escapeHtml_(brand)}.
      </div>
    </div>
  </div>`;

  // Plain text fallback
  let body =
    `Hi,\n\n` +
    `Thanks for your ${isRenew ? "renewal" : "purchase"} on ${brand}.\n\n` +
    `Order ID: ${orderId}\n` +
    `Service: ${service}\n` +
    `Plan: ${plan}\n` +
    `Amount: ${currency} ${Math.round(amount)}\n\n` +
    (mode === "RENEWED"
      ? "Renewal confirmed. Your subscription has been extended.\n"
      : (mode === "MANUAL_PENDING"
          ? "Payment received. Activation will be done manually soon.\n"
          : "Payment verified. Your subscription is active.\n")) +
    `\nSupport: ${supportEmail}\n` +
    `— ${fromName}\n`;

  // Send email (HTML + attachment)
  try {
    const opts = { name: fromName, htmlBody: htmlBody, replyTo: supportEmail, from: supportEmail };
    if (pdfBlob) opts.attachments = [pdfBlob];

    GmailApp.sendEmail(email, subject, body, opts);
    logEmail_(orderId, email, mode, "SENT", "");
  } catch (e) {
    logEmail_(orderId, email, mode, "FAILED", String(e));
  }

  // Invoice log
  try {
    const invLog = sh_(TAB_INVOICE_LOGS);
    invLog.appendRow([
      invNo || genInvoiceNo_(),
      orderId,
      now_(),
      amount,
      currency,
      pdfFileUrl || "",
      (mode !== "") ? "YES" : "NO",
      (pdfBlob ? "SENT" : "LOGGED"),
      (pdfBlob ? "" : "PDF_NOT_GENERATED")
    ]);
  } catch (_) {}
}


function logEmail_(orderId, toEmail, emailType, status, err) {
  try {
    const log = sh_(TAB_EMAIL_LOGS);
    log.appendRow(["LOG-" + Date.now().toString().slice(-6), now_(), orderId, toEmail, emailType, status, err || ""]);
  } catch (_) {}
}

/**
 * ✅ Dedupe: has email already been SENT for this OrderID + Type?
 */
function wasEmailSent_(orderId, emailType) {
  try {
    const log = sh_(TAB_EMAIL_LOGS);
    const { idx } = headerIndex_(log);

    if (idx["OrderID"] == null || idx["Type"] == null || idx["Status"] == null) return false;

    const last = log.getLastRow();
    if (last < 2) return false;

    const data = log.getRange(2, 1, last - 1, log.getLastColumn()).getValues();
    const oid = String(orderId || "").trim();
    const typ = String(emailType || "").trim().toUpperCase();

    for (let i = data.length - 1; i >= 0; i--) {
      const r = data[i];
      if (String(r[idx["OrderID"]] || "").trim() !== oid) continue;
      if (String(r[idx["Type"]] || "").trim().toUpperCase() !== typ) continue;
      if (String(r[idx["Status"]] || "").trim().toUpperCase() === "SENT") return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}


/**
 * Read access from SUBSCRIPTIONS by OrderID.
 */
function getAccessFromSubsByOrderId_(orderId) {
  try {
    const subs = sh_(TAB_SUBS);
    const rn = getRowByValue_(subs, "OrderID", orderId);
    if (!rn) return {};

    const { idx } = headerIndex_(subs);

    function gv(col) {
      if (idx[col] == null) return "";
      return subs.getRange(rn, idx[col] + 1).getValue();
    }

    return {
      user: String(gv("LoginId") || ""),
      pass: String(gv("Password") || ""),
      profileNumber: gv("ProfileNumber") || "",
      profileName: String(gv("ProfileName") || ""),
      profilePin: String(gv("ProfilePIN") || "")
    };
  } catch (e) {
    return {};
  }
}


/**
 * ✅ Renewal info helper (SubID + expiry dates) for a given renew order
 * - If RenewSubID passed, read that sub row
 * - Else: try match SUBSCRIPTIONS row by OrderID (after your renew flow sets OrderID)
 */
function getRenewInfoByOrder_(orderId, renewSubId) {
  try {
    const subs = sh_(TAB_SUBS);
    const { idx } = headerIndex_(subs);

    let rn = 0;
    let subId = String(renewSubId || "").trim();

    if (subId) rn = getRowByValue_(subs, "SubID", subId);
    if (!rn) rn = getRowByValue_(subs, "OrderID", orderId);
    if (!rn) return { subId: subId, oldExpiry: "", newExpiry: "" };

    if (!subId && idx["SubID"] != null) subId = String(subs.getRange(rn, idx["SubID"] + 1).getValue() || "").trim();

    // We can’t always know old expiry unless you store it somewhere.
    // But in your renew flow you *do* have oldExpiry in code; if you later add a column, we can show it.
    const newExpiry = (idx["ExpiryDate"] != null) ? subs.getRange(rn, idx["ExpiryDate"] + 1).getValue() : "";

    return {
      subId: subId,
      oldExpiry: "",
      newExpiry: newExpiry ? String(newExpiry) : ""
    };
  } catch (e) {
    return { subId: String(renewSubId || ""), oldExpiry: "", newExpiry: "" };
  }
}


/**
 * Generate PDF invoice (updated: can show renewal meta)
 */
function generateInvoicePdf_(invNo, orderId, service, plan, amount, currency, customerEmail, extraKey, extraVal, meta) {
  meta = meta || {};
  const fromName = String(getSetting_("FROM_NAME", "FluxFilm Support"));
  const supportEmail = String(getSetting_("SUPPORT_EMAIL", "support@fluxfilm.in"));
  const issuedAt = now_();

  const isRenew = !!meta.isRenew;

  const extraLine = (extraKey && extraVal)
    ? `<tr><td style="padding:6px 0;color:#444;">${escapeHtml_(extraKey)}</td><td style="padding:6px 0;text-align:right;color:#444;">${escapeHtml_(extraVal)}</td></tr>`
    : ``;

  const renewLine = isRenew
    ? `
      <tr><td style="padding:6px 0;color:#444;">Type</td><td style="padding:6px 0;text-align:right;color:#444;"><b>Renewal</b></td></tr>
      ${meta.subId ? `<tr><td style="padding:6px 0;color:#444;">SubID</td><td style="padding:6px 0;text-align:right;color:#444;">${escapeHtml_(meta.subId)}</td></tr>` : ``}
      ${meta.newExpiry ? `<tr><td style="padding:6px 0;color:#444;">New expiry</td><td style="padding:6px 0;text-align:right;color:#444;">${escapeHtml_(meta.newExpiry)}</td></tr>` : ``}
    `
    : `<tr><td style="padding:6px 0;color:#444;">Type</td><td style="padding:6px 0;text-align:right;color:#444;"><b>New Purchase</b></td></tr>`;

  const html = `
  <html>
    <body style="font-family: Arial, sans-serif; color:#111;">
      <div style="max-width:720px; margin:0 auto; padding:24px;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start;">
          <div>
            <h2 style="margin:0;">FluxFilm</h2>
            <div style="margin-top:6px; color:#555;">${escapeHtml_(fromName)}</div>
            <div style="color:#555;">Support: ${escapeHtml_(supportEmail)}</div>
          </div>
          <div style="text-align:right;">
            <div style="font-size:14px; color:#555;">Invoice</div>
            <div style="font-size:20px; font-weight:bold;">${escapeHtml_(invNo)}</div>
            <div style="font-size:12px; color:#777; margin-top:4px;">${issuedAt}</div>
          </div>
        </div>

        <hr style="margin:18px 0; border:none; border-top:1px solid #ddd;" />

        <table style="width:100%; border-collapse:collapse;">
          <tr><td style="padding:6px 0; color:#444;">Order ID</td><td style="padding:6px 0; text-align:right;">${escapeHtml_(orderId)}</td></tr>
          <tr><td style="padding:6px 0; color:#444;">Customer</td><td style="padding:6px 0; text-align:right;">${escapeHtml_(customerEmail)}</td></tr>
          ${renewLine}
        </table>

        <h3 style="margin:18px 0 8px;">Details</h3>
        <table style="width:100%; border-collapse:collapse;">
          <tr>
            <th style="text-align:left; padding:8px 0; border-bottom:1px solid #eee;">Item</th>
            <th style="text-align:right; padding:8px 0; border-bottom:1px solid #eee;">Amount</th>
          </tr>
          <tr>
            <td style="padding:10px 0;">
              <div style="font-weight:bold;">${escapeHtml_(service)} — ${escapeHtml_(plan)}</div>
              <div style="color:#666; font-size:12px;">Digital service</div>
            </td>
            <td style="padding:10px 0; text-align:right; font-weight:bold;">${escapeHtml_(currency)} ${Math.round(amount)}</td>
          </tr>
          ${extraLine}
          <tr>
            <td style="padding:10px 0; border-top:1px solid #eee; font-weight:bold;">Total</td>
            <td style="padding:10px 0; border-top:1px solid #eee; text-align:right; font-weight:bold;">${escapeHtml_(currency)} ${Math.round(amount)}</td>
          </tr>
        </table>

        <div style="margin-top:18px; font-size:12px; color:#666;">
          This is a system-generated invoice for your FluxFilm transaction.
        </div>
      </div>
    </body>
  </html>`;

  const out = HtmlService.createHtmlOutput(html).getAs(MimeType.PDF);
  out.setName(`${invNo}_${orderId}.pdf`);
  return out;
}


/**
 * Save PDF to Drive
 */
function saveInvoicePdfToDrive_(invNo, pdfBlob) {
  try {
    const folderId = String(getSetting_("INVOICE_FOLDER_ID", "")).trim();
    let file;
    if (folderId) {
      const folder = DriveApp.getFolderById(folderId);
      file = folder.createFile(pdfBlob);
    } else {
      file = DriveApp.createFile(pdfBlob);
    }
    file.setName(pdfBlob.getName() || `${invNo}.pdf`);
    return { id: file.getId(), url: file.getUrl() };
  } catch (e) {
    return { id: "", url: "" };
  }
}

function escapeHtml_(s) {
  s = String(s || "");
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
