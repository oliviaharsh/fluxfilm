function getResumePaymentByPhone(phone){
  const last = getLastOrderByPhone(phone);
  if(!last) return null;

  const upiVpa = String(getSetting_("UPI_VPA","YOURUPI@BANK"));
  const payee = String(getSetting_("UPI_PAYEE_NAME","FluxFilm"));
  const currency = String(last.currency || getSetting_("CURRENCY","INR"));

  const tn = encodeURIComponent(last.orderId);
  const amt = Math.max(0, Number(last.finalAmount || 0));

  const upiLink =
    "upi://pay?pa=" + encodeURIComponent(upiVpa) +
    "&pn=" + encodeURIComponent(payee) +
    "&am=" + encodeURIComponent(amt) +
    "&cu=" + encodeURIComponent(currency) +
    "&tn=" + tn;

  return {
    orderId:last.orderId,
    service:last.service,
    plan:last.plan,
    finalAmount:amt,
    currency:currency,
    upiLink:upiLink
  };
}

function getLastOrderByPhone(phone) {
  phone = String(phone || "").replace(/\D/g, "");
  if (!phone) return null;

  const ss = getCore_();
  const sh = ss.getSheetByName("ORDERS");
  const vals = sh.getDataRange().getValues();
  const head = vals[0].map(String);

  const iOrderId = head.indexOf("OrderID");
  const iService = head.indexOf("Service");
  const iPlan = head.indexOf("Plan");
  const iPhone = head.indexOf("Phone");
  const iStatus = head.indexOf("Status");
  const iFinal = head.indexOf("FinalAmount");
  const iCurrency = head.indexOf("Currency");

  for (let r = vals.length - 1; r >= 1; r--) {
    if (String(vals[r][iPhone]).replace(/\D/g, "") !== phone) continue;

    const st = String(vals[r][iStatus]).toUpperCase();
    if (st === "CREATED" || st === "PAID") {
      return {
        orderId: vals[r][iOrderId],
        service: vals[r][iService],
        plan: vals[r][iPlan],
        finalAmount: Number(vals[r][iFinal] || 0),
        currency: vals[r][iCurrency] || "INR"
      };
    }
  }
  return null;
}
