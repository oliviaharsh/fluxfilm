function DEBUG_getMySubscriptions() {
  const TEST_PHONE = "9818196079"; // 👈 put a real phone from SUBSCRIPTIONS

  const phone = String(TEST_PHONE).replace(/\D/g, "");
  const sh = sh_(TAB_SUBS);
  const { idx } = headerIndex_(sh);
  const rows = sh.getDataRange().getValues();

  Logger.log("Phone searching for: " + phone);
  Logger.log("Total rows in SUBSCRIPTIONS: " + (rows.length - 1));

  let found = 0;

  for (let r = 1; r < rows.length; r++) {
    const rowPhone = String(rows[r][idx["Phone"]] || "").replace(/\D/g, "");

    if (rowPhone === phone) {
      found++;
      Logger.log("MATCH @ row " + (r + 1));
      Logger.log({
        SubID: rows[r][idx["SubID"]],
        OrderID: rows[r][idx["OrderID"]],
        Service: rows[r][idx["Service"]],
        Plan: rows[r][idx["Plan"]],
        Status: rows[r][idx["Status"]],
        StartDate: rows[r][idx["StartDate"]],
        ExpiryDate: rows[r][idx["ExpiryDate"]]
      });
    }
  }

  Logger.log("TOTAL MATCHES FOUND: " + found);
}


function DEBUG_manageFetch() {
  const r = getMySubscriptions("9818196079");
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}


function TEST_manage(){
  Logger.log(mg_getMySubscriptions_("9818196079"));
}


function ADMIN_markPaidAndFulfill(orderId, txnRef) {
  const oid = String(orderId || "").trim();
  const ref = String(txnRef || "").trim();
  if (!oid) return { ok:false, message:"OrderID required" };

  const sh = sh_(TAB_ORDERS);
  const row = getRowByValue_(sh, "OrderID", oid);
  if (!row) return { ok:false, message:"Order not found: " + oid };

  const h = headerIndex_(sh).idx;
  const get = (name) => (h[name] != null ? sh.getRange(row, h[name] + 1).getValue() : "");
  const set = (name, val) => { if (h[name] != null) sh.getRange(row, h[name] + 1).setValue(val); };

  const curFulfill = String(get("FulfillmentStatus") || "").toUpperCase().trim();

  if (curFulfill === "FULFILLED") {
    const orderType = String(get("OrderType") || "").toUpperCase().trim();
    if (orderType !== "RENEW") {
      return { ok:true, message:"Already fulfilled. Not re-fulfilling.", orderId: oid };
    }
    // RENEW orders: clear FULFILLED flag so fulfillRenew_ can re-run and fix expiry if needed
    set("FulfillmentStatus", "PENDING");
  }
  set("Status", "PAID");
  if (ref) set("TxnRef", ref);
  set("VerifiedAt", new Date());
  set("Error", "");

  updateRenewalStatusByOrder_(oid, "PAID");

  if (!curFulfill) set("FulfillmentStatus", "PENDING");

  return fulfillOrder(oid);
}


function RUN_ONCE() {
  return ADMIN_markPaidAndFulfill("FF5331271");
}








function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("FluxFilm Admin")
    .addItem("Mark Paid + Fulfill", "ADMIN_menuMarkPaidAndFulfill")
    .addSeparator()
    .addItem("Create Paid + Fulfill Order", "ADMIN_promptCreatePaidAndFulfillOrder")
    .addItem("Create Paid + Fulfill Existing Customer", "ADMIN_promptCreatePaidAndFulfillExistingCustomer")
    .addItem("Create Renew Order", "ADMIN_promptCreateRenewOrder")
    .addSeparator()
    .addItem("Run Test Order", "RUN_ONCE")
    .addToUi();
}

function ADMIN_menuMarkPaidAndFulfill() {
  const ui = SpreadsheetApp.getUi();

  const orderRes = ui.prompt("Mark Paid + Fulfill", "Enter Order ID", ui.ButtonSet.OK_CANCEL);
  if (orderRes.getSelectedButton() !== ui.Button.OK) return;

  const orderId = String(orderRes.getResponseText() || "").trim();
  if (!orderId) {
    ui.alert("Order ID is required.");
    return;
  }

  const txnRes = ui.prompt("Mark Paid + Fulfill", "Enter TxnRef (optional)", ui.ButtonSet.OK_CANCEL);
  if (txnRes.getSelectedButton() !== ui.Button.OK) return;

  const txnRef = String(txnRes.getResponseText() || "").trim();

  const result = ADMIN_markPaidAndFulfill(orderId, txnRef);

  ui.alert(
    result && result.ok
      ? "✅ Success\n\n" + (result.message || "")
      : "❌ Failed\n\n" + ((result && result.message) || "Unknown error")
  );
}

function ADMIN_promptCreatePaidAndFulfillOrder() {
  const ui = SpreadsheetApp.getUi();

  const service = ui.prompt("Create + Fulfill Order", "Enter Service", ui.ButtonSet.OK_CANCEL);
  if (service.getSelectedButton() !== ui.Button.OK) return;
  const serviceVal = String(service.getResponseText() || "").trim();

  const plan = ui.prompt("Create + Fulfill Order", "Enter Plan", ui.ButtonSet.OK_CANCEL);
  if (plan.getSelectedButton() !== ui.Button.OK) return;
  const planVal = String(plan.getResponseText() || "").trim();

  let extraFieldKey = "";
  let extraFieldValue = "";

  if (serviceVal.toLowerCase().includes("prime")) {
    const device = ui.prompt(
      "Prime Device Type",
      "Enter TV or NON_TV",
      ui.ButtonSet.OK_CANCEL
    );
    if (device.getSelectedButton() !== ui.Button.OK) return;

    extraFieldKey = "DeviceType";
    extraFieldValue = String(device.getResponseText() || "").trim().toUpperCase();

    if (extraFieldValue !== "TV" && extraFieldValue !== "NON_TV") {
      ui.alert("Invalid device type. Please enter exactly TV or NON_TV.");
      return;
    }
  }

  const name = ui.prompt("Create + Fulfill Order", "Enter Customer Full Name", ui.ButtonSet.OK_CANCEL);
  if (name.getSelectedButton() !== ui.Button.OK) return;
  const nameVal = String(name.getResponseText() || "").trim();

  const phone = ui.prompt("Create + Fulfill Order", "Enter Phone Number", ui.ButtonSet.OK_CANCEL);
  if (phone.getSelectedButton() !== ui.Button.OK) return;
  const phoneVal = String(phone.getResponseText() || "").trim();

  const email = ui.prompt("Create + Fulfill Order", "Enter Email", ui.ButtonSet.OK_CANCEL);
  if (email.getSelectedButton() !== ui.Button.OK) return;
  const emailVal = String(email.getResponseText() || "").trim();

  const coupon = ui.prompt("Create + Fulfill Order", "Enter Coupon Code (leave blank if none)", ui.ButtonSet.OK_CANCEL);
  if (coupon.getSelectedButton() !== ui.Button.OK) return;
  const couponVal = String(coupon.getResponseText() || "").trim();

  const notes = ui.prompt("Create + Fulfill Order", "Enter Notes (optional)", ui.ButtonSet.OK_CANCEL);
  if (notes.getSelectedButton() !== ui.Button.OK) return;
  const notesVal = String(notes.getResponseText() || "").trim();

  const txn = ui.prompt("Create + Fulfill Order", "Enter TxnRef (leave blank if none)", ui.ButtonSet.OK_CANCEL);
  if (txn.getSelectedButton() !== ui.Button.OK) return;
  const txnVal = String(txn.getResponseText() || "").trim();

  const res = createOrder({
    service: serviceVal,
    plan: planVal,
    name: nameVal,
    email: emailVal,
    phone: phoneVal,
    couponCode: couponVal,
    notes: notesVal,
    extraFieldKey: extraFieldKey,
    extraFieldValue: extraFieldValue
  });

  if (!res || !res.ok) {
    ui.alert("Order Create Failed", res && res.message ? res.message : "Unknown error", ui.ButtonSet.OK);
    return res;
  }

  const fulfillRes = ADMIN_markPaidAndFulfill(res.orderId, txnVal);

  ui.alert(
    "Done ✅",
    "OrderID: " + res.orderId +
    "\nCreate: " + (res.ok ? "OK" : "FAILED") +
    "\nFulfill: " + ((fulfillRes && fulfillRes.ok) ? "OK" : "FAILED") +
    "\nMessage: " + (fulfillRes && fulfillRes.message ? fulfillRes.message : ""),
    ui.ButtonSet.OK
  );

  return {
    create: res,
    fulfill: fulfillRes
  };
}

function ADMIN_promptCreateRenewOrder() {
  const ui = SpreadsheetApp.getUi();

  const sub = ui.prompt("Create Renew Order", "Enter SubID to renew", ui.ButtonSet.OK_CANCEL);
  if (sub.getSelectedButton() !== ui.Button.OK) return;

  const subId = String(sub.getResponseText() || "").trim();
  if (!subId) {
    ui.alert("SubID is required.");
    return;
  }

  const res = createRenewOrder(subId);

  if (!res || !res.ok) {
    ui.alert("Renew Order Failed", res && res.message ? res.message : "Unknown error", ui.ButtonSet.OK);
    return res;
  }

  ui.alert(
    "Renew Order Created ✅",
    "OrderID: " + res.orderId +
    "\nSubID: " + subId +
    "\nFinal Amount: " + res.finalAmount,
    ui.ButtonSet.OK
  );

  return res;
}

function ADMIN_markPaidAndFulfill(orderId, txnRef) {
  const oid = String(orderId || "").trim();
  const ref = String(txnRef || "").trim();
  if (!oid) return { ok:false, message:"OrderID required" };

  const sh = sh_(TAB_ORDERS);
  const row = getRowByValue_(sh, "OrderID", oid);
  if (!row) return { ok:false, message:"Order not found: " + oid };

  const h = headerIndex_(sh).idx;

  const get = (name) =>
    h[name] != null ? sh.getRange(row, h[name] + 1).getValue() : "";

  const set = (name, val) => {
    if (h[name] != null) sh.getRange(row, h[name] + 1).setValue(val);
  };

  const curFulfill = String(get("FulfillmentStatus") || "").toUpperCase().trim();

  // NEW — allows admin to re-run stuck RENEW orders:
  if (curFulfill === "FULFILLED") {
    const orderType = String(get("OrderType") || "").toUpperCase().trim();
    if (orderType !== "RENEW") {
      return { ok:true, message:"Already fulfilled. Not re-fulfilling.", orderId: oid };
    }
    // RENEW: reset to PENDING so fulfillRenew_ can re-run and fix the expiry
    set("FulfillmentStatus", "PENDING");
  }

  set("Status", "PAID");
  if (ref) set("TxnRef", ref);
  set("VerifiedAt", new Date());
  set("Error", "");

  if (typeof updateRenewalStatusByOrder_ === "function") {
    updateRenewalStatusByOrder_(oid, "PAID");
  }

  if (!curFulfill) set("FulfillmentStatus", "PENDING");

  return fulfillOrder(oid);
}

function ADMIN_promptCreatePaidAndFulfillExistingCustomer() {
  const ui = SpreadsheetApp.getUi();

  const phoneRes = ui.prompt("Existing Customer Order", "Enter Customer Phone Number", ui.ButtonSet.OK_CANCEL);
  if (phoneRes.getSelectedButton() !== ui.Button.OK) return;

  const phoneVal = String(phoneRes.getResponseText() || "").replace(/\D/g, "").slice(-10);
  if (!phoneVal) {
    ui.alert("Phone number is required.");
    return;
  }

  const customer = ADMIN_findCustomerByPhone_(phoneVal);
  if (!customer.ok) {
    ui.alert("Customer Not Found", customer.message, ui.ButtonSet.OK);
    return;
  }

  const service = ui.prompt("Existing Customer Order", "Enter Service", ui.ButtonSet.OK_CANCEL);
  if (service.getSelectedButton() !== ui.Button.OK) return;
  const serviceVal = String(service.getResponseText() || "").trim();

  const plan = ui.prompt("Existing Customer Order", "Enter Plan", ui.ButtonSet.OK_CANCEL);
  if (plan.getSelectedButton() !== ui.Button.OK) return;
  const planVal = String(plan.getResponseText() || "").trim();

  let extraFieldKey = "";
  let extraFieldValue = "";

  if (serviceVal.toLowerCase().includes("prime")) {
    const device = ui.prompt("Prime Device Type", "Enter TV or NON_TV", ui.ButtonSet.OK_CANCEL);
    if (device.getSelectedButton() !== ui.Button.OK) return;

    extraFieldKey = "DeviceType";
    extraFieldValue = String(device.getResponseText() || "").trim().toUpperCase();

    if (extraFieldValue !== "TV" && extraFieldValue !== "NON_TV") {
      ui.alert("Invalid device type. Enter exactly TV or NON_TV.");
      return;
    }
  }

  const coupon = ui.prompt("Existing Customer Order", "Enter Coupon Code (leave blank if none)", ui.ButtonSet.OK_CANCEL);
  if (coupon.getSelectedButton() !== ui.Button.OK) return;
  const couponVal = String(coupon.getResponseText() || "").trim();

  const notes = ui.prompt("Existing Customer Order", "Enter Notes (optional)", ui.ButtonSet.OK_CANCEL);
  if (notes.getSelectedButton() !== ui.Button.OK) return;
  const notesVal = String(notes.getResponseText() || "").trim();

  const txn = ui.prompt("Existing Customer Order", "Enter TxnRef (leave blank if none)", ui.ButtonSet.OK_CANCEL);
  if (txn.getSelectedButton() !== ui.Button.OK) return;
  const txnVal = String(txn.getResponseText() || "").trim();

  const confirm = ui.alert(
    "Confirm Existing Customer",
    "Name: " + customer.name +
    "\nEmail: " + customer.email +
    "\nPhone: " + phoneVal +
    "\n\nCreate paid order?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const res = createOrder({
    service: serviceVal,
    plan: planVal,
    name: customer.name,
    email: customer.email,
    phone: phoneVal,
    couponCode: couponVal,
    notes: notesVal,
    extraFieldKey: extraFieldKey,
    extraFieldValue: extraFieldValue
  });

  if (!res || !res.ok) {
    ui.alert("Order Create Failed", res && res.message ? res.message : "Unknown error", ui.ButtonSet.OK);
    return res;
  }

  const fulfillRes = ADMIN_markPaidAndFulfill(res.orderId, txnVal);

  ui.alert(
    "Done ✅",
    "OrderID: " + res.orderId +
    "\nCustomer: " + customer.name +
    "\nEmail: " + customer.email +
    "\nCreate: OK" +
    "\nFulfill: " + ((fulfillRes && fulfillRes.ok) ? "OK" : "FAILED") +
    "\nMessage: " + (fulfillRes && fulfillRes.message ? fulfillRes.message : ""),
    ui.ButtonSet.OK
  );

  return {
    customer: customer,
    create: res,
    fulfill: fulfillRes
  };
}


function ADMIN_findCustomerByPhone_(phone) {
  const ph = String(phone || "").replace(/\D/g, "").slice(-10);
  if (!ph) return { ok:false, message:"Phone required." };

  // 1) Search SUBSCRIPTIONS first
  const fromSubs = ADMIN_findCustomerInSheetByPhone_(TAB_SUBS, ph);
  if (fromSubs.ok) return fromSubs;

  // 2) If not found, search ORDERS
  const fromOrders = ADMIN_findCustomerInSheetByPhone_(TAB_ORDERS, ph);
  if (fromOrders.ok) return fromOrders;

  return {
    ok:false,
    message:
      "No customer found with phone: " + ph +
      "\n\nChecked SUBSCRIPTIONS and ORDERS." +
      "\nMake sure the phone number is saved in Phone column."
  };
}

function ADMIN_findCustomerInSheetByPhone_(sheetName, ph) {
  const sh = sh_(sheetName);
  const { idx } = headerIndex_(sh);

  if (idx["Phone"] == null) return { ok:false, message: sheetName + " missing Phone header." };
  if (idx["Name"] == null) return { ok:false, message: sheetName + " missing Name header." };
  if (idx["Email"] == null) return { ok:false, message: sheetName + " missing Email header." };

  const last = sh.getLastRow();
  if (last < 2) return { ok:false, message:"No rows in " + sheetName };

  const rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();

  // bottom to top = latest record first
  for (let i = rows.length - 1; i >= 0; i--) {
    const rowPhone = String(rows[i][idx["Phone"]] || "").replace(/\D/g, "").slice(-10);

    if (rowPhone !== ph) continue;

    const name = String(rows[i][idx["Name"]] || "").trim();
    const email = String(rows[i][idx["Email"]] || "").trim();

    if (!name) return { ok:false, message:"Customer found in " + sheetName + " but Name is blank." };
    if (!email) return { ok:false, message:"Customer found in " + sheetName + " but Email is blank." };

    return {
      ok:true,
      source: sheetName,
      phone: ph,
      name: name,
      email: email
    };
  }

  return { ok:false, message:"Not found in " + sheetName };
}


function RUN_ONCE() {
  return ADMIN_markPaidAndFulfill("FF5331271", "TEST");
}












function buildCustomersSheet() {
  const ss = getCore_(); // uses your existing CORE_SS_ID
  const subs = ss.getSheetByName("SUBSCRIPTIONS");
  const orders = ss.getSheetByName("ORDERS");

  let customers = ss.getSheetByName("CUSTOMERS");
  if (!customers) customers = ss.insertSheet("CUSTOMERS");

  const headers = [
    "CustomerID","MemberSince","UpdatedAt","Name","Phone","Email",
    "LastOrderID","TotalOrders","TotalSpent","lastActivity","Notes","Status"
  ];

  customers.clearContents();
  customers.getRange(1, 1, 1, headers.length).setValues([headers]);

  const sData = subs.getDataRange().getValues();
  const sHead = sData[0].map(String);
  const si = {};
  sHead.forEach((h, i) => si[h.trim()] = i);

  const oData = orders.getDataRange().getValues();
  const oHead = oData[0].map(String);
  const oi = {};
  oHead.forEach((h, i) => oi[h.trim()] = i);

  const map = {};

  function normPhone(p) {
    const s = String(p || "").replace(/\D/g, "");
    return s.length > 10 ? s.slice(-10) : s;
  }

  function normEmail(e) {
    return String(e || "").trim().toLowerCase();
  }

  function asDate(v) {
    if (v instanceof Date && !isNaN(v)) return v;
    const d = new Date(v);
    return isNaN(d) ? null : d;
  }

  function asNumber(v) {
    const n = Number(String(v || "").replace(/[,₹ ]/g, ""));
    return isNaN(n) ? 0 : n;
  }

  function getCustomerId(phone) {
    return "CUS-" + String(phone || "").slice(-6);
  }

  // 1) Build customer base from SUBSCRIPTIONS
  for (let r = 1; r < sData.length; r++) {
    const row = sData[r];
    const phone = normPhone(row[si.Phone]);
    if (!phone) continue;

    const email = normEmail(row[si.Email]);
    const name = String(row[si.Name] || "").trim();
    const orderId = String(row[si.OrderID] || "").trim();
    const start = asDate(row[si.StartDate]);
    const expiry = asDate(row[si.ExpiryDate]);
    const status = String(row[si.Status] || "").trim().toUpperCase();

    if (!map[phone]) {
      map[phone] = {
        customerId: getCustomerId(phone),
        memberSince: start || "",
        updatedAt: new Date(),
        name,
        phone,
        email,
        lastOrderId: orderId,
        totalOrders: 0,
        totalSpent: 0,
        lastActivity: start || "",
        notes: "",
        status: "INACTIVE",
        hasActive: false
      };
    }

    const c = map[phone];

    if (!c.name && name) c.name = name;
    if (!c.email && email) c.email = email;

    if (start && (!c.memberSince || start < c.memberSince)) c.memberSince = start;

    const activityDate = expiry || start;
    if (activityDate && (!c.lastActivity || activityDate > c.lastActivity)) {
      c.lastActivity = activityDate;
      c.lastOrderId = orderId;
    }

    if (status === "ACTIVE") c.hasActive = true;
  }

  // 2) Add order count + paid spend from ORDERS
  for (let r = 1; r < oData.length; r++) {
    const row = oData[r];
    const phone = normPhone(row[oi.Phone]);
    if (!phone) continue;

    const status = String(row[oi.Status] || "").trim().toUpperCase();
    const fulfill = String(row[oi.FulfillmentStatus] || "").trim().toUpperCase();
    const finalAmount = asNumber(row[oi.FinalAmount]);
    const createdAt = asDate(row[oi.CreatedAt]);
    const verifiedAt = asDate(row[oi.VerifiedAt]);
    const orderId = String(row[oi.OrderID] || "").trim();

if (!map[phone]) continue;

    const c = map[phone];

    if (!c.name && oi.Name != null) c.name = String(row[oi.Name] || "").trim();
    if (!c.email && oi.Email != null) c.email = normEmail(row[oi.Email]);

    if (status === "PAID" || status === "FULFILLED" || fulfill === "FULFILLED" || fulfill === "MANUAL_PENDING") {
      c.totalOrders += 1;
      c.totalSpent += finalAmount; // FinalAmount already excludes discount
    }

    const activityDate = verifiedAt || createdAt;
    if (activityDate && (!c.lastActivity || activityDate > c.lastActivity)) {
      c.lastActivity = activityDate;
      c.lastOrderId = orderId;
    }

    if (createdAt && (!c.memberSince || createdAt < c.memberSince)) {
      c.memberSince = createdAt;
    }
  }

  // 3) Final rows
  const rows = Object.values(map).map(c => {
    c.status = c.hasActive ? "ACTIVE" : "INACTIVE";

    return [
      c.customerId,
      c.memberSince,
      new Date(),
      c.name,
      c.phone,
      c.email || "",
      c.lastOrderId,
      c.totalOrders,
      c.totalSpent,
      c.lastActivity,
      c.notes,
      c.status
    ];
  });

  rows.sort((a, b) => String(a[4]).localeCompare(String(b[4])));

  if (rows.length) {
    customers.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  customers.autoResizeColumns(1, headers.length);

  return `CUSTOMERS built ✅ ${rows.length} unique customers`;
}


