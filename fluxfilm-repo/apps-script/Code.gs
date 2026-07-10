/*************** FluxFilm Backend — Config ***************/

// Core sheet (this spreadsheet)
const CORE_SS_ID = "1lvIaCUpQWQEwOjrPZ17SvQ1GXAWE-Lc8O5b1MJgYhYU";

// Bank spreadsheet (your EquitasTx spreadsheet) — put your bank spreadsheet ID here
// IMPORTANT: Bank sheet must be shared as Editor to support@fluxfilm.in
const BANK_SS_ID = "1bN1qYI3esCxKitony9k5DR8ZfPksjC9Nh_bfkp7ZWSc";
const BANK_SHEET_NAME = "EquitasTx";

// Tab names (must match your STEP 1/2 tabs)
const TAB_SETTINGS = "SETTINGS";
const TAB_PLANS = "PLANS";
const TAB_ORDERS = "ORDERS";
const TAB_SUBS = "SUBSCRIPTIONS";
const TAB_CUS = "CUSTOMERS";
const TAB_RENEWALS = "RENEWALS";
const TAB_INV_ACCOUNTS = "INVENTORY_ACCOUNTS";
const TAB_INV_PROFILES = "INVENTORY_PROFILES";
const TAB_INV_CAPACITY = "INVENTORY_CAPACITY";
const TAB_MANUAL = "MANUAL_TASKS";
const TAB_EMAIL_LOGS = "EMAIL_LOGS";
const TAB_INVOICE_LOGS = "INVOICE_LOGS";
const TAB_AI_LOGS = "AI_LOGS";
const TAB_COUPONS = "COUPONS";
const TAB_COUPON_USAGE = "COUPON_USAGE";
const TAB_TRENDING = "TRENDING";


// Basic statuses
const ORDER_STATUS = {
  CREATED: "CREATED",
  PAID: "PAID",
  CANCELLED: "CANCELLED",
};

const FULFILL_STATUS = {
  PENDING: "PENDING",
  FULFILLED: "FULFILLED",
  MANUAL_PENDING: "MANUAL_PENDING",
  FAILED: "FAILED",
};

const SUB_STATUS = {
  ACTIVE: "ACTIVE",
  EXPIRING: "EXPIRING",
  EXPIRED: "EXPIRED",
  SUSPENDED: "SUSPENDED",
};
