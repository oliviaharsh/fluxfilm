/*************** FluxFilm Backend — Coins (FRESH, SELF-HEALING) ***************/

const TAB_WALLET = "WALLET";
const TAB_COINS_LEDGER = "COINS_LEDGER";

/**
 * Public: get wallet by phone for UI
 * Always returns ok:true (never throws).
 */
function getWalletByPhone(phone){
  try{
    const ph = normalizePhoneCoins_(phone);
    if(!ph) return bad_("Phone required.");

    // Ensure sheets + headers exist (self-heal)
    ensureCoinsSheets_();

    const wallet = getCore_().getSheetByName(TAB_WALLET);
    const { idx } = headerIndex_(wallet);

    const rn = getRowByValueSafe_(wallet, "Phone", ph);
    if(!rn){
      // no wallet row yet -> return zeros
      return ok_({
        phone: ph,
        coinsBalance: 0,
        coinsLifetime: 0,
        lastEarnedAt: "",
        lastSpentAt: "",
        lastEvent: ""
      });
    }

    return ok_({
      phone: ph,
      coinsBalance: asNumber_(wallet.getRange(rn, idx["CoinsBalance"]+1).getValue()),
      coinsLifetime: asNumber_(wallet.getRange(rn, idx["CoinsLifetime"]+1).getValue()),
      lastEarnedAt: (idx["LastEarnedAt"]!=null) ? wallet.getRange(rn, idx["LastEarnedAt"]+1).getValue() : "",
      lastSpentAt: (idx["LastSpentAt"]!=null) ? wallet.getRange(rn, idx["LastSpentAt"]+1).getValue() : "",
      lastEvent: (idx["LastEvent"]!=null) ? String(wallet.getRange(rn, idx["LastEvent"]+1).getValue()||"") : ""
    });

  }catch(e){
    return bad_("Could not load wallet: " + (e && e.message ? e.message : e));
  }
}

/**
 * Award coins for an order (idempotent)
 * event: "NEW_PURCHASE" | "RENEWAL" | ...
 */
function awardCoinsForOrder_(event, orderId, phone, service, plan, amount, note){
  try{
    const ph = normalizePhoneCoins_(phone);
    const oid = String(orderId || "").trim();
    if(!ph || !oid) return;

    ensureCoinsSheets_();

    // idempotent check
    if (coinsLedgerHas_(event, oid)) return;

    const coins = computeCoins_(event, amount);
    if (coins <= 0) return;

    const ss = getCore_();
    const wallet = ss.getSheetByName(TAB_WALLET);
    const ledger = ss.getSheetByName(TAB_COINS_LEDGER);

    const w = upsertWallet_(wallet, ph);

    const newLifetime = asNumber_(w.coinsLifetime) + coins;
    const newBalance  = asNumber_(w.coinsBalance) + coins;

    setCellByHeader_(wallet, w.rowNum, "CoinsLifetime", newLifetime);
    setCellByHeader_(wallet, w.rowNum, "CoinsBalance", newBalance);

    if (hasHeaderSafe_(wallet, "LastEarnedAt")) setCellByHeader_(wallet, w.rowNum, "LastEarnedAt", now_());
    if (hasHeaderSafe_(wallet, "LastEvent")) setCellByHeader_(wallet, w.rowNum, "LastEvent", `${String(event||"")}:${oid}`);

    appendCoinsLedger_(ledger, {
      ts: now_(),
      event: String(event||""),
      orderId: oid,
      phone: ph,
      service: String(service||""),
      plan: String(plan||""),
      amount: asNumber_(amount),
      coinsDelta: coins,
      balanceAfter: newBalance,
      note: note || ""
    });

  }catch(e){
    // never block purchase flow
  }
}

/**
 * Optional: spend coins (future)
 */
function spendCoins_(phone, coinsToSpend, event, spendRef, note){
  try{
    const ph = normalizePhoneCoins_(phone);
    const coins = Math.max(0, Math.floor(asNumber_(coinsToSpend)));
    const ev = String(event || "").trim();
    const ref = String(spendRef || "").trim();

    if(!ph) return bad_("Phone required.");
    if(coins <= 0) return bad_("Invalid coins.");
    if(!ev) return bad_("Event required.");

    ensureCoinsSheets_();

    // idempotent if ref present
    if(ref && coinsLedgerHas_(ev, ref)) return bad_("Already processed.");

    const ss = getCore_();
    const wallet = ss.getSheetByName(TAB_WALLET);
    const ledger = ss.getSheetByName(TAB_COINS_LEDGER);

    const w = upsertWallet_(wallet, ph);
    const bal = asNumber_(w.coinsBalance);
    if(bal < coins) return bad_("Not enough coins.");

    const newBalance = bal - coins;

    setCellByHeader_(wallet, w.rowNum, "CoinsBalance", newBalance);
    if (hasHeaderSafe_(wallet, "LastSpentAt")) setCellByHeader_(wallet, w.rowNum, "LastSpentAt", now_());
    if (hasHeaderSafe_(wallet, "LastEvent")) setCellByHeader_(wallet, w.rowNum, "LastEvent", `${ev}:${ref || ""}`);

    appendCoinsLedger_(ledger, {
      ts: now_(),
      event: ev,
      orderId: ref,
      phone: ph,
      service: "",
      plan: "",
      amount: 0,
      coinsDelta: -coins,
      balanceAfter: newBalance,
      note: note || ""
    });

    return ok_({ phone: ph, spent: coins, balance: newBalance });

  }catch(e){
    return bad_("Spend error: " + (e && e.message ? e.message : e));
  }
}

/* -------------------- SELF-HEAL / INTERNALS -------------------- */

function ensureCoinsSheets_(){
  const ss = getCore_();

  // WALLET
  let w = ss.getSheetByName(TAB_WALLET);
  if(!w) w = ss.insertSheet(TAB_WALLET);
  ensureHeaders_(w, ["Phone","CoinsBalance","CoinsLifetime","LastEarnedAt","LastSpentAt","LastEvent"]);

  // COINS_LEDGER
  let l = ss.getSheetByName(TAB_COINS_LEDGER);
  if(!l) l = ss.insertSheet(TAB_COINS_LEDGER);
  ensureHeaders_(l, ["Ts","Event","OrderID","Phone","Service","Plan","Amount","CoinsDelta","BalanceAfter","Note"]);
}

function ensureHeaders_(sheet, wanted){
  const lastCol = Math.max(sheet.getLastColumn(), wanted.length);
  if(sheet.getLastRow() < 1) sheet.appendRow(new Array(lastCol).fill(""));

  const row1 = sheet.getRange(1,1,1,lastCol).getValues()[0];
  const existing = row1.map(h=>String(h||"").trim());

  // build new header row preserving existing where possible
  const set = new Set(existing.filter(Boolean));
  let changed = false;

  for(const h of wanted){
    if(!set.has(h)){
      // put in first empty cell, else append at end
      let placed = false;
      for(let i=0;i<existing.length;i++){
        if(!existing[i]){
          existing[i] = h;
          placed = true;
          changed = true;
          break;
        }
      }
      if(!placed){
        existing.push(h);
        changed = true;
      }
    }
  }

  if(changed){
    sheet.getRange(1,1,1,existing.length).setValues([existing]);
  }
}

function normalizePhoneCoins_(p){
  let s = String(p || "").replace(/\D/g, "");
  if(s.length > 10) s = s.slice(-10);
  return s;
}

function hasHeaderSafe_(sheet, h){
  try{
    const { idx } = headerIndex_(sheet);
    return idx[h] != null;
  }catch(e){
    return false;
  }
}

/**
 * Safe row lookup (won’t throw if headers missing)
 */
function getRowByValueSafe_(sheet, headerName, value){
  try{
    return getRowByValue_(sheet, headerName, value);
  }catch(e){
    return null;
  }
}

function computeCoins_(event, amount){
  const amt = asNumber_(amount || 0);

  const minAmt = asNumber_(getSetting_("COINS_MIN_ORDER_AMOUNT", 0));
  if (amt < minAmt) return 0;

  const per100 = asNumber_(getSetting_("COINS_PER_100", 5)); // default: 5 coins per ₹100
  if (per100 <= 0) return 0;

  const base = Math.floor((amt / 100) * per100);

  const multNew   = asNumber_(getSetting_("COINS_NEW_MULTIPLIER", 1.0));
  const multRenew = asNumber_(getSetting_("COINS_RENEW_MULTIPLIER", 1.0));

  const ev = String(event||"").toUpperCase();
  const mult = (ev.includes("RENEW")) ? multRenew : multNew;

  return Math.max(0, Math.floor(base * mult));
}

function upsertWallet_(walletSheet, ph){
  const rn = getRowByValueSafe_(walletSheet, "Phone", ph);
  const { idx } = headerIndex_(walletSheet);

  if (rn){
    return {
      rowNum: rn,
      coinsLifetime: walletSheet.getRange(rn, idx["CoinsLifetime"]+1).getValue(),
      coinsBalance: walletSheet.getRange(rn, idx["CoinsBalance"]+1).getValue()
    };
  }

  const row = new Array(walletSheet.getLastColumn()).fill("");

  function put(h,v){ if(idx[h]!=null) row[idx[h]] = v; }

  put("Phone", ph);
  put("CoinsLifetime", 0);
  put("CoinsBalance", 0);
  if(idx["LastEarnedAt"]!=null) put("LastEarnedAt", "");
  if(idx["LastSpentAt"]!=null) put("LastSpentAt", "");
  if(idx["LastEvent"]!=null) put("LastEvent", "");

  walletSheet.appendRow(row);
  const newRn = walletSheet.getLastRow();

  return { rowNum: newRn, coinsLifetime: 0, coinsBalance: 0 };
}

function appendCoinsLedger_(ledger, p){
  const { idx } = headerIndex_(ledger);
  const row = new Array(ledger.getLastColumn()).fill("");

  function put(h,v){ if(idx[h]!=null) row[idx[h]] = v; }

  put("Ts", p.ts);
  put("Event", p.event);
  put("OrderID", p.orderId);
  put("Phone", p.phone);
  put("Service", p.service);
  put("Plan", p.plan);
  put("Amount", p.amount);
  put("CoinsDelta", p.coinsDelta);
  put("BalanceAfter", p.balanceAfter);
  put("Note", p.note);

  ledger.appendRow(row);
}

function coinsLedgerHas_(event, orderId){
  try{
    const ss = getCore_();
    const ledger = ss.getSheetByName(TAB_COINS_LEDGER);
    if(!ledger) return false;

    const { idx } = headerIndex_(ledger);
    const last = ledger.getLastRow();
    if(last < 2) return false;

    const ev = String(event||"").trim().toUpperCase();
    const oid = String(orderId||"").trim();

    const N = 800;
    const start = Math.max(2, last - N + 1);
    const data = ledger.getRange(start, 1, last - start + 1, ledger.getLastColumn()).getValues();

    for(let i=data.length-1;i>=0;i--){
      const r = data[i];
      const rEv = String(r[idx["Event"]]||"").trim().toUpperCase();
      const rOid = String(r[idx["OrderID"]]||"").trim();
      if(rEv === ev && rOid === oid) return true;
    }
    return false;
  }catch(e){
    return false;
  }
}




function TEST_wallet(){
  Logger.log(getWalletByPhone("9818196079"));
}

