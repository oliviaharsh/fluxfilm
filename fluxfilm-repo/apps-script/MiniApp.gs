function adminGetSubs(){
  // READ ONLY - safe
  const sh = SpreadsheetApp.getActive().getSheetByName("SUBSCRIPTIONS");
  if(!sh) return { ok:false, message:"Missing sheet: SUBSCRIPTIONS" };

  const values = sh.getDataRange().getValues();
  if(values.length < 2) return { ok:true, rows:[], count:0 };

  const headers = values[0].map(h => String(h||"").trim());
  const rows = [];

  for(let i=1;i<values.length;i++){
    const r = values[i];
    // skip fully empty rows
    if(r.join("").trim() === "") continue;

    const obj = {};
    for(let c=0;c<headers.length;c++){
      const key = headers[c] || ("Col"+(c+1));
      obj[key] = r[c];
    }
    rows.push(obj);
  }

  return { ok:true, count: rows.length, rows: rows };
}










