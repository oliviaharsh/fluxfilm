/*************** FluxFilm Backend — FAQ ***************/

function getFaqs(){
  try{
    const sh = sh_("FAQ");
    const { idx } = headerIndex_(sh);

    const required = ["Category","Question","Answer","Active","Sort"];
    for (const h of required){
      if (idx[h] == null) return { ok:false, message:"FAQ sheet missing header: " + h };
    }

    const last = sh.getLastRow();
    if (last < 2) return { ok:true, result:{ faqs: [] } };

    const rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const out = [];

    for (const r of rows){
      const active = String(r[idx["Active"]] || "").trim().toUpperCase();
      if (active !== "TRUE") continue;

      out.push({
        category: String(r[idx["Category"]] || "").trim(),
        question: String(r[idx["Question"]] || "").trim(),
        answer: String(r[idx["Answer"]] || "").trim(),
        sort: Number(r[idx["Sort"]] || 9999)
      });
    }

    out.sort((a,b)=>{
      const ac = (a.category || "").toLowerCase();
      const bc = (b.category || "").toLowerCase();
      if (ac < bc) return -1;
      if (ac > bc) return 1;
      return (a.sort - b.sort);
    });

    return { ok:true, result:{ faqs: out } };

  }catch(e){
    return { ok:false, message:"getFaqs error: " + (e && e.message ? e.message : e) };
  }
}
