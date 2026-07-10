/*************** FluxFilm — Netflix Household Helper ***************/

/**
 * Called by frontend Tools tab.
 * Looks up the Netflix account email in INVENTORY_ACCOUNTS sheet
 * and returns the HouseholdLink column value for that account.
 *
 * INVENTORY_ACCOUNTS sheet needs a column called: HouseholdLink
 * (add it yourself — paste the household fix URL for each account row)
 *
 * Returns:
 *   { ok:true,  link:"https://..." }   — found
 *   { ok:false, message:"..." }        — not found or missing column
 */
function getNetflixHouseholdLink(email) {
  try {
    const em = String(email || "").trim().toLowerCase();
    if (!em || !em.includes("@")) return bad_("Valid email is required.");

    const sh = sh_(TAB_INV_ACCOUNTS);  // uses your existing INVENTORY_ACCOUNTS sheet
    const { idx } = headerIndex_(sh);

    if (idx["LoginId"] == null) return bad_("INVENTORY_ACCOUNTS missing header: LoginId");
    if (idx["HouseholdLink"] == null) return bad_("INVENTORY_ACCOUNTS missing header: HouseholdLink — please add this column.");

    const last = sh.getLastRow();
    if (last < 2) return bad_("No accounts found.");

    const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();

    for (const row of data) {
      const loginId = String(row[idx["LoginId"]] || "").trim().toLowerCase();
      if (loginId === em) {
        const link = String(row[idx["HouseholdLink"]] || "").trim();
        if (!link) return bad_("Household link not set for this account. Contact support.");
        return ok_({ link });
      }
    }

    return bad_("No Netflix account found with that email. Make sure you enter the account login email, not your personal email.");

  } catch (e) {
    return bad_("getNetflixHouseholdLink error: " + (e && e.message ? e.message : e));
  }
}