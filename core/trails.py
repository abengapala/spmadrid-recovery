"""
core/trails.py
Rebuilds the Trails Upload sheet as a single-day snapshot.

FIXED (this revision):
- CRITICAL: the DRR filter used `date_parsed <= report_date` instead of
  `== report_date`. That meant an account with NO activity today would
  silently show its most recent OLDER entry as if it happened today
  (with the old date printed inside REMARKS, but ACTION DATE still set
  to today) — the exact opposite of the confirmed rule: no activity on
  the report date means REMARKS/ACTION CODE stay BLANK. Fixed to exact
  date match, and `blanked_pns` now correctly includes every account
  with no entry on report_date, not just accounts with zero DRR history
  ever.
- Accounts on hold (PULLED_OUT TAG != EXISTING / blank) are now excluded
  from Trails Upload entirely, matching the confirmed rule. Previously
  there was no such check at all.
- ACTION CODE here stays purely chronological (latest entry on the exact
  report date) — this is intentionally different from DAILY's RANK-based
  selection in core/remarks.py. Do not import or reuse that logic here.
"""
from __future__ import annotations

from typing import Optional

import pandas as pd

from core.action_codes import resolve_action_code
from core.remarks import format_remark, is_held

TRAILS_COLUMNS = [
    "FINANCIER ID", "APPLICATION ID", "CUSTOMER ID", "USER ID",
    "ACTION DATE", "ACTION TIME", "ACTION CODE", "CONTACT MODE",
    "PERSON CONTACTED", "PLACE CONTACTED", "CURRENCY", "ACTION AMOUNT",
    "NEXT ACTION DATE", "NEXT ACTION TIME", "REMINDER MODE",
    "CONTACTED BY", "REMARKS",
]


def _fmt_date(ts: pd.Timestamp) -> str:
    """Format a Timestamp as MM.DD.YYYY for Trails Upload."""
    return ts.strftime("%m.%d.%Y")


def rebuild_trails(
    daily_df: pd.DataFrame,
    drr_df: Optional[pd.DataFrame],
    existing_trails_df: pd.DataFrame,
    report_date: pd.Timestamp,
) -> tuple[pd.DataFrame, list[str]]:
    """
    Rebuild Trails Upload for the report date.

    For each ACTIVE (non-held) account in daily_df:
    - ACTION DATE = report_date, NEXT ACTION DATE = report_date + 1 day
    - Look for DRR entries dated EXACTLY report_date (not "on or before")
    - If found: REMARKS = that day's latest entry formatted, ACTION CODE
      = resolved code for that same entry (chronological, not RANK-based)
    - If NOT found: REMARKS = blank, ACTION CODE = blank — never
      backfilled with an older date's activity
    - Held accounts (FOR HOLD / PULLED OUT / REPO / etc.) are excluded
      from Trails Upload entirely — no row is written for them.
    - FINANCIER ID, CUSTOMER ID, USER ID, CONTACT MODE, etc. preserved
      from the existing row when the account already has one.

    Returns: (new_trails_df, blanked_account_pns)
    """
    next_date = report_date + pd.Timedelta(days=1)
    action_date_str = _fmt_date(report_date)
    next_date_str = _fmt_date(next_date)

    existing_lookup: dict[str, dict] = {}
    for _, row in existing_trails_df.iterrows():
        app_id = str(row.get("APPLICATION ID", "")).strip()
        if app_id and app_id != "nan":
            try:
                app_id = str(int(float(app_id)))
            except (ValueError, TypeError):
                pass
            existing_lookup[app_id] = row.to_dict()

    blanked_pns: list[str] = []
    rows = []

    for _, acct_row in daily_df.iterrows():
        pn_raw = acct_row.get("PN", "")
        pn = _normalize_pn(pn_raw)
        if not pn:
            continue

        if is_held(acct_row.get("PULLED_OUT TAG")):
            continue  # excluded from Trails Upload entirely

        existing = existing_lookup.get(pn, {})

        remarks_str = ""
        action_code_str = ""

        if drr_df is not None and not drr_df.empty:
            acct_drr_today = drr_df[
                (drr_df["account_no"] == pn) &
                (drr_df["date_parsed"].notna()) &
                (drr_df["date_parsed"].dt.date == report_date.date())   # EXACT date only
            ].sort_values(["date_parsed", "time_str"], ascending=False)

            if not acct_drr_today.empty:
                latest = acct_drr_today.iloc[0]
                remark_text = str(latest.get("remark", "")).strip()
                raw_code = str(latest.get("action_code", "")).strip()
                src = str(latest.get("remark_source", "RAW")).strip()

                if remark_text and remark_text.lower() != "nan":
                    remarks_str = format_remark(report_date, remark_text)
                action_code_str = resolve_action_code(raw_code, src)
            else:
                blanked_pns.append(pn)
        else:
            # No DRR at all this run — every account is blank, not stale-filled
            blanked_pns.append(pn)

        new_row = {
            "FINANCIER ID":    _preserve(existing, "FINANCIER ID"),
            "APPLICATION ID":  pn,   # kept as string — see note below
            "CUSTOMER ID":     _preserve(existing, "CUSTOMER ID"),
            "USER ID":         _preserve(existing, "USER ID"),
            "ACTION DATE":     action_date_str,
            "ACTION TIME":     None,
            "ACTION CODE":     action_code_str or None,
            "CONTACT MODE":    _preserve(existing, "CONTACT MODE"),
            "PERSON CONTACTED":_preserve(existing, "PERSON CONTACTED"),
            "PLACE CONTACTED": _preserve(existing, "PLACE CONTACTED"),
            "CURRENCY":        None,
            "ACTION AMOUNT":   None,
            "NEXT ACTION DATE":next_date_str,
            "NEXT ACTION TIME":None,
            "REMINDER MODE":   None,
            "CONTACTED BY":    _preserve(existing, "CONTACTED BY"),
            "REMARKS":         remarks_str or None,
        }
        rows.append(new_row)

    df = pd.DataFrame(rows, columns=TRAILS_COLUMNS)
    return df, blanked_pns


def _preserve(existing: dict, col: str):
    """Return the existing value for a column if non-blank, else None."""
    v = existing.get(col)
    if v is None:
        return None
    s = str(v).strip()
    if s in ("", "nan", "None", "NaN"):
        return None
    return v


def _normalize_pn(val) -> str:
    """
    Normalize account numbers as strings, never as Python int/float.
    Account numbers here run 12-15+ digits, which is at real risk of
    float-precision corruption if ever coerced through a float — always
    keep as string end-to-end, including in the output APPLICATION ID
    column (previously cast with int(pn), which is unnecessary and risky
    for very large account numbers).
    """
    if val is None:
        return ""
    try:
        f = float(str(val))
        if f == int(f):
            return str(int(f))
    except (ValueError, TypeError):
        pass
    return str(val).strip()
