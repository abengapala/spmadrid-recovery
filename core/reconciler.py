"""
core/reconciler.py
Account list reconciliation logic.
Pure diff — no mutations. The UI decides what to apply.

FIXED (this revision):
- build_new_account_row now fills the confirmed constant defaults for a
  brand-new account instead of leaving them None:
    COLL TAGGING  -> "MARJELVIN DIAZ"
    PULLED_OUT TAG -> "EXISTING"
    COLLECTOR     -> "ALBERT LEGARA"
  (AGENCY was already correctly defaulted to "SP MADRID".)
  Previously these three were left blank for every new account, which
  meant a fresh endorsement would come in missing its collector tag and
  (worse) with no PULLED_OUT TAG at all — which would make the account
  invisible to the is_held() check in remarks.py/trails.py, since that
  check only skips a row when the tag is explicitly something other
  than "EXISTING"/blank; a blank was already being treated as "not
  held", so this wasn't a correctness bug for hold-detection, but it did
  mean every new account silently launched without its collector name.
- PN is now consistently normalized as a STRING throughout (was already
  string via _normalize_pn, but build_new_account_row previously passed
  the raw `pn` value straight through, which is fine since callers pass
  the normalized string — no behavior change here, just confirming it
  stays a string all the way into the DAILY row dict, never int()).
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

KEEPS_FILE = os.path.join(os.path.dirname(__file__), "..", "kept_accounts.json")

DEFAULT_COLL_TAGGING = "MARJELVIN DIAZ"
DEFAULT_PULLED_OUT_TAG = "EXISTING"
DEFAULT_COLLECTOR = "ALBERT LEGARA"
DEFAULT_AGENCY = "SP MADRID"
DEFAULT_PRODUCT = "01 AL - AUTO LOAN"


# ---------------------------------------------------------------------------
# Persistence — manual keeps
# ---------------------------------------------------------------------------

def load_kept_accounts() -> set[str]:
    """Load the set of manually-kept account PNs from the sidecar JSON."""
    path = os.path.abspath(KEEPS_FILE)
    if not os.path.exists(path):
        return set()
    try:
        with open(path, "r") as f:
            data = json.load(f)
        return set(str(p) for p in data.get("kept", []))
    except Exception:
        return set()


def save_kept_accounts(keeps: set[str]) -> None:
    """Persist the manually-kept account set to the sidecar JSON."""
    path = os.path.abspath(KEEPS_FILE)
    with open(path, "w") as f:
        json.dump({"kept": sorted(keeps)}, f, indent=2)


# ---------------------------------------------------------------------------
# Reconcile result dataclass
# ---------------------------------------------------------------------------

@dataclass
class ReconcileResult:
    current_accounts: list[str] = field(default_factory=list)
    new_candidates: list[str] = field(default_factory=list)
    manually_added: list[str] = field(default_factory=list)
    pn_to_name: dict[str, str] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Core reconcile function
# ---------------------------------------------------------------------------

def reconcile(
    daily_df: pd.DataFrame,
    db_df: Optional[pd.DataFrame],
    kept_accounts: set[str],  # kept for backward compat, unused in new logic
) -> ReconcileResult:
    """
    The current report's account list is ALWAYS the master.
    - All accounts in the current report are retained (current_accounts).
    - Accounts in the current report but NOT in the DB are flagged as
      manually_added (informational only — they are still kept).
    - Accounts in the DB but NOT in the current report are new_candidates
      (user must explicitly choose to add them).

    Nothing in the current report is ever auto-dropped.
    """
    result = ReconcileResult()
    pn_to_name: dict[str, str] = {}

    current_pns: set[str] = set()
    for _, row in daily_df.iterrows():
        pn = _normalize_pn(row.get("PN", ""))
        name = str(row.get("NAME", "")).strip()
        if pn:
            current_pns.add(pn)
            pn_to_name[pn] = name

    db_pns: set[str] = set()
    if db_df is not None and not db_df.empty:
        for _, row in db_df.iterrows():
            pn = _normalize_pn(row.get("PN_NO", ""))
            name = str(row.get("CUST_NAME", "")).strip()
            if pn:
                db_pns.add(pn)
                pn_to_name[pn] = pn_to_name.get(pn) or name

    result.pn_to_name = pn_to_name
    result.current_accounts = sorted(current_pns)

    if db_df is None or db_df.empty:
        return result

    result.new_candidates = sorted(db_pns - current_pns)
    result.manually_added = sorted(current_pns - db_pns)

    return result


def build_new_account_row(
    pn: str,
    db_df: pd.DataFrame,
    drr_df: Optional[pd.DataFrame],
    field_df: Optional[pd.DataFrame],
    report_date: pd.Timestamp,
) -> dict:
    """
    Build a DAILY-shaped row for a brand-new account using DB fields
    + whatever DRR/field history is available. Fields with no source
    stay blank EXCEPT the confirmed standing defaults below, which are
    always applied to a new account regardless of source data:
        COLL TAGGING   = "MARJELVIN DIAZ"
        PULLED_OUT TAG = "EXISTING"
        COLLECTOR      = "ALBERT LEGARA"
        AGENCY         = "SP MADRID"
        PRODUCT        = "01 AL - AUTO LOAN" (unless DB gives a real value)
    """
    db_row = db_df[db_df["PN_NO"] == pn]
    if db_row.empty:
        db_row_dict = {}
    else:
        db_row_dict = db_row.iloc[0].to_dict()

    def db(col, default=None):
        v = db_row_dict.get(col)
        return v if pd.notna(v) and str(v).strip() not in ("", "nan") else default

    status_remarks = ""
    if drr_df is not None and not drr_df.empty:
        acct_drr = drr_df[drr_df["account_no"] == pn].copy()
        if not acct_drr.empty:
            from core.remarks import build_remark_entries
            entries = build_remark_entries(acct_drr)
            status_remarks = ", ".join(entries)

    fv_remarks = ""
    if field_df is not None and not field_df.empty:
        acct_fv = field_df[field_df["pn"] == pn].sort_values("date_parsed", ascending=False)
        if not acct_fv.empty:
            fv_remarks = acct_fv.iloc[0].get("FV REMARK", "") or ""

    return {
        "DATE": report_date,
        "PRODUCT": db("PRODUCT", DEFAULT_PRODUCT),
        "PN": pn,
        "NAME": db("CUST_NAME", ""),
        "OB": db("OUTSTANDING_BALANCE"),
        "DPD": db("DPD"),
        "BUCKET": None,
        "ENDO DATE": db("ENDS_DATE"),
        "FLOWING DATE": db("FLOWING_DATE"),
        "COLL TAGGING": DEFAULT_COLL_TAGGING,
        "ACTION CODE": None,
        "STATUS REMARKS": status_remarks,
        "FV REMARKS": fv_remarks,
        "CLIENT STATUS": None,
        "ADDRESS STATUS": None,
        "UNIT STATUS": None,
        "PULLED_OUT TAG": DEFAULT_PULLED_OUT_TAG,
        "PULLED_OUT DATE": None,
        "RFD": None,
        "GEO": db("GEO_TAG"),
        "AGENCY": DEFAULT_AGENCY,
        "COLLECTOR": DEFAULT_COLLECTOR,
    }


def apply_reconcile(
    daily_df: pd.DataFrame,
    db_df: Optional[pd.DataFrame],
    drr_df: Optional[pd.DataFrame],
    field_df: Optional[pd.DataFrame],
    report_date: pd.Timestamp,
    result: ReconcileResult,
    pns_to_add: list[str],
    pns_to_drop: Optional[list[str]] = None,
    new_kept: Optional[set[str]] = None,
) -> pd.DataFrame:
    """
    Apply reconcile decisions:
    - All existing accounts are kept (no drops).
    - Adds rows only for pns_to_add (new accounts the user chose to include).
    """
    df = daily_df.copy()

    if pns_to_add and db_df is not None:
        new_rows = []
        for pn in pns_to_add:
            row = build_new_account_row(pn, db_df, drr_df, field_df, report_date)
            new_rows.append(row)
        new_df = pd.DataFrame(new_rows)
        df = pd.concat([df, new_df], ignore_index=True)

    return df


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize_pn(val) -> str:
    """Normalize a PN value to a clean string (strip .0 from floats)."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    try:
        f = float(str(val))
        if f == int(f):
            return str(int(f))
    except (ValueError, TypeError):
        pass
    return str(val).strip()
