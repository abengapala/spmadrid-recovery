"""
core/load2.py
DRR loader that uses the DRR's clean 'FINAL CMS REMARK' column (falling back
to the raw 'FINAL REMARK' text only if the clean one is blank), plus the
extra columns Trails Upload / PTP need (Card No. -> CUSTOMER ID, PTP Amount,
PTP Date) and a REPO AI source flag.
"""
from __future__ import annotations

import io
import re

import pandas as pd

from core.clean import clean_text
from core.rawdrr import ensure_labelled

EXCLUDE_STATUSES = {"BP", "NEW", "REACTIVE", "ABORTED", "LOCKED", "UNLOCKED", "SMS FAILED"}
EXCLUDE_REMARKS = ["NEW ASSIGNMENT", "UPDATES WHEN CASE", "SYSTEM AUTO PD", "NEW CONTACT DETAILS"]


def _s(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def load_drr_clean(file_bytes: bytes, warnings: list | None = None, ref: dict | None = None) -> pd.DataFrame:
    warnings = warnings if warnings is not None else []
    df = pd.read_excel(io.BytesIO(file_bytes), dtype={"ACCOUNT": str, "Account No.": str, "Card No.": str})
    df.columns = [str(c).strip() for c in df.columns]
    df = ensure_labelled(df, warnings, ref)
    acct_col = "Account No." if "Account No." in df.columns else "ACCOUNT"
    df = df[df[acct_col].notna() & (df[acct_col].astype(str).str.strip() != "")].copy()
    n0 = len(df)

    out = pd.DataFrame()
    out["account_no"] = df[acct_col].astype(str).str.strip()
    out["date_parsed"] = pd.to_datetime(df["Date"], errors="coerce")
    out["time_str"] = df["Time"].astype(str)
    cms = df["FINAL CMS REMARK"].apply(_s)
    raw = df["FINAL REMARK"].apply(_s)
    base = cms.where(cms != "", raw)
    out["remark"] = base.apply(clean_text)
    out["raw_remark"] = raw
    out["action_code"] = df["CMS STATUS"].apply(_s)
    out["remark_source"] = "CMS"
    out["rank"] = pd.to_numeric(df["RANK"], errors="coerce")
    out["card_no"] = df["Card No."].apply(_s)
    out["ptp_amount"] = pd.to_numeric(df["PTP Amount"], errors="coerce").replace(0, float("nan"))
    out["ptp_date"] = pd.to_datetime(df["PTP Date"], dayfirst=True, errors="coerce")
    out["is_repo_ai"] = raw.str.contains(r"REPO\s*AI", case=False, regex=True)
    out = out.reset_index(drop=True)

    # noise rows
    st = out["action_code"].str.upper()
    rm = out["remark"].str.upper()
    mask = ~st.isin(EXCLUDE_STATUSES)
    for p in EXCLUDE_REMARKS:
        mask &= ~rm.str.contains(p, na=False)
    removed = int((~mask).sum())
    out = out[mask].reset_index(drop=True)
    if removed:
        warnings.append(f"DRR: removed {removed} noise rows (system statuses/remarks)")

    # weekend rows (confirmed rule: no weekend effort)
    wk = out["date_parsed"].notna() & (out["date_parsed"].dt.weekday >= 5)
    if wk.any():
        warnings.append(f"DRR: excluded {int(wk.sum())} weekend-dated rows")
    out = out[~wk].reset_index(drop=True)

    # rows with no usable remark text
    blank = out["remark"].str.strip() == ""
    if blank.any():
        warnings.append(f"DRR: dropped {int(blank.sum())} rows with blank remark")
    out = out[~blank].reset_index(drop=True)
    warnings.append(f"DRR rows: {n0} read -> {len(out)} used")
    return out
