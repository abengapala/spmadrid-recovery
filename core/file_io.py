"""
core/file_io.py
Handles decryption, encryption, and smart loading of Excel files.

FIXED (this revision):
- load_drr now KEEPS the RANK column (was being dropped before the final
  column selection, which silently made rank-based ACTION CODE selection
  for DAILY impossible downstream).
- load_drr now excludes Saturday/Sunday-dated entries by default
  (confirmed rule: no legitimate agent effort falls on a weekend; entries
  with those dates are logging artifacts, not real activity). Excluded
  rows are reported via `warnings`, never silently dropped without a trace.
"""
from __future__ import annotations

import io
import re
from typing import Optional

import msoffcrypto
import openpyxl
import pandas as pd

DEFAULT_PASSWORD = "cbs1234"

# ---------------------------------------------------------------------------
# Decrypt / Encrypt
# ---------------------------------------------------------------------------

def decrypt_workbook(file_bytes: bytes, password: str = DEFAULT_PASSWORD) -> io.BytesIO:
    """Decrypt an msoffcrypto-protected .xlsx and return a BytesIO buffer."""
    buf_in = io.BytesIO(file_bytes)
    try:
        office_file = msoffcrypto.OfficeFile(buf_in)
        office_file.load_key(password=password)
        buf_out = io.BytesIO()
        office_file.decrypt(buf_out)
        buf_out.seek(0)
        return buf_out
    except Exception as e:
        raise ValueError(f"Failed to decrypt workbook — wrong password or corrupt file: {e}") from e


def encrypt_workbook(wb: openpyxl.Workbook, password: str = DEFAULT_PASSWORD) -> bytes:
    """Save an openpyxl workbook and re-encrypt it with msoffcrypto. Returns bytes."""
    plain_buf = io.BytesIO()
    wb.save(plain_buf)
    plain_buf.seek(0)

    encrypted_buf = io.BytesIO()
    office_file = msoffcrypto.OfficeFile(plain_buf)
    office_file.encrypt(password, encrypted_buf)
    encrypted_buf.seek(0)
    return encrypted_buf.read()


# ---------------------------------------------------------------------------
# Smart sheet loading
# ---------------------------------------------------------------------------

def _normalize_col(name: str) -> str:
    """Uppercase + strip for fuzzy column matching."""
    return str(name).strip().upper()


def _find_header_row(ws, expected_cols: list[str], max_scan: int = 15) -> int:
    """
    Scan the first `max_scan` rows to find the row whose cells contain
    at least 60 % of the expected column names.  Returns 0-indexed row index.
    """
    expected_upper = {_normalize_col(c) for c in expected_cols}
    for idx, row in enumerate(ws.iter_rows(max_row=max_scan, values_only=True)):
        row_vals = {_normalize_col(str(v)) for v in row if v is not None}
        if len(row_vals & expected_upper) / max(len(expected_upper), 1) >= 0.6:
            return idx
    return 0  # fall back to row 0


def load_sheet_as_df(
    buf: io.BytesIO,
    sheet_name: str,
    expected_cols: Optional[list[str]] = None,
    header_row: Optional[int] = None,
) -> pd.DataFrame:
    """
    Load a single sheet from a BytesIO buffer into a DataFrame.
    - If header_row is given, use it directly.
    - Otherwise use smart header detection.
    - Drops rows that are completely blank.
    """
    buf.seek(0)
    wb = openpyxl.load_workbook(buf, read_only=True, data_only=True)

    if sheet_name not in wb.sheetnames:
        available = wb.sheetnames
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {available}")

    ws = wb[sheet_name]

    if header_row is None and expected_cols:
        header_row = _find_header_row(ws, expected_cols)
    elif header_row is None:
        header_row = 0

    buf.seek(0)
    df = pd.read_excel(buf, sheet_name=sheet_name, header=header_row)

    # Normalize column names (strip whitespace)
    df.columns = [str(c).strip() for c in df.columns]

    # Drop completely blank rows
    df = df.dropna(how="all").reset_index(drop=True)

    return df


def load_report_sheets(
    file_bytes: bytes, password: str = DEFAULT_PASSWORD
) -> dict[str, pd.DataFrame]:
    """
    Decrypt the main report and return all 4 sheets as DataFrames.
    Keys: 'DAILY', 'Trails Upload', 'PTP INVENTORY', 'ACTION CODE'
    """
    buf = decrypt_workbook(file_bytes, password)

    sheets = {}
    for sheet in ["DAILY", "Trails Upload", "PTP INVENTORY", "ACTION CODE"]:
        sheets[sheet] = load_sheet_as_df(buf, sheet)

    return sheets


def load_report_workbook(file_bytes: bytes, password: str = DEFAULT_PASSWORD) -> openpyxl.Workbook:
    """
    Decrypt and return the raw openpyxl Workbook, WITH formulas/formatting intact
    (data_only=False — we need real cell objects to edit in place and preserve
    styling, not just their last-calculated values).
    """
    buf = decrypt_workbook(file_bytes, password)
    return openpyxl.load_workbook(buf, data_only=False)


# ---------------------------------------------------------------------------
# Source file loaders
# ---------------------------------------------------------------------------

def load_drr(
    file_bytes: bytes,
    sheet_name: Optional[str] = None,
    warnings: Optional[list] = None,
    exclude_weekends: bool = True,
) -> pd.DataFrame:
    """
    Load DRR file or DRR sheet from a multi-sheet workbook.
    - Auto-detects sheet matching 'DRR' if multi-sheet.
    - Dynamic header row scanning.
    - Flexible column aliases for Account No, Date, Remark, Status, RANK.
    - Normalizes account numbers and parsed dates.
    - Excludes Saturday/Sunday-dated rows by default (confirmed: no real
      agent effort happens on a weekend — these are logging artifacts).
    """
    buf = io.BytesIO(file_bytes)

    target_sheet = sheet_name
    try:
        xl = pd.ExcelFile(buf)
        avail_sheets = xl.sheet_names
        if target_sheet:
            found = next((s for s in avail_sheets if s.strip().upper() == target_sheet.strip().upper()), None)
            target_sheet = found if found else avail_sheets[0]
        else:
            # Check for sheet named DRR or containing DRR
            found = next((s for s in avail_sheets if "DRR" in s.strip().upper()), None)
            target_sheet = found if found else avail_sheets[0]
    except Exception:
        target_sheet = 0

    buf.seek(0)
    # Dynamic header search
    raw = pd.read_excel(buf, sheet_name=target_sheet, header=None)
    header_idx = 0
    expected_drr_cols = ["ACCOUNT", "PN", "DATE", "REMARK", "STATUS", "CMS"]
    for i, row in raw.head(30).iterrows():
        row_str = " ".join([str(v).strip().upper() for v in row if pd.notna(v)])
        matches = sum(1 for e in expected_drr_cols if e in row_str)
        if matches >= 2:
            header_idx = i
            break

    buf.seek(0)
    df = pd.read_excel(buf, sheet_name=target_sheet, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)

    # Column alias matching helper
    def find_col(candidates: list[str]) -> Optional[str]:
        for cand in candidates:
            # Exact match
            for col in df.columns:
                if col.strip().upper() == cand.upper():
                    return col
            # Normalized match (no punctuation, no spaces)
            cand_clean = re.sub(r"[^A-Z0-9]", "", cand.upper())
            for col in df.columns:
                col_clean = re.sub(r"[^A-Z0-9]", "", col.upper())
                if col_clean == cand_clean:
                    return col
        return None

    acct_col = find_col([
        "Account No.", "Account No", "Account Number", "Account_No",
        "Account", "PN", "PN_NO", "PN NO", "Acct No", "Acct_No", "Acct No."
    ])
    if not acct_col:
        raise ValueError(
            f"DRR file/sheet missing 'Account No.' or 'PN' column. Available columns: {list(df.columns)}"
        )

    # Drop rows with no account number
    df = df[df[acct_col].notna() & (df[acct_col].astype(str).str.strip() != "")].copy()
    df = df.reset_index(drop=True)

    # Normalize account number to string
    df["account_no"] = df[acct_col].apply(
        lambda x: str(int(float(x))) if _is_numeric(x) else str(x).strip()
    )

    # Find Date column
    date_col = find_col(["Date", "DATE", "Action Date", "Activity Date", "Date Time", "Date_Time"])
    if date_col:
        df["date_parsed"] = df[date_col].apply(_parse_date_flexible)
    else:
        df["date_parsed"] = pd.NaT

    # Find Remarks and Status
    final_remark_col = find_col(["FINAL REMARK", "CLEAN REMARKS", "CLEANED REMARK", "CMS REMARK"])
    raw_remark_col = find_col(["Remark", "REMARK", "Remarks", "REMARKS", "STATUS REMARKS", "STATUS REMARK"])

    cms_status_col = find_col(["CMS STATUS", "CMS_STATUS", "ACTION CODE", "ACTION_CODE"])
    raw_status_col = find_col(["Status", "STATUS", "RAW STATUS", "ACTION"])

    # RANK column — used ONLY for DAILY's ACTION CODE selection (never for
    # Trails Upload, which stays purely chronological). May legitimately be
    # absent (raw-format DRR exports don't always include it); downstream
    # code must treat a missing/empty rank as "no priority info available"
    # rather than erroring.
    rank_col = find_col(["RANK", "Rank"])
    if rank_col:
        df["rank"] = pd.to_numeric(df[rank_col], errors="coerce")
    else:
        df["rank"] = pd.NA

    if final_remark_col and cms_status_col:
        df["remark"] = df[final_remark_col].fillna(df[raw_remark_col] if raw_remark_col else "")
        df["action_code"] = df[cms_status_col].fillna(df[raw_status_col] if raw_status_col else "")
        df["remark_source"] = "CMS"
    elif raw_remark_col or raw_status_col:
        df["remark"] = df[raw_remark_col] if raw_remark_col else pd.Series([""] * len(df))
        df["action_code"] = df[raw_status_col] if raw_status_col else pd.Series([""] * len(df))
        df["remark_source"] = "RAW"
    else:
        df["remark"] = pd.Series([""] * len(df))
        df["action_code"] = pd.Series([""] * len(df))
        df["remark_source"] = "RAW"

    # Time column
    time_col = find_col(["Time", "TIME", "Action Time", "TIME STR"])
    if time_col:
        df["time_str"] = df[time_col].astype(str)
    else:
        df["time_str"] = ""

    # --- DRR cleaning: filter out noise statuses and system remarks ---
    EXCLUDE_STATUSES = {"BP", "NEW", "REACTIVE", "ABORTED", "LOCKED", "UNLOCKED", "SMS FAILED"}
    EXCLUDE_REMARKS = ["NEW ASSIGNMENT", "UPDATES WHEN CASE", "SYSTEM AUTO PD", "NEW CONTACT DETAILS"]

    before_clean = len(df)

    status_upper = df["action_code"].astype(str).str.strip().str.upper()
    status_mask = ~status_upper.isin(EXCLUDE_STATUSES)

    remark_upper = df["remark"].astype(str).str.strip().str.upper()
    remark_mask = pd.Series([True] * len(df))
    for pattern in EXCLUDE_REMARKS:
        remark_mask = remark_mask & ~remark_upper.str.contains(pattern, na=False)

    df = df[status_mask & remark_mask].reset_index(drop=True)
    after_clean = len(df)

    if before_clean != after_clean:
        removed = before_clean - after_clean
        if warnings is not None:
            warnings.append(
                f"ℹ️ DRR cleaned: removed {removed} noise rows "
                f"(system statuses like BP/New/Locked and system remarks like New Assignment/System Auto PD)."
            )

    # --- Weekend exclusion (confirmed rule) ---
    if exclude_weekends:
        before_weekend = len(df)
        has_date = df["date_parsed"].notna()
        is_weekend = has_date & (df["date_parsed"].dt.weekday >= 5)  # Sat=5, Sun=6
        weekend_rows = df[is_weekend]
        df = df[~is_weekend].reset_index(drop=True)
        removed_weekend = before_weekend - len(df)
        if removed_weekend and warnings is not None:
            sample_accts = weekend_rows["account_no"].unique()[:5].tolist()
            warnings.append(
                f"⚠️ Excluded {removed_weekend} weekend-dated DRR row(s) (Sat/Sun) — "
                f"not counted as effort. Sample accounts: {sample_accts}"
            )

    return df[["account_no", "date_parsed", "time_str", "remark", "action_code", "remark_source", "rank"]].copy()


def load_field_file(
    file_bytes: bytes,
    sheet_name: Optional[str] = "OVERALL FIELD",
    header_row: Optional[int] = None,
) -> tuple[pd.DataFrame, list[str]]:
    """
    Load the OVERALL FIELD sheet from the monitoring .xlsm or a standalone export.
    Returns (df, warnings).
    Validates that PN looks like account numbers and DATE looks like dates.
    Detects OVERALL vs incremental by date span.
    """
    warnings: list[str] = []
    buf = io.BytesIO(file_bytes)

    target_sheet = sheet_name
    try:
        xl = pd.ExcelFile(buf)
        avail_sheets = xl.sheet_names
        if target_sheet:
            found = next((s for s in avail_sheets if s.strip().upper() == target_sheet.strip().upper()), None)
            if not found:
                found = next((s for s in avail_sheets if "FIELD" in s.strip().upper()), None)
            target_sheet = found if found else avail_sheets[0]
        else:
            target_sheet = avail_sheets[0]
    except Exception:
        target_sheet = 0

    buf.seek(0)
    raw = pd.read_excel(buf, sheet_name=target_sheet, header=None)

    expected_field_cols = ["PN", "DATE", "FV REMARK", "CLIENT STATUS", "ADDRESS STATUS", "UNIT STATUS"]
    detected_header = 0
    if header_row is not None and header_row < len(raw):
        detected_header = header_row
    else:
        for i, row in raw.head(20).iterrows():
            row_upper = {str(v).strip().upper() for v in row if pd.notna(v)}
            matches = sum(1 for e in expected_field_cols if e in row_upper)
            if matches >= 2:
                detected_header = i
                break

    buf.seek(0)
    df = pd.read_excel(buf, sheet_name=target_sheet, header=detected_header)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)

    # --- raw field export support -------------------------------------
    # FV REMARK = TEXT(DATE,"mm.dd.yyyy") & " " & Message (the user's clean-field formula)
    if "DATE" in df.columns and "Message" in df.columns:
        _d = pd.to_datetime(df["DATE"], errors="coerce")
        _built = (_d.dt.strftime("%m.%d.%Y") + " " + df["Message"].fillna("").astype(str)).str.replace("\u00e2\u20ac\u2122", "", regex=False).str.strip()
        _built = _built.where(_d.notna() & df["Message"].notna() & (df["Message"].astype(str).str.strip() != ""), "")
        if "FV REMARK" not in df.columns:
            df["FV REMARK"] = _built
            warnings.append("Field file: FV REMARK built from DATE + Message")
        else:
            _blank = df["FV REMARK"].isna() | (df["FV REMARK"].astype(str).str.strip().isin(["", "0", "nan"]))
            if (_blank & (_built != "")).any():
                df.loc[_blank, "FV REMARK"] = _built[_blank].replace("", pd.NA)
                warnings.append(f"Field file: built {int((_blank & (_built != '')).sum())} blank FV REMARK from DATE + Message")
    _valid = lambda x: bool(re.match(r"^\d{8,20}$", str(x).strip().split(".")[0]))
    if "PN" in df.columns:
        for _alt in ["ACCOUNT NUMBER", "NEW PN"]:
            if _alt in df.columns:
                _bad = ~df["PN"].apply(_valid)
                _fix = _bad & df[_alt].apply(_valid)
                if _fix.any():
                    df.loc[_fix, "PN"] = df.loc[_fix, _alt].astype(str).str.split(".").str[0]
                    warnings.append(f"Field file: {int(_fix.sum())} rows with no PN matched using {_alt}")

    if "PN" not in df.columns:
        for alt in ["PN_NO", "PN NO", "ACCOUNT NO", "ACCOUNT_NO", "ACCOUNT", "ACCT NO"]:
            if alt in df.columns:
                df["PN"] = df[alt]
                break

    if "PN" not in df.columns:
        raise ValueError("Field file missing 'PN' column. Cannot match to DAILY sheet.")

    pn_all = df["PN"].dropna().astype(str).str.strip()
    pn_all = pn_all[pn_all.str.lower() != "nan"]
    if len(pn_all) > 0:
        pn_numeric_count = pn_all.apply(lambda p: bool(re.match(r"^\d{8,20}$", p))).sum()
        pn_numeric_ratio = pn_numeric_count / len(pn_all)
        if pn_numeric_ratio < 0.5:
            warnings.append(
                "⚠️ COLUMN ALIGNMENT WARNING: Most 'PN' values don't look like account numbers "
                f"(only {pn_numeric_count}/{len(pn_all)} numeric). Check the field file header row."
            )
        elif pn_numeric_ratio < 1.0:
            non_numeric = pn_all[~pn_all.apply(lambda p: bool(re.match(r"^\d{8,20}$", p)))].tolist()
            warnings.append(
                f"ℹ️ Field file contains {len(non_numeric)} non-account-number rows in PN column "
                f"(e.g. {non_numeric[:3]}) — these rows will be skipped when matching accounts."
            )

    date_col = None
    for d_candidate in ["DATE", "DATE TIME", "FIELD DATE", "VISIT DATE"]:
        if d_candidate in df.columns:
            date_col = d_candidate
            break

    if not date_col:
        warnings.append("ℹ️ No DATE or DATE TIME column found in field file — FV REMARKS will be sorted without dates.")

    df["pn"] = df["PN"].apply(
        lambda x: str(int(float(x))) if _is_numeric(x) else str(x).strip()
    )
    df["_pn_valid"] = df["pn"].apply(lambda p: bool(re.match(r"^\d{8,20}$", p)))

    use_col = date_col or "DATE"
    if use_col in df.columns:
        df["date_parsed"] = df[use_col].apply(_parse_date_flexible)
    else:
        df["date_parsed"] = pd.NaT

    ptp_amt_col = next((c for c in ["PTP AMOUNT", "PTP_AMOUNT", "PTP AMT"] if c in df.columns), None)
    if ptp_amt_col:
        df["ptp_amount"] = pd.to_numeric(df[ptp_amt_col], errors="coerce").replace(0, float("nan"))
    else:
        df["ptp_amount"] = float("nan")

    ptp_date_col = next((c for c in ["PTP-Date", "PTP DATE", "PTP_DATE"] if c in df.columns), None)
    if ptp_date_col:
        df["ptp_date"] = df[ptp_date_col].apply(_clean_ptp_date)
    else:
        df["ptp_date"] = pd.NaT

    valid_dates = df["date_parsed"].dropna()
    if len(valid_dates) >= 2:
        span_days = (valid_dates.max() - valid_dates.min()).days
        df.attrs["is_overall"] = span_days > 30
        if df.attrs["is_overall"]:
            warnings.append(
                f"ℹ️ Field file spans {span_days} days — treating as OVERALL (full-history rebuild)."
            )
    else:
        df.attrs["is_overall"] = False

    keep = ["pn", "date_parsed", "FV REMARK", "CLIENT STATUS", "ADDRESS STATUS",
            "UNIT STATUS", "RFD", "ptp_amount", "ptp_date", "_pn_valid"]
    for c in keep:
        if c not in df.columns:
            match = next((orig for orig in df.columns if orig.upper() == c.upper()), None)
            if match:
                df[c] = df[match]
            else:
                df[c] = None

    result = df[keep].copy()
    result = result[result["_pn_valid"] == True].drop(columns=["_pn_valid"]).reset_index(drop=True)

    return result, warnings


def load_database(
    file_bytes: bytes,
    sheet_name: Optional[str] = None,
) -> tuple[pd.DataFrame, list[str]]:
    """
    Load the account database.
    - Auto-detects 'DATABASE' sheet if part of a multi-sheet workbook.
    - Auto-detects header row.
    - Standardizes column names (e.g. PN / PN_NO, CUST_NAME / NAME).
    - Filters PLACEMENT == RECOVERY, AGENCY == SP MADRID if those columns exist.
    - Deduplicates by PN_NO.
    Returns (df, warnings).
    """
    warnings: list[str] = []
    buf = io.BytesIO(file_bytes)

    target_sheet = sheet_name
    try:
        xl = pd.ExcelFile(buf)
        avail_sheets = xl.sheet_names
        if target_sheet:
            found = next((s for s in avail_sheets if s.strip().upper() == target_sheet.strip().upper()), None)
            target_sheet = found if found else avail_sheets[0]
        else:
            found = next((s for s in avail_sheets if s.strip().upper() in ("DATABASE", "DB", "ACCOUNTS")), None)
            if found:
                target_sheet = found
                warnings.append(f"ℹ️ Reading database from sheet '{target_sheet}'.")
            else:
                target_sheet = avail_sheets[0]
    except Exception:
        target_sheet = 0

    expected = ["PN_NO", "PN", "CUST_NAME", "NAME", "OUTSTANDING_BALANCE", "DPD", "AGENCY",
                "PLACEMENT", "ENDS_DATE", "FLOWING_DATE", "GEO_TAG", "CUST_ID"]

    buf.seek(0)
    raw = pd.read_excel(buf, sheet_name=target_sheet, header=None)
    header_idx = 0
    for i, row in raw.head(25).iterrows():
        row_upper = {str(v).strip().upper() for v in row if pd.notna(v)}
        matches = sum(1 for e in expected if e.upper() in row_upper)
        if matches >= 3:
            header_idx = i
            break

    buf.seek(0)
    df = pd.read_excel(buf, sheet_name=target_sheet, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)

    if header_idx > 0:
        warnings.append(f"ℹ️ Database header detected on row {header_idx + 1}.")

    alias_map = {
        "PN": "PN_NO",
        "ACCOUNT NO": "PN_NO",
        "ACCOUNT_NO": "PN_NO",
        "ACCT NO": "PN_NO",
        "CUSTOMER NAME": "CUST_NAME",
        "NAME": "CUST_NAME",
        "CLIENT NAME": "CUST_NAME",
        "OB": "OUTSTANDING_BALANCE",
        "BALANCE": "OUTSTANDING_BALANCE",
        "ENDO DATE": "ENDS_DATE",
        "ENDORSEMENT DATE": "ENDS_DATE",
        "FLOWING DATE": "FLOWING_DATE",
        "GEO": "GEO_TAG",
        "GEOTAG": "GEO_TAG",
    }
    for old_col, new_col in alias_map.items():
        if new_col not in df.columns:
            match = next((c for c in df.columns if c.strip().upper() == old_col), None)
            if match:
                df[new_col] = df[match]

    if "PN_NO" not in df.columns:
        raise ValueError(f"Database file missing PN / PN_NO column. Found columns: {list(df.columns)}")

    before = len(df)
    if "PLACEMENT" in df.columns:
        recovery_mask = df["PLACEMENT"].astype(str).str.strip().str.upper() == "RECOVERY"
        if recovery_mask.any():
            df = df[recovery_mask].copy()

    if "AGENCY" in df.columns:
        agency_mask = df["AGENCY"].astype(str).str.strip().str.upper().str.contains("MADRID", na=False)
        if agency_mask.any():
            df = df[agency_mask].copy()

    after_filter = len(df)
    if before != after_filter:
        warnings.append(f"ℹ️ Database: {before} rows → {after_filter} after RECOVERY / SP MADRID filters.")

    df["PN_NO"] = df["PN_NO"].apply(
        lambda x: str(int(float(x))) if _is_numeric(x) else str(x).strip()
    )
    df = df[df["PN_NO"] != ""].copy()
    before_dedup = len(df)
    df = df.drop_duplicates(subset="PN_NO").reset_index(drop=True)
    if len(df) < before_dedup:
        warnings.append(f"ℹ️ Removed {before_dedup - len(df)} duplicate PN entries from database.")

    return df, warnings


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _is_numeric(val) -> bool:
    try:
        float(str(val))
        return True
    except (ValueError, TypeError):
        return False


def _parse_date_flexible(val) -> Optional[pd.Timestamp]:
    """Parse both datetime objects and DD-MM-YYYY strings."""
    if pd.isna(val) if not isinstance(val, str) else False:
        return None
    if isinstance(val, pd.Timestamp):
        return val
    try:
        return pd.Timestamp(val)
    except Exception:
        pass
    try:
        return pd.Timestamp(str(val).strip(), dayfirst=True)
    except Exception:
        return None


def _clean_ptp_date(val) -> Optional[pd.Timestamp]:
    """Return None for placeholder PTP dates (0, time(0,0), 1900-01-01 epoch)."""
    if val is None or (isinstance(val, float) and val == 0):
        return None
    import datetime as dt
    if isinstance(val, dt.time):
        return None
    ts = _parse_date_flexible(val)
    if ts is None:
        return None
    if ts.year <= 1900:
        return None
    return ts
