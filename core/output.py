"""
core/output.py
Assembles the final workbook by editing the DECRYPTED SOURCE WORKBOOK
in place, rather than building a brand-new workbook with hardcoded
styling.

REWRITTEN (this revision):
The previous version built a fresh openpyxl.Workbook() from scratch and
wrote every cell with a fixed, uniform style (navy header fill, thin grey
borders everywhere). That does NOT match the actual report's real
formatting, and it throws away whatever the source file's own
(hand-maintained) styling looked like.

What this file does instead — matching the manual process exactly:
- Start from the already-open `source_wb` (the decrypted original,
  loaded with formulas/styles intact via file_io.load_report_workbook).
- For each account that already has a row in a sheet, update ONLY the
  cell VALUES for columns that changed — its existing formatting
  (borders, fonts, fills, column widths) is left completely alone,
  because we never touch a cell's style when we're not touching its
  row for the first time.
- For a genuinely NEW row (new account added to DAILY/Trails, or a new
  PTP INVENTORY entry), the row's cell style is explicitly copied from
  the immediately preceding row before any value is written — this is
  the fix for the recurring "new rows have no border" bug that kept
  showing up when rows were appended without inheriting formatting.
- The ACTION CODE sheet is left completely untouched — it's a static
  reference sheet, never regenerated.
- Sheet order is preserved as it exists in the source workbook (do not
  reorder sheets — a prior version of this tool rebuilt sheet order from
  scratch, which is unnecessary and risks not matching the original).
"""
from __future__ import annotations

import copy
from typing import Optional

import openpyxl
import pandas as pd

from core.file_io import encrypt_workbook, DEFAULT_PASSWORD

DAILY_COLUMNS = [
    "DATE", "PRODUCT", "PN", "NAME", "OB", "DPD", "BUCKET",
    "ENDO DATE", "FLOWING DATE", "COLL TAGGING", "ACTION CODE",
    "STATUS REMARKS", "FV REMARKS", "CLIENT STATUS", "ADDRESS STATUS",
    "UNIT STATUS", "PULLED_OUT TAG", "PULLED_OUT DATE", "RFD",
    "GEO", "AGENCY", "COLLECTOR",
]

TRAILS_COLUMNS = [
    "FINANCIER ID", "APPLICATION ID", "CUSTOMER ID", "USER ID",
    "ACTION DATE", "ACTION TIME", "ACTION CODE", "CONTACT MODE",
    "PERSON CONTACTED", "PLACE CONTACTED", "CURRENCY", "ACTION AMOUNT",
    "NEXT ACTION DATE", "NEXT ACTION TIME", "REMINDER MODE",
    "CONTACTED BY", "REMARKS",
]

PTP_COLUMNS = DAILY_COLUMNS  # PTP INVENTORY shares DAILY's column layout


# ---------------------------------------------------------------------------
# Header discovery
# ---------------------------------------------------------------------------

def _header_map(ws, max_col: int = 30) -> dict[str, int]:
    """Map header name -> 1-based column index, read from row 1."""
    mapping = {}
    for c in range(1, max_col + 1):
        val = ws.cell(row=1, column=c).value
        if val:
            mapping[str(val).strip()] = c
    return mapping


def _last_data_row(ws, key_col: int) -> int:
    """
    Return the row number of the last row that actually has a value in
    key_col (skips any trailing fully-blank rows some source files have).
    Returns 1 (the header row) if there is no data at all yet.
    """
    last = 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=key_col).value not in (None, ""):
            last = r
    return last


def _copy_row_style(ws, src_row: int, dst_row: int, num_cols: int) -> None:
    """Copy every cell's style from src_row to dst_row, column by column."""
    for c in range(1, num_cols + 1):
        src_cell = ws.cell(row=src_row, column=c)
        dst_cell = ws.cell(row=dst_row, column=c)
        dst_cell._style = copy.copy(src_cell._style)


def _clean_value(val):
    """Convert pandas NaN/NaT to a plain None; pass everything else through."""
    if val is None:
        return None
    if isinstance(val, float) and pd.isna(val):
        return None
    if isinstance(val, pd.Timestamp):
        if pd.isna(val):
            return None
        return val.to_pydatetime()
    s = str(val)
    if s in ("nan", "NaT", "None"):
        return None
    return val


# ---------------------------------------------------------------------------
# Generic "sync a DataFrame into an existing sheet" routine
# ---------------------------------------------------------------------------

def _sync_sheet(
    ws,
    df: pd.DataFrame,
    columns: list[str],
    key_col_name: str,
) -> None:
    """
    Update `ws` in place so its data rows match `df`:
    - Existing rows (matched by the key column, e.g. PN / APPLICATION ID)
      get their cell VALUES updated for every column — formatting is
      never touched on an update.
    - Rows present in `df` but not yet in the sheet are appended, with
      their style explicitly copied from the last existing data row
      first, then values written.
    - Rows present in the sheet but NOT in `df` are left as-is (this
      function never deletes rows — dropping an account is a decision
      handled explicitly elsewhere, never as a side effect of syncing).
    """
    header = _header_map(ws)
    # Make sure every expected column has a header cell; if the sheet is
    # missing one entirely (shouldn't happen with the real template, but
    # don't crash silently), append it.
    next_free_col = max(header.values(), default=0) + 1
    for col_name in columns:
        if col_name not in header:
            ws.cell(row=1, column=next_free_col, value=col_name)
            header[col_name] = next_free_col
            next_free_col += 1

    key_col_idx = header[key_col_name]

    # Build PN/APPLICATION ID -> row number lookup from the sheet as it
    # currently stands (normalized the same way as the DataFrame's key).
    existing_row_by_key: dict[str, int] = {}
    for r in range(2, ws.max_row + 1):
        raw = ws.cell(row=r, column=key_col_idx).value
        key = _normalize_key(raw)
        if key:
            existing_row_by_key[key] = r

    last_row = _last_data_row(ws, key_col_idx)

    for _, row in df.iterrows():
        raw_key = row.get(key_col_name if key_col_name in df.columns else "PN")
        key = _normalize_key(raw_key)
        if not key:
            continue

        target_row = existing_row_by_key.get(key)

        if target_row is None:
            # Genuinely new row — copy style from the last data row FIRST.
            last_row += 1
            target_row = last_row
            if target_row > 2:
                _copy_row_style(ws, target_row - 1, target_row, len(header))
            existing_row_by_key[key] = target_row

        for col_name in columns:
            col_idx = header[col_name]
            # NOTE: ws.cell(..., value=X) is a no-op when X is None (openpyxl
            # only assigns when value is not None), so clearing a cell back
            # to blank requires setting .value directly rather than passing
            # value=None through the value= kwarg.
            ws.cell(row=target_row, column=col_idx).value = _clean_value(row.get(col_name))


def _normalize_key(val) -> str:
    """Normalize an account-number-like key to a clean string for matching."""
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", ""):
        return ""
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
    except (ValueError, TypeError):
        pass
    return s


# ---------------------------------------------------------------------------
# PTP INVENTORY — append-only (rows accumulate, never matched/updated)
# ---------------------------------------------------------------------------

def _append_ptp_rows(ws, new_rows_df: pd.DataFrame, columns: list[str]) -> None:
    """
    PTP INVENTORY rows are pure append: an account can legitimately have
    more than one row over time (separate PTP events), so we never match
    against an existing row here — every row in new_rows_df is a row that
    isn't in the sheet yet and should be added, with style copied from
    the preceding row.
    """
    if new_rows_df is None or new_rows_df.empty:
        return

    header = _header_map(ws)
    key_col_idx = header.get("PN", 3)
    last_row = _last_data_row(ws, key_col_idx)

    for _, row in new_rows_df.iterrows():
        last_row += 1
        if last_row > 2:
            _copy_row_style(ws, last_row - 1, last_row, len(header))
        for col_name in columns:
            col_idx = header.get(col_name)
            if col_idx:
                ws.cell(row=last_row, column=col_idx, value=_clean_value(row.get(col_name)))


# ---------------------------------------------------------------------------
# Main entry points
# ---------------------------------------------------------------------------

def build_output_workbook(
    source_wb: openpyxl.Workbook,
    daily_df: pd.DataFrame,
    trails_df: pd.DataFrame,
    ptp_new_rows_df: Optional[pd.DataFrame] = None,
) -> openpyxl.Workbook:
    """
    Edit `source_wb` in place and return it.

    - DAILY and Trails Upload are synced from daily_df / trails_df:
      existing rows get their values refreshed, new rows are appended
      with copied styling, nothing is ever deleted here.
    - PTP INVENTORY only receives NEW rows via ptp_new_rows_df (pass the
      already-decided "add these" rows — never the full inventory,
      since existing PTP rows are historical record and must never be
      rewritten or reordered).
    - ACTION CODE sheet is untouched.
    - Sheet order and every sheet's existing formatting/column widths
      are preserved exactly as they were in the source file.
    """
    if "DAILY" not in source_wb.sheetnames:
        raise ValueError("Source workbook has no 'DAILY' sheet.")
    if "Trails Upload" not in source_wb.sheetnames:
        raise ValueError("Source workbook has no 'Trails Upload' sheet.")
    if "PTP INVENTORY" not in source_wb.sheetnames:
        raise ValueError("Source workbook has no 'PTP INVENTORY' sheet.")

    ws_daily = source_wb["DAILY"]
    ws_trails = source_wb["Trails Upload"]
    ws_ptp = source_wb["PTP INVENTORY"]

    _sync_sheet(ws_daily, daily_df, DAILY_COLUMNS, key_col_name="PN")
    _sync_sheet(ws_trails, trails_df, TRAILS_COLUMNS, key_col_name="APPLICATION ID")

    if ptp_new_rows_df is not None and not ptp_new_rows_df.empty:
        _append_ptp_rows(ws_ptp, ptp_new_rows_df, PTP_COLUMNS)

    # ACTION CODE sheet: intentionally left completely untouched.

    return source_wb


def generate_download(
    source_wb: openpyxl.Workbook,
    daily_df: pd.DataFrame,
    trails_df: pd.DataFrame,
    ptp_new_rows_df: Optional[pd.DataFrame],
    password: str = DEFAULT_PASSWORD,
) -> bytes:
    """Edit the source workbook in place, encrypt, and return the .xlsx as bytes."""
    wb = build_output_workbook(source_wb, daily_df, trails_df, ptp_new_rows_df)
    return encrypt_workbook(wb, password)
