"""
core/remarks.py
STATUS REMARKS and FV REMARKS update logic.

FIXED (this revision):
- DAILY's ACTION CODE is now chosen by RANK (lowest number = highest
  priority) when RANK data is available, NOT by plain recency. Trails
  Upload's ACTION CODE stays purely chronological — that logic lives in
  core/trails.py and is intentionally NOT touched here. These two must
  never be conflated again.
- Accounts whose PULLED_OUT TAG is anything other than "EXISTING" (FOR
  HOLD / PULLED OUT / REPO / any future hold-style tag) are now skipped
  entirely — their STATUS REMARKS, ACTION CODE, and FV REMARKS are left
  completely untouched. This was previously unhandled, so held accounts
  would get silently overwritten on every run.
- New: REPO AI-sourced STATUS REMARKS entries are mirrored into FV
  REMARKS, but ONLY when they say something genuinely new that FV REMARKS
  doesn't already contain (confirmed rule — duplicated events should not
  be added twice).
- New: when FV REMARKS is the placeholder "AWAITING FIELD STATUS" (no
  real field visit yet), ADDRESS STATUS is set to the literal value
  "(FOR FIELD VISIT)" — not "FFV", not "NEGATIVE".
- New: stale FV REMARKS (no matching activity in the current field/DRR
  update window) get reset back to "AWAITING FIELD STATUS" rather than
  left indefinitely. "Stale" here means: the account's FV REMARKS holds
  content, but neither this run's field file nor a REPO AI DRR entry
  supplied anything current for it. This is a judgment call the tool
  flags for review rather than applying silently — see `stale_fv_pns` in
  update_all_remarks' return value.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

import pandas as pd

# Date prefix pattern: MM.DD.YYYY (with optional space after)
_DATE_PREFIX_RE = re.compile(r"(\d{2}\.\d{2}\.\d{4})\s+")
# Split lookahead: split on ", " followed by a date prefix
_SPLIT_RE = re.compile(r",\s*(?=\d{2}\.\d{2}\.\d{4}\s)")

FV_PLACEHOLDER = "AWAITING FIELD STATUS"
ADDRESS_PLACEHOLDER = "(FOR FIELD VISIT)"

# PULLED_OUT TAG values that mean "don't touch this account"
HOLD_TAG_VALUES_EXCLUDED = {"EXISTING", "", None}  # anything NOT in this set = held


def is_held(pulled_out_tag) -> bool:
    """True if this account's PULLED_OUT TAG marks it as on-hold (skip entirely)."""
    if pulled_out_tag is None:
        return False
    val = str(pulled_out_tag).strip()
    if val == "" or val.lower() == "nan":
        return False
    return val.upper() != "EXISTING"


@dataclass
class RemarkEntry:
    date_str: str       # "MM.DD.YYYY"
    text: str           # full remark text (without the date prefix)
    full: str           # "MM.DD.YYYY text"
    date: Optional[pd.Timestamp] = None

    @property
    def sort_key(self) -> pd.Timestamp:
        return self.date if self.date is not None else pd.Timestamp.min


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

def parse_remarks_string(s: str) -> list[RemarkEntry]:
    """
    Split a remarks string into individual RemarkEntry objects.
    Splits on ", " only when followed immediately by a MM.DD.YYYY date prefix.
    This avoids breaking on commas inside remark text.
    """
    if not s or not isinstance(s, str) or s.strip() == "" or s.strip().lower() == "nan":
        return []

    parts = _SPLIT_RE.split(s.strip())
    entries = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        m = _DATE_PREFIX_RE.match(part)
        if m:
            date_str = m.group(1)
            text = part[m.end():].strip()
            ts = _parse_mm_dd_yyyy(date_str)
            entries.append(RemarkEntry(date_str=date_str, text=text, full=part, date=ts))
        else:
            if entries:
                entries[-1] = RemarkEntry(
                    date_str=entries[-1].date_str,
                    text=entries[-1].text + " " + part,
                    full=entries[-1].full + " " + part,
                    date=entries[-1].date,
                )
            else:
                entries.append(RemarkEntry(date_str="", text=part, full=part, date=None))

    return entries


def _parse_mm_dd_yyyy(s: str) -> Optional[pd.Timestamp]:
    """Parse a MM.DD.YYYY string into a Timestamp."""
    try:
        parts = s.split(".")
        if len(parts) == 3:
            m, d, y = int(parts[0]), int(parts[1]), int(parts[2])
            return pd.Timestamp(year=y, month=m, day=d)
    except Exception:
        pass
    return None


def format_remark(date: pd.Timestamp, text: str) -> str:
    """Format a single remark entry as MM.DD.YYYY text."""
    return f"{date.strftime('%m.%d.%Y')} {text.strip()}"


# ---------------------------------------------------------------------------
# Build new remark entries from DRR
# ---------------------------------------------------------------------------

def build_remark_entries(acct_drr: pd.DataFrame) -> list[str]:
    """
    Convert DRR rows for one account into formatted remark strings.
    Returns list of "MM.DD.YYYY <remark>", newest first.
    Deduplicates by (date, remark text).
    """
    entries = []
    seen = set()
    sorted_drr = acct_drr.sort_values("date_parsed", ascending=False)

    for _, row in sorted_drr.iterrows():
        date = row.get("date_parsed")
        remark = str(row.get("remark", "")).strip()

        if not isinstance(date, pd.Timestamp) or pd.isna(date):
            continue
        if not remark or remark.lower() == "nan":
            continue

        key = (date.date(), remark.lower())
        if key in seen:
            continue
        seen.add(key)

        entries.append(format_remark(date, remark))

    return entries


# ---------------------------------------------------------------------------
# RANK-based ACTION CODE selection (DAILY ONLY — never used for Trails)
# ---------------------------------------------------------------------------

def _rank_key(rank_val) -> float:
    """Lower = higher priority. Missing/unparseable rank sorts last (999)."""
    try:
        if rank_val is None or pd.isna(rank_val):
            return 999.0
        return float(rank_val)
    except (TypeError, ValueError):
        return 999.0


def select_daily_action_code(acct_drr: pd.DataFrame) -> Optional[str]:
    """
    Choose the ACTION CODE to show on the DAILY sheet for one account,
    from all of that account's DRR rows within the current update window.

    Priority order:
    1. Lowest RANK value (most significant status — e.g. REPO/PTP-type
       outcomes rank lower/more-important than a routine no-answer call).
    2. Ties (including "no RANK data at all") break by most recent
       date+time.

    This is DELIBERATELY DIFFERENT from Trails Upload's ACTION CODE
    (core/trails.py), which is always purely chronological. Do not merge
    this logic with that file's.
    """
    if acct_drr is None or acct_drr.empty:
        return None

    df = acct_drr.copy()
    df["_rank_sort"] = df["rank"].apply(_rank_key) if "rank" in df.columns else 999.0

    # Build a single sortable timestamp for the tiebreak (date + time_str
    # is best-effort; if time can't be parsed, entries still sort by date).
    def _dt_key(row):
        d = row.get("date_parsed")
        if not isinstance(d, pd.Timestamp) or pd.isna(d):
            return pd.Timestamp.min
        t = str(row.get("time_str", "") or "")
        try:
            parsed_t = pd.to_datetime(t).time()
            return pd.Timestamp.combine(d.date(), parsed_t)
        except Exception:
            return d

    df["_dt_sort"] = df.apply(_dt_key, axis=1)
    df = df.sort_values(["_rank_sort", "_dt_sort"], ascending=[True, False])

    if df.empty:
        return None

    best = df.iloc[0]
    from core.action_codes import resolve_action_code
    return resolve_action_code(
        str(best.get("action_code", "")),
        str(best.get("remark_source", "RAW")),
    )


# ---------------------------------------------------------------------------
# Update STATUS REMARKS
# ---------------------------------------------------------------------------

def update_status_remarks(
    existing: str,
    acct_drr: pd.DataFrame,
    cutoff_date: pd.Timestamp,
) -> tuple[str, bool]:
    """
    Update the STATUS REMARKS string for one account.

    Logic:
    1. Parse existing → keep entries with date STRICTLY BEFORE cutoff_date.
    2. Build new entries from DRR (any date on/after cutoff), newest first.
    3. Join: new_entries + kept_old_entries, comma-separated.

    This means any date on/after the cutoff is fully REPLACED by the new
    DRR's version of that date, never duplicated alongside an old copy.

    Returns: (new_remarks_string, had_new_entries: bool)
    """
    existing_entries = parse_remarks_string(existing)

    kept_old = [
        e for e in existing_entries
        if e.date is not None and e.date < cutoff_date
    ]

    if acct_drr is not None and not acct_drr.empty:
        new_drr = acct_drr[
            acct_drr["date_parsed"].notna() &
            (acct_drr["date_parsed"] >= cutoff_date)
        ]
        new_entries = build_remark_entries(new_drr)
    else:
        new_entries = []

    had_new = len(new_entries) > 0

    all_entries = new_entries + [e.full for e in kept_old]
    result = ", ".join(e for e in all_entries if e)

    return result, had_new


# ---------------------------------------------------------------------------
# REPO AI → FV REMARKS mirroring (conditional on genuinely new content)
# ---------------------------------------------------------------------------

def extract_repo_ai_entries(status_remarks: str) -> list[str]:
    """Return the individual STATUS REMARKS entries whose source is REPO AI."""
    entries = parse_remarks_string(status_remarks)
    return [e.full for e in entries if "REPO AI" in e.text.upper()]


def merge_repo_ai_into_fv(fv_remarks: str, status_remarks: str) -> tuple[str, bool]:
    """
    Add REPO AI-sourced STATUS REMARKS entries into FV REMARKS, but only
    the ones that say something genuinely new — i.e. skip any REPO AI
    entry whose text is already substantively present in FV REMARKS
    (avoids duplicating the same recovered/surrender event twice).

    Returns: (new_fv_remarks, changed)
    """
    repo_entries = extract_repo_ai_entries(status_remarks)
    if not repo_entries:
        return fv_remarks, False

    existing_lower = (fv_remarks or "").lower()

    def _is_redundant(entry_full: str) -> bool:
        # Pull out the REMARKS ... payload after the "REMARKS" keyword if
        # present, since that's the actual substantive content to compare;
        # fall back to the whole entry text otherwise.
        m = re.search(r"REMARKS\s+(.*)", entry_full, re.IGNORECASE)
        payload = (m.group(1) if m else entry_full).strip().lower()
        if not payload:
            return True
        # Consider it redundant if a meaningful chunk of the payload
        # already appears in the existing FV REMARKS text.
        snippet = payload[:40]
        return snippet in existing_lower

    new_entries = [e for e in repo_entries if not _is_redundant(e)]
    if not new_entries:
        return fv_remarks, False

    base = fv_remarks if fv_remarks and fv_remarks.strip() and fv_remarks.strip() != FV_PLACEHOLDER else ""
    combined = ", ".join(new_entries + ([base] if base else []))
    return combined, True


# ---------------------------------------------------------------------------
# Update FV REMARKS
# ---------------------------------------------------------------------------

def update_fv_remarks(
    existing: str,
    acct_field: pd.DataFrame,
    is_overall: bool,
) -> tuple[str, bool]:
    """
    Update FV REMARKS for one account from the field file.

    - If is_overall: rebuild entirely from field entries (newest first).
    - Else: prepend new entries, skip duplicates.

    Returns: (new_fv_remarks_string, changed: bool)
    """
    if acct_field is None or acct_field.empty:
        return existing, False

    sorted_fv = acct_field.sort_values("date_parsed", ascending=False)
    field_entries = []
    for _, row in sorted_fv.iterrows():
        fv_remark = str(row.get("FV REMARK", "")).strip()
        if fv_remark and fv_remark.lower() not in ("nan", "0", ""):
            field_entries.append(fv_remark)

    if not field_entries:
        return existing, False

    if is_overall:
        new_str = ", ".join(field_entries)
        changed = new_str != (existing or "").strip()
        return new_str, changed

    existing_entries = parse_remarks_string(existing)
    existing_fulls = {e.full.strip().lower() for e in existing_entries}

    new_to_prepend = [e for e in field_entries if e.strip().lower() not in existing_fulls]
    if not new_to_prepend:
        return existing, False

    old_parts = [e.full for e in existing_entries]
    all_parts = new_to_prepend + old_parts
    new_str = ", ".join(p for p in all_parts if p)
    return new_str, True


# ---------------------------------------------------------------------------
# Update status columns from field data
# ---------------------------------------------------------------------------

def update_status_columns(
    daily_row: dict,
    latest_field_row: Optional[pd.Series],
) -> dict:
    """
    Overwrite CLIENT STATUS / ADDRESS STATUS / UNIT STATUS / RFD
    only when the field value is non-blank, non-zero, non-null.
    Never overwrites a real value with a blank one.
    """
    if latest_field_row is None:
        return daily_row

    row = daily_row.copy()
    for col in ["CLIENT STATUS", "ADDRESS STATUS", "UNIT STATUS", "RFD"]:
        field_val = latest_field_row.get(col)
        if _is_real_value(field_val):
            row[col] = str(field_val).strip()

    return row


def apply_address_status_placeholder(row: dict) -> dict:
    """
    If FV REMARKS is the placeholder (no real field visit yet), ADDRESS
    STATUS must be the literal placeholder "(FOR FIELD VISIT)" — not
    "FFV" and not "NEGATIVE". Only applies when FV REMARKS is exactly
    the placeholder; a real (even old) FV REMARKS value leaves ADDRESS
    STATUS alone here.
    """
    fv = str(row.get("FV REMARKS", "") or "").strip()
    if fv == FV_PLACEHOLDER or fv == "":
        row = row.copy()
        row["ADDRESS STATUS"] = ADDRESS_PLACEHOLDER
    return row


def _is_real_value(val) -> bool:
    """True if val is a non-blank, non-zero, non-null value."""
    if val is None:
        return False
    if isinstance(val, float) and (pd.isna(val) or val == 0):
        return False
    s = str(val).strip()
    return s not in ("", "0", "nan", "NaN", "None")


# ---------------------------------------------------------------------------
# Batch update for the full DAILY dataframe
# ---------------------------------------------------------------------------

def update_all_remarks(
    daily_df: pd.DataFrame,
    drr_df: Optional[pd.DataFrame],
    field_df: Optional[pd.DataFrame],
    cutoff_date: pd.Timestamp,
    report_date: pd.Timestamp,
) -> tuple[pd.DataFrame, list[str], list[str], list[str], list[str]]:
    """
    Apply STATUS REMARKS, DAILY ACTION CODE, FV REMARKS, and the
    ADDRESS STATUS placeholder rule to the entire DAILY dataframe.
    Accounts on hold (PULLED_OUT TAG != EXISTING / blank) are skipped
    entirely — nothing about them is touched, and they're not counted
    in any of the returned lists.

    Returns:
        - updated_df
        - zero_activity_pns: active accounts with no new DRR entries
        - changed_fv_pns: active accounts whose FV REMARKS changed
        - repo_ai_added_pns: accounts where a REPO AI entry was mirrored
          into FV REMARKS this run
        - stale_fv_pns: active accounts that have FV REMARKS content but
          got no fresh field-file or REPO AI activity this run — flagged
          for review, NOT auto-reset (resetting is a decision the caller/
          UI should confirm, since "stale" has no hard day-count cutoff)
    """
    df = daily_df.copy()
    is_overall = field_df.attrs.get("is_overall", False) if field_df is not None else False

    zero_activity_pns: list[str] = []
    changed_fv_pns: list[str] = []
    repo_ai_added_pns: list[str] = []
    stale_fv_pns: list[str] = []

    for idx, row in df.iterrows():
        pn = str(row.get("PN", "")).strip()
        if not pn or pn == "nan":
            continue

        if is_held(row.get("PULLED_OUT TAG")):
            continue  # never touch held accounts

        acct_drr = pd.DataFrame()
        if drr_df is not None and not drr_df.empty:
            acct_drr = drr_df[drr_df["account_no"] == pn]

        # --- STATUS REMARKS ---
        had_new = False
        if not acct_drr.empty or (drr_df is not None):
            new_remarks, had_new = update_status_remarks(
                str(row.get("STATUS REMARKS", "") or ""),
                acct_drr,
                cutoff_date,
            )
            df.at[idx, "STATUS REMARKS"] = new_remarks
            if not had_new:
                zero_activity_pns.append(pn)

            # --- DAILY ACTION CODE: RANK-based, current window only ---
            window_drr = acct_drr[
                acct_drr["date_parsed"].notna() &
                (acct_drr["date_parsed"] >= cutoff_date)
            ] if not acct_drr.empty else acct_drr
            ac = select_daily_action_code(window_drr)
            if ac is not None:
                df.at[idx, "ACTION CODE"] = ac

        # --- FV REMARKS from field file ---
        fv_changed_this_run = False
        if field_df is not None and not field_df.empty:
            acct_field = field_df[field_df["pn"] == pn].sort_values("date_parsed", ascending=False)
            if not acct_field.empty:
                new_fv, changed = update_fv_remarks(
                    str(row.get("FV REMARKS", "") or ""),
                    acct_field,
                    is_overall,
                )
                df.at[idx, "FV REMARKS"] = new_fv
                if changed:
                    changed_fv_pns.append(pn)
                    fv_changed_this_run = True

                updated_row = update_status_columns(dict(df.iloc[idx]), acct_field.iloc[0])
                for col in ["CLIENT STATUS", "ADDRESS STATUS", "UNIT STATUS", "RFD"]:
                    df.at[idx, col] = updated_row.get(col, row.get(col))

        # --- REPO AI → FV REMARKS mirroring (only if not already updated
        #     by real field data above, and only when genuinely new) ---
        current_status_remarks = str(df.at[idx, "STATUS REMARKS"] or "")
        current_fv = str(df.at[idx, "FV REMARKS"] or "")
        merged_fv, repo_changed = merge_repo_ai_into_fv(current_fv, current_status_remarks)
        if repo_changed:
            df.at[idx, "FV REMARKS"] = merged_fv
            repo_ai_added_pns.append(pn)
            fv_changed_this_run = True

        # --- Staleness flag: has FV content but nothing fresh this run ---
        fv_now = str(df.at[idx, "FV REMARKS"] or "").strip()
        if fv_now and fv_now != FV_PLACEHOLDER and not fv_changed_this_run:
            stale_fv_pns.append(pn)

        # --- ADDRESS STATUS placeholder rule (applied after FV updates) ---
        row_dict = apply_address_status_placeholder(dict(df.iloc[idx]))
        df.at[idx, "ADDRESS STATUS"] = row_dict["ADDRESS STATUS"]

        # Update DATE to report_date
        df.at[idx, "DATE"] = report_date

    return df, zero_activity_pns, changed_fv_pns, repo_ai_added_pns, stale_fv_pns


def reset_stale_fv_remarks(daily_df: pd.DataFrame, pns_to_reset: list[str]) -> pd.DataFrame:
    """
    Explicitly reset FV REMARKS (and the dependent ADDRESS STATUS
    placeholder) to the placeholder state for the given PNs. This is a
    SEPARATE, explicit step the UI should call only after the user
    confirms which flagged stale accounts should actually be reset —
    never applied automatically inside update_all_remarks.
    """
    df = daily_df.copy()
    pn_set = set(pns_to_reset)
    for idx, row in df.iterrows():
        pn = str(row.get("PN", "")).strip()
        if pn in pn_set and not is_held(row.get("PULLED_OUT TAG")):
            df.at[idx, "FV REMARKS"] = FV_PLACEHOLDER
            df.at[idx, "ADDRESS STATUS"] = ADDRESS_PLACEHOLDER
    return df
