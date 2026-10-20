"""
core/rebuild2.py
STATUS REMARKS full-history rebuild (clean text) and FV REMARKS rebuild from
the whole field file.
"""
from __future__ import annotations

import re
from typing import Optional

import pandas as pd

from core.clean import clean_text
from core.remarks import parse_remarks_string, FV_PLACEHOLDER

_EARLIEST = pd.Timestamp("2026-08-01")
_LATEST = pd.Timestamp("2026-12-31")
_DATE_PFX = re.compile(r"^\s*(\d{1,2})\.(\d{2})\.(\d{4})\s+")


def _fmt(ts: pd.Timestamp) -> str:
    return ts.strftime("%m.%d.%Y")


_EMBED = re.compile(r"\s+(?=\d{1,2}\.\d{2}\.\d{4}\s)")


def _flatten(entries, stats):
    """Split entries that were glued together without a comma (an embedded
    MM.DD.YYYY token starts a new entry)."""
    from core.remarks import RemarkEntry, _parse_mm_dd_yyyy
    flat = []
    for e in entries:
        parts = _EMBED.split(e.text) if e.date is not None else [e.text]
        if len(parts) == 1:
            flat.append(e)
            continue
        stats["glued_split"] = stats.get("glued_split", 0) + len(parts) - 1
        flat.append(RemarkEntry(date_str=e.date_str, text=parts[0].strip(), full=f"{e.date_str} {parts[0].strip()}", date=e.date))
        for p in parts[1:]:
            m = _DATE_PFX.match(p + " ")
            if not m:
                flat.append(RemarkEntry(date_str=e.date_str, text=p.strip(), full=p.strip(), date=e.date))
                continue
            mm, dd, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
            ts = pd.Timestamp(year=yy, month=mm, day=dd)
            body = (p + " ")[m.end():].strip()
            flat.append(RemarkEntry(date_str=ts.strftime("%m.%d.%Y"), text=body, full=f"{ts.strftime('%m.%d.%Y')} {body}", date=ts))
    return flat


def rebuild_status_history(
    existing: str,
    drr_acct: pd.DataFrame,
    cutoff: pd.Timestamp,
    stats: dict,
) -> str:
    """
    Whole STATUS REMARKS rebuilt in clean form:
      - old entries (dated before cutoff) are cleaned with core.clean and kept
      - entries on/after cutoff come fresh from the DRR (clean column)
      - exact duplicates (same date + same clean text) collapsed
      - entries whose day/month were transposed (e.g. 02.10.2026 sitting
        between 10.02 and 09.30) are re-dated when the swap puts them in order
    Returns the new newest-first, comma-separated string.
    """
    # single-digit months ("9.04.2026") are not recognised as a date prefix
    existing = re.sub(r"(^|,\s*)(\d)\.(\d{2})\.(\d{4})(?=\s)", r"\g<1>0\2.\3.\4", str(existing or ""))
    entries = _flatten(parse_remarks_string(existing), stats)

    # --- fix transposed day/month dates (out of order in a newest-first list) ---
    dates = [e.date for e in entries]
    fixed_dates = list(dates)
    for i, d in enumerate(dates):
        if d is None:
            continue
        nxt = next((x for x in dates[i + 1:] if x is not None and x != d), None)
        prv = next((x for x in reversed(fixed_dates[:i]) if x is not None), None)
        if d < _EARLIEST and d.day <= 12:
            try:
                sw = pd.Timestamp(year=d.year, month=d.day, day=d.month)
            except Exception:
                continue
            if _EARLIEST <= sw <= _LATEST:
                fixed_dates[i] = sw
                stats["date_fixed"] = stats.get("date_fixed", 0) + 1
            continue
        if nxt is not None and d < nxt and d.day <= 12:
            try:
                sw = pd.Timestamp(year=d.year, month=d.day, day=d.month)
            except Exception:
                continue
            if sw >= nxt and (prv is None or sw <= prv):
                fixed_dates[i] = sw
                stats["date_fixed"] = stats.get("date_fixed", 0) + 1

    kept: list[tuple[Optional[pd.Timestamp], str]] = []
    seen = set()
    for e, d in zip(entries, fixed_dates):
        if d is not None and d >= cutoff:
            stats["old_replaced"] = stats.get("old_replaced", 0) + 1
            continue  # replaced by fresh DRR window
        if d is None:
            kept.append((None, e.full))  # undated fragment: left untouched
            stats["undated_untouched"] = stats.get("undated_untouched", 0) + 1
            continue
        c = clean_text(e.text)
        if c != e.text:
            stats["cleaned"] = stats.get("cleaned", 0) + 1
        key = (d.date(), c.lower())
        if key in seen:
            stats["dupes_collapsed"] = stats.get("dupes_collapsed", 0) + 1
            continue
        seen.add(key)
        kept.append((d, f"{_fmt(d)} {c}"))

    kept.sort(key=lambda x: (x[0] is not None, x[0] if x[0] is not None else pd.Timestamp.min), reverse=True)
    new_entries: list[str] = []
    if drr_acct is not None and not drr_acct.empty:
        w = drr_acct[drr_acct["date_parsed"].notna() & (drr_acct["date_parsed"] >= cutoff)].copy()
        w = w.sort_values(["date_parsed", "time_str"], ascending=[False, False])
        seen_new = set()
        for _, r in w.iterrows():
            c = str(r["remark"]).strip()
            if not c:
                continue
            key = (r["date_parsed"].date(), c.lower())
            if key in seen_new:
                continue
            seen_new.add(key)
            new_entries.append(f"{_fmt(r['date_parsed'])} {c}")
    stats["new_entries"] = stats.get("new_entries", 0) + len(new_entries)

    return ", ".join([*new_entries, *[t for _, t in kept]])


# ---------------------------------------------------------------------------
# FV REMARKS
# ---------------------------------------------------------------------------

def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def _entry_date_key(text: str):
    m = _DATE_PFX.match(text)
    if not m:
        return None
    try:
        return pd.Timestamp(year=int(m.group(3)), month=int(m.group(1)), day=int(m.group(2)))
    except Exception:
        return None


def rebuild_fv(existing: str, field_rows: pd.DataFrame, stats: dict) -> tuple[str, bool]:
    """
    Whole field file is authoritative for field-visit entries: every field
    entry for the account (newest first, duplicates collapsed) is used. Existing
    FV entries that are NOT in the field file (REPO AI / SRC FIELD lines mirrored
    from the DRR earlier) are kept, merged in date order. Returns (new, changed).
    """
    cur = "" if existing is None else str(existing).strip()
    cur_entries = [] if cur in ("", FV_PLACEHOLDER) else [e.full for e in parse_remarks_string(cur)]

    f_entries: list[tuple[pd.Timestamp, str]] = []
    seen = set()
    fr = field_rows.sort_values("date_parsed", ascending=False, na_position="last")
    for _, r in fr.iterrows():
        t = str(r.get("FV REMARK", "")).strip()
        if not t or t.lower() in ("nan", "0"):
            continue
        d = _entry_date_key(t) or (r["date_parsed"] if pd.notna(r["date_parsed"]) else None)
        if _entry_date_key(t) is None and d is not None:
            t = f"{_fmt(d)} {t}"
        k = (d.date() if d is not None else None, _norm(t)[:80])
        if k in seen:
            continue
        seen.add(k)
        f_entries.append((d, t))
    if not f_entries:
        return cur, False

    field_keys = {(d.date() if d is not None else None, _norm(t)[:40]) for d, t in f_entries}
    field_norm40 = {_norm(t)[:40] for _, t in f_entries}
    retained: list[tuple[Optional[pd.Timestamp], str]] = []
    for e in cur_entries:
        d = _entry_date_key(e)
        n40 = _norm(e)[:40]
        if n40 in field_norm40:
            continue  # same entry, superseded by the field version
        if "AWAITING FIELD STATUS" in e.upper() and len(e) < 40:
            continue
        retained.append((d, e))

    merged = f_entries + retained
    merged.sort(key=lambda x: (x[0] if x[0] is not None else pd.Timestamp.min), reverse=True)
    new = ", ".join(t for _, t in merged)
    stats["fv_retained_nonfield"] = stats.get("fv_retained_nonfield", 0) + len(retained)
    return new, new != cur
