"""
core/clean.py
Normalizes STATUS REMARKS entry text to the clean form used by the DRR's
FINAL CMS REMARK column.

The DRR's clean column is, in practice, just the human remark payload
("CBR", "NOT RINGING", "EMAIL SENT", "Sent SMS Template") without the phone
number / RFD / SRC scaffolding. Older entries in the report still carry that
scaffolding in several inconsistent spellings, so the same extraction is
applied to them. Anything that doesn't match a known pattern is returned
UNCHANGED (never guessed at).
"""
from __future__ import annotations

import re

_MARKER_RE = re.compile(r"REM(?:ARKS|ARS|AKRS)\s*[:;]?\s*(.*)$", re.IGNORECASE | re.DOTALL)
_EMAIL_RE = re.compile(r"^\s*Sent Email\s*-\s*Notice of Uncontactable\s*$", re.IGNORECASE)
_SMS_RE = re.compile(r"^\s*(?:\d{9,12}\s+)?Sent SMS\s*$", re.IGNORECASE)
_DIALER_RE = re.compile(r"^\s*(?:\d{9,12}\s+)?(Ringing no answer|Line busy)\s*$", re.IGNORECASE)
_LEAD_DATES_RE = re.compile(r"^(?:\d{1,2}\.\d{2}\.\d{4}\s+)+")


def clean_text(text: str) -> str:
    """Return the clean form of one remark's text (date prefix NOT included)."""
    if text is None:
        return ""
    t = str(text).strip()
    if not t:
        return t

    # repeated date tokens inside the text ("09.17.2026 09.17.2026 ...")
    t = _LEAD_DATES_RE.sub("", t).strip()

    if _EMAIL_RE.match(t):
        return "EMAIL SENT"
    if _SMS_RE.match(t):
        return "Sent SMS Template"
    m = _DIALER_RE.match(t)
    if m:
        return m.group(1).upper()
    if t.lower() == "client sms responsive":
        return "CLIENT SMS REPLY"

    for _ in range(3):
        m = _MARKER_RE.search(t)
        if not m:
            break
        payload = m.group(1).strip()
        if not payload:
            break
        t = payload
        # keep unwrapping only while scaffolding is still present
        if not re.search(r"\bSRC\b|\bSOURCE\b|\bRFD\b|\|", t, re.IGNORECASE):
            break
    return t


def looks_raw(text: str) -> bool:
    """True if text still carries RFD/SRC/SOURCE scaffolding after cleaning."""
    t = str(text)
    return bool(re.search(r"\bSRC\b|\bSOURCE\b|\bRFD\b|\|", t))
