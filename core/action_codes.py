"""
core/action_codes.py
Maps a raw DRR action code / remark into the ACTION CODE legend vocabulary.

Per the confirmed rule (Streamlit_Automation_Spec.md section 5):
- When CMS STATUS is already present (the DRR's pre-cleaned column), use it
  directly — it already reflects the legend mapping.
- Only when CMS STATUS is unavailable and a raw Status must be classified:
    * Contains "EMAIL" -> EMAIL
    * Contains "SMS" -> SMS
    * Starts with "CALL" AND contains "WRONG NUMBER" or "NOT IN SERVICE" -> NIS_KOR_CBR
    * Everything else -> leave as the raw DRR text, don't force a mapping
      that doesn't fit. This was an explicit decision, not an oversight.
"""
from __future__ import annotations


def resolve_action_code(raw_code: str, source: str) -> str:
    """
    source == "CMS": raw_code is already the cleaned CMS STATUS value —
    use it as-is (trimmed).
    source == "RAW" (or anything else, including unverified sources):
    apply the confirmed fallback mapping rules to a raw Status string.
    """
    code = (raw_code or "").strip()
    if not code:
        return ""

    if source == "CMS":
        return code

    upper = code.upper()
    if "EMAIL" in upper:
        return "EMAIL"
    if "SMS" in upper:
        return "SMS"
    if upper.startswith("CALL") and ("WRONG NUMBER" in upper or "NOT IN SERVICE" in upper):
        return "NIS_KOR_CBR"

    # No clean legend equivalent — leave as the raw DRR text rather than
    # forcing it into a code that doesn't fit.
    return code
