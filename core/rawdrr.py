"""
core/rawdrr.py
Lets the tool take the RAW CMS DRR export. Ports the user's Excel logic:
  - FINAL CMS REMARK  = the LET(...) clean-remark formula
  - CMS STATUS / RANK / RPC = lookup by Status (table learned from the user's labelled DRR)
Columns the user already filled in are honoured; only missing ones are computed.
"""
from __future__ import annotations
import re
import pandas as pd

# Status -> (CMS STATUS, RANK, RPC), taken from the user's labelled DRR (Oct 15-16)
STATUS_MAP = {
 "BUSY": ("NIS_KOR_CBR", 3, "NOT RPC"), "RNA": ("NIS_KOR_CBR", 3, "NOT RPC"),
 "CALL - FOLLOW UP COMPLIANT": ("FPTP", 1.75, "RPC"), "CALL - FOLLOW UP KOR": ("FPTP", 1.75, "RPC"),
 "CALL - FOLLOW UP LMTRC": ("NIS_KOR_CBR", 1.75, "RPC"), "CALL - FOLLOW UP UNCONTACTABLE": ("NIS_KOR_CBR", 1.75, "RPC"),
 "CALL - NEG_BUSY": ("NIS_KOR_CBR", 2.75, "NOT RPC"), "CALL - NEG_CBR": ("NIS_KOR_CBR", 2.75, "NOT RPC"),
 "CALL - NEG_DROPPED": ("NIS_KOR_CBR", 2.75, "NOT RPC"), "CALL - NEG_KOR": ("NIS_KOR_CBR", 2.5, "NOT RPC"),
 "CALL - NEG_LEAVE MSG TO 3RD PARTY": ("LM", 2.75, "NOT RPC"), "CALL - NEG_NOT IN SERVICE": ("NIS_KOR_CBR", 2.75, "NOT RPC"),
 "CALL - NEG_UNATTENDED": ("NIS_KOR_CBR", 2.75, "NOT RPC"), "CALL - NEG_WRONG NUMBER": ("NIS_KOR_CBR", 2.75, "NOT RPC"),
 "CALL - POS_BUSY": ("NIS_KOR_CBR", 2.5, "NOT RPC"), "CALL - POS_DROPPED": ("NIS_KOR_CBR", 2.5, "NOT RPC"),
 "CALL - POS_KOR": ("NIS_KOR_CBR", 2.5, "NOT RPC"), "CALL - POS_LEAVE MSG TO 3RD PARTY": ("LM", 2.5, "NOT RPC"),
 "CALL - PTP FULL UPDATE": ("ECA_B3_PTP", 1.5, "RPC"), "CALL - PTP PARTIAL": ("ECA_B3_PTP", 1.5, "RPC"),
 "CALL - UNDERNEGO": ("NEGO", 1.75, "RPC"), "CALL - UNIT_CARNAPPED": ("NEGO", 1.75, "RPC"),
 "EMAIL - FOLLOW UP COMPLIANT": ("FPTP", 1.75, "RPC"), "EMAIL - PTP_PARTIAL": ("ECA_B3_PTP", 1.5, "RPC"),
 "EMAIL - RESPONSIVE": ("NEGO", 1.75, "RPC"), "EMAIL SENT": ("EMAIL", 3, "NOT RPC"),
 "FIELD - FOLLOW UP COMPLIANT": ("FPTP", 1.75, "RPC"), "FIELD - MOVED_OUT": ("MOVED_OUT_NLR", 2, "NOT RPC"),
 "FIELD - RESULT": ("FFV", 3, "NOT RPC"), "FIELD REQUEST - FOR REVISIT": ("FFV", 3, "NOT RPC"),
 "FIELD REQUEST - NEW_ADDRESS": ("FFV", 3, "NOT RPC"), "REPO AI - UNDERNEGO": ("NEGO", 1.75, "RPC"),
 "SKIP - NEGATIVE": ("SKIP_SOCIAL_MEDIA", 2.75, "NOT RPC"), "SKIP - SMEDIA ACCOUNT": ("SKIP_SOCIAL_MEDIA", 2, "NOT RPC"),
 "SMS - FOLLOW UP COMPLIANT": ("FPTP", 1.75, "RPC"), "SMS - NEG_SENT MESSAGE": ("SMS", 2.75, "NOT RPC"),
 "SMS REPLY": ("SMS", 1.75, "RPC"), "SMS REPLIED": ("SMS", 1.75, "RPC"), "SMS SENT": ("SMS", 3, "NOT RPC"),
 "BULK SMS SENT": ("SMS", 3, "NOT RPC"),
 "VIBER - DELIVERED": ("SMS", 2, "NOT RPC"), "VIBER - FOLLOW UP COMPLIANT": ("FPTP", 1.75, "RPC"),
 "VIBER - KEPT_PARTIAL": ("WITH_PTP", 1, "RPC"), "VIBER - PARTIAL": ("ECA_B3_PTP", 2, "NOT RPC"),
 "VIBER - POS_SENT A MESSAGE": ("SMS", 2.5, "NOT RPC"), "VIBER - PTP FULL UPDATE": ("ECA_B3_PTP", 1.5, "RPC"),
 "VIBER - PTP PARTIAL": ("ECA_B3_PTP", 1.5, "RPC"), "VIBER - RESPONSIVE": ("NEGO", 1.75, "RPC"),
 "VIBER - UNIT DAMAGE OR WRECK": ("NEGO", 1.75, "RPC"),
}

_STRIP1 = ["/", "\\", "'", "(", ")"]
_STRIP2 = ["#", "_", ":", "[", "]"]
_STRIP3 = ["+", "&", "<", ">", "|"]
_STRIP4 = ["%", "?", "=", "OBD", "INB"]


def _s(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)


def let_clean(status: str, remark: str, rpc: str) -> str:
    """Port of the user's =LET(...) clean-remark formula."""
    st, rem = _s(status).strip(), _s(remark)
    u = st.upper()
    if u == "SKIP - SMEDIA ACCOUNT": original = "Skip smed done no info gathered"
    elif u == "EMAIL SENT": original = "EMAIL SENT"
    elif u in ("BULK SMS SENT", "SMS SENT", "SMS"): original = " Sent SMS Template"
    elif u in ("SMS REPLY", "SMS REPLY AS"): original = " CLIENT SMS REPLY"
    elif u in ("PM", "PU"): original = " Call connected No Response"
    elif u == "RNA": original = " Ringing no answer"
    elif u == "BUSY": original = " Line busy"
    elif u == "DROPPED": original = " Call dropped"
    else: original = rem
    step = original
    for ch in _STRIP1 + _STRIP2 + _STRIP3 + _STRIP4:
        step = step.replace(ch, "")
    after = original.split("|", 1)[1] if "|" in original else step
    cleaned = after
    if _s(rpc).upper() == "NOT RPC":
        i = after.upper().find("REMARKS:")
        if i >= 0 and "REMARKS:" in after:
            cleaned = after.split("REMARKS:", 1)[1]
    return cleaned.strip()


import io, os
_REF = None


def load_reference(file_bytes: bytes | None = None) -> dict:
    """Status -> (CMS STATUS, RANK, RPC) from the user's REFERENCE_BASED file (first match wins,
    like a lookup). With no file, the copy shipped with the tool is used."""
    global _REF
    if file_bytes is None and _REF is not None:
        return _REF
    if file_bytes is None:
        d = pd.read_csv(os.path.join(os.path.dirname(__file__), "reference_table.csv"))
        d.columns = ["key", "code", "rank", "rpc"]
    else:
        d = pd.read_excel(io.BytesIO(file_bytes)).iloc[:, [0, 1, 4, 7]]
        d.columns = ["key", "code", "rank", "rpc"]
        d["key"] = d["key"].apply(lambda s: re.sub(r"\s+", " ", str(s).strip().upper()))
        d = d.drop_duplicates("key", keep="first")
    ref = {r.key: (str(r.code), float(r["rank"]), str(r.rpc)) for _, r in d.iterrows()}
    if file_bytes is None:
        _REF = ref
    return ref


def classify(status: str, ref: dict | None = None):
    """(CMS STATUS, RANK, RPC) for a Status; second value of the tuple is None if unknown."""
    key = re.sub(r"\s+", " ", _s(status).strip().upper())
    ref = ref if ref is not None else load_reference()
    if key in ref:
        return ref[key], True
    if key in STATUS_MAP:
        return STATUS_MAP[key], True
    # keyword fallback so an unseen promise status is never silently dropped
    if "PTP" in key or "KEPT" in key:
        return ("WITH_PTP" if "KEPT" in key else "ECA_B3_PTP", 1.5, "RPC"), False
    if "NEGO" in key:
        return ("NEGO", 1.75, "RPC"), False
    return ("", 9, "NOT RPC"), False


def ensure_labelled(df: pd.DataFrame, warnings: list, ref: dict | None = None) -> pd.DataFrame:
    """Add CMS STATUS / RANK / RPC / FINAL CMS REMARK when the raw export doesn't have them."""
    df = df.copy()
    sysrow = df["Status"].isna() & df["Remark"].astype(str).str.contains("Sub Special Status Change", case=False, na=False)
    if sysrow.any():
        warnings.append(f"Raw DRR: dropped {int(sysrow.sum())} 'Sub Special Status Change' system rows")
        df = df[~sysrow].reset_index(drop=True)
    need = [c for c in ["CMS STATUS", "RANK", "RPC"] if c not in df.columns or df[c].isna().all()]
    unknown = {}
    if need:
        cls = df["Status"].apply(lambda s: classify(s, ref))
        if "CMS STATUS" in need: df["CMS STATUS"] = cls.apply(lambda t: t[0][0])
        if "RANK" in need: df["RANK"] = cls.apply(lambda t: t[0][1])
        if "RPC" in need: df["RPC"] = cls.apply(lambda t: t[0][2])
        for s, (t, known) in zip(df["Status"], cls):
            if not known: unknown[_s(s)] = unknown.get(_s(s), 0) + 1
        warnings.append("Raw DRR: CMS STATUS / RANK / RPC computed from Status (" + ", ".join(need) + ")")
    if unknown:
        warnings.append("⚠️ Statuses not in the known table (please check): " +
                        "; ".join(f"{k} x{v}" for k, v in unknown.items()))
    if "FINAL CMS REMARK" not in df.columns or df["FINAL CMS REMARK"].isna().all():
        df["FINAL CMS REMARK"] = [let_clean(s, r, p) for s, r, p in zip(df["Status"], df["Remark"], df["RPC"])]
        warnings.append("Raw DRR: FINAL CMS REMARK computed with your clean-remark formula")
    if "FINAL REMARK" not in df.columns:
        df["FINAL REMARK"] = df["Remark"]
    return df
