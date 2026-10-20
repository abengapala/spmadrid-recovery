import io, pandas as pd, streamlit as st
from pipeline import run, precheck

st.set_page_config(page_title="SP MADRID Recovery Report", layout="wide")
st.title("SP MADRID – Recovery Status Report")
st.caption("Upload your latest report, the RAW DRR export and the field result file. The tool applies your clean-remark formula, CMS STATUS / RANK / RPC and the field date + message clean-up itself.")

with st.sidebar:
    pw = st.text_input("Report password", "cbs1234", type="password")
    mode = st.radio("If your status differs from the field result",
                    ["Field wins only if newer than your latest FV entry (recommended)",
                     "Field always wins", "Keep my value"], index=0)
    refup = st.file_uploader("Reference table (optional)", type="xlsx", help="REFERENCE_BASED.xlsx. Upload only if it changed; a copy is built in.")
    use_date = st.checkbox("Set the as-of date myself")
    rd = st.date_input("As-of date") if use_date else None
    st.markdown("**No field result?** Leave it. The whole field file is re-read every run, so it fills in once the field sends it.")

c1, c2, c3 = st.columns(3)
rep = c1.file_uploader("1. Latest report (.xlsx)", type="xlsx")
drr = c2.file_uploader("2. DRR – all dates since last report", type="xlsx")
fld = c3.file_uploader("3. Whole field file", type="xlsx")

if rep and drr and fld:
    try:
        pc = precheck(rep.getvalue(), drr.getvalue(), fld.getvalue(), pw, refup.getvalue() if refup else None)
        st.subheader("Check before running")
        st.write(f"**Report:** {pc['accounts']} accounts, last updated as of {pc['report_asof']:%b %d, %Y}.  "
                 f"**DRR:** {pc['drr_rows']:,} usable rows, {pc['drr_from']:%b %d} to {pc['drr_to']:%b %d, %Y}.  "
                 f"**Field:** {pc['field_rows']:,} rows, latest {pc['field_latest']:%b %d, %Y}.")
        if pc["missing_dates"]:
            st.error("The DRR has no rows for these weekdays since the last report: " + ", ".join(str(d) for d in pc["missing_dates"]) + ". Their effort would be missing.")
        for w in pc["warnings"]:
            if w.startswith("⚠️"): st.warning(w)
    except Exception as e:
        st.warning(f"Could not pre-check the files: {e}")
if st.button("Update report", type="primary", disabled=not (rep and drr and fld)):
    try:
        with st.spinner("Updating…"):
            m = "field" if mode.startswith("Field always") else "mine" if mode.startswith("Keep") else "newer"
            res = run(rep.getvalue(), drr.getvalue(), fld.getvalue(), pw, rd, m, refup.getvalue() if refup else None)
        n = res["notes"]
        st.success(f"Done. As of {n['report_date']:%b %d, %Y}. DRR dates: {', '.join(str(d) for d in n['drr_dates'])}. {res['accounts']} accounts.")
        st.download_button("Download updated report", res["file"],
                           f"SP_MADRID_Recovery_Status_Report_as_of_{n['report_date']:%B_%d_%Y}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        for w in res["warnings"]:
            if w.startswith("⚠️"): st.warning(w)
        t1, t2, t3, t4, t5, t6 = st.tabs(["Trails Upload", "No effort", "No field result", "Status changes", "Conflicts", "PTP INVENTORY"])
        with t1:
            tr = res["trails"]; st.write(f"{len(tr)} rows"); st.dataframe(tr.drop(columns=["txt"]).assign(remark=tr["txt"]), use_container_width=True)
        with t2:
            st.write(f"{len(res['new_endo'])} new endo accounts with no effort (expected). Genuine no-effort accounts:")
            st.dataframe(res["no_effort"], use_container_width=True)
        with t3:
            st.write(f"{len(res['no_field'])} accounts have no field result in the file.")
            st.dataframe(res["no_field"], use_container_width=True)
        with t4:
            st.dataframe(res["changed"], use_container_width=True)
        with t5:
            st.write("Your value differs from the field and was kept. Please review.")
            st.dataframe(res["conflicts"], use_container_width=True)
        with t6:
            st.write(f"{len(res['new_ptp'])} accounts were added to PTP INVENTORY (PTP / NEGO / KEPT on DAILY, not in the inventory yet):")
            st.dataframe(res["new_ptp"], use_container_width=True)
            st.write("PTP / NEGO in this DRR. Accounts not on DAILY are reported only; add to PTP INVENTORY yourself after review.")
            st.dataframe(res["candidates"], use_container_width=True)
        with st.expander("Cleaning stats / warnings"):
            st.json(res["stats"]); st.write(res["warnings"])
    except Exception as e:
        st.error(f"Could not update: {e}")
