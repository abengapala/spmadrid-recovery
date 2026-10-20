# SP MADRID Recovery Report – Update Tool

## Run
    pip install -r requirements.txt
    streamlit run app.py

## Each update (process)
1. Send/upload 3 files: latest report (password cbs1234), RAW DRR export (no helper columns needed) with ALL dates since the last report, WHOLE field file (PN column filled in; PN 0 rows cannot be matched).
2. Click "Update report". As-of date = last DRR date (override in sidebar).
3. Review tabs: No effort, No field result, Conflicts, PTP / NEGO.
4. Download the report. Add any new PTP accounts to PTP INVENTORY yourself.
5. Next time, upload the downloaded report as "latest report". Accounts with no field result fill in automatically once they appear in the whole field file.

## What it does
DAILY: clean STATUS REMARKS history, ACTION CODE by DRR rank, FV REMARKS rebuilt from whole field, CLIENT/ADDRESS/UNIT STATUS and RFD from field, "(FOR FIELD VISIT)" when no field visit.
Trails Upload: effort on as-of date + PTP/NEGO/KEPT accounts. Held accounts and weekend DRR rows are skipped.

## Raw files
Raw DRR: FINAL CMS REMARK (your LET formula), CMS STATUS, RANK and RPC are computed from Status. Columns you already filled in are used as-is. Unknown statuses are flagged for you to confirm; add them to STATUS_MAP in core/rawdrr.py.
Field: FV REMARK = date(mm.dd.yyyy) + Message when the clean column is missing.
