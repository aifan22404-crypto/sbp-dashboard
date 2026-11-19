import os, openpyxl
from openpyxl.styles import Alignment, PatternFill
R=os.path.expanduser("~/mnt/SBP_Dashboard")
wb=openpyxl.load_workbook(f"{R}/05_red_flags/red_flags_log.xlsx"); ws=wb["red_flags"]
for r in ws.iter_rows(min_row=2):
    if r[0].value=="RF02":
        r[2].value="🟢 Low"; r[2].fill=PatternFill("solid",fgColor="C6E0B4")
        r[13].value="EXPLAINED 2026-10-16 — DBD data error, not a real event"
        r[7].value+=(" RESOLVED: DBD firm-level list for Apr 2566 shows Narathiwat 10 (not 119), Kalasin 19 (not 511), Surin 32 (not 338). National total identical (6,041) but the aggregate file "
                     "put other provinces' counts under the wrong names: Narathiwat got Nakhon Ratchasima's 119, Kalasin got Chonburi's 511, Surin got Phuket's 338, Chumphon got Khon Kaen's 101. "
                     "Every other month 2565–2568 matches exactly after cleaning.")
        r[12].value="Use firm-list counts for Apr 2566 (clean DBD file corrected). Lesson: always cross-check aggregate stats against micro data."
        r[14].value="'A spike in one month turned out to be a mislabeled table in the official dataset — checked against the firm list.'"
ws.append(["RF22","2026-10-16","🟢 Low","B. Source quality","DBD monthly firm lists","All","2567–2568",
 "Late-2567/2568 files pad with up to 2,972 blank rows, switch province format to 'จ.xxx', and Dec 2568 mixes date formats; the Jun 2569 dissolution resource in the catalog points to a new-firms file.",
 "Naive counts would be wrong (e.g. Nov 2568 shows 8,159 rows but 5,554 real firms).","Publishing pipeline changes.","—","Cleaning rules applied; totals then match DBD's aggregate statistics.","Always drop blank rows / normalize names before counting.","Explained","—"])
for c in ws[ws.max_row]: c.alignment=Alignment(wrap_text=True,vertical="top")
ws.cell(ws.max_row,3).fill=PatternFill("solid",fgColor="C6E0B4")
wb.save(f"{R}/05_red_flags/red_flags_log.xlsx")
wb=openpyxl.load_workbook(f"{R}/00_admin/source_log.xlsx"); s=wb["source_log"]
s.append(["DBD04","DBD_{NewFirms|DissolvedFirms}_list_<YYYYMM>_dl20261016.csv (96)","02_business_dbd/{new|dissolved}_firms_monthly","กรมพัฒนาธุรกิจการค้า (DBD) — downloaded in Thailand","openapi.dbd.go.th/juristic_person/{registration|dissolution}/99_YYYYMM_{1|2}.csv","2026-10-16","Jan 2565 – Dec 2568","Firm (head-office district)","All 77; used PTN, YLA, NWT, SKA (+4 districts)","csv","P1,P4, red flags","Pass with caution","426,356 firm rows nationwide; 8,426 in our provinces. Blank rows / 'จ.' prefix cleaned (RF22). Resolved RF02."])
s.append(["FI-DoF","DoF_*_dl20261016 (12 CSV + 2 XLSX)","05_fisheries/DoF","กรมประมง — downloaded in Thailand","catalog.fisheries.go.th (see _DOWNLOAD_LINKS_DoF_BLOCKED.txt)","2026-10-16","2564–2567 (aquaculture 2559+, vessels 2559–2568)","Province","All; used PTN, NWT, SKA, YLA","csv/xlsx","P1,P3,P4","Pass","Marine capture by sea area only (national context)."])
s.append(["DBD-T","02_clean/DBD_firms_by_business_type_2565-2568_clean.xlsx","02_clean","Built from DBD04",None,"2026-10-16","2565–2568","Province + Songkhla 4 districts","PTN,YLA,NWT,SKA,SKA-4","xlsx","P1,P4","Pass","TSIC sections, top divisions, by district."])
s.append(["FI-C2","02_clean/Fisheries_DoF_4prov_2562-2567_clean.xlsx","02_clean","Built from FI-DoF",None,"2026-10-16","2562–2568","Province","PTN,NWT,SKA,YLA","xlsx","P1,P3,P4","Pass","Landings, species, aquaculture, vessels."])
for r in s.iter_rows(min_row=2):
    if r[0].value=="FI-X": r[11].value="Resolved"; r[12].value="DoF data obtained via contact in Thailand 2026-10-16 (see FI-DoF)."
wb.save(f"{R}/00_admin/source_log.xlsx")
wb=openpyxl.load_workbook(f"{R}/00_admin/project_stats.xlsx"); s=wb["summary"]
def setv(r,v,n=None):
    s.cell(r,2).value=v
    if n is not None: s.cell(r,3).value=n
setv(2,333,"223 + 110 จากไทย (DBD 96 + กรมประมง 14)")
setv(3,161,"+2 กรมประมง"); setv(4,962,None)
setv(6,172,"+108 จากไทย"); setv(7,73563,"+18,776 แถวในพื้นที่ (DBD 8,426 + กรมประมง 10,350) — นับเฉพาะแถวที่ใช้"); setv(8,1039047,"+192,403")
setv(13,3605,"ประมาณ 50 แถว/หน้า — ค่าประมาณ"); setv(14,7105,"3,500 PDF + 3,605 ประมาณการ")
setv(16,14,"… + DBD แยกประเภทธุรกิจ, กรมประมง")
setv(17,19,s.cell(17,3).value+", กรมประมง")
setv(18,"2026-10-16 (Thailand batch: DBD firm lists + DoF)")
for r in s.iter_rows(min_row=2):
    if r[0].value and str(r[0].value).startswith("Red flags"): r[1].value=22; r[2].value="🔴4 / 🟡10 / 🟢8 — RF02 resolved"
    if r[0].value and str(r[0].value).startswith("แถว UCDP"):
        r[0].value="แถวข้อมูลระดับประเทศที่สแกน (ไม่นับเป็นหน้า)"; r[1].value=1001731; r[2].value="UCDP 417,968 + DBD 426,356 + กรมประมง 157,407 — ใช้จริงเฉพาะแถวในพื้นที่"
i=wb["inventory"]
i.append(["02_business_dbd","DBD_*_list_* (96 CSV)","DBD (โหลดจากไทย)","CSV",None,None,8426,"Active","2026-10-16"])
i.append(["05_fisheries/DoF","DoF_* (12 CSV + 2 Excel)","กรมประมง (โหลดจากไทย)","CSV+Excel",None,2,10350,"Active","2026-10-16"])
i.append(["02_clean","DBD_firms_by_business_type_2565-2568_clean.xlsx","Built (Claude)","Excel (clean)",None,5,None,"Output","2026-10-16"])
i.append(["02_clean","Fisheries_DoF_4prov_2562-2567_clean.xlsx","Built (Claude)","Excel (clean)",None,7,None,"Output","2026-10-16"])
wb["milestones"].append(["2026-10-16","Thailand batch (via family member): DBD firm lists 96 + DoF 14 — RF02 solved; scanned rows pass 1 million",110,None,18776])
wb.save(f"{R}/00_admin/project_stats.xlsx")
wb=openpyxl.load_workbook(f"{R}/00_admin/data_checklist.xlsx"); c=wb["checklist"]
for r in c.iter_rows(min_row=2):
    if r[0].value==2: r[7].value="Done (incl. TSIC)"; c.cell(r[0].row,10).value="Firm lists 2565–2568 → by business type + Songkhla 4 districts. RF02 resolved."
    if r[0].value==7: r[7].value="Done"; r[7].fill=PatternFill("solid",fgColor="C6E0B4"); c.cell(r[0].row,10).value="FMO ports + DoF province landings/aquaculture/vessels 2564–2567."
wb.save(f"{R}/00_admin/data_checklist.xlsx")
print("ok")
