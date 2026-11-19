import os, openpyxl
from openpyxl.styles import Alignment, PatternFill
R=os.path.expanduser("~/mnt/SBP_Dashboard")
wb=openpyxl.load_workbook(f"{R}/05_red_flags/red_flags_log.xlsx"); ws=wb["red_flags"]
row=["RF19","2026-10-10","🟡 Medium","B. Source inconsistency","DOPA Yala (provincial data catalog)","Yala","2568–2569",
"Same office publishes two tables: 2568 = 71 incidents in one, 21 in the other; 2569 = 22 vs 10. 2565–2567 match (20, 26, 21). Table B is titled 'ร้อยละ' (percent) but unit says 'ครั้ง' (times).",
"Two official numbers for the same year from the same office.","Table B may be a partial-year/earlier extract or mislabeled.","Under-reporting in one version.",
"Ask ที่ทำการปกครองจังหวัดยะลา which is final; compare with UCDP (Yala 2025: 8 lethal events, 12 deaths — up from 6/7 in 2024, supports the higher number).",
"Use Table A (71) with a note; show UCDP alongside.","Open","'Official Yala figures for 2025 differ between two tables (71 vs 21).'" ]
ws.append(row)
for c in ws[ws.max_row]: c.alignment=Alignment(wrap_text=True,vertical="top")
ws.cell(ws.max_row,3).fill=PatternFill("solid",fgColor="FFE699")
wb.save(f"{R}/05_red_flags/red_flags_log.xlsx")
wb=openpyxl.load_workbook(f"{R}/00_admin/source_log.xlsx"); s=wb["source_log"]
s.append(["SEC01","UCDP_GED_v26.1_global_1989-2025_dl20261010.zip + codebook PDF","08_security","Uppsala Conflict Data Program (UCDP), Uppsala University","ucdp.uu.se/downloads/ged/ged261-csv.zip","2026-10-10","1989–2025 (used 2019–2025)","Event (village coordinates) → district","Thailand: PTN, YLA, NWT, SKA","csv (zip) + pdf","P3,P4","Pass","Lethal events only (≥1 death). 417,968 global rows scanned; 2,909 Thailand; 202 in our 4 provinces 2019–2025."])
s.append(["SEC02","{NWT|YLA|SKA}_DOPA_Unrest_*_dl20261010.csv (6 files)","08_security","ที่ทำการปกครองจังหวัด นราธิวาส/ยะลา/สงขลา via provincial gdcatalog","narathiwat/yala/songkhla.gdcatalog.go.th (see _DOWNLOAD_LINKS.txt)","2026-10-10","2555–2569","NWT district×type; YLA & SKA province","NWT, YLA, SKA (whole prov)","csv","P3,P4","Pass with caution","All incident types incl. non-lethal. Pattani none published. Yala tables disagree (RF19). Songkhla 2566 '-'."])
s.append(["SEC-C","02_clean/Security_4prov_2562-2568_clean.xlsx","02_clean","Built from SEC01–SEC02",None,"2026-10-10","2019–2025","Province + District","PTN,YLA,NWT,SKA(4 districts)","xlsx","P3,P4","Pass","Comparison sheet UCDP vs official; district hotspots; event list with coordinates."])
s.append(["GPP-C","02_clean/GPP_4prov_2019-2024_clean.xlsx","02_clean","Built from GPP01 (NESDC edition 2024)",None,"2026-10-10","2019–2024p","Province","PTN,YLA,NWT,SKA(whole)","xlsx","P1,P2,P4","Pass","Summary, 19 sectors long, structure 2021 vs 2024 + real CAGR."])
for r in s.iter_rows(min_row=2):
    if r[0].value=="TO02": r[11].value="Pass"; r[12].value="Cleaned 2026-10-10 → Tourism clean file, sheet monthly_2568_vs_2567. Monthly sums = annual (checked)."
wb.save(f"{R}/00_admin/source_log.xlsx")
wb=openpyxl.load_workbook(f"{R}/00_admin/project_stats.xlsx"); s=wb["summary"]
def setv(r,v,n=None):
    s.cell(r,2).value=v
    if n is not None: s.cell(r,3).value=n
setv(2,188,"180 + 8 ความไม่สงบ (UCDP zip + codebook + 6 CSV ทางการ)")
setv(6,63,"15 DBD + 1 ชายแดน + 40 ความยากจน + 7 ความไม่สงบ")
setv(7,42895,"+3,383 ความไม่สงบ (474 ทางการ + 2,909 แถว UCDP ประเทศไทย)")
setv(8,632588,"+147,053 ความไม่สงบ")
setv(10,2968,"1,556 + 940 ประมง + 29 BOI + 391 ความยากจน + 52 UCDP codebook"); setv(11,2417,None)
setv(13,2986,"ประมาณ 50 แถว/หน้า — ค่าประมาณ"); setv(14,5954,"2,968 PDF + 2,986 ประมาณการ")
setv(16,10,"แรงงาน, DBD, ชายแดน, เกษตร, ประมง, ท่องเที่ยว(+รายเดือน), BOI, ความยากจน, ความไม่สงบ, GPP")
setv(17,15,s.cell(17,3).value+", UCDP (Uppsala Univ.), ที่ทำการปกครองจังหวัด")
setv(18,"2026-10-10 (Security + GPP + tourism monthly)")
for r in s.iter_rows(min_row=2):
    if r[0].value and str(r[0].value).startswith("Red flags"): r[1].value=19; r[2].value="05_red_flags/red_flags_log.xlsx — 🔴5 / 🟡8 / 🟢6"
s.append(["แถว UCDP ทั่วโลกที่สแกน (ไม่นับเป็นหน้า)",417968,"ใช้จริง 2,909 แถวประเทศไทย / 202 เหตุการณ์ในพื้นที่ — ไม่นับรวมในยอดหน้า เพื่อไม่ให้ตัวเลขเกินจริง"])
i=wb["inventory"]
i.append(["08_security","UCDP_GED_v26.1 (zip 1 CSV) + codebook","UCDP Uppsala","CSV+PDF",52,None,2909,"Active","2026-10-10"])
i.append(["08_security","*_DOPA_Unrest_* (6 CSV)","ที่ทำการปกครองจังหวัด","CSV",None,None,474,"Active","2026-10-10"])
i.append(["02_clean","Security_4prov_2562-2568_clean.xlsx","Built (Claude)","Excel (clean)",None,7,None,"Output","2026-10-10"])
i.append(["02_clean","GPP_4prov_2019-2024_clean.xlsx","Built (Claude)","Excel (clean)",None,4,None,"Output","2026-10-10"])
wb["milestones"].append(["2026-10-10","Security incidents (UCDP + official) cleaned; GPP 4-province clean; tourism monthly 2568 cleaned",8,52,3383])
wb.save(f"{R}/00_admin/project_stats.xlsx")
wb=openpyxl.load_workbook(f"{R}/00_admin/data_checklist.xlsx"); c=wb["checklist"]
for r in c.iter_rows(min_row=2):
    if r[0].value==3: r[5].value="UCDP GED (Uppsala) + ที่ทำการปกครองจังหวัด (DSW public data ends Oct 2564)"; r[7].value="Done"; r[7].fill=PatternFill("solid",fgColor="C6E0B4"); r[8].value="02_clean/Security_4prov_2562-2568_clean.xlsx"; c.cell(r[0].row,10).value="UCDP = lethal only; official incl. non-lethal; Pattani official missing; RF19."
    if r[0].value==1: r[7].value="Done"; r[7].fill=PatternFill("solid",fgColor="C6E0B4"); r[8].value="02_clean/GPP_4prov_2019-2024_clean.xlsx"
    if r[0].value==8: c.cell(r[0].row,10).value="Monthly 2568 cleaned (seasonality)."
    if r[0].value==11: r[7].value="Links ready"; c.cell(r[0].row,10).value="Annual 2564–2568 PDFs + BOT regional stats (customs by checkpoint)."
    if r[0].value==12: c.cell(r[0].row,10).value="Budget Bureau provincial allocation (5 budget types) FY2564–2569 links ready. SBPAC opendata server down."
wb.save(f"{R}/00_admin/data_checklist.xlsx")
print("ok")
