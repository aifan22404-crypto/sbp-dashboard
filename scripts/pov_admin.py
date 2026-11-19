import os, openpyxl
from openpyxl.styles import Alignment, PatternFill
R=os.path.expanduser("~/mnt/SBP_Dashboard")
# red flags
wb=openpyxl.load_workbook(f"{R}/05_red_flags/red_flags_log.xlsx"); ws=wb["red_flags"]
new=[
["RF17","2026-10-10","🔴 High","A. Definition clash","NESDC poverty vs TPMAP","Narathiwat, Pattani, Yala","2566–2568",
 "Narathiwat poor: NESDC 18.38% (2566) vs TPMAP 0.31% (TPMAP2566) — ~60× apart. Ranks flip: TPMAP2568 puts Yala #2 & Pattani #27; NESDC 2567 puts Yala #2, Pattani #3.",
 "Two official 'poverty' numbers for the same place disagree massively.",
 "Different things: NESDC = consumption below poverty line (sample survey); TPMAP = จปฐ. households failing basic needs AND on welfare-card list (targeting list).",
 "TPMAP is also a performance target for officials ('eradicate poverty') → pressure to report fewer poor.",
 "Compare TPMAP poor counts with welfare-card holders per province (Ministry of Finance) and with NESDC.",
 "NESDC = poverty LEVEL & trend. TPMAP = WHERE (district pockets) and WHICH dimension, within one year only. Never put both on one chart as 'poverty rate'.",
 "Explained (definitions) — keep as caution","'Two official poverty measures count different things.'"],
["RF18","2026-10-10","🔴 High","C. Outlier / possible reporting effect","TPMAP (จปฐ.)","All 4, esp. Songkhla 4 districts","TPMAP 2564–2569",
 "National target poor 1,025,782 (2565) → 211,739 (2566, −79%) → 813,054 (2567). District swings: Saba Yoi 31 → 2,378 → 60 → 138 → 10 → 47; Thepha 12 (2566) → 6,223 (2568); Na Thawi 6 people in 2568. Persons surveyed stay ~50,000 every year.",
 "The population surveyed is stable but the number of poor jumps 100× up and down.",
 "TPMAP 2566 used a different welfare-card base / criteria; จปฐ. collection quality varies by district volunteers.",
 "Numbers entered to hit 'zero poor' targets, or inflated to attract programmes — จปฐ. is self-reported to local officials.",
 "Ask CDD/TPMAP for method notes per year; compare district results with SES/NESDC and with welfare-card counts; field check Thepha (11.8% vs neighbours <1%).",
 "No year-to-year TPMAP trends. Use latest year (TPMAP 2569) for district map, flagged as administrative data.",
 "Open","'Local poverty lists jump wildly year to year — we only use them to see where, not how fast.'"],
]
for row in new:
    ws.append(row)
    for c in ws[ws.max_row]: c.alignment=Alignment(wrap_text=True,vertical="top")
    ws.cell(ws.max_row,3).fill=PatternFill("solid",fgColor="F8CBAD")
for r in ws.iter_rows(min_row=2):
    if r[0].value=="RF16": r[7].value+=" UPDATE 2026-10-10: NESDC poverty Yala 9.04% (2566) → 25.54% (2567) → 18.05% (2568) — same noise."
wb.save(f"{R}/05_red_flags/red_flags_log.xlsx")
# source log
wb=openpyxl.load_workbook(f"{R}/00_admin/source_log.xlsx"); s=wb["source_log"]
s.append(["PV01","TPMAP_{poor|common}_allprov_TPMAP<year>  + TPMAP_poor_amphur_<PROV>_TPMAP<year> (36 CSV) + datadict","10_poverty_income","TPMAP (NECTEC/NSTDA + สศช.), data = จปฐ. (CDD) × welfare card (MOF)","tpmap.in.th/public/data/<yy>/…","2026-10-10","TPMAP 2564–2569 (จปฐ. 2563–2568)","Province + District","All 77 prov; districts of SKA/PTN/YLA/NWT","csv","P2,P4","Pass with caution","Targeting list, not poverty rate. Big year-to-year breaks (RF18)."])
s.append(["PV02","NESDC_PovertyInequality_report_2568_dl20261010.pdf","10_poverty_income","สศช. (NESDC)","nesdc.go.th/wp-content/uploads/2026/09/poverty-report-2568-FY.pdf","2026-10-10","2551–2568","Province (Appendix ก.6)","All 77","pdf (text)","P2,P3,P4","Pass","Official poverty rate by province 2551–2568, 139 pages."])
s.append(["PV03","NSO_IncomeDistribution_byProvince_report_2566_dl20261010.pdf","10_poverty_income","สสช. (NSO)","catalogapi.nso.go.th/api/doc/department/D10/SD10_04/SD10_04_252_1.pdf","2026-10-10","2566","Province","All 77","pdf (text)","P2","Pass","SES 2566 provincial income/expenditure/debt, 252 pages. 2564 edition: NSO server error (HTTP 500)."])
s.append(["PV04","NSO_SES_Gini_byProvince_2556-2566_dl20261009.csv","10_poverty_income","สสช. (NSO)","catalogapi.nso.go.th/api/index?table=SFD_01_Table_0011_01&format=csv&div=1","2026-10-09","2556–2566","Province","All (NWT missing 2560+)","csv","P2","Pass with caution","Small samples → volatile (RF16)."])
s.append(["PV-C","02_clean/Poverty_4prov_2564-2568_clean.xlsx","02_clean","Built from PV01–PV04",None,"2026-10-10","2021–2025","Province + District","PTN,YLA,NWT,SKA (+4 districts)","xlsx","P2,P4","Pass","SES income sheet pending (waiting for NSO jj/jA/jm + SFD_SPB0801)."])
wb.save(f"{R}/00_admin/source_log.xlsx")
# stats
wb=openpyxl.load_workbook(f"{R}/00_admin/project_stats.xlsx"); s=wb["summary"]
def setv(r,v,n=None):
    s.cell(r,2).value=v
    if n is not None: s.cell(r,3).value=n
setv(2,174,"134 + 40 ความยากจน (36 CSV TPMAP + 1 Gini + 1 datadict + 2 PDF)")
setv(3,153,"+1 TPMAP datadict"); setv(4,945,None); setv(5,105422,"+55 TPMAP datadict")
setv(6,53,"15 DBD + 1 ชายแดน + 37 ความยากจน"); setv(7,1936,"195 DBD + 73 ชายแดน + 1,668 ความยากจน"); setv(8,85751,"+71,531 ความยากจน")
setv(10,2916,"1,556 + 940 ประมง + 29 BOI + 391 ความยากจน"); setv(11,2365,None)
setv(13,2151,"ประมาณ 50 แถว/หน้า (+34 จากความยากจน)"); setv(14,5067,"2,916 PDF + 2,151 ประมาณการ — ทะลุ 5,000 แล้ว")
setv(16,8,"แรงงาน + DBD + การค้าชายแดน + เกษตร + ประมง + ท่องเที่ยว + BOI + ความยากจน")
setv(17,13,s.cell(17,3).value+", TPMAP (NECTEC/สวทช.)")
setv(18,"2026-10-10 (Poverty)")
for r in s.iter_rows(min_row=2):
    if r[0].value and str(r[0].value).startswith("Red flags"): r[1].value=18; r[2].value="05_red_flags/red_flags_log.xlsx — 🔴5 / 🟡7 / 🟢6"
i=wb["inventory"]
i.append(["10_poverty_income","TPMAP_* (36 CSV + datadict)","TPMAP / NECTEC","CSV",None,2,1723,"Active","2026-10-10"])
i.append(["10_poverty_income","NESDC_PovertyInequality_report_2568","สศช.","PDF",139,None,None,"Active","2026-10-10"])
i.append(["10_poverty_income","NSO_IncomeDistribution_byProvince_report_2566","สสช.","PDF",252,None,None,"Active","2026-10-10"])
i.append(["10_poverty_income","NSO_SES_Gini_byProvince_2556-2566","สสช.","CSV",None,None,462,"Active","2026-10-10"])
i.append(["02_clean","Poverty_4prov_2564-2568_clean.xlsx","Built (Claude)","Excel (clean)",None,8,None,"Output","2026-10-10"])
wb["milestones"].append(["2026-10-10","Poverty: NESDC provincial poverty rate 2551–2568 + TPMAP district poor 2564–2569 cleaned; total pages passed 5,000",40,391,None])
wb.save(f"{R}/00_admin/project_stats.xlsx")
# checklist
wb=openpyxl.load_workbook(f"{R}/00_admin/data_checklist.xlsx"); c=wb["checklist"]
for r in c.iter_rows(min_row=2):
    if r[0].value==10:
        r[5].value="สศช. (poverty report ก.6), TPMAP, สสช. SES/Gini"; r[7].value="Partial — poverty done, SES income pending"; r[7].fill=PatternFill("solid",fgColor="FFE699")
        r[8].value="02_clean/Poverty_4prov_2564-2568_clean.xlsx"; c.cell(r[0].row,10).value="Need NSO jj/jA/jm + SFD_SPB0801 for household income/debt. See RF16–18."
wb.save(f"{R}/00_admin/data_checklist.xlsx")
print("ok")
