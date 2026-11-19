import os, openpyxl
from openpyxl.styles import PatternFill
R=os.path.expanduser("~/mnt/SBP_Dashboard")
wb=openpyxl.load_workbook(f"{R}/00_admin/source_log.xlsx"); s=wb["source_log"]
s.append(["PV05","NSO_SES_{Income|Expenditure|DebtByPurpose}_byProvince_*_dl20261010.xlsx (3 files)","10_poverty_income","สสช. (NSO) — statistics_and_indicators, branch 309","nso.go.th/nsoweb/downloadFile/stat_impt/{jj|jA|jm}/file_xls_th","2026-10-10","2547–2566 (exp 2555–2566)","Province","All 77 + regions","xlsx","P2,P4","Pass","Replaces D1 (NSO catalog upload broken, HTTP 500). 2566 values identical to PV06."])
s.append(["PV06","NSO_SES_{IncExpDebt|IncomeBySource|DebtByPurpose}_byProvince_2566-2568_dl20261010.csv (3 files)","10_poverty_income","สสช. (NSO) catalogapi","catalogapi.nso.go.th/api/index?table={SFD_SPB0801|SFD_SPB0802_66|SFD_SPB0806}&format=csv","2026-10-10","2566, 2568","Province × socio-economic class","All 77","csv","P2,P3,P4","Pass with caution","Source/purpose tables only by class → province totals household-weighted (reproduce official totals ±1 baht). Province samples small (RF16)."])
for r in s.iter_rows(min_row=2):
    if r[0].value=="PV-C": r[12].value="Sheets: SES summary 2564/2565/2566/2568, long series 2547–2566, income by source, debt by purpose & informal debt, class mix, NESDC poverty, TPMAP province/district, Gini."
wb.save(f"{R}/00_admin/source_log.xlsx")
wb=openpyxl.load_workbook(f"{R}/00_admin/project_stats.xlsx"); s=wb["summary"]
def setv(r,v,n=None):
    s.cell(r,2).value=v
    if n is not None: s.cell(r,3).value=n
setv(2,180,"134 + 46 ความยากจน/รายได้ (39 CSV + 4 Excel + 2 PDF + 1 Gini) — ไม่นับ 2 ไฟล์ซ้ำใน _duplicates")
setv(3,156,"1 GPP + 137 LFS + 1 ชายแดน + 7 เกษตร + 6 ท่องเที่ยว + 4 ความยากจน"); setv(4,951,None); setv(5,106191,"+824 ความยากจน (datadict 55 + ตาราง สสช. 769)")
setv(6,56,"15 DBD + 1 ชายแดน + 40 ความยากจน"); setv(7,39512,"195 DBD + 73 ชายแดน + 39,244 ความยากจน"); setv(8,485535,"+471,315 ความยากจน")
setv(13,2918,"ประมาณ 50 แถว/หน้า — เป็นค่าประมาณ"); setv(14,5834,"2,916 PDF + 2,918 ประมาณการ")
setv(18,"2026-10-10 (Poverty + SES income)")
i=wb["inventory"]
i.append(["10_poverty_income","NSO_SES_* (3 Excel, 2547–2566)","สสช.","Excel",None,6,769,"Active","2026-10-10"])
i.append(["10_poverty_income","NSO_SES_* (3 CSV, 2566 & 2568)","สสช.","CSV",None,None,37576,"Active","2026-10-10"])
i.append(["10_poverty_income/_duplicates","data (3).csv, data (4).csv","สสช.","CSV (duplicate)",None,None,None,"Duplicate — not counted","2026-10-10"])
wb["milestones"].append(["2026-10-10","Household income / expenditure / debt (NSO SES 2564–2568) cleaned — poverty category complete",6,None,38345])
wb.save(f"{R}/00_admin/project_stats.xlsx")
wb=openpyxl.load_workbook(f"{R}/00_admin/data_checklist.xlsx"); c=wb["checklist"]
for r in c.iter_rows(min_row=2):
    if r[0].value==10:
        r[7].value="Done"; r[7].fill=PatternFill("solid",fgColor="C6E0B4"); c.cell(r[0].row,10).value="NESDC poverty 2551–2568, SES income/exp/debt 2564–2568, TPMAP district. See RF16–18."
wb.save(f"{R}/00_admin/data_checklist.xlsx")
wb=openpyxl.load_workbook(f"{R}/05_red_flags/red_flags_log.xlsx"); ws=wb["red_flags"]
for r in ws.iter_rows(min_row=2):
    if r[0].value=="RF16":
        r[7].value+=" UPDATE 2: Songkhla debt fall = house/land loans 109,463 → 19,161 and business loans 58,704 → 1,498 (a few big-loan households in the 2566 sample)."
        r[13].value="Partly explained (sample outliers) — keep caution"
wb.save(f"{R}/05_red_flags/red_flags_log.xlsx")
print("ok")
