import os, openpyxl
from openpyxl.styles import Alignment, PatternFill
R=os.path.expanduser("~/mnt/SBP_Dashboard")
wb=openpyxl.load_workbook(f"{R}/05_red_flags/red_flags_log.xlsx"); ws=wb["red_flags"]
for r in ws.iter_rows(min_row=2):
    if r[0].value=="RF07":
        r[7].value+=(" UPDATE 2026-10-10 (budget data): capital budget allocated FY2565→FY2566: Yala 4,946→4,005 MB (−19%), Narathiwat 3,924→4,033 MB (+3%). "
                     "FY2567 budget enacted late (2024) → Oct–Dec 2023 capital spending squeezed.")
        r[12].value="Yala: largely explained (capital budget cut −19% + FY2567 delay). Narathiwat: budget flat but construction −21% → needs disbursement data for FY2566 (only FY2568 published)."
        r[13].value="Partly explained"
new=[
["RF20","2026-10-10","🟡 Medium","D. Money vs outcome (watch)","Budget Bureau provincial allocation","Pattani","FY2569",
 "Pattani total allocation jumps 9,577 → 14,083 MB (+47%). Almost all from Office of the Prime Minister (incl. ISOC/กอ.รมน.): 1,522 → 5,995 MB, booked as 'other expenditure'. Capital share falls to 29.7% (lowest of all provinces/years).",
 "A security-agency line quadruples in one year and dominates the province's budget.","Re-booking of ISOC Region 4 personnel/operations to Pattani (HQ location), not new spending in Pattani's economy.",
 "Security budget growing while development (capital) share shrinks.",
 "Read the FY2569 Pattani PDF line items under สำนักนายกรัฐมนตรี (which programme); compare with FY2569 budget law for ISOC.",
 "Report 'development budget' (capital + subsidies) separately from security lines.","Open","'Pattani's 2026 budget jumps 47%, mostly a security-agency line — not necessarily money for the local economy.'"],
["RF21","2026-10-10","🟡 Medium","D. Money vs outcome","Comptroller General (Narathiwat Treasury)","Narathiwat","FY2568",
 "By end of FY2568 (Sep 2025): capital budget disbursed 3,000 of 4,606 MB = 65.1%, while recurrent budget 99.2%.",
 "One-third of investment money for the province not spent within the fiscal year.","Normal carry-over (เหลื่อมปี) of construction contracts; flood/security delays.",
 "Projects stalled or not delivered.","Check district/agency rows with lowest disbursement in sheet NWT_disbursement_FY2568; follow carry-over in FY2569 file.",
 "Use as indicator of implementation capacity (P2).","Open","'Only 65% of Narathiwat's 2025 investment budget was spent by year-end.'"]]
for row in new:
    ws.append(row)
    for c in ws[ws.max_row]: c.alignment=Alignment(wrap_text=True,vertical="top")
    ws.cell(ws.max_row,3).fill=PatternFill("solid",fgColor="FFE699")
wb.save(f"{R}/05_red_flags/red_flags_log.xlsx")
wb=openpyxl.load_workbook(f"{R}/00_admin/source_log.xlsx"); s=wb["source_log"]
s.append(["BUD01","BB_ProvBudgetAlloc_5types_{PROV}_FY{year}_dl20261010.pdf (23 PDF)","12_sbpac_budget","สำนักงบประมาณ (Budget Bureau)","bb.go.th/web/budget/province/province_bud{yy}/PDF/5bud/...","2026-10-10","FY2564–2569","Province","PTN,YLA,NWT,SKA (Yala FY2569 missing)","pdf (text)","P2","Pass","Allocation (plan), not disbursement. Totals = sum of 5 types (checked)."])
s.append(["BUD02","CGD_NWT_Disbursement_byUnit_FY2568_dl20261010.csv","12_sbpac_budget","กรมบัญชีกลาง / สำนักงานคลังจังหวัดนราธิวาส","narathiwat.gdcatalog.go.th (dataset ผลการเบิกจ่ายเงินประจำปีงบประมาณ)","2026-10-10","FY2568 monthly","District × unit × project","NWT","csv","P2","Pass","Only Narathiwat publishes. Sep 2568 = year-end cumulative."])
s.append(["BOT01","BOT_SRO_Annual{Press|Table}_{year}_dl20261010.pdf (8) + 3 BOT_RG_SR_*.xlsx","11_reports_bot","ธนาคารแห่งประเทศไทย สำนักงานภาคใต้","bot.or.th (see _DOWNLOAD_LINKS.txt)","2026-10-10","2564–2568","Region (14 southern provinces)","South","pdf + xlsx","All (context)","Pass","Customs table is by product for whole South, not by checkpoint."])
s.append(["BUD-C","02_clean/Budget_4prov_FY2564-2569_clean.xlsx","02_clean","Built from BUD01–02",None,"2026-10-10","FY2564–2569","Province","PTN,YLA,NWT,SKA","xlsx","P2","Pass",None])
s.append(["BOT-C","02_clean/BOT_South_indicators_2562-2568_clean.xlsx","02_clean","Built from BOT01",None,"2026-10-10","2562–2568","Region","South","xlsx","All","Pass",None])
wb.save(f"{R}/00_admin/source_log.xlsx")
wb=openpyxl.load_workbook(f"{R}/00_admin/project_stats.xlsx"); s=wb["summary"]
def setv(r,v,n=None):
    s.cell(r,2).value=v
    if n is not None: s.cell(r,3).value=n
setv(2,223,"188 + 35 (BOT 11 + งบประมาณ 24)")
setv(3,159,"+3 BOT"); setv(4,960,None); setv(5,106455,"+264 BOT")
setv(6,64,"+1 การเบิกจ่ายนราธิวาส"); setv(7,54787,"+11,892 การเบิกจ่ายนราธิวาส"); setv(8,846644,"+214,056")
setv(10,3500,"… + 532 (BOT 8 PDF + งบประมาณ 23 PDF)"); setv(11,2949,None)
setv(13,3229,"ประมาณ 50 แถว/หน้า — ค่าประมาณ"); setv(14,6729,"3,500 PDF + 3,229 ประมาณการ")
setv(16,12,"… + ความไม่สงบ, GPP, งบประมาณ, BOT ภาคใต้")
setv(17,18,s.cell(17,3).value+", สำนักงบประมาณ, ธปท. สำนักงานภาคใต้, กรมบัญชีกลาง/คลังจังหวัด")
setv(18,"2026-10-10 (Budget + BOT — data collection phase 1 complete)")
for r in s.iter_rows(min_row=2):
    if r[0].value and str(r[0].value).startswith("Red flags"): r[1].value=21; r[2].value="05_red_flags/red_flags_log.xlsx — 🔴5 / 🟡10 / 🟢6"
i=wb["inventory"]
i.append(["11_reports_bot","BOT_SRO_* (8 PDF) + BOT_RG_SR_* (3 Excel)","ธปท. สภต.","PDF+Excel",None,9,264,"Active","2026-10-10"])
i.append(["12_sbpac_budget","BB_ProvBudgetAlloc_* (23 PDF) + CGD_NWT_* (1 CSV)","สำนักงบประมาณ / กรมบัญชีกลาง","PDF+CSV",None,None,11892,"Active","2026-10-10"])
i.append(["02_clean","Budget_4prov_FY2564-2569_clean.xlsx","Built (Claude)","Excel (clean)",None,3,None,"Output","2026-10-10"])
i.append(["02_clean","BOT_South_indicators_2562-2568_clean.xlsx","Built (Claude)","Excel (clean)",None,4,None,"Output","2026-10-10"])
wb["milestones"].append(["2026-10-10","Budget allocation FY2564–2569 + BOT southern reports cleaned → data collection phase 1 complete (12 clean datasets)",35,532,12156])
wb.save(f"{R}/00_admin/project_stats.xlsx")
wb=openpyxl.load_workbook(f"{R}/00_admin/data_checklist.xlsx"); c=wb["checklist"]
for r in c.iter_rows(min_row=2):
    if r[0].value in (11,12):
        r[7].value="Done"; r[7].fill=PatternFill("solid",fgColor="C6E0B4")
        r[8].value="02_clean/BOT_South_indicators_2562-2568_clean.xlsx" if r[0].value==11 else "02_clean/Budget_4prov_FY2564-2569_clean.xlsx"
        if r[0].value==12: c.cell(r[0].row,10).value="Budget Bureau allocation FY2564–2569 (Yala FY69 missing) + NWT disbursement FY68. SBPAC opendata unreachable from UAE."
wb.save(f"{R}/00_admin/data_checklist.xlsx")
print("ok")
