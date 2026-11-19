import os, openpyxl
R=os.path.expanduser("~/mnt/SBP_Dashboard/00_admin")
wb=openpyxl.load_workbook(f"{R}/source_log.xlsx"); ws=wb["source_log"]
ws.append(["BOI01","BOI_Region5_LowerSouth7_InvestmentSummary_<year>_<period>_dl20261009.pdf (6 files: 2564–2568 Jan-Dec + 2569 Jan-Jun)","03_investment_boi","BOI ศูนย์เศรษฐกิจการลงทุนภาคที่ 5 (สงขลา)","hatyai.boi.go.th/uploads/investment_file_download/... (see _DOWNLOAD_LINKS.txt)","2026-10-09","2564–2568 (+2569 H1)","Province","7 lower-south provinces incl. PTN,YLA,NWT,SKA","pdf (text)","P1,P2,P4","Pass","Approved projects (not realised investment). Songkhla = whole province. catalog.boi.go.th blocked from UAE; regional office site works."])
ws.append(["BOI-C","02_clean/BOI_4prov_2564-2568_clean.xlsx","02_clean","Built from BOI01",None,"2026-10-09","2021–2025 (+2026 H1)","Province","PTN,YLA,NWT,SKA","xlsx","P1,P2,P4","Pass","Sheets: province_approvals (24 rows, cross-checked), business_types, sbp_special_measures (overlap — do not sum), notes."])
wb.save(f"{R}/source_log.xlsx")
wb=openpyxl.load_workbook(f"{R}/project_stats.xlsx"); s=wb["summary"]
def setv(row,val,note=None):
    s.cell(row,2).value=val
    if note is not None: s.cell(row,3).value=note
setv(2,134,"128 + 6 PDF BOI")
setv(10,2525,"1,556 + 940 ประมง + 29 BOI")
setv(11,1974,None)
setv(14,4642,"2,525 PDF + 2,117 ประมาณการ — มีส่วนที่เป็นประมาณการ")
setv(16,7,"แรงงาน + DBD + การค้าชายแดน + เกษตร + ประมง + ท่องเที่ยว + BOI")
setv(17,12,"สศช., สสช., DSW, สนง.อุตสาหกรรมกลุ่มจังหวัดฯ, DBD, พาณิชย์จังหวัด×3, สศก., องค์การสะพานปลา, กระทรวงการท่องเที่ยวและกีฬา, BOI ภาคที่ 5 สงขลา")
setv(18,"2026-10-09 (BOI)")
i=wb["inventory"]
i.append(["03_investment_boi","BOI_Region5_LowerSouth7_InvestmentSummary_* (6 PDF)","BOI ศูนย์เศรษฐกิจการลงทุนภาคที่ 5","PDF",29,None,None,"Active","2026-10-09"])
i.append(["02_clean","BOI_4prov_2564-2568_clean.xlsx","Built (Claude)","Excel (clean)",None,4,None,"Output","2026-10-09"])
m=wb["milestones"]
m.append(["2026-10-09","Investment (BOI Region 5, 2564–2568 + 2569 H1) collected & cleaned — approvals by province, business types, SBP measures",6,29,None])
wb.save(f"{R}/project_stats.xlsx")
for r in s.iter_rows(values_only=True): print(r)
