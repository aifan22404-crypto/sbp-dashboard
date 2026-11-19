import os, openpyxl
from openpyxl.styles import Alignment, PatternFill
A=os.path.expanduser("~/mnt/SBP_Dashboard/00_admin")
wb=openpyxl.load_workbook(f"{A}/data_checklist.xlsx"); ws=wb["checklist"]
ws.cell(1,10).value="Notes (updated 2026-10-09)"
upd={
1:("GPP 4 จังหวัด 2021-2025 แยกรายสาขา, ราคาประจำปี + CVM, per capita","สศช. nesdc.go.th → บัญชีประชาชาติ → GPP","Collected — clean pending","01_raw/01_economy_gpp/NESDC_GPP_edition2024_allprov_1995-2024_dl20261008.xlsx","Edition 2024 = latest (covers to 2024p). GPP 2025 not released yet. Still need a 4-province clean file."),
2:(None,"gdcatalog.go.th → openapi.dbd.go.th (โหลดผ่าน VPN มือถือ)","Done (no TSIC split)","02_clean/DBD_juristic_4prov_2564-2568_clean.xlsx","By business type (TSIC) still missing. See red flag RF02 (Apr-2566 spike)."),
3:(None,None,"Partial — codebooks only","01_raw/08_security/ (3 codebooks)","Incident data needs email request to DSW."),
4:("นำเข้า/ส่งออก รายด่าน รายปี 2021-2025: สุไหงโก-ลก, ตากใบ, บูเก๊ะตา, เบตง, บ้านประกอบ (+สะเดา/ปาดังฯ เป็นตัวเทียบ)","สนง.พาณิชย์จังหวัด (narathiwat.gdcatalog, yala.moc.go.th, songkhla.moc.go.th) — customs/dft blocked","Done","02_clean/BorderTrade_3prov_2563-2568_clean.xlsx","Narathiwat = total only (no per-checkpoint). Top-10 goods not yet."),
5:(None,"nso.go.th (LFS provincial tables)","Partial — LFS T1 done","02_clean/NSO_LFS_T1_4prov_2564Q1-2569Q2_clean.xlsx","Employment by industry, informal labour, district-level not yet. See RF01 (weight break 2565)."),
6:(None,"catalog.oae.go.th + agriapi.nabc.go.th","Done","02_clean/Agriculture_4prov_2563-2568_clean.xlsx","Fruit 2568 preliminary → use to 2567 (RF10). Prices national, not provincial."),
7:(None,"องค์การสะพานปลา new.fishmarket.co.th (Plan B) — กรมประมงถูกบล็อก","Partial — FMO ports only","02_clean/Fisheries_FMO_ports_FY2565-2568_clean.xlsx","DoF data: Irfan getting it himself. FY2564 & FY2566 H2 missing. Fiscal years."),
8:("นักท่องเที่ยวไทย vs ต่างชาติ (มาเลเซีย), รายได้, อัตราเข้าพัก — ระวังฐานต่ำปี 2021",None,"Done","02_clean/Tourism_4prov_2564-2568_clean.xlsx","2568 monthly file not yet cleaned (seasonality)."),
9:(None,"hatyai.boi.go.th (BOI ภาคที่ 5) — catalog.boi.go.th blocked","Done","02_clean/BOI_4prov_2564-2568_clean.xlsx","Approvals only (not realised). +2569 H1 bonus. See RF03/RF04."),
10:(None,None,"In progress",None,"Started 2026-10-09."),
11:("รายงานเศรษฐกิจภาคใต้รายไตรมาส 2021-2025",None,"Not started",None,None),
12:(None,None,"Partial","01_raw/12_sbpac_budget/DIPROM_IndustryPlan_SouthBorderCluster_2566-2570_rev2567_dl20261008.pdf","Budget/disbursement data still needed (for RF07)."),
}
for r in ws.iter_rows(min_row=2):
    n=r[0].value
    if n in upd:
        what,where,st,fn,note=upd[n]
        if what: r[4].value=what
        if where: r[5].value=where
        r[7].value=st
        if fn: r[8].value=fn
        ws.cell(r[0].row,10).value=note
ws.append([13,"T3","05_red_flags","จัดซื้อจัดจ้างภาครัฐ","สัญญา/ผู้ชนะประมูลรายจังหวัด — ใช้ตรวจ red flags (เช่น RF02)","ACT Ai (actai.co), e-GP gprocurement.go.th","P2, Red flags","Not started",None,"Optional — only for checking red flags."])
col={"Done":"C6E0B4","Partial":"FFE699","Collected":"FFE699","In progress":"BDD7EE","Not started":"F2F2F2"}
for r in ws.iter_rows(min_row=2):
    for k,v in col.items():
        if str(r[7].value).startswith(k): r[7].fill=PatternFill("solid",fgColor=v)
    for c in r: c.alignment=Alignment(wrap_text=True,vertical="top")
ws.column_dimensions["J"].width=50; ws.column_dimensions["I"].width=45
wb.save(f"{A}/data_checklist.xlsx")

wb=openpyxl.load_workbook(f"{A}/data_dictionary.xlsx"); ws=wb["data_dictionary"]
fix={"gpp_current":"2021-2024 (2025 not released)","gpp_cvm":"2021-2024 (2025 not released)","gpp_growth_yoy":"2021-2024","gpp_per_capita":"2021-2024 (2025 not released)",
"biz_new":"2021-2025","biz_closed":"2021-2025","incidents":"2021-2025","border_trade_value":"2021-2025","unemployment_rate":"2021-2025 (quarterly to 2026Q2)"}
for r in ws.iter_rows(min_row=2):
    if r[0].value in fix: r[5].value=fix[r[0].value]
    if r[0].value=="border_trade_value": r[3].value="Provincial Commerce Offices (MOC)"; r[4].value="Province / Checkpoint"
new=[
("biz_new_capital","ทุนจดทะเบียนของนิติบุคคลตั้งใหม่","ล้านบาท/ปี","DBD","Province","2021-2025","P1,P4"),
("biz_net_formation","ตั้งใหม่ − เลิกกิจการ","ราย/ปี","Calculated (DBD)","Province","2021-2025","P1,P4"),
("employed","ผู้มีงานทำ (ระวัง RF01: น้ำหนักเปลี่ยนปี 2565)","คน","NSO LFS","Province","2021-2025","P2"),
("lfp_rate","อัตราการมีส่วนร่วมในกำลังแรงงาน","%","NSO LFS","Province","2021-2025","P2"),
("boi_projects","โครงการที่ได้รับอนุมัติส่งเสริม BOI","โครงการ/ปี","BOI ภาค 5","Province","2021-2025 (+2026H1)","P1,P2"),
("boi_investment","เงินลงทุนที่ได้รับอนุมัติ (ไม่ใช่เงินที่ใช้จริง)","ล้านบาท/ปี","BOI ภาค 5","Province","2021-2025 (+2026H1)","P1,P2"),
("boi_jobs","การจ้างงานของโครงการที่อนุมัติ","คน/ปี","BOI ภาค 5","Province","2021-2025 (+2026H1)","P1,P2"),
("crop_output","ผลผลิตพืชหลัก (ยาง ปาล์ม ทุเรียน มังคุด เงาะ ลองกอง มะพร้าว)","ตัน/ปี","OAE","Province","2021-2025 (fruit ใช้ถึง 2024)","P1,P3,P4"),
("farm_price","ราคาที่เกษตรกรขายได้ (ยางแผ่นดิบ RSS3, ปาล์มทะลาย)","บาท/กก.","OAE (national)","National","2021-2025 (monthly)","P3,P4"),
("fish_landing_volume","ปริมาณสัตว์น้ำผ่านท่าองค์การสะพานปลา","ตัน/ปีงบ","FMO","Port","FY2022-2025 (FY2023 partial)","P1,P4"),
("fish_landing_value","มูลค่าสัตว์น้ำผ่านท่า","ล้านบาท/ปีงบ","FMO","Port","FY2022-2025 (FY2023 partial)","P1,P4"),
("tourist_visitors","ผู้เยี่ยมเยือน (ไทย+ต่างชาติ)","คน/ปี","MOTS","Province","2021-2025","P1,P3,P4"),
("tourism_revenue","รายได้จากการท่องเที่ยว","ล้านบาท/ปี","MOTS","Province","2021-2025","P1,P3,P4"),
("hotel_occupancy","อัตราการเข้าพักโรงแรม","%","MOTS","Province","2021-2025","P1"),
("malaysian_hotel_guests","ผู้เข้าพักสัญชาติมาเลเซีย","คน/ปี","MOTS","Province","2021-2025","P1,P3"),
]
have={r[0].value for r in ws.iter_rows(min_row=2)}
for n in new:
    if n[0] not in have: ws.append(list(n))
wb.save(f"{A}/data_dictionary.xlsx")
print("ok", ws.max_row)
