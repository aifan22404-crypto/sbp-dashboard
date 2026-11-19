import os,json,csv,collections,openpyxl
from openpyxl.styles import Font,PatternFill,Alignment
R=os.path.expanduser("~/mnt/SBP_Dashboard"); O=f"{R}/03_analysis/00_master_panel"
P=json.load(open("/tmp/panel.json")); rows=P["rows"]; D=P["dict"]
COLS=["area","area_level","scope","year_BE","year_CE","period_type","domain","indicator_code","indicator_th","indicator_en","value","unit","source_id","clean_file","quality_flag","used_in"]
dup=collections.Counter((r[0],r[3],r[7]) for r in rows); dups=[k for k,v in dup.items() if v>1]
print("duplicates:",len(dups),dups[:5])
with open(f"{O}/SBP_master_panel_long.csv","w",encoding="utf-8-sig",newline="") as f:
    w=csv.writer(f); w.writerow(COLS); w.writerows(rows)
wb=openpyxl.Workbook(); H=Font(bold=True,color="FFFFFF"); FL=PatternFill("solid",fgColor="1F4E78")
def sheet(ws,c,d,widths=None):
    ws.append(c)
    for x in ws[1]: x.font=H; x.fill=FL; x.alignment=Alignment(wrap_text=True,vertical="center")
    for r in d: ws.append(r)
    for i,cc in enumerate(c): ws.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=(widths or {}).get(i,max(10,min(36,len(str(cc))+3)))
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
rd=wb.active; rd.title="README"
for l in ["SBP Master Panel — Southern Thailand Regional Economic Dashboard",
"One shared table for P1–P4. Every number traces back to a clean file (02_clean) and a source ID (00_admin/source_log.xlsx).",
f"Built 2026-10-16 from 14 clean datasets: {len(rows):,} values, {len(D)} indicators, years 2563–2569 (core scope 2564–2568).",
"",
"HOW TO USE",
"• panel_long = the master table (one row = one area × one year × one indicator). Filter by indicator_code, area, year_BE.",
"• indicator_dictionary = what each code means, unit, source, which Project uses it.",
"• coverage = which indicators have data in which years (number of areas).",
"• wide_core = the 32 most-used indicators as columns, one row per area-year (quick charts / Excel pivot).",
"• Same data as CSV: SBP_master_panel_long.csv (paste into any Claude Project or Looker/Power BI).",
"",
"RULES",
"• scope: 'in scope' = Pattani, Yala, Narathiwat, Songkhla-4districts (Chana, Thepha, Saba Yoi, Na Thawi) and their sum. 'context' = Songkhla whole province. 'benchmark' = South / Whole Kingdom.",
"• quality_flag ≠ 'ok' means read the flag first (red flag IDs → 05_red_flags/red_flags_log.xlsx). Never chart a flagged series without its note.",
"• period_type: calendar / fiscal (Oct–Sep) / survey (SES, NESDC) / จปฐ. survey year (TPMAP). Do not mix fiscal and calendar on one axis without a label.",
"• Years: year_BE = Thai Buddhist year; year_CE = year_BE − 543. GPP label '2024p' = BE 2567 preliminary.",
"• Songkhla whole-province values are context only; where 4-district data exists it is a separate area 'Songkhla-4districts' (DBD firms, UCDP, border trade Ban Prakob)."]: rd.append([l])
rd.column_dimensions["A"].width=140; rd["A1"].font=Font(bold=True,size=14)
sheet(wb.create_sheet("panel_long"),COLS,rows,{8:34,9:34,13:30,14:30})
# dictionary
by=collections.defaultdict(list)
for r in rows: by[r[7]].append(r)
drows=[]
for code,(dom,th,en,unit,src,cf,use) in sorted(D.items(),key=lambda x:(x[1][0],x[0])):
    rs=by.get(code,[])
    if not rs: continue
    drows.append([code,dom,th,en,unit,src,cf,use,len(rs),", ".join(sorted({r[0] for r in rs})),f"{min(r[3] for r in rs)}–{max(r[3] for r in rs)}",sum(1 for r in rs if r[14]!="ok")])
sheet(wb.create_sheet("indicator_dictionary"),["indicator_code","domain","indicator_th","indicator_en","unit","source_id","clean_file","used_in","n_values","areas","years_BE","n_flagged"],drows,{2:34,3:34,6:30,9:40})
# coverage
yrs=list(range(2563,2570)); cov=[]
for code in [d[0] for d in drows]:
    rs=by[code]; c=collections.Counter(r[3] for r in rs if r[2] in ("in scope",))
    cov.append([code,D[code][0]]+[c.get(y,0) or None for y in yrs])
sheet(wb.create_sheet("coverage"),["indicator_code","domain"]+[f"{y} (# in-scope areas)" for y in yrs],cov)
# wide core
CORE=["gpp_current_mb","gpp_real_growth_pct","gpp_per_capita_baht","gpp_share_agri_pct","gpp_share_pubadmin_pct","gpp_share_educ_pct","gpp_share_manuf_pct","gpp_share_accom_pct","unemployment_rate_pct","lfp_rate_pct","dbd_new_firms","dbd_dissolved","dbd_net_formation","boi_projects","boi_investment_mb","boi_jobs","tour_visitors","tour_revenue_mb","tour_occupancy_pct","border_trade_mb","fish_landings_t","fish_vessels_licensed","poverty_rate_pct","hh_income_month_baht","hh_debt_baht","hh_informal_debt_baht","sec_ucdp_lethal_events","sec_ucdp_deaths","sec_official_incidents","budget_total_mb","budget_capital_mb","budget_capital_share_pct"]
val={(r[0],r[3],r[7]):r[10] for r in rows}
areas=["Pattani","Yala","Narathiwat","Songkhla-4districts","3 provinces (PTN+YLA+NWT)","Songkhla","South region","Whole Kingdom"]
wide=[]
for a in areas:
    for y in range(2564,2570):
        vals=[val.get((a,y,c)) for c in CORE]
        if any(v is not None for v in vals): wide.append([a,y,y-543]+vals)
sheet(wb.create_sheet("wide_core"),["area","year_BE","year_CE"]+CORE,wide)
wb.save(f"{O}/SBP_master_panel_2564-2568.xlsx")
print(len(drows),"indicators with data;",len(wide),"wide rows; flagged values:",sum(1 for r in rows if r[14]!="ok"))
# show core snapshot 2568 for in-scope
for a in ["Pattani","Yala","Narathiwat"]:
    print(a,{c:val.get((a,2567,c)) for c in ["gpp_per_capita_baht","poverty_rate_pct","unemployment_rate_pct","sec_ucdp_deaths","budget_capital_share_pct"]})
