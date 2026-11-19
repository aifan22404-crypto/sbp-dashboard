import os,openpyxl,csv,collections,json,re
from openpyxl.styles import Font,PatternFill,Alignment
R=os.path.expanduser("~/mnt/SBP_Dashboard"); C=f"{R}/02_clean"
TH2EN={"ปัตตานี":"Pattani","ยะลา":"Yala","นราธิวาส":"Narathiwat","สงขลา":"Songkhla"}
LEVEL={"Pattani":"province","Yala":"province","Narathiwat":"province","Songkhla":"province (context: whole)","Songkhla-4districts":"district group (in scope)","3 provinces (PTN+YLA+NWT)":"province group","South region":"region","Whole Kingdom":"national"}
SCOPE={"Pattani":"in scope","Yala":"in scope","Narathiwat":"in scope","Songkhla":"context (whole province)","Songkhla-4districts":"in scope","3 provinces (PTN+YLA+NWT)":"in scope","South region":"benchmark","Whole Kingdom":"benchmark"}
rows=[]; DICT={}
def wb(f): return openpyxl.load_workbook(f"{C}/{f}",read_only=True,data_only=True)
def sh(f,s):
    L=list(wb(f)[s].iter_rows(values_only=True)); h=L[0]; return [dict(zip(h,r)) for r in L[1:] if any(x is not None for x in r)]
def add(area,yBE,code,val,period="calendar",flag="ok"):
    if val is None or val=="" : return
    try: val=float(val)
    except: return
    d=DICT[code]; area=TH2EN.get(area,area)
    rows.append([area,LEVEL.get(area,"other"),SCOPE.get(area,"?"),int(yBE),int(yBE)-543,period,d[0],code,d[1],d[2],round(val,4),d[3],d[4],d[5],flag,d[6]])
def ind(code,domain,th,en,unit,src,cleanf,use): DICT[code]=(domain,th,en,unit,src,cleanf,use)
# ---------- ECONOMY (GPP) ----------
ind("gpp_current_mb","Economy","ผลิตภัณฑ์มวลรวมจังหวัด (ราคาประจำปี)","GPP, current prices","million baht","GPP01","GPP_4prov_2019-2024_clean.xlsx","P1,P2,P4")
ind("gpp_real_growth_pct","Economy","อัตราการเติบโตที่แท้จริงของ GPP","GPP real growth (CVM)","%","GPP01","GPP_4prov_2019-2024_clean.xlsx","P1,P2,P3")
ind("gpp_per_capita_baht","Economy","GPP ต่อหัว","GPP per capita","baht/person/year","GPP01","GPP_4prov_2019-2024_clean.xlsx","P2")
ind("population_1000","Economy","ประชากร (ประมาณการ สศช.)","Population (NESDC estimate)","1,000 persons","GPP01","GPP_4prov_2019-2024_clean.xlsx","All")
for r in sh("GPP_4prov_2019-2024_clean.xlsx","gpp_summary"):
    fl="preliminary (2024p)" if str(r["nesdc_label"]).endswith("p") else "ok"
    for c,k in [("gpp_current_mb","gpp_current_MB"),("gpp_real_growth_pct","real_growth_%"),("gpp_per_capita_baht","gpp_per_capita_baht"),("population_1000","population_1000")]:
        add(r["province"],r["year_BE"],c,r[k],flag=fl if c!="gpp_real_growth_pct" or r["province"]!="Yala" or r["year_BE"]!=2566 else "caution RF06 (base effect)")
SECT={"Agriculture, forestry and fishing":"agri","Manufacturing":"manuf","Construction":"constr","Wholesale and retail trade and repair of motor vehicles and motorcycle":"trade","Accommodation and food service activities":"accom","Public administration and defence; compulsory social security":"pubadmin","Education":"educ","Transportation and storage":"transport","Human health and social work activities":"health","Mining and quarrying":"mining"}
for s,c in SECT.items():
    ind(f"gpp_share_{c}_pct","Economy",f"สัดส่วนสาขา {c} ใน GPP",f"GPP share: {s}","% of GPP","GPP01","GPP_4prov_2019-2024_clean.xlsx","P1,P2")
    ind(f"gpp_real_growth_{c}_pct","Economy",f"การเติบโตที่แท้จริงสาขา {c}",f"Real growth: {s}","%","GPP01","GPP_4prov_2019-2024_clean.xlsx","P1,P3")
for r in sh("GPP_4prov_2019-2024_clean.xlsx","sectors_long"):
    c=SECT.get(r["sector"])
    if c:
        yBE=r["year_CE"]+543; fl="preliminary (2024p)" if str(r["nesdc_label"]).endswith("p") else "ok"
        add(r["province"],yBE,f"gpp_share_{c}_pct",r["share_%"],flag=fl)
        cf="caution RF07 (watch)" if c=="constr" and yBE==2566 and r["province"] in ("Yala","Narathiwat") else fl
        add(r["province"],yBE,f"gpp_real_growth_{c}_pct",r["real_growth_%"],flag=cf)
# ---------- LABOUR ----------
ind("unemployment_rate_pct","Labour","อัตราการว่างงาน (เฉลี่ย 4 ไตรมาส)","Unemployment rate (4-quarter avg)","%","LFS01","NSO_LFS_T1_4prov_2564Q1-2569Q2_clean.xlsx","P2,P3")
ind("lfp_rate_pct","Labour","อัตราการมีส่วนร่วมในกำลังแรงงาน","Labour force participation rate","%","LFS01","NSO_LFS_T1_4prov_2564Q1-2569Q2_clean.xlsx","P2")
ind("employed_persons","Labour","ผู้มีงานทำ","Employed persons","persons","LFS01","NSO_LFS_T1_4prov_2564Q1-2569Q2_clean.xlsx","P2")
for r in sh("NSO_LFS_T1_4prov_2564Q1-2569Q2_clean.xlsx","annual_avg"):
    add(r["province"],r["year_be"],"unemployment_rate_pct",r["unemployment_rate_pct"])
    add(r["province"],r["year_be"],"lfp_rate_pct",r["lfp_rate_pct"])
    add(r["province"],r["year_be"],"employed_persons",r["employed"],flag="caution RF01 (weight break 2564→2565; compare levels from 2565 only)")
# ---------- BUSINESS (DBD) ----------
for c,th,en,u in [("dbd_new_firms","นิติบุคคลตั้งใหม่","New juristic persons","firms"),("dbd_new_capital_mb","ทุนจดทะเบียนตั้งใหม่","Registered capital of new firms","million baht"),("dbd_dissolved","นิติบุคคลเลิกกิจการ","Dissolved juristic persons","firms"),("dbd_net_formation","ตั้งใหม่สุทธิ","Net firm formation","firms")]:
    ind(c,"Business",th,en,u,"DBD01-03 / DBD04","DBD_juristic_4prov_2564-2568_clean.xlsx","P1,P4")
for r in sh("DBD_juristic_4prov_2564-2568_clean.xlsx","annual"):
    fl="corrected Apr-2566 (RF02 resolved)" if r["year_be"]==2566 else "ok"
    add(r["province"],r["year_be"],"dbd_new_firms",r["new_reg"],flag=fl); add(r["province"],r["year_be"],"dbd_new_capital_mb",r["new_capital_mb"],flag="caution: Apr-2566 capital by province uncorrected" if r["year_be"]==2566 else "ok")
    add(r["province"],r["year_be"],"dbd_dissolved",r["dissolved"]); add(r["province"],r["year_be"],"dbd_net_formation",r["net_formation"],flag=fl)
for r in sh("DBD_firms_by_business_type_2565-2568_clean.xlsx","by_year_area_section"):
    s=r["TSIC_section"]
    for k,c,en in [("new_firms","new","New firms"),("dissolved","dissolved","Dissolved firms")]:
        code=f"dbd_{c}_sec{s}"
        if code not in DICT: ind(code,"Business",f"{'ตั้งใหม่' if c=='new' else 'เลิกกิจการ'} หมวด {s} {r['section_th']}",f"{en}, TSIC section {s}","firms","DBD04","DBD_firms_by_business_type_2565-2568_clean.xlsx","P1,P4")
        add(r["area"],r["year_BE"],code,r[k])
    if r["area"]=="Songkhla-4districts" or r["area"].startswith("3 provinces"):
        pass
# SKA4 & 3prov totals
tot=collections.defaultdict(lambda:[0,0])
for r in sh("DBD_firms_by_business_type_2565-2568_clean.xlsx","by_year_area_section"):
    if r["area"] in ("Songkhla-4districts","3 provinces (PTN+YLA+NWT)"): tot[(r["area"],r["year_BE"])][0]+=r["new_firms"]; tot[(r["area"],r["year_BE"])][1]+=r["dissolved"]
for (a,y),(n,d) in tot.items(): add(a,y,"dbd_new_firms",n,flag="from firm lists (DBD04)"); add(a,y,"dbd_dissolved",d,flag="from firm lists (DBD04)"); add(a,y,"dbd_net_formation",n-d,flag="from firm lists (DBD04)")
# ---------- INVESTMENT (BOI) ----------
for c,k,th,en,u in [("boi_projects","projects","โครงการ BOI ที่อนุมัติ","BOI approved projects","projects"),("boi_investment_mb","investment_MB","เงินลงทุนที่ได้รับอนุมัติ","BOI approved investment","million baht"),("boi_jobs","jobs","การจ้างงานของโครงการ BOI","BOI approved jobs","persons")]:
    ind(c,"Investment",th,en,u,"BOI01","BOI_4prov_2564-2568_clean.xlsx","P1,P2")
for r in sh("BOI_4prov_2564-2568_clean.xlsx","province_approvals"):
    if r["period"]!="Jan-Dec": continue
    fl="caution RF03/RF04 (approval ≠ realised; solar-heavy)" if r["year_BE"]==2568 else "approval ≠ realised (RF04)"
    for c,k in [("boi_projects","projects"),("boi_investment_mb","investment_MB"),("boi_jobs","jobs")]: add(r["province_en"],r["year_BE"],c,r[k],flag=fl)
# ---------- TOURISM ----------
for c,k,th,en,u in [("tour_visitors","visitors","ผู้เยี่ยมเยือน","Visitors","persons"),("tour_visitors_foreign","visitors_foreign","ผู้เยี่ยมเยือนชาวต่างชาติ","Foreign visitors","persons"),("tour_revenue_mb","revenue_mb","รายได้จากการท่องเที่ยว","Tourism revenue","million baht"),("tour_occupancy_pct","occupancy_rate_pct","อัตราการเข้าพัก","Hotel occupancy","%"),("tour_hotel_guests_malaysian","hotel_guests_malaysian","ผู้เข้าพักชาวมาเลเซีย","Malaysian hotel guests","persons"),("tour_spend_per_day_baht","avg_spend_per_visitor_baht_day","ค่าใช้จ่ายเฉลี่ยต่อคนต่อวัน","Average spend per visitor per day","baht")]:
    ind(c,"Tourism",th,en,u,"TO01","Tourism_4prov_2564-2568_clean.xlsx","P1,P3,P4")
for r in sh("Tourism_4prov_2564-2568_clean.xlsx","annual"):
    for c,k in [("tour_visitors","visitors"),("tour_visitors_foreign","visitors_foreign"),("tour_revenue_mb","revenue_mb"),("tour_occupancy_pct","occupancy_rate_pct"),("tour_hotel_guests_malaysian","hotel_guests_malaysian"),("tour_spend_per_day_baht","avg_spend_per_visitor_baht_day")]:
        fl="COVID base year" if r["year_be"]==2564 else ("caution RF09 (estimate)" if r["province"]=="Yala" and c=="tour_occupancy_pct" else "ok")
        add(r["province"],r["year_be"],c,r.get(k),flag=fl)
# ---------- BORDER TRADE ----------
ind("border_trade_mb","Trade","มูลค่าการค้าชายแดน (ส่งออก+นำเข้า)","Border trade value (export+import)","million baht","BT01-03","BorderTrade_3prov_2563-2568_clean.xlsx","P1,P3")
for r in sh("BorderTrade_3prov_2563-2568_clean.xlsx","annual"):
    area="Songkhla-4districts" if "Na Thawi" in str(r["province"]) else r["province"]
    fl="INCOMPLETE year" if r.get("note") and "INCOMPLETE" in str(r["note"]) else ("checkpoint: "+str(r["checkpoint"]))
    add(area,r["year_be"],"border_trade_mb",r["total_mb"],flag=fl)
# ---------- AGRICULTURE ----------
for r in sh("Agriculture_4prov_2563-2568_clean.xlsx","production_annual"):
    crop=re.sub(r"\W+","_",str(r["crop"]).lower()).strip("_"); code=f"agri_output_{crop}_t"
    if code not in DICT: ind(code,"Agriculture",f"ผลผลิต {r['crop']}",f"Output: {r['crop']}","tonnes (coconut: as source)","AG01","Agriculture_4prov_2563-2568_clean.xlsx","P1,P3,P4")
    fl="PRELIMINARY — do not use (RF10)" if r["year_be"]==2568 and r["crop"] not in ("Rubber","Oil palm") else "ok"
    add(r["province"],r["year_be"],code,r["output_ton"],flag=fl)
for r in sh("Agriculture_4prov_2563-2568_clean.xlsx","price_annual_avg"):
    code="price_rubber_rss3_baht_kg" if "ยาง" in str(r["product"]) else "price_oilpalm_baht_kg"
    if code not in DICT: ind(code,"Agriculture","ราคา"+str(r["product"]),"Farm-gate price: "+("rubber RSS3" if "ยาง" in str(r["product"]) else "oil palm FFB"),"baht/kg","AG02","Agriculture_4prov_2563-2568_clean.xlsx","P3,P4")
    add("Whole Kingdom",r["year_be"],code,r["avg_price_baht_per_kg"],flag="ok" if r["n_months"]==12 else f"partial year ({r['n_months']} months)")
# ---------- FISHERIES ----------
ind("fish_landings_t","Fisheries","สัตว์น้ำเค็มขึ้นท่า (กรมประมง)","Saltwater landings (DoF)","tonnes","FI-DoF","Fisheries_DoF_4prov_2562-2567_clean.xlsx","P1,P3,P4")
ind("fish_landings_value_mb","Fisheries","มูลค่าสัตว์น้ำขึ้นท่า","Landings value (DoF)","million baht","FI-DoF","Fisheries_DoF_4prov_2562-2567_clean.xlsx","P1,P4")
ind("fish_vessels_licensed","Fisheries","เรือประมงพาณิชย์ที่ได้รับอนุญาต","Licensed commercial fishing vessels","vessels","FI-DoF","Fisheries_DoF_4prov_2562-2567_clean.xlsx","P3,P4")
ind("fmo_port_volume_t","Fisheries","ปริมาณสัตว์น้ำผ่านท่า อสป. (ปีงบ)","FMO port volume (fiscal year)","tonnes","FI01","Fisheries_FMO_ports_FY2565-2568_clean.xlsx","P1,P4")
for r in sh("Fisheries_DoF_4prov_2562-2567_clean.xlsx","landings_province_total"):
    add(r["province"],r["year_BE"],"fish_landings_t",r["tonnes"]); add(r["province"],r["year_BE"],"fish_landings_value_mb",r["value_MB"])
for r in sh("Fisheries_DoF_4prov_2562-2567_clean.xlsx","vessels_licensed"):
    if r["year_BE"]>=2563: add(r["province_of_license"],r["year_BE"],"fish_vessels_licensed",r["vessels"])
for r in sh("Fisheries_FMO_ports_FY2565-2568_clean.xlsx","annual_FY"):
    if "Songkhla" in str(r["port"]): continue
    fl="INCOMPLETE" if "INCOMPLETE" in str(r["basis"]) else "fiscal year (Oct–Sep)"
    add(r["port"],r["fiscal_year_be"],"fmo_port_volume_t",r["volume_tonnes"],period="fiscal",flag=fl)
# ---------- POVERTY / HOUSEHOLD ----------
ind("poverty_rate_pct","Poverty","สัดส่วนคนจน (สศช.)","Poverty headcount ratio (NESDC)","% of population","PV02","Poverty_4prov_2564-2568_clean.xlsx","P2,P4")
ind("poverty_rank_77","Poverty","อันดับความยากจน (1=จนที่สุด)","Poverty rank of 77 (1=poorest)","rank","PV02","Poverty_4prov_2564-2568_clean.xlsx","P2")
for r in sh("Poverty_4prov_2564-2568_clean.xlsx","nesdc_poverty_rate"):
    fl="caution RF16 (sample noise)" if r["area_en"]=="Yala" else "ok"
    add(r["area_en"],r["year_BE"],"poverty_rate_pct",r["poverty_rate_pct"],period="survey",flag=fl); add(r["area_en"],r["year_BE"],"poverty_rank_77",r["rank_of_77_(1=poorest)"],period="survey")
for c,k,th,en,u in [("hh_income_month_baht","income_baht_per_month","รายได้เฉลี่ยต่อเดือนต่อครัวเรือน","Household income per month","baht"),("hh_expenditure_month_baht","expenditure_baht_per_month","ค่าใช้จ่ายเฉลี่ยต่อเดือนต่อครัวเรือน","Household expenditure per month","baht"),("hh_debt_baht","debt_baht_per_household","หนี้สินเฉลี่ยต่อครัวเรือน","Household debt","baht"),("hh_exp_to_income_pct","exp_to_income_pct","สัดส่วนค่าใช้จ่ายต่อรายได้","Expenditure / income","%")]:
    ind(c,"Poverty",th,en,u,"PV05/PV06","Poverty_4prov_2564-2568_clean.xlsx","P2,P4")
for r in sh("Poverty_4prov_2564-2568_clean.xlsx","ses_household_summary"):
    for c,k in [("hh_income_month_baht","income_baht_per_month"),("hh_expenditure_month_baht","expenditure_baht_per_month"),("hh_debt_baht","debt_baht_per_household"),("hh_exp_to_income_pct","exp_to_income_pct")]:
        add(r["area_en"],r["year_BE"],c,r[k],period="survey",flag="caution RF16 (province sample)" if r["area_en"] in TH2EN.values() else "ok")
ind("hh_informal_debt_baht","Poverty","หนี้นอกระบบเฉลี่ยต่อครัวเรือน","Informal debt per household","baht","PV06","Poverty_4prov_2564-2568_clean.xlsx","P3,P4")
ind("hh_transfers_share_pct","Poverty","สัดส่วนรายได้จากเงินช่วยเหลือ","Share of income from transfers","%","PV06","Poverty_4prov_2564-2568_clean.xlsx","P2")
ind("hh_farm_income_share_pct","Poverty","สัดส่วนรายได้จากการเกษตร","Share of income from farming","%","PV06","Poverty_4prov_2564-2568_clean.xlsx","P2,P4")
for r in sh("Poverty_4prov_2564-2568_clean.xlsx","ses_debt_by_purpose"): add(r["province"],r["year_BE"],"hh_informal_debt_baht",r["INFORMAL_debt"],period="survey",flag="household-weighted (method in notes)")
for r in sh("Poverty_4prov_2564-2568_clean.xlsx","ses_income_by_source"):
    add(r["province"],r["year_BE"],"hh_transfers_share_pct",r["transfers_%"],period="survey"); add(r["province"],r["year_BE"],"hh_farm_income_share_pct",r["farm_%"],period="survey")
ind("tpmap_target_poor_rate_pct","Poverty","สัดส่วนคนจนเป้าหมาย TPMAP","TPMAP target-poor rate","% of จปฐ. persons","PV01","Poverty_4prov_2564-2568_clean.xlsx","P2 (where only)")
for r in sh("Poverty_4prov_2564-2568_clean.xlsx","tpmap_province"):
    add(r["province_en"],r["jpt_survey_year"],"tpmap_target_poor_rate_pct",r["target_poor_rate_pct"],period="จปฐ. survey year",flag="caution RF17/RF18 — not a poverty rate; no trends")
ind("gini_income","Poverty","สัมประสิทธิ์ความไม่เสมอภาค (รายได้)","Gini coefficient (income)","0–1","PV04","Poverty_4prov_2564-2568_clean.xlsx","P2")
for r in sh("Poverty_4prov_2564-2568_clean.xlsx","nso_gini"):
    if r["year_BE"]>=2562: add(r["province_th"],r["year_BE"],"gini_income",r["gini_income"],period="survey",flag="caution RF16")
# ---------- SECURITY ----------
ind("sec_ucdp_lethal_events","Security","เหตุการณ์ที่มีผู้เสียชีวิต (UCDP)","Lethal conflict events (UCDP)","events","SEC01","Security_4prov_2562-2568_clean.xlsx","P3,P4")
ind("sec_ucdp_deaths","Security","ผู้เสียชีวิตจากความไม่สงบ (UCDP)","Conflict deaths (UCDP best estimate)","persons","SEC01","Security_4prov_2562-2568_clean.xlsx","P3,P4")
ind("sec_official_incidents","Security","เหตุการณ์ความไม่สงบทุกประเภท (ทางการ)","Unrest incidents, all types (official)","incidents","SEC02","Security_4prov_2562-2568_clean.xlsx","P3,P4")
for r in sh("Security_4prov_2562-2568_clean.xlsx","ucdp_annual_province"):
    add(r["province"],r["year_BE"],"sec_ucdp_lethal_events",r["lethal_events"],flag="Songkhla = 4 districts only" if r["province"]=="Songkhla" else "ok")
    add(r["province"],r["year_BE"],"sec_ucdp_deaths",r["deaths_best"],flag="Songkhla = 4 districts only" if r["province"]=="Songkhla" else "ok")
for r in sh("Security_4prov_2562-2568_clean.xlsx","official_annual"):
    add("Narathiwat",r["year_BE"],"sec_official_incidents",r["Narathiwat_incidents"])
    add("Yala",r["year_BE"],"sec_official_incidents",r["Yala_incidents_tableA"],flag="caution RF19 (two tables differ)" if r["year_BE"]>=2568 else "ok")
    add("Songkhla",r["year_BE"],"sec_official_incidents",r["Songkhla_incidents_wholeprov"])
# fix: UCDP Songkhla rows are 4-district only → relabel area
for row in rows:
    if row[7].startswith("sec_ucdp") and row[0]=="Songkhla": row[0]="Songkhla-4districts"; row[1]=LEVEL[row[0]]; row[2]=SCOPE[row[0]]; row[14]="ok"
# ---------- BUDGET ----------
for c,k,th,en,u in [("budget_total_mb","total_MB","งบประมาณที่จัดสรรลงจังหวัด","Budget allocated to province","million baht"),("budget_capital_mb","CAPITAL_investment_MB","งบลงทุนที่จัดสรร","Capital budget allocated","million baht"),("budget_capital_share_pct","capital_share_%","สัดส่วนงบลงทุน","Capital share of allocation","%")]:
    ind(c,"Public budget",th,en,u,"BUD01","Budget_4prov_FY2564-2569_clean.xlsx","P2")
for r in sh("Budget_4prov_FY2564-2569_clean.xlsx","allocation_5types"):
    fl="caution RF20 (security line jump)" if r["province"]=="Pattani" and r["FY_BE"]==2569 else "allocation ≠ disbursement"
    for c,k in [("budget_total_mb","total_MB"),("budget_capital_mb","CAPITAL_investment_MB"),("budget_capital_share_pct","capital_share_%")]: add(r["province"],r["FY_BE"],c,r[k],period="fiscal",flag=fl)
# ---------- REGION (BOT) ----------
ind("bot_private_investment_index","Region","ดัชนีการลงทุนภาคเอกชน ภาคใต้","Private investment index, South (2561=100)","index","BOT01","BOT_South_indicators_2562-2568_clean.xlsx","Context")
for r in sh("BOT_South_indicators_2562-2568_clean.xlsx","private_investment_index"):
    if r["index"]=="ดัชนีการลงทุนภาคเอกชน": add("South region",r["year_BE"],"bot_private_investment_index",r["value (2561=100)"],flag="preliminary" if "p" in str(r["label"]) else "ok")
ind("bot_customs_exports_mb","Region","มูลค่าส่งออกผ่านด่านศุลกากรภาคใต้","Exports via southern customs","million baht","BOT01","BOT_South_indicators_2562-2568_clean.xlsx","Context")
for r in sh("BOT_South_indicators_2562-2568_clean.xlsx","customs_trade_south_yearly"):
    if r["item"]=="สินค้าออก": add("South region",r["year_BE"],"bot_customs_exports_mb",r["value_MB_or_tonnes"])
# keep 2563–2569
rows=[r for r in rows if 2563<=r[3]<=2569]
for r in rows:
    if r[3]==2569 and r[14]=="ok": r[14]="partial/forward year"
    if r[3]==2563 and r[14]=="ok": r[14]="base year (before scope)"
rows.sort(key=lambda r:(r[6],r[7],r[0],r[3]))
COLS=["area","area_level","scope","year_BE","year_CE","period_type","domain","indicator_code","indicator_th","indicator_en","value","unit","source_id","clean_file","quality_flag","used_in"]
json.dump({"rows":rows,"dict":DICT},open("/tmp/panel.json","w"),ensure_ascii=False)
print(len(rows),"rows",len(DICT),"indicators")
print(collections.Counter(r[6] for r in rows))
print(collections.Counter(r[0] for r in rows))
