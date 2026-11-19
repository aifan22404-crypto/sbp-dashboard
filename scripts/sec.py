import os,json,csv,collections,openpyxl,shutil
from openpyxl.styles import Font,PatternFill,Alignment
R=os.path.expanduser("~/mnt/SBP_Dashboard"); D=f"{R}/01_raw/08_security"
# ---- rename raw
ren={"ged261-csv.zip":"UCDP_GED_v26.1_global_1989-2025_dl20261010.zip","ged261.pdf":"UCDP_GED_v26.1_codebook_dl20261010.pdf",
"unrest-2.csv":"NWT_DOPA_Unrest_byDistrict_byType_2562-2569_dl20261010.csv","untitled.csv":"YLA_DOPA_Unrest_annual_2560-2569_dl20261010.csv",
"untitled (1).csv":"YLA_DOPA_Unrest_annual_tableB_2565-2569_dl20261010.csv","dataset_03_38.csv":"SKA_DOPA_Unrest_incidents_2555-2568_dl20261010.csv",
"dataset_03_39.csv":"SKA_DOPA_Unrest_injured_2555-2568_dl20261010.csv","dataset_03_40.csv":"SKA_DOPA_Unrest_deaths_2555-2568_dl20261010.csv"}
for a,b in ren.items():
    if os.path.exists(f"{D}/{a}"): os.rename(f"{D}/{a}",f"{D}/{b}")
with open(f"{D}/_DOWNLOAD_LINKS.txt","a",encoding="utf-8") as f:
    f.write("\n## RENAMED 2026-10-10\n"+"".join(f"# {a} -> {b}\n" for a,b in ren.items()))
AMP={"Bacho":"บาเจาะ","Chanae":"จะแนะ","Cho Ai Rong":"เจาะไอร้อง","Mueang Narathiwat":"เมืองนราธิวาส","Rangae":"ระแงะ","Rueso":"รือเสาะ","Si Sakhon":"ศรีสาคร","Sukhirin":"สุคิริน","Sungai Kolok":"สุไหงโก-ลก","Sungai Padi":"สุไหงปาดี","Tak Bai":"ตากใบ","Yingo":"ยี่งอ","Waeng":"แว้ง",
"Kapho":"กะพ้อ","Khok Pho":"โคกโพธิ์","Mae Lan":"แม่ลาน","Mai Kaen":"ไม้แก่น","Mayo":"มายอ","Mueang Pattani":"เมืองปัตตานี","Nong Chik":"หนองจิก","Panare":"ปะนาเระ","Sai Buri":"สายบุรี","Thung Yang Daeng":"ทุ่งยางแดง","Yarang":"ยะรัง","Yaring":"ยะหริ่ง",
"Chana":"จะนะ","Saba Yoi":"สะบ้าย้อย","Sadao":"สะเดา","Thepha":"เทพา","Na Thawi":"นาทวี","Hat Yai":"หาดใหญ่",
"Bannang Sata":"บันนังสตา","Kabang":"กาบัง","Krong Penang":"กรงปินัง","Mueang Yala":"เมืองยะลา","Raman":"รามัน","Than To":"ธารโต","Yaha":"ยะหา","Betong":"เบตง"}
PV={"Narathiwat province":"Narathiwat","Pattani province":"Pattani","Songkhla province":"Songkhla","Yala province":"Yala"}
SK4={"Chana","Saba Yoi","Thepha","Na Thawi"}
th=json.load(open('/tmp/ucdp_th.json'))
ev=[]
for x in th:
    if x['adm_1'] not in PV or int(x['year'])<2019: continue
    d=x['adm_2'].replace(" district","").strip(); p=PV[x['adm_1']]
    scope="IN SCOPE" if p!="Songkhla" else ("IN SCOPE (4 districts)" if d in SK4 else ("UNKNOWN district" if not d else "outside scope"))
    if p!="Songkhla" and not d: scope="IN SCOPE (district unknown)"
    ev.append([int(x['year']),int(x['year'])+543,x['date_start'][:10],p,d or None,AMP.get(d) if d else None,scope,x['where_coordinates'],int(x['where_prec']),
        float(x['latitude']),float(x['longitude']),{"1":"state-based (insurgents vs state)","3":"one-sided (vs civilians)"}.get(x['type_of_violence'],x['type_of_violence']),
        x['side_a'],x['side_b'],int(x['deaths_a']),int(x['deaths_b']),int(x['deaths_civilians']),int(x['deaths_unknown']),int(x['best']),int(x['low']),int(x['high']),x['id']])
ev.sort(key=lambda r:r[2])
# annual
ann=collections.defaultdict(lambda:[0,0,0])
for r in ev:
    if r[6].startswith("IN SCOPE"):
        k=(r[0],r[3]); ann[k][0]+=1; ann[k][1]+=r[18]; ann[k][2]+=r[16]
annual=[[y,y+543,p,*ann.get((y,p),[0,0,0])] for y in range(2019,2026) for p in ["Pattani","Yala","Narathiwat","Songkhla"]]
dist=collections.defaultdict(lambda:[0,0])
for r in ev:
    if r[6].startswith("IN SCOPE") and 2021<=r[0]<=2025: k=(r[3],r[4] or "(unknown)",r[5]); dist[k][0]+=1; dist[k][1]+=r[18]
district=sorted([[*k,*v] for k,v in dist.items()],key=lambda r:-r[4])
# official
def rcsv(fn,enc="utf-8-sig"): return list(csv.reader(open(f"{D}/{fn}",encoding=enc)))
nwt=[r for r in rcsv("NWT_DOPA_Unrest_byDistrict_byType_2562-2569_dl20261010.csv") if r and r[0].isdigit()]
nwt_rows=[[int(r[0]),r[2],r[4],int(r[5] or 0)] for r in nwt]
nwt_year=collections.Counter(); [nwt_year.update({r[0]:r[3]}) for r in nwt_rows]
yla=[[int(r[1]),int(r[3])] for r in rcsv("YLA_DOPA_Unrest_annual_2560-2569_dl20261010.csv","cp874")[1:] if r]
ylb=[[int(r[1]),int(r[3])] for r in rcsv("YLA_DOPA_Unrest_annual_tableB_2565-2569_dl20261010.csv","cp874")[1:] if r]
def sk(fn): return {int(r[0]):(None if r[2].strip()=="-" else int(r[2])) for r in rcsv(fn)[1:] if r and r[0].isdigit()}
ska_i=sk("SKA_DOPA_Unrest_incidents_2555-2568_dl20261010.csv"); ska_inj=sk("SKA_DOPA_Unrest_injured_2555-2568_dl20261010.csv"); ska_d=sk("SKA_DOPA_Unrest_deaths_2555-2568_dl20261010.csv")
off=[]
for y in range(2560,2570):
    off.append([y,y-543,nwt_year.get(y),dict(yla).get(y),dict(ylb).get(y),ska_i.get(y),ska_d.get(y),ska_inj.get(y)])
# comparison 2564-2568
cmp_=[]
for y in range(2564,2569):
    ce=y-543
    cmp_.append([y,ce,ann[(ce,"Narathiwat")][0],ann[(ce,"Narathiwat")][1],nwt_year.get(y),ann[(ce,"Yala")][0],ann[(ce,"Yala")][1],dict(yla).get(y),dict(ylb).get(y),ann[(ce,"Pattani")][0],ann[(ce,"Pattani")][1],
        sum(1 for r in ev if r[0]==ce and r[6]=="IN SCOPE (4 districts)"),ska_i.get(y),ska_d.get(y)])
# write
wb=openpyxl.Workbook(); wb.remove(wb.active)
H=Font(bold=True,color="FFFFFF"); FL=PatternFill("solid",fgColor="1F4E78")
def sheet(n,c,d):
    w=wb.create_sheet(n); w.append(c)
    for x in w[1]: x.font=H; x.fill=FL; x.alignment=Alignment(wrap_text=True)
    for r in d: w.append(r)
    for i,cc in enumerate(c): w.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=max(10,min(36,len(str(cc))+2))
    w.freeze_panes="A2"
sheet("compare_2564-2568",["year_BE","year_CE","NWT_ucdp_events","NWT_ucdp_deaths","NWT_official_incidents","YLA_ucdp_events","YLA_ucdp_deaths","YLA_official_A","YLA_official_B","PTN_ucdp_events","PTN_ucdp_deaths","SKA4_ucdp_events","SKA_official_incidents_wholeprov","SKA_official_deaths_wholeprov"],cmp_)
sheet("ucdp_annual_province",["year_CE","year_BE","province","lethal_events","deaths_best","civilian_deaths"],annual)
sheet("ucdp_by_district_2021-2025",["province","district_en","district_th","lethal_events","deaths_best"],district)
sheet("ucdp_events_2019-2025",["year_CE","year_BE","date","province","district_en","district_th","scope","place","where_precision(1=exact)","lat","lon","violence_type","side_a","side_b","deaths_side_a","deaths_side_b","deaths_civilians","deaths_unknown","deaths_best","deaths_low","deaths_high","ucdp_id"],ev)
sheet("official_annual",["year_BE","year_CE","Narathiwat_incidents","Yala_incidents_tableA","Yala_incidents_tableB","Songkhla_incidents_wholeprov","Songkhla_deaths","Songkhla_injured"],off)
sheet("official_NWT_district_type",["year_BE","district","type","incidents"],nwt_rows)
n=wb.create_sheet("notes")
for l in ["UCDP GED v26.1 (Uppsala Conflict Data Program, 2026): events of ORGANIZED violence with ≥1 death. Non-lethal bombings, arson, injuries-only attacks are NOT included → UCDP undercounts 'incidents' but is consistent and independent. Cite: Davies, Pettersson & Öberg (2026) JPR; Sundberg & Melander (2013).",
"An 'event' in UCDP = one actor-dyad, one place, one day. deaths_best = best estimate. Filtered: Thailand, 4 provinces, 2019–2025. 2025 Thai–Cambodian border events excluded (different conflict, other provinces).",
"Songkhla: only Chana, Thepha, Saba Yoi, Na Thawi are in scope (Sadao event 2022 marked outside scope).",
"Official counts: Provincial Administration Offices (ที่ทำการปกครองจังหวัด) via provincial data catalogs. Count incidents of any kind (incl. non-lethal). Narathiwat has district × type; Yala and Songkhla only provincial totals; Songkhla = WHOLE province. Pattani: no official dataset published.",
"Yala official tables disagree for 2568 (71 vs 21) and 2569 (22 vs 10) — same office (RF19).",
"Thai years: official = BE (likely calendar year; not stated). UCDP = calendar year CE.",
"Deep South Watch public monthly datasheets end Oct 2564 → full DSW data 2565+ only by request (optional)."]: n.append([l])
n.column_dimensions["A"].width=160
wb.save(f"{R}/02_clean/Security_4prov_2562-2568_clean.xlsx")
json.dump({"cmp":cmp_,"dist":district[:12],"annual":annual},open("/tmp/sec.json","w"),ensure_ascii=False)
for r in cmp_: print(r)
print(district[:12])
print(len(ev),"events")
