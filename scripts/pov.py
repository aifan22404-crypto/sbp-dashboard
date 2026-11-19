import os, re, csv, shutil, json, openpyxl, subprocess
from openpyxl.styles import Font, PatternFill, Alignment
R=os.path.expanduser("~/mnt/SBP_Dashboard"); D=f"{R}/01_raw/10_poverty_income"
PROV={"90":"Songkhla","94":"Pattani","95":"Yala","96":"Narathiwat"}
TH={"Songkhla":"สงขลา","Pattani":"ปัตตานี","Yala":"ยะลา","Narathiwat":"นราธิวาส"}
AB={"90":"SKA","94":"PTN","95":"YLA","96":"NWT"}
# ---------- 1) rename raw ----------
ren={}
for yy in range(64,70):
    Y=2500+yy
    ren[f"{yy}_province_11111.csv"]=f"TPMAP_poor_allprov_TPMAP{Y}_dl20261010.csv"
    ren[f"{yy}_province_common.csv"]=f"TPMAP_common_allprov_TPMAP{Y}_dl20261010.csv"
    for p in PROV: ren[f"{yy}_{p}_amphur_11111.csv"]=f"TPMAP_poor_amphur_{AB[p]}_TPMAP{Y}_dl20261010.csv"
ren["TPMAP_datadict_plan13.xlsx"]="TPMAP_datadict_plan13_dl20261010.xlsx"
ren["SD10_04_252_1.pdf"]="NSO_IncomeDistribution_byProvince_report_2566_dl20261010.pdf"
ren["poverty-report-2568-FY.pdf"]="NESDC_PovertyInequality_report_2568_dl20261010.pdf"
done=[]
for a,b in ren.items():
    if os.path.exists(f"{D}/{a}"): os.rename(f"{D}/{a}",f"{D}/{b}"); done.append(a)
g=f"{R}/00_admin/data.csv"
if os.path.exists(g):
    shutil.move(g,f"{D}/NSO_SES_Gini_byProvince_2556-2566_dl20261009.csv"); ren["00_admin/data.csv"]="NSO_SES_Gini_byProvince_2556-2566_dl20261009.csv"; done.append("data.csv")
with open(f"{D}/_DOWNLOAD_LINKS.txt","a",encoding="utf-8") as f:
    f.write("\n## RENAMED 2026-10-10 (original -> project name)\n")
    for a in done: f.write(f"# {a} -> {ren[a if a!='data.csv' else '00_admin/data.csv']}\n")
    f.write("# D1 (2564 provincial report) NOT downloadable: catalog.nso.go.th returns HTTP 500 for every uploaded year (2558–2564) — NSO server fault, checked 2026-10-10.\n")
print("renamed",len(done))
# ---------- 2) NESDC table ก.6 ----------
txt=subprocess.run(["pdftotext","-layout",f"{D}/NESDC_PovertyInequality_report_2568_dl20261010.pdf","-"],capture_output=True,text=True).stdout.splitlines()
s=[i for i,l in enumerate(txt) if "ตารางที่ ก.6" in l][0]; e=[i for i,l in enumerate(txt) if "ตารางที่ ก.7" in l][0]
years=list(range(2551,2569)); region=None; rows=[]
regmap={"กรุงเทพมหานคร":"Bangkok","ภาคกลาง":"Central","ภาคเหนือ":"North","ภาคตะวันออก":"Northeast","ภาคใต้":"South"}
for l in txt[s:e]:
    nums=re.findall(r"(?<![\d.])\d+\.\d{2}(?![\d.])",l)
    head=re.split(r"\s{2,}",l.strip())
    for k,v in regmap.items():
        if l.strip().startswith(k) and not l.strip().startswith("กรุงเทพมหานคร   กรุงเทพ") or l.strip().startswith(k+" "): pass
    st=l.strip()
    for k,v in regmap.items():
        if st.startswith(k) and (k!="กรุงเทพมหานคร" or st.count("กรุงเทพมหานคร")==2): region=v
    if not nums: continue
    label=re.sub(r"[\d.\s]+$","",st)
    for k in regmap: 
        if label.startswith(k) and label!=k: label=label[len(k):].strip()
    label=label.replace("เฉียงเหนือ","").strip()
    if "ทั่วประเทศ" in label: label="ทั่วประเทศ"; reg="Whole Kingdom"
    else: reg=region
    if label=="รวม": label=f"รวม ({reg})"
    vals=[float(x) for x in nums]; ys=years[-len(vals):]
    rows.append((reg,label,dict(zip(ys,vals))))
prov=[r for r in rows if not r[1].startswith("รวม") and r[1]!="ทั่วประเทศ"]
print("NESDC rows",len(rows),"provinces",len(prov))
def rank(y,name):
    vs=sorted([(r[2].get(y),r[1]) for r in prov if r[2].get(y) is not None],reverse=True)
    return [n for v,n in vs].index(name)+1 if name in [n for v,n in vs] else None
# ---------- 3) TPMAP ----------
def rd(fn):
    with open(f"{D}/{fn}",encoding="utf-8-sig") as f: return list(csv.DictReader(f))
names={}
tp_prov=[]; tp_amp=[]; nat=[]
AMP={"9001":"เมืองสงขลา","9002":"สทิงพระ","9003":"จะนะ","9004":"นาทวี","9005":"เทพา","9006":"สะบ้าย้อย","9007":"ระโนด","9008":"กระแสสินธุ์","9009":"รัตภูมิ","9010":"สะเดา","9011":"หาดใหญ่","9012":"นาหม่อม","9013":"ควนเนียง","9014":"บางกล่ำ","9015":"สิงหนคร","9016":"คลองหอยโข่ง",
"9401":"เมืองปัตตานี","9402":"โคกโพธิ์","9403":"หนองจิก","9404":"ปะนาเระ","9405":"มายอ","9406":"ทุ่งยางแดง","9407":"สายบุรี","9408":"ไม้แก่น","9409":"ยะหริ่ง","9410":"ยะรัง","9411":"กะพ้อ","9412":"แม่ลาน",
"9501":"เมืองยะลา","9502":"เบตง","9503":"บันนังสตา","9504":"ธารโต","9505":"ยะหา","9506":"รามัน","9507":"กาบัง","9508":"กรงปินัง",
"9601":"เมืองนราธิวาส","9602":"ตากใบ","9603":"บาเจาะ","9604":"ยี่งอ","9605":"ระแงะ","9606":"รือเสาะ","9607":"ศรีสาคร","9608":"แว้ง","9609":"สุคิริน","9610":"สุไหงโก-ลก","9611":"สุไหงปาดี","9612":"จะแนะ","9613":"เจาะไอร้อง"}
SCOPE4={"9003","9004","9005","9006"}
def f(x):
    try: return float(x)
    except: return None
for yy in range(64,70):
    Y=2500+yy
    P=rd(f"TPMAP_poor_allprov_TPMAP{Y}_dl20261010.csv"); C={r["province_ID"]:r for r in rd(f"TPMAP_common_allprov_TPMAP{Y}_dl20261010.csv")}
    tot_poor=sum(f(r["poor.JPT.MOFval.CNT"]) or 0 for r in P); tot_pop=sum(f(C[r["province_ID"]]["HOUSEMEMBER_CNT"]) or 0 for r in P if r["province_ID"] in C)
    ranked=sorted(P,key=lambda r:-(f(r["JPT.MOFval.pov.rate"]) or 0)); rk={r["province_ID"]:i+1 for i,r in enumerate(ranked)}
    nat.append((Y,Y-1,tot_poor,tot_pop,round(tot_poor/tot_pop*100,3) if tot_pop else None,len(P)))
    for r in P:
        pid=r["province_ID"].strip('"')
        if pid in PROV:
            c=C.get(pid,{})
            tp_prov.append([Y,Y-1,PROV[pid],TH[PROV[pid]],f(c.get("household.CNT")),f(c.get("HOUSEMEMBER_CNT")),f(r["poor.household.CNT"]),f(r["poor.JPT.CNT"]),f(r["poor.JPT.MOFval.CNT"]),round(f(r["JPT.MOFval.pov.rate"]),3),rk[r["province_ID"]],len(P),
                f(r["poor.JPT.MOFval.health"]),f(r["poor.JPT.MOFval.living"]),f(r["poor.JPT.MOFval.education"]),f(r["poor.JPT.MOFval.income"]),f(r["poor.JPT.MOFval.accessibility"]),f(r["poor.JPT.MOFval.ind29.CNT"]),f(r["poor.JPT.MOFval.ind27.CNT"])])
    for p in PROV:
        for r in rd(f"TPMAP_poor_amphur_{AB[p]}_TPMAP{Y}_dl20261010.csv"):
            a=r["amphur_ID"].strip('"'); rate=f(r["JPT.MOFval.pov.rate"]); n=f(r["poor.JPT.MOFval.CNT"])
            surveyed=round(n/rate*100) if rate else None
            tp_amp.append([Y,Y-1,PROV[p],a,AMP.get(a,"?"),"IN SCOPE" if (p!="90" or a in SCOPE4) else "context only (Songkhla, outside 4 districts)",
                f(r["poor.household.CNT"]),f(r["poor.JPT.CNT"]),n,round(rate,3) if rate is not None else None,surveyed,
                f(r["poor.JPT.MOFval.health"]),f(r["poor.JPT.MOFval.living"]),f(r["poor.JPT.MOFval.education"]),f(r["poor.JPT.MOFval.income"]),f(r["poor.JPT.MOFval.accessibility"])])
print("tpmap prov",len(tp_prov),"amp",len(tp_amp))
# ---------- 4) Gini ----------
gini=[r for r in rd("NSO_SES_Gini_byProvince_2556-2566_dl20261009.csv") if r["PROVINCE"] in TH.values()]
# ---------- 5) write clean ----------
H=Font(bold=True,color="FFFFFF"); FL=PatternFill("solid",fgColor="1F4E78")
def sheet(wb,name,cols,data):
    ws=wb.create_sheet(name); ws.append(cols)
    for c in ws[1]: c.font=H; c.fill=FL; c.alignment=Alignment(wrap_text=True,vertical="center")
    for r in data: ws.append(r)
    for i,c in enumerate(cols): ws.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=max(10,min(40,len(str(c))+2))
    ws.freeze_panes="A2"; return ws
wb=openpyxl.Workbook(); wb.remove(wb.active)
nd=[]
for y in range(2564,2569):
    for nm in ["ปัตตานี","ยะลา","นราธิวาส","สงขลา"]:
        r=[x for x in prov if x[1]==nm][0]
        en=[k for k,v in TH.items() if v==nm][0]
        nd.append([y,y-543,en,nm,r[2].get(y),rank(y,nm),len([x for x in prov if x[2].get(y) is not None]),"WHOLE province" if en=="Songkhla" else "in scope"])
    for lab in ["รวม (South)","ทั่วประเทศ"]:
        r=[x for x in rows if x[1]==lab][0]; nd.append([y,y-543,lab.replace("รวม (South)","South region").replace("ทั่วประเทศ","Whole Kingdom"),lab,r[2].get(y),None,None,"benchmark"])
sheet(wb,"nesdc_poverty_rate",["year_BE","year_CE","area_en","area_th","poverty_rate_pct","rank_of_77_(1=poorest)","n_provinces","note"],nd)
ws=sheet(wb,"nesdc_all_provinces_2551-2568",["region","province"]+years,[[r[0],r[1]]+[r[2].get(y) for y in years] for r in rows])
sheet(wb,"tpmap_province",["TPMAP_year","jpt_survey_year","province_en","province_th","jpt_households","jpt_persons","poor_households","poor_persons_MPI","target_poor_persons","target_poor_rate_pct","rank_of_77_(1=highest)","n_prov",
 "target_poor_health","target_poor_living","target_poor_education","target_poor_income","target_poor_access","fail_ind29_income_per_person","fail_ind27_age15-59_job_income"],tp_prov)
sheet(wb,"tpmap_district",["TPMAP_year","jpt_survey_year","province_en","amphur_ID","amphur_th","scope","poor_households","poor_persons_MPI","target_poor_persons","target_poor_rate_pct","persons_surveyed_(derived)",
 "target_poor_health","target_poor_living","target_poor_education","target_poor_income","target_poor_access"],tp_amp)
sheet(wb,"tpmap_national",["TPMAP_year","jpt_survey_year","target_poor_persons","jpt_persons","target_poor_rate_pct","n_prov"],nat)
sheet(wb,"nso_gini",["year_BE","province_th","gini_income"],[[int(r["YEAR"]),r["PROVINCE"],f(r["VALUE"])] for r in gini])
sheet(wb,"ses_income_PENDING",["status"],[["Waiting for NSO tables jj/jA/jm (2547–2566) + SFD_SPB0801 (2566, 2568). Will add income / expenditure / debt per household here."]])
wn=wb.create_sheet("notes")
for l in [
"TWO DIFFERENT POVERTY MEASURES — never mix them:",
"(1) NESDC poverty rate (sheet nesdc_*): % of people whose consumption is below the poverty line. Source: NESDC 'รายงานการวิเคราะห์สถานการณ์ความยากจนและความเหลื่อมล้ำ ปี 2568', Appendix table ก.6 (p.101–102), from NSO SES sample survey. Province level only. Year = survey year. Table title says 2551–2567 but the table has a 2568 column (kept).",
"(2) TPMAP 'target poor' (sheets tpmap_*): people in จปฐ. (CDD basic-needs census of rural households) who fail MPI indicators AND are in households registered for the state welfare card. A TARGETING LIST, not a poverty rate. Has DISTRICT level.",
"TPMAP year N uses จปฐ. survey of year N-1 (e.g. TPMAP 2569 = จปฐ. 2568). Column jpt_survey_year = N-1. Our scope 2564–2568 → TPMAP 2565–2569; TPMAP 2564 kept as base.",
"Songkhla: NESDC = whole province (context only). TPMAP district sheet marks Chana(9003), Na Thawi(9004), Thepha(9005), Saba Yoi(9006) as IN SCOPE.",
"SES province estimates are noisy (small samples) — e.g. Yala 9.04% (2566) → 25.54% (2567). Use 2–3 year averages. See red flag RF16/RF17.",
"Gini: NSO SES via catalogapi (SFD_01_Table_0011_01); Narathiwat missing for 2560–2566 in source.",
]: wn.append([l])
wn.column_dimensions["A"].width=160
out=f"{R}/02_clean/Poverty_4prov_2564-2568_clean.xlsx"; wb.save(out); print("saved",out)
json.dump({"nd":nd,"prov":[r for r in tp_prov],"nat":nat,"amp_scope":[r for r in tp_amp if r[0]==2569 and r[5]=="IN SCOPE"]},open("/tmp/pov.json","w"),ensure_ascii=False)
