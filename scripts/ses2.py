import os,json,openpyxl
from openpyxl.styles import Font,PatternFill,Alignment
R=os.path.expanduser("~/mnt/SBP_Dashboard"); D=f"{R}/01_raw/10_poverty_income"
EN={"ปัตตานี":"Pattani","ยะลา":"Yala","นราธิวาส":"Narathiwat","สงขลา":"Songkhla","ทั่วราชอาณาจักร":"Whole Kingdom","ภาคใต้":"South region"}
ses=json.load(open("/tmp/ses.json"))
def longtab(fn,has_code):
    ws=openpyxl.load_workbook(f"{D}/{fn}",data_only=True)["ข้อมูล"]; rows=list(ws.iter_rows(values_only=True))
    hdr=[r for r in rows if r and any(isinstance(x,int) and 2540<x<2570 for x in r)][0]
    yrs=[x for x in hdr if isinstance(x,int)]
    out={}
    for r in rows:
        vals=[x for x in r if x is not None]
        names=[x for x in vals if isinstance(x,str)]
        nums=[x for x in vals if isinstance(x,(int,float)) and not isinstance(x,bool)]
        if not names or len(nums)<len(yrs): continue
        nm=[n for n in names if n.strip() in EN]
        if nm: out[nm[-1].strip()]=dict(zip(yrs,nums[-len(yrs):]))
    return yrs,out
yi,inc=longtab("NSO_SES_Income_byProvince_2547-2566_dl20261010.xlsx",True)
ye,exp=longtab("NSO_SES_Expenditure_byProvince_2555-2566_dl20261010.xlsx",False)
# debt: total rows only
ws=openpyxl.load_workbook(f"{D}/NSO_SES_DebtByPurpose_byProvince_2547-2566_dl20261010.xlsx",data_only=True)["ข้อมูล"]
rows=list(ws.iter_rows(values_only=True)); hdr=rows[2]; yd=[x for x in hdr if isinstance(x,int)]
debt={}; cur=None
ny=len(yd)
for r in rows[3:]:
    r=list(r)
    if isinstance(r[1],str) and r[1].strip(): cur=r[1].strip()
    purpose=r[2].strip() if isinstance(r[2],str) else None
    nums=r[3:3+ny]
    if cur and purpose: debt.setdefault(cur,{})[purpose]={y:v for y,v in zip(yd,nums) if isinstance(v,(int,float))}
print("income keys",list(inc),"exp",list(exp),"debt",list(debt))
for p in ["ปัตตานี","ยะลา","นราธิวาส","สงขลา"]:
    print(p,"inc2564",inc[p][2564],"exp2564",exp[p][2564],"debt2564",debt[p]["หนี้สินทั้งสิ้น"][2564], list(debt[p].keys()))
wb=openpyxl.load_workbook(f"{R}/02_clean/Poverty_4prov_2564-2568_clean.xlsx")
if "ses_income_PENDING" in wb.sheetnames: del wb["ses_income_PENDING"]
H=Font(bold=True,color="FFFFFF"); FL=PatternFill("solid",fgColor="1F4E78")
def sheet(name,cols,data,idx=None):
    if name in wb.sheetnames: del wb[name]
    w=wb.create_sheet(name,idx); w.append(cols)
    for c in w[1]: c.font=H; c.fill=FL; c.alignment=Alignment(wrap_text=True)
    for r in data: w.append(r)
    for i,c in enumerate(cols): w.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=max(11,min(38,len(str(c))+2))
    w.freeze_panes="A2"
order=["ปัตตานี","ยะลา","นราธิวาส","สงขลา","ภาคใต้","ทั่วราชอาณาจักร"]
summ=[]
for y in (2564,2565,2566,2568):
    for p in order:
        if y==2568:
            k=f"2568|{p}"
            if k not in ses: summ.append([y,y-543,EN[p],p,None,None,None,None,"2568 national/region not in province file"]); continue
            c=ses[k]["c1"]; i,e,d=c["รายได้ทั้งสิ้นต่อเดือน"],c["ค่าใช้จ่ายทั้งสิ้นต่อเดือน"],c["หนี้สินเฉลี่ยต่อครัวเรือนทั้งสิ้น"]; src="NSO catalogapi SFD_SPB0801"
        else:
            i=inc.get(p,{}).get(y); e=exp.get(p,{}).get(y); d=debt.get(p,{}).get("หนี้สินทั้งสิ้น",{}).get(y); src="NSO stat tables (jj/jA/jm)"
        summ.append([y,y-543,EN[p],p,round(i) if i else None,round(e) if e else None,round(d) if d else None,round(e/i*100,1) if i and e else None,src])
sheet("ses_household_summary",["year_BE","year_CE","area_en","area_th","income_baht_per_month","expenditure_baht_per_month","debt_baht_per_household","exp_to_income_pct","source"],summ,0)
# long series
lr=[]
for p in order:
    for y in sorted(set(yi)|set(ye)):
        lr.append([y,EN[p],inc.get(p,{}).get(y),exp.get(p,{}).get(y),debt.get(p,{}).get("หนี้สินทั้งสิ้น",{}).get(y)])
sheet("ses_long_series_2547-2566",["year_BE","area","income_month","expenditure_month","debt_household"],lr)
# income by source
SRC=["ค่าจ้างและเงินเดือน","กำไรสุทธิจากการทำธุรกิจ","กำไรสุทธิจากการทำการเกษตร","เงินที่ได้รับเป็นการช่วยเหลือ","รายได้จากทรัพย์สิน"]
ib=[]
for y in ("2566","2568"):
    for p in order[:4]:
        v=ses[f"{y}|{p}"]["inc"]; tot=v["รายได้ทั้งสิ้นต่อเดือน|รายได้ทั้งสิ้นต่อเดือน|รายได้ทั้งสิ้นต่อเดือน"]
        g=lambda s:v.get(f"รายได้ประจำ|รายได้ที่เป็นตัวเงิน|{s}",0)
        nonc=v.get("รายได้ประจำ|รายได้ที่ไม่เป็นตัวเงิน|รายได้ที่ไม่เป็นตัวเงิน",0); irr=v.get("รายได้ไม่ประจำ (ที่เป็นตัวเงิน)|รายได้ไม่ประจำ (ที่เป็นตัวเงิน)|รายได้ไม่ประจำ (ที่เป็นตัวเงิน)",0)
        row=[int(y),EN[p],tot]+[g(s) for s in SRC]+[nonc,irr]+[round(g(s)/tot*100,1) for s in SRC]
        ib.append(row)
sheet("ses_income_by_source",["year_BE","province","total_income_month","wages","business_profit","farm_profit","transfers_assistance","property_income","non_cash_income","irregular_cash_income",
 "wages_%","business_%","farm_%","transfers_%","property_%"],ib)
# debt by purpose / formal-informal
PUR=["ใช้จ่ายอุปโภค บริโภคอื่น ๆ ในครัวเรือน","ใช้ซื้อ/เช่าซื้อบ้านและ/หรือที่ดิน","ใช้ในการศึกษา","ใช้ในการทำธุรกิจ","ใช้ในการทำเกษตร","หนี้อื่น ๆ"]
db=[]
for y in ("2566","2568"):
    for p in order[:4]:
        v=ses[f"{y}|{p}"]["debt"]; tot=v["จำนวนหนี้สินเฉลี่ยต่อครัวเรือน|จำนวนหนี้สินเฉลี่ยต่อครัวเรือน"]
        f_=v.get("จำแนกตามแหล่งเงินกู้|หนี้ในระบบ",0); inf=v.get("จำแนกตามแหล่งเงินกู้|หนี้นอกระบบ",0)
        db.append([int(y),EN[p],tot]+[v.get(f"จำแนกตามวัตถุประสงค์|{k}",0) for k in PUR]+[f_,inf,round(inf/tot*100,1) if tot else None])
sheet("ses_debt_by_purpose",["year_BE","province","debt_per_household","consumption","house_land","education","business","farming","other","formal_debt","INFORMAL_debt","informal_%"],db)
# class mix
cm=[]
for y in ("2566","2568"):
    for p in order[:4]:
        for k,v in sorted(ses[f"{y}|{p}"]["cls"].items(),key=lambda x:-x[1]): cm.append([int(y),EN[p],k,v])
sheet("ses_household_class_mix",["year_BE","province","socio_economic_class","share_of_households_%"],cm)
n=wb["notes"]
for l in ["SES (sheets ses_*): NSO Household Socio-Economic Survey, province level. 2564 & 2566 (+2565 expenditure) from NSO official tables (statistics_and_indicators branch 309: jj/jA/jm); 2568 from catalogapi SFD_SPB0801. 2566 values in both sources match exactly (checked).",
"ses_income_by_source / ses_debt_by_purpose: NSO publishes these only by socio-economic class; province totals = household-weighted average across classes (weights = households per class from SFD_SPB0806). Check: weighted totals reproduce official province totals to ±1 baht.",
"Songkhla debt 2566 (227,079) → 2568 (55,626): almost all of the fall is house/land (109,463 → 19,161) and business loans (58,704 → 1,498) — a few big-loan households in the 2566 sample. Sample noise, not a real collapse (RF16).",
"Raw duplicates: 'data (4).csv' = same as C1, 'data (3).csv' = same Gini file → moved to 01_raw/10_poverty_income/_duplicates (md5 identical)."]: n.append([l])
wb.save(f"{R}/02_clean/Poverty_4prov_2564-2568_clean.xlsx")
json.dump({"summ":summ,"ib":ib,"db":db},open("/tmp/ses2.json","w"),ensure_ascii=False)
for r in summ: print(r[:8])
print(); [print(r) for r in ib]; print(); [print(r) for r in db]
