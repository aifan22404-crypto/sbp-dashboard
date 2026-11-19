import csv,collections,os,json
D=os.path.expanduser("~/mnt/SBP_Dashboard/01_raw/10_poverty_income")
P=["ปัตตานี","ยะลา","นราธิวาส","สงขลา"]
def rd(f): return list(csv.DictReader(open(f"{D}/{f}",encoding="utf-8-sig")))
C1=rd("NSO_SES_IncExpDebt_byProvince_2566-2568_dl20261010.csv")
C2=rd("NSO_SES_IncomeBySource_byProvince_2566-2568_dl20261010.csv")
C3=rd("NSO_SES_DebtByPurpose_byProvince_2566-2568_dl20261010.csv")
fv=lambda x: float(x) if x not in ("",None) else 0.0
# household weights by class2
W=collections.defaultdict(float)
for r in C3:
    if r["purpose_source_bor"]=="จำนวนครัวเรือนทั้งสิ้น": W[(r["year"],r["province"],r["soc_eco_class2"])]+=fv(r["value"])
out={}
for y in ("2566","2568"):
  for p in P:
    tot={r["mthincome_mthexp_totaldebt_pctexptoincome"]:fv(r["value"]) for r in C1 if r["year"]==y and r["province"]==p and r["soc_eco_class1"]=="รวมทั้งสิ้น"}
    hh=sum(v for k,v in W.items() if k[0]==y and k[1]==p)
    inc=collections.defaultdict(float)
    for r in C2:
        if r["year"]==y and r["province"]==p:
            w=W.get((y,p,r["soc_eco_class2"]),0)
            key=(r["source_income1"],r["source_income2"],r["source_income3"])
            inc[key]+=fv(r["value"])*w
    inc={k:v/hh for k,v in inc.items()}
    debt=collections.defaultdict(float)
    for r in C3:
        if r["year"]==y and r["province"]==p and r["unit"]=="บาท":
            w=W.get((y,p,r["soc_eco_class2"]),0); debt[(r["hhdebt_totaldebt_purpose_source"],r["purpose_source_bor"])]+=fv(r["value"])*w
    debt={k:v/hh for k,v in debt.items()}
    # class shares
    cls=collections.defaultdict(float)
    for k,v in W.items():
        if k[0]==y and k[1]==p: cls[k[2]]+=v
    out[f"{y}|{p}"]={"c1":tot,"hh":hh,"inc":{"|".join(k):round(v) for k,v in inc.items()},"debt":{"|".join(k):round(v) for k,v in debt.items()},"cls":{k:round(v/hh*100,1) for k,v in cls.items()}}
json.dump(out,open("/tmp/ses.json","w"),ensure_ascii=False)
for k,v in out.items():
    i=v["inc"]; d=v["debt"]
    print(k,"hh",round(v["hh"]),"C1 inc",v["c1"].get("รายได้ทั้งสิ้นต่อเดือน"),"weighted inc",i.get("รายได้ทั้งสิ้นต่อเดือน|รายได้ทั้งสิ้นต่อเดือน|รายได้ทั้งสิ้นต่อเดือน"),
      "| C1 debt",v["c1"].get("หนี้สินเฉลี่ยต่อครัวเรือนทั้งสิ้น"),"weighted debt",d.get("จำนวนหนี้สินเฉลี่ยต่อครัวเรือน|จำนวนหนี้สินเฉลี่ยต่อครัวเรือน"))
    print("   inc src:",{kk.split("|")[2]:vv for kk,vv in i.items() if kk.startswith("รายได้ประจำ|รายได้ที่เป็นตัวเงิน")})
    print("   debt:",{kk.split("|")[1]:vv for kk,vv in d.items() if not kk.startswith("จำนวนหนี้")})
