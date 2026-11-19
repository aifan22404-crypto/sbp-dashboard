import os,json,openpyxl
from openpyxl.styles import Font,PatternFill,Alignment
R=os.path.expanduser("~/mnt/SBP_Dashboard")
D={tuple(k.split("|")):v for k,v in json.load(open("/tmp/gpp.json")).items()}
YK=['2019','2020r','2021r','2022r','2023r','2024p']; Y={k:int(k[:4]) for k in YK}
PS=["Pattani","Yala","Narathiwat","Songkhla"]
TOP=["Agriculture, forestry and fishing","Mining and quarrying","Manufacturing","Electricity, gas, steam and air conditioning supply","Water supply; sewerage, waste management and remediation activities","Construction","Wholesale and retail trade and repair of motor vehicles and motorcycle","Transportation and storage","Accommodation and food service activities","Information and communication","Financial and insurance activities","Real estate activities","Professional, scientific and technical activities","Administrative and support service activities","Public administration and defence; compulsory social security","Education","Human health and social work activities","Arts, entertainment and recreation","Other service activities"]
G="Gross provincial product (GPP)"
def g(p,k,l,y): return D.get((p,k,l),{}).get(y)
summ=[]
for p in PS:
    for i,y in enumerate(YK):
        cur=g(p,"current",G,y); cvm=g(p,"cvm","Gross provincial product (CVMs)",y)
        pc=g(p,"current","GPP Per capita (Baht)",y); pop=g(p,"current","Population (1,000 persons)",y)
        prev=YK[i-1] if i else None
        rg=round((cvm/g(p,"cvm","Gross provincial product (CVMs)",prev)-1)*100,2) if prev and cvm else None
        ng=round((cur/g(p,"current",G,prev)-1)*100,2) if prev and cur else None
        summ.append([Y[y],Y[y]+543,y,p,round(cur,1) if cur else None,round(cvm,1) if cvm else None,rg,ng,round(pc) if pc else None,round(pop,1) if pop else None,"WHOLE province" if p=="Songkhla" else "in scope"])
sec=[]; grow=[]; struct=[]
for p in PS:
    tot24=g(p,"current",G,'2024p'); tot21=g(p,"current",G,'2021r')
    for s in TOP:
        for i,y in enumerate(YK):
            v=g(p,"current",s,y); t=g(p,"current",G,y)
            c=g(p,"cvm",s,y); cp=g(p,"cvm",s,YK[i-1]) if i else None
            sec.append([Y[y],y,p,s,round(v,1) if v is not None else None,round(v/t*100,2) if v is not None and t else None,round((c/cp-1)*100,2) if c and cp else None])
        v21,v24=g(p,"current",s,'2021r'),g(p,"current",s,'2024p'); c21,c24=g(p,"cvm",s,'2021r'),g(p,"cvm",s,'2024p')
        cagr=round(((c24/c21)**(1/3)-1)*100,2) if c21 and c24 and c21>0 else None
        struct.append([p,s,round(v24,1) if v24 else None,round(v24/tot24*100,2) if v24 else None,round(v21/tot21*100,2) if v21 else None,
                       round(v24/tot24*100-v21/tot21*100,2) if v21 and v24 else None,cagr])
wb=openpyxl.Workbook(); wb.remove(wb.active)
H=Font(bold=True,color="FFFFFF"); FL=PatternFill("solid",fgColor="1F4E78")
def sheet(n,c,d):
    w=wb.create_sheet(n); w.append(c)
    for x in w[1]: x.font=H; x.fill=FL; x.alignment=Alignment(wrap_text=True)
    for r in d: w.append(r)
    for i,cc in enumerate(c): w.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=max(10,min(45,len(str(cc))+2))
    w.freeze_panes="A2"
sheet("gpp_summary",["year_CE","year_BE","nesdc_label","province","gpp_current_MB","gpp_real_CVM_MB","real_growth_%","nominal_growth_%","gpp_per_capita_baht","population_1000","scope"],summ)
sheet("sector_structure_2021_vs_2024",["province","sector","value_2024_MB","share_2024_%","share_2021_%","share_change_pp","real_CAGR_2021-2024_%"],sorted(struct,key=lambda r:(r[0],-(r[3] or 0))))
sheet("sectors_long",["year_CE","nesdc_label","province","sector","value_current_MB","share_%","real_growth_%"],sec)
n=wb.create_sheet("notes")
for l in ["Source: NESDC GPP edition 2024 (latest; released 2025-26), file NESDC_GPP_edition2024_allprov_1995-2024_dl20261008.xlsx, sheet SO.",
"Years: 2020r–2023r = revised, 2024p = preliminary (will be revised in next edition). GPP 2025 not yet published.",
"Real growth from chain volume measures (reference year 2002). Chain-volume sectors are NOT additive → use for growth only, not for summing.",
"Songkhla = whole province (incl. Hat Yai, Sadao). NESDC does not publish district GPP → our 4 Songkhla districts cannot be isolated; Songkhla is context only.",
"real_CAGR_2021-2024 uses 2021r as base (COVID year) — high values for accommodation/food reflect recovery, not new structural growth (see RF06).",
"Per capita uses NESDC population estimates (not registration)."]: n.append([l])
n.column_dimensions["A"].width=150
wb.save(f"{R}/02_clean/GPP_4prov_2019-2024_clean.xlsx")
for r in summ: 
    if r[0]>=2021: print(r[:10])
for p in PS:
    top=sorted([r for r in struct if r[0]==p],key=lambda r:-(r[3] or 0))[:4]
    fast=sorted([r for r in struct if r[0]==p and r[6] is not None and (r[3] or 0)>=1],key=lambda r:-r[6])[:3]
    print(p,"TOP",[(r[1][:22],r[3]) for r in top],"FAST",[(r[1][:22],r[6]) for r in fast])
