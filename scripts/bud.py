import os,json,csv,collections,openpyxl,re
from openpyxl.styles import Font,PatternFill,Alignment
R=os.path.expanduser("~/mnt/SBP_Dashboard"); D=f"{R}/01_raw/12_sbpac_budget"; B=f"{R}/01_raw/11_reports_bot"
bud=json.load(open("/tmp/bud.json"))
AB={"Songkhla":"SKA","Pattani":"PTN","Yala":"YLA","Narathiwat":"NWT"}
ren={}
for f,p,y,n in bud: ren[f]=f"BB_ProvBudgetAlloc_5types_{AB[p]}_FY{y}_dl20261010.pdf"
ren["untitled.csv"]="CGD_NWT_Disbursement_byUnit_FY2568_dl20261010.csv"
for a,b in ren.items():
    if os.path.exists(f"{D}/{a}"): os.rename(f"{D}/{a}",f"{D}/{b}")
with open(f"{D}/_DOWNLOAD_LINKS.txt","a",encoding="utf-8") as fh:
    fh.write("\n## RENAMED 2026-10-10 (year read from inside each PDF)\n"+"".join(f"# {a} -> {b}\n" for a,b in ren.items())+"# MISSING: Yala FY2569 (04950000) — not in folder.\n")
bren={"Press_SO_Y_64.pdf":"BOT_SRO_AnnualPress_2564","Press_SO_Y_65.pdf":"BOT_SRO_AnnualPress_2565","Press_SO_2566.pdf":"BOT_SRO_AnnualPress_2566","Table_SO_2566.pdf":"BOT_SRO_AnnualTable_2566",
"Press_SO_Y_67.pdf":"BOT_SRO_AnnualPress_2567","Table_SO_Y_67.pdf":"BOT_SRO_AnnualTable_2567","Press_SO_68.pdf":"BOT_SRO_AnnualPress_2568","Table_SO_68.pdf":"BOT_SRO_AnnualTable_2568",
"RG_SR_042_S3_TH_ALL.XLSX":"BOT_RG_SR_042_S3_CustomsTradeSouth_byProduct","RG_SR_039_S3_TH_ALL.XLSX":"BOT_RG_SR_039_S3_PrivateInvestmentIndexSouth","RG_SR_038_S2_TH_ALL.XLSX":"BOT_RG_SR_038_S2_PrivateConsumptionIndexSouth"}
for a,b in bren.items():
    ext=os.path.splitext(a)[1].lower()
    if os.path.exists(f"{B}/{a}"): os.rename(f"{B}/{a}",f"{B}/{b}_dl20261010{ext}")
with open(f"{B}/_DOWNLOAD_LINKS.txt","a",encoding="utf-8") as fh:
    fh.write("\n## RENAMED 2026-10-10\n"+"".join(f"# {a} -> {b}_dl20261010{os.path.splitext(a)[1].lower()}\n" for a,b in bren.items()))
# ---- budget sheet
rows=[]
for f,p,y,n in sorted(bud,key=lambda r:(r[1],r[2])):
    pers,op,inv,sub,oth,tot=n
    rows.append([int(y),int(y)-543,p,round(pers/1e6,2),round(op/1e6,2),round(inv/1e6,2),round(sub/1e6,2),round(oth/1e6,2),round(tot/1e6,2),round(inv/tot*100,1)])
pop={"Pattani":652.6,"Yala":486.3,"Narathiwat":743.0,"Songkhla":1714.0}
for r in rows: r.append(round(r[8]*1e6/(pop[r[2]]*1000)))
# ---- Narathiwat disbursement FY2568
C=list(csv.DictReader(open(f"{D}/CGD_NWT_Disbursement_byUnit_FY2568_dl20261010.csv",encoding="utf-8-sig")))
f2=lambda x: float(str(x).replace(",","") or 0)
sep=[r for r in C if r["เดือน"].strip()=="กันยายน"]
agg=collections.defaultdict(lambda:[0.0,0.0])
for r in sep: k=(r["อำเภอ"].strip(),r["ประเภทรายจ่าย"].strip()); agg[k][0]+=f2(r["งบจัดสรร"]); agg[k][1]+=f2(r["เบิกจ่าย"])
dis=[[k[0],k[1],round(v[0]/1e6,2),round(v[1]/1e6,2),round(v[1]/v[0]*100,1) if v[0] else None] for k,v in sorted(agg.items(),key=lambda x:-x[1][0])]
tot=collections.defaultdict(lambda:[0.0,0.0])
for k,v in agg.items(): tot[k[1]][0]+=v[0]; tot[k[1]][1]+=v[1]
print({k:(round(v[0]/1e6),round(v[1]/1e6),round(v[1]/v[0]*100,1)) for k,v in tot.items()})
# ---- BOT yearly
def yearly(fn):
    ws=openpyxl.load_workbook(f"{B}/{fn}",read_only=True,data_only=True)["Yearly"]; L=list(ws.iter_rows(values_only=True))
    hdr=[str(x).strip() if x is not None else "" for x in L[5]]
    out=[]
    for r in L[6:]:
        if r and r[1] and str(r[0]).strip().isdigit():
            for i,h in enumerate(hdr):
                if h[:4].isdigit() and int(h[:4])>=2562 and r[i] not in (None,""):
                    out.append([int(h[:4]),h,str(r[1]).strip(),float(r[i])])
    return out
cus=yearly("BOT_RG_SR_042_S3_CustomsTradeSouth_byProduct_dl20261010.xlsx")
inv=yearly("BOT_RG_SR_039_S3_PrivateInvestmentIndexSouth_dl20261010.xlsx")
con=yearly("BOT_RG_SR_038_S2_PrivateConsumptionIndexSouth_dl20261010.xlsx")
H=Font(bold=True,color="FFFFFF"); FL=PatternFill("solid",fgColor="1F4E78")
def mk(path,sheets,notes):
    wb=openpyxl.Workbook(); wb.remove(wb.active)
    for n,c,d in sheets:
        w=wb.create_sheet(n); w.append(c)
        for x in w[1]: x.font=H; x.fill=FL; x.alignment=Alignment(wrap_text=True)
        for r in d: w.append(r)
        for i,cc in enumerate(c): w.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=max(10,min(40,len(str(cc))+2))
        w.freeze_panes="A2"
    w=wb.create_sheet("notes")
    for l in notes: w.append([l])
    w.column_dimensions["A"].width=150; wb.save(path)
mk(f"{R}/02_clean/Budget_4prov_FY2564-2569_clean.xlsx",[
 ("allocation_5types",["FY_BE","FY_CE","province","personnel_MB","operating_MB","CAPITAL_investment_MB","subsidies_MB","other_MB","total_MB","capital_share_%","total_baht_per_person"],rows),
 ("NWT_disbursement_FY2568",["district","expense_type","allocated_MB","disbursed_MB_by_Sep2568","disbursed_%"],dis)],
 ["Allocation: Budget Bureau 'รายงานแสดงแผนการจัดสรรงบประมาณไปจังหวัด' (5 budget types), one PDF per province per fiscal year (Oct–Sep). Totals read from the 'รวมทั้งสิ้น' row; checked: 5 types sum to total in every file.",
  "Allocation = PLANNED budget located in the province (by agency HQ location). Not disbursement. Yala is high in operating budget partly because regional agencies (e.g. SBPAC HQ) sit in Yala.",
  "Missing: Yala FY2569 (not downloaded).","Per-person uses NESDC 2024 population (constant) — rough comparison only.",
  "NWT disbursement: Comptroller General / Narathiwat Provincial Treasury, FY2568 monthly cumulative by unit; sheet uses September 2568 (year-end) rows. Only Narathiwat publishes this.",
  "FY2567 budget law was enacted late (2024) → Oct 2023–Apr 2024 spending limited; relevant to RF07."])
mk(f"{R}/02_clean/BOT_South_indicators_2562-2568_clean.xlsx",[
 ("customs_trade_south_yearly",["year_BE","label","item","value_MB_or_tonnes"],cus),
 ("private_investment_index",["year_BE","label","index","value (2561=100)"],inv),
 ("private_consumption_index",["year_BE","label","item","value"],con)],
 ["Bank of Thailand Southern Region Office (สภต.) regional statistics. Region = 14 southern provinces (NOT our 4 provinces only) → context for P1–P4.",
  "Customs table = all southern customs houses by PRODUCT (not by checkpoint). It does NOT replace checkpoint-level border trade (we already have MOC provincial data).",
  "r = revised, p = preliminary. Annual narrative reports 2564–2568 are in 01_raw/11_reports_bot (PDF)."])
json.dump({"rows":rows,"tot":{k:v for k,v in tot.items()},"cus":[c for c in cus if c[2] in ("สินค้าออก","สินค้าเข้า","ยาง","น้ำมันพืชและไขมัน")],"inv":[c for c in inv if c[2]=="ดัชนีการลงทุนภาคเอกชน"]},open("/tmp/bud2.json","w"),ensure_ascii=False)
for r in rows: print(r)
