import csv,collections,os,openpyxl,json
from openpyxl.styles import Font,PatternFill,Alignment
R=os.path.expanduser("~/mnt/SBP_Dashboard"); D=f"{R}/01_raw/05_fisheries/DoF"
EN={"ปัตตานี":"Pattani","ยะลา":"Yala","นราธิวาส":"Narathiwat","สงขลา":"Songkhla"}
f=lambda x: float(str(x).replace(",","") or 0) if str(x).strip() not in ("","-") else 0.0
def rd(fn): return list(csv.DictReader(open(f"{D}/{fn}",encoding="utf-8-sig",newline="")))
land=collections.defaultdict(lambda:[0.0,0.0]); sp=collections.defaultdict(lambda:[0.0,0.0]); nat=collections.defaultdict(lambda:[0.0,0.0])
for y in range(2564,2568):
    for r in rd(f"DoF_landing_{y}_dl20261016.csv"):
        p=r["จังหวัด"].strip(); t=f(r["ปริมาณ (ตัน)"]); v=f(r["มูลค่า (พันบาท)"])
        nat[y][0]+=t; nat[y][1]+=v
        if p in EN:
            land[(y,EN[p],r["ประเภทการทำการประมง"].strip())][0]+=t; land[(y,EN[p],r["ประเภทการทำการประมง"].strip())][1]+=v
            sp[(EN[p],r["ชนิดสัตว์น้ำ"].strip())][0]+=t; sp[(EN[p],r["ชนิดสัตว์น้ำ"].strip())][1]+=v
land_rows=[[k[0],k[0]-543,k[1],k[2],round(v[0]),round(v[1]/1000,1),round(v[1]/v[0],1) if v[0] else None] for k,v in sorted(land.items())]
tot=collections.defaultdict(lambda:[0.0,0.0])
for k,v in land.items(): tot[(k[0],k[1])][0]+=v[0]; tot[(k[0],k[1])][1]+=v[1]
land_tot=[[y,y-543,p,round(v[0]),round(v[1]/1000,1),round(v[0]/nat[y][0]*100,2),round(v[1]/1000/ (nat[y][1]/1000)*100,2)] for (y,p),v in sorted(tot.items())]
top=[]
for p in ["Pattani","Narathiwat","Songkhla","Yala"]:
    s=sorted([(k[1],v) for k,v in sp.items() if k[0]==p],key=lambda x:-x[1][1])[:10]
    for n,v in s: top.append([p,n,round(v[0]),round(v[1]/1000,1)])
aq=collections.defaultdict(lambda:[0.0,0.0,0.0,0.0])
for r in rd("DoF_dofd07_05_0101_04_aquaculture_dl20261016.csv"):
    p=r["จังหวัด"].strip()
    if p in EN and r["ปี"].strip().isdigit() and int(r["ปี"])>=2562:
        k=(int(r["ปี"]),EN[p],r["ประเภทการเพาะเลี้ยง"].strip()); a=aq[k]; a[0]+=f(r["จำนวนฟาร์ม"]); a[1]+=f(r["เนื้อที่เลี้ยง"]); a[2]+=f(r["ปริมาณ"]); a[3]+=f(r["มูลค่า"])
aq_rows=[[k[0],k[1],k[2],*[round(x,1) for x in v]] for k,v in sorted(aq.items())]
ves=collections.defaultdict(float); vesz=collections.defaultdict(float)
for r in rd("DoF_dofd07_05_0101_08_thaivessel_dl20261016.csv"):
    p=r["จังหวัดทีออกใบอนุญาต"].strip()
    if p in EN: ves[(int(r["ปี"]),EN[p])]+=f(r["จำนวนเรือ"]); vesz[(int(r["ปี"]),EN[p],r["ขนาดเรือประมง"].strip())]+=f(r["จำนวนเรือ"])
ves_rows=[[k[0],k[1],int(v)] for k,v in sorted(ves.items())]
vesz_rows=[[k[0],k[1],k[2],int(v)] for k,v in sorted(vesz.items())]
wb=openpyxl.Workbook(); wb.remove(wb.active); H=Font(bold=True,color="FFFFFF"); FL=PatternFill("solid",fgColor="1F4E78")
def sheet(n,c,d):
    w=wb.create_sheet(n); w.append(c)
    for x in w[1]: x.font=H; x.fill=FL; x.alignment=Alignment(wrap_text=True)
    for r in d: w.append(r)
    for i,cc in enumerate(c): w.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=max(10,min(40,len(str(cc))+2))
    w.freeze_panes="A2"
sheet("landings_province_total",["year_BE","year_CE","province","tonnes","value_MB","share_of_national_tonnes_%","share_of_national_value_%"],land_tot)
sheet("landings_by_fishery_type",["year_BE","year_CE","province","fishery_type","tonnes","value_MB","baht_per_kg"],land_rows)
sheet("top10_species_2564-2567",["province","species","tonnes_4yr","value_MB_4yr"],top)
sheet("aquaculture_2562+",["year_BE","province","aquaculture_type","farms","area_rai","output_tonnes","value_thousand_baht_or_as_source"],aq_rows)
sheet("vessels_licensed",["year_BE","province_of_license","vessels"],ves_rows)
sheet("vessels_by_size",["year_BE","province","vessel_size","vessels"],vesz_rows)
n=wb.create_sheet("notes")
for l in ["Source: Department of Fisheries (กรมประมง) open data catalog, downloaded in Thailand 2026-10-16 (14 files). Latest year = 2567; 2568 not yet published.",
"Landings = saltwater catch landed at piers in each province (all piers, not only FMO). Unit tonnes; value converted from thousand baht to million baht.",
"Marine capture files (marine_2564–2567) are by sea area / gear, not by province → kept in raw as national context only.",
"Aquaculture values: unit as in source file (check header in raw before quoting). Vessels = licensed commercial vessels by licensing province.",
"Songkhla = whole province (DoF does not publish district)."]: n.append([l])
n.column_dimensions["A"].width=150
wb.save(f"{R}/02_clean/Fisheries_DoF_4prov_2562-2567_clean.xlsx")
for r in land_tot: print(r)
print([r for r in top if r[0]=="Pattani"][:5])
print(ves_rows[-12:])
