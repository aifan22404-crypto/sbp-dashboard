import os,openpyxl
from openpyxl.styles import Font,PatternFill,Alignment
R=os.path.expanduser("~/mnt/SBP_Dashboard")
wb=openpyxl.load_workbook(f"{R}/01_raw/06_tourism/MOTS_DomesticTourism_byProvince_monthly_2568R_dl20261009.xlsx",read_only=True,data_only=True)
P={"ปัตตานี":"Pattani","ยะลา":"Yala","นราธิวาส":"Narathiwat","สงขลา":"Songkhla"}
MET=["occupancy_pct","hotel_guests","visitors","visitors_thai","visitors_foreign","revenue_mb","revenue_thai_mb","revenue_foreign_mb"]
out=[]; chk={}
for m,sn in enumerate(wb.sheetnames,1):
    for r in wb[sn].iter_rows(values_only=True):
        idx=next((i for i,x in enumerate(r[:3]) if isinstance(x,str) and x.strip()),None) if r else None
        if idx is not None and r[idx].strip() in P:
            vals=list(r[idx+1:idx+25])
            row=[2568,2025,(m if m<=12 else "Jan-Dec"),P[r[idx].strip()]]
            for i,k in enumerate(MET): row+= [vals[3*i],vals[3*i+1]]
            out.append(row)
print(len(out))
# check: sum of months vs annual sheet
for p in P.values():
    mon=[r for r in out if r[3]==p and r[2]!="Jan-Dec"]; tot=[r for r in out if r[3]==p and r[2]=="Jan-Dec"][0]
    print(p,"visitors sum",round(sum(r[8] for r in mon)),"annual",tot[8],"| revenue sum",round(sum(r[14] for r in mon),1),"annual",tot[14])
cols=["year_BE","year_CE","month","province"]+[f"{k}_{y}" for k in MET for y in ("2568","2567")]
w=openpyxl.load_workbook(f"{R}/02_clean/Tourism_4prov_2564-2568_clean.xlsx")
if "monthly_2568_vs_2567" in w.sheetnames: del w["monthly_2568_vs_2567"]
ws=w.create_sheet("monthly_2568_vs_2567"); ws.append(cols)
for c in ws[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="1F4E78"); c.alignment=Alignment(wrap_text=True)
for r in out: ws.append(r)
ws.freeze_panes="E2"
w["notes"].append(["monthly_2568_vs_2567: MOTS monthly 2568R file (13 sheets: 12 months + Jan–Dec). Each metric shows 2568R and 2567R side by side. Use for seasonality (Ramadan, school holidays, Malaysian holidays)."])
w.save(f"{R}/02_clean/Tourism_4prov_2564-2568_clean.xlsx")
import collections
for p in P.values():
    mon=sorted([r for r in out if r[3]==p and r[2]!="Jan-Dec"],key=lambda r:r[2])
    print(p,"visitors by month:",[round(r[8]/1000) for r in mon],"| foreign:",[round(r[12]/1000,1) for r in mon])
