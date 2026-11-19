import os,openpyxl,json
from openpyxl.styles import Font,PatternFill,Alignment
R=os.path.expanduser("~/mnt/SBP_Dashboard")
wb=openpyxl.load_workbook(f"{R}/01_raw/01_economy_gpp/NESDC_GPP_edition2024_allprov_1995-2024_dl20261008.xlsx",read_only=True)
rows=list(wb['SO'].iter_rows(values_only=True))
PROV={"0308 - SONGKHLA":"Songkhla","0310 - YALA":"Yala","0312 - NARATHIWAT":"Narathiwat","0314 - PATTANI":"Pattani"}
starts=[i for i,r in enumerate(rows) if r[0] and str(r[0]).startswith("GROSS")]
data={}
for k,s in enumerate(starts):
    e=starts[k+1] if k+1<len(starts) else len(rows)
    prov=str(rows[s+1][0]).strip()
    if prov not in PROV: continue
    kind="current" if "CURRENT" in rows[s][0] else "cvm"
    hdr=[r for r in rows[s:s+6] if r[1] and str(r[1]).strip().startswith("19")][0]
    yi={str(y).strip():j for j,y in enumerate(hdr) if y}
    yk=[y for y in yi if y[:4].isdigit() and int(y[:4])>=2019]
    for r in rows[s+3:e]:
        if r[0] and isinstance(r[0],str) and r[yi[yk[0]]] is not None and not r[0].startswith(("Note","GROSS")):
            lab=" ".join(r[0].split())
            for y in yk:
                v=r[yi[y]]
                if isinstance(v,(int,float)): data.setdefault((PROV[prov],kind,lab),{})[y]=v
yrs=sorted({y for d in data.values() for y in d},key=lambda y:int(y[:4]))
print(yrs); labs=[]
for k in data:
    if k[0]=="Pattani" and k[1]=="current": labs.append(k[2])
print(labs)
json.dump({"|".join(k):v for k,v in data.items()},open("/tmp/gpp.json","w"))
