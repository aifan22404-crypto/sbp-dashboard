import os,openpyxl
R=os.path.expanduser("~/mnt/SBP_Dashboard"); fn=f"{R}/02_clean/DBD_juristic_4prov_2564-2568_clean.xlsx"
wb=openpyxl.load_workbook(fn)
fix={"Narathiwat":10,"Yala":11,"Songkhla":72}
m=wb["monthly"]; h=[c.value for c in m[1]]
for r in m.iter_rows(min_row=2):
    v={h[i]:c for i,c in enumerate(r)}
    if v["year_be"].value==2566 and v["month"].value==4 and v["province"].value in fix:
        old=v["new_reg"].value; v["new_reg"].value=fix[v["province"].value]; v["net_formation"].value=v["new_reg"].value-(v["dissolved"].value or 0)
        print("fixed",v["province"].value,old,"->",v["new_reg"].value)
a=wb["annual"]; h=[c.value for c in a[1]]
for r in a.iter_rows(min_row=2):
    v={h[i]:c for i,c in enumerate(r)}
    if v["year_be"].value==2566 and v["province"].value in fix:
        delta={"Narathiwat":10-119,"Yala":11-21,"Songkhla":72-19}[v["province"].value]
        v["new_reg"].value+=delta; v["net_formation"].value=v["new_reg"].value-v["dissolved"].value
        if v["new_capital_mb"].value and v["new_reg"].value: v["avg_capital_per_new_firm_mb"].value=round(v["new_capital_mb"].value/v["new_reg"].value,2)
        print("annual",v["province"].value,v["new_reg"].value)
wb["notes"].append(["CORRECTED 2026-10-16: April 2566 new registrations for Narathiwat (119→10), Yala (21→11), Songkhla (19→72) replaced with counts from DBD's firm-level list. DBD's aggregate statistics file had province labels scrambled for that month (RF02). Capital for that month not corrected (firm list capital available in DBD_firms_by_business_type file)."])
wb.save(fn)
