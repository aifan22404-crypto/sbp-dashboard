import json,collections,re,os,openpyxl
from openpyxl.styles import Font,PatternFill,Alignment
R=os.path.expanduser("~/mnt/SBP_Dashboard")
D=json.load(open("/tmp/dbd_firms.json")); F=D["firms"]
SEC=[(1,3,"A","เกษตร ป่าไม้ ประมง"),(5,9,"B","เหมืองแร่"),(10,33,"C","การผลิต (อุตสาหกรรม)"),(35,35,"D","ไฟฟ้า ก๊าซ"),(36,39,"E","น้ำ จัดการของเสีย"),(41,43,"F","ก่อสร้าง"),(45,47,"G","ค้าส่ง ค้าปลีก ซ่อมยานยนต์"),(49,53,"H","ขนส่ง คลังสินค้า"),(55,56,"I","ที่พักและอาหาร"),(58,63,"J","ข้อมูลข่าวสารและสื่อสาร"),(64,66,"K","การเงิน ประกันภัย"),(68,68,"L","อสังหาริมทรัพย์"),(69,75,"M","วิชาชีพ วิทยาศาสตร์ เทคนิค"),(77,82,"N","บริหารและบริการสนับสนุน"),(84,84,"O","บริหารราชการ"),(85,85,"P","การศึกษา"),(86,88,"Q","สุขภาพ สังคมสงเคราะห์"),(90,93,"R","ศิลปะ บันเทิง นันทนาการ"),(94,96,"S","บริการอื่น ๆ"),(97,98,"T","ครัวเรือน"),(99,99,"U","องค์การระหว่างประเทศ")]
def sec(code):
    m=re.match(r"\d{2}",code or "")
    if not m: return ("?","ไม่ระบุ")
    d=int(m.group())
    for a,b,s,n in SEC:
        if a<=d<=b: return (s,n)
    return ("?","ไม่ระบุ")
EN={"ปัตตานี":"Pattani","ยะลา":"Yala","นราธิวาส":"Narathiwat","สงขลา":"Songkhla"}
SK4={"จะนะ","เทพา","สะบ้าย้อย","นาทวี"}
def amp(a): return re.sub(r"^(อ\.|อำเภอ)\s*","",a or "").strip()
def capf(x):
    try: return float(str(x).replace(",",""))
    except: return 0.0
def areas(r):
    out=[EN[r["จังหวัด"]]]
    if r["จังหวัด"]=="สงขลา" and amp(r.get("อำเภอ")) in SK4: out.append("Songkhla-4districts")
    if r["จังหวัด"]!="สงขลา": out.append("3 provinces (PTN+YLA+NWT)")
    return out
cnt=collections.defaultdict(lambda:[0,0.0,0])  # new, capital, dissolved
divdesc=collections.defaultdict(collections.Counter); divc=collections.defaultdict(lambda:[0,0])
dist=collections.defaultdict(lambda:[0,0])
for kind in ("new","dis"):
    for r in F[kind]:
        y=int(r["_ym"][:4])+543; s=sec(r.get("รหัสวัตถุประสงค์"))
        for a in areas(r):
            k=(y,a,s[0],s[1])
            if kind=="new": cnt[k][0]+=1; cnt[k][1]+=capf(r.get("ทุนจดทะเบียน"))
            else: cnt[k][2]+=1
        div=(r.get("รหัสวัตถุประสงค์") or "")[:2]
        if r["จังหวัด"]!="สงขลา" or amp(r.get("อำเภอ")) in SK4:
            divdesc[div][(r.get("วัตถุประสงค์") or "")[:60]]+=1
            divc[div][0 if kind=="new" else 1]+=1
        dist[(EN[r["จังหวัด"]],amp(r.get("อำเภอ")),kind)][0]+=1
sec_rows=[[k[0],k[0]-543,k[1],k[2],k[3],v[0],round(v[1]/1e6,2),v[2],v[0]-v[2]] for k,v in sorted(cnt.items())]
# totals per area-year for shares
tot=collections.Counter()
for r in sec_rows: tot[(r[0],r[2])]+=r[5]
for r in sec_rows: r.append(round(r[5]/tot[(r[0],r[2])]*100,1) if tot[(r[0],r[2])] else None)
# 4-year summary by section for 3 provinces + SK4
summ=collections.defaultdict(lambda:[0,0,0.0])
for r in sec_rows:
    if r[2] in ("3 provinces (PTN+YLA+NWT)","Songkhla-4districts"): k=(r[2],r[3],r[4]); summ[k][0]+=r[5]; summ[k][1]+=r[7]; summ[k][2]+=r[6]
summ_rows=sorted([[k[0],k[1],k[2],v[0],v[1],v[0]-v[1],round(v[1]/v[0]*100,1) if v[0] else None,round(v[2],1)] for k,v in summ.items()],key=lambda r:(r[0],-r[3]))
div_rows=sorted([[d,divdesc[d].most_common(1)[0][0] if divdesc[d] else "",v[0],v[1],v[0]-v[1]] for d,v in divc.items()],key=lambda r:-r[2])[:40]
dist_rows=[]
keys={(a,b) for (a,b,_) in dist}
for a,b in sorted(keys): dist_rows.append([a,b,"IN SCOPE" if a!="Songkhla" or b in SK4 else "context",dist[(a,b,"new")][0],dist[(a,b,"dis")][0]])
wb=openpyxl.Workbook(); wb.remove(wb.active); H=Font(bold=True,color="FFFFFF"); FL=PatternFill("solid",fgColor="1F4E78")
def sheet(n,c,d):
    w=wb.create_sheet(n); w.append(c)
    for x in w[1]: x.font=H; x.fill=FL; x.alignment=Alignment(wrap_text=True)
    for r in d: w.append(r)
    for i,cc in enumerate(c): w.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=max(10,min(46,len(str(cc))+2))
    w.freeze_panes="A2"
sheet("summary_2565-2568_by_section",["area","TSIC_section","section_th","new_firms_4yr","dissolved_4yr","net","dissolved_per_100_new","new_capital_MB_4yr"],summ_rows)
sheet("by_year_area_section",["year_BE","year_CE","area","TSIC_section","section_th","new_firms","new_capital_MB","dissolved","net","share_of_new_%"],sec_rows)
sheet("top40_divisions_inscope",["TSIC_2digit","example_description","new_firms_4yr","dissolved_4yr","net"],div_rows)
sheet("by_district",["province","district","scope","new_firms_2565-2568","dissolved_2565-2568"],dist_rows)
n=wb.create_sheet("notes")
for l in ["Source: DBD monthly firm lists (รายชื่อนิติบุคคลจดทะเบียนตั้งใหม่ / เลิกกิจการ) Jan 2565–Dec 2568, openapi.dbd.go.th via gdcatalog; downloaded in Thailand 2026-10-16 (96 files, 440,862 firm rows nationwide).",
"Business type = TSIC code in 'รหัสวัตถุประสงค์' (stated objective at registration) → grouped to TSIC sections A–U by first 2 digits. Stated objective ≠ actual activity.",
"Songkhla-4districts = Chana, Thepha, Saba Yoi, Na Thawi (by head-office district). First time we have DBD data at our exact Songkhla scope.",
"Cleaning: late-2567/2568 files pad with blank rows and use 'จ.' prefix on province names → blanks removed, prefix stripped. After cleaning, firm-list totals equal DBD's aggregate statistics in every month except April 2566 (see RF02) and Jan 2565 Songkhla (110 vs 111).",
"Dissolution date = registered dissolution, not the day trading stopped. 2564 (2021) firm lists not published in this series."]: n.append([l])
n.column_dimensions["A"].width=150
wb.save(f"{R}/02_clean/DBD_firms_by_business_type_2565-2568_clean.xlsx")
json.dump({"summ":summ_rows,"div":div_rows[:15]},open("/tmp/tsic.json","w"),ensure_ascii=False)
for r in summ_rows: 
    if r[3]>=15 or r[4]>=15: print(r)
print(); [print(r) for r in div_rows[:15]]
