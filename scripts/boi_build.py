import json, os, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
R=os.path.expanduser("~/mnt/SBP_Dashboard")
D=f"{R}/01_raw/03_investment_boi"
ren={"th-knoptuwz0269.pdf":(2564,"Jan-Dec"),"th-bcdehkmnsu28.pdf":(2565,"Jan-Dec"),"th-cghilnpqrt29.pdf":(2566,"Jan-Dec"),
     "th-efhlsuwyz089.pdf":(2567,"Jan-Dec"),"th-glmqtvy13456.pdf":(2568,"Jan-Dec"),"th-degituvwxy06.pdf":(2569,"Jan-Jun")}
newname={k:f"BOI_Region5_LowerSouth7_InvestmentSummary_{y}_{p}_dl20261009.pdf" for k,(y,p) in ren.items()}
for k,v in newname.items():
    if os.path.exists(f"{D}/{k}"): os.rename(f"{D}/{k}",f"{D}/{v}")
with open(f"{D}/_DOWNLOAD_LINKS.txt","a",encoding="utf-8") as f:
    f.write("\n## RENAMED 2026-10-09 (original server name -> project name)\n")
    for k,v in newname.items(): f.write(f"# {k} -> {v}\n")
res=json.load(open("/tmp/boi.json"))["res"]
EN={"Songkhla":"สงขลา","Pattani":"ปัตตานี","Yala":"ยะลา","Narathiwat":"นราธิวาส"}
tot={2564:(39,18001.95,4990),2565:(40,7359.40,1467),2566:(53,9065.69,3219),2567:(40,4685.27,2049),2568:(42,12180.96,1370),2569:(15,1337.17,355)}
acts={
(2564,"Songkhla"):"ถุงมือยางทางการแพทย์ (5) ถุงมือยางธรรมชาติ น้ำยางข้น (4) ยางแท่ง ยางผสม ไฟฟ้าแสงอาทิตย์ (2) เนื้อปลาบด (2) สัตว์น้ำ/อาหารแช่แข็ง สลักภัณฑ์ ระบบฟาร์มอัจฉริยะโรงเรือนไก่ อาหารทะเลแปรรูปแช่แข็ง โปรตีนแอนติเจน สารสกัดจากเซรั่ม HIGH FRUCTOSE SYRUP",
(2564,"Pattani"):"ผลิตภัณฑ์จากเนื้อปลาแปรรูป",(2564,"Yala"):"น้ำยางข้น",(2564,"Narathiwat"):"น้ำมันปาล์มดิบ และน้ำยางข้น",
(2565,"Songkhla"):"ไฟฟ้าแสงอาทิตย์ (8) ไฟฟ้าชีวมวล น้ำยางข้น (3) ถุงมือยางตรวจโรค ไม้วีเนียร์ ผลิตภัณฑ์/ชิ้นส่วนโลหะ อาหารบรรจุภาชนะผนึก อาหารพร้อมรับประทาน เนื้อสัตว์แปรรูป แผ่นใยไม้อัด MDF เครื่องจักรอุตสาหกรรม ถ้วยกระดาษ ชิ้นส่วนเหล็ก",
(2565,"Pattani"):"ไฟฟ้าแสงอาทิตย์ (2)",(2565,"Yala"):"ไฟฟ้าแสงอาทิตย์ และยางแท่ง",(2565,"Narathiwat"):"ยางแท่ง",
(2566,"Songkhla"):"น้ำยางข้น (2) เส้นหมี่ ไฟฟ้าแสงอาทิตย์ (9) ไฟฟ้าจากขยะ ยาสีฟัน อาหารสัตว์ (2) ยางแท่ง ถุงยางอนามัย ท่อไฮดรอลิค เครื่องจักรอุตสาหกรรม อาหารบรรจุภาชนะผนึก อาหารแช่แข็ง อาหารทะเลแปรรูป (2) กล่องกระดาษ (2) ศูนย์ดูแลผู้สูงอายุ เซรามิกส์ Rubber Seal ถุงมือยางทางการแพทย์ Formalin",
(2566,"Pattani"):"ยางแท่ง (2) พืชผักผลไม้แช่แข็ง โรงพยาบาล น้ำมันปาล์ม",(2566,"Yala"):"โรงพยาบาล ไฟฟ้าจากขยะ ไฟฟ้าแสงอาทิตย์",(2566,"Narathiwat"):"ไฟฟ้าแสงอาทิตย์ (2)",
(2567,"Songkhla"):"ไฟฟ้าแสงอาทิตย์ (10) ไฟฟ้าชีวมวล (2) อาหารสัตว์เลี้ยง (2) อาหารทะเลกึ่งสำเร็จรูปแช่แข็ง (2) สัตว์น้ำแช่แข็ง (2) ไฟฟ้าจากขยะ ห้องเย็น ส่วนผสมอาหารสัตว์ ผลิตภัณฑ์กระดาษ เชื้อเพลิงจากขยะ ยางแท่ง แผ่นยิปซั่ม",
(2567,"Pattani"):"ไฟฟ้าแสงอาทิตย์",(2567,"Yala"):"ไฟฟ้าชีวมวล (2) ไฟฟ้าแสงอาทิตย์ เครื่องอัดอากาศ",(2567,"Narathiwat"):"— (ไม่มีโครงการ)",
(2568,"Songkhla"):"ไฟฟ้าแสงอาทิตย์ (10) ผลิตภัณฑ์กระดาษ (2) ถุงมือยางตรวจโรค ชิ้นส่วนโลหะขึ้นรูป น้ำยางข้น (3) ยางแท่ง/ยางผสม ผงถ่าน/ถ่านชิ้น ผลิตภัณฑ์โลหะ เนื้อปลาบดแปรรูป ห้องเย็น น้ำมันปลา/น้ำนึ่งปลาเข้มข้น พุดดิ้ง ถุงยางอนามัย ชิ้นส่วนเครื่องสับไม้ ปะเก็นโลหะ ขนถ่ายตู้คอนเทนเนอร์",
(2568,"Pattani"):"ไฟฟ้าแสงอาทิตย์ (2) และการผลิตน้ำกะทิ",(2568,"Yala"):"ไฟฟ้าแสงอาทิตย์ และผลไม้แช่แข็ง",(2568,"Narathiwat"):"ไฟฟ้าแสงอาทิตย์ (3)",
(2569,"Songkhla"):"ไฟฟ้าแสงอาทิตย์ (3) กิจการสนับสนุนการค้าและการลงทุน น้ำมันปาล์มดิบ ถุงมือยางทางการแพทย์",
(2569,"Pattani"):"— (ไม่มีโครงการ)",(2569,"Yala"):"การสำรวจแร่",(2569,"Narathiwat"):"— (ไม่มีโครงการ)"}
meas=[(2564,"Jan-Dec","SBP industrial development policy",6,504.45,106,"น้ำยางข้น (2) ไฟฟ้าแสงอาทิตย์ (2) รากไม้สับ น้ำมันปาล์มดิบ"),
(2564,"Jan-Dec","Model city (มั่นคง มั่งคั่ง ยั่งยืน)",2,338.82,574,"น้ำยางข้น และผลิตภัณฑ์จากเนื้อปลาแปรรูป"),
(2565,"Jan-Dec","SBP industrial development policy",8,402.35,159,"ไฟฟ้าแสงอาทิตย์ (5) ยางแท่ง น้ำยางข้น ถ้วยกระดาษ"),
(2565,"Jan-Dec","Model city (มั่นคง มั่งคั่ง ยั่งยืน)",1,55.66,2,"ไฟฟ้าแสงอาทิตย์"),
(2566,"Jan-Dec","SBP industrial development policy",7,2574.55,197,"ไฟฟ้าแสงอาทิตย์ (4) น้ำมันปาล์ม ไฟฟ้าจากขยะ วงจรสื่อสัญญาณใต้น้ำระหว่างประเทศ"),
(2566,"Jan-Dec","Model city (มั่นคง มั่งคั่ง ยั่งยืน)",3,137.14,31,"พืชผักผลไม้แช่แข็ง น้ำมันปาล์ม ไฟฟ้าแสงอาทิตย์"),
(2567,"Jan-Dec","SBP investment measure",9,1684.03,442,"ไฟฟ้าแสงอาทิตย์ (3) ไฟฟ้าชีวมวล (4) ผลิตภัณฑ์กระดาษ อาหารสัตว์"),
(2567,"Jan-Dec","Model city (SBP)",2,52.30,7,"ไฟฟ้าแสงอาทิตย์ และเครื่องอัดอากาศ"),
(2568,"Jan-Dec","SBP industrial development policy",8,5883.96,None,"ไฟฟ้าแสงอาทิตย์ (8) — รายงานไม่ระบุการจ้างงาน"),
(2568,"Jan-Dec","Model city (SBP)",4,2324.82,179,"ไฟฟ้าแสงอาทิตย์ ชิ้นส่วนโลหะขึ้นรูป น้ำกะทิ ผลไม้แช่แข็ง"),
(2569,"Jan-Jun","SBP investment measure",1,5.00,11,"การสำรวจแร่")]
H=Font(bold=True,color="FFFFFF"); F=PatternFill("solid",fgColor="1F4E78")
def head(ws,cols):
    ws.append(cols)
    for c in ws[1]: c.font=H; c.fill=F; c.alignment=Alignment(wrap_text=True,vertical="center")
wb=openpyxl.Workbook()
ws=wb.active; ws.title="province_approvals"
head(ws,["year_BE","year_CE","period","province_en","province_th","in_scope_note","projects","investment_MB","jobs","share_of_lower_south_investment_%","lower_south7_projects","lower_south7_investment_MB","lower_south7_jobs","MB_per_job","source_file"])
for y,p,pe,n,mb,j,src in sorted(res,key=lambda r:(r[0],["Songkhla","Pattani","Yala","Narathiwat"].index(r[2]))):
    t=tot[y]; per=p.replace("–","-")
    ws.append([y,y-543,per,pe,EN[pe],"WHOLE province (not 4 districts)" if pe=="Songkhla" else "whole province = in scope",n,mb,j,round(mb/t[1]*100,2),t[0],t[1],t[2],round(mb/j,2) if j else None,newname[src]])
ws2=wb.create_sheet("business_types")
head(ws2,["year_BE","province_en","province_th","projects","approved_activities (count in brackets)"])
cnt={(r[0],r[2]):r[3] for r in res}
for y in range(2564,2570):
    for pe in ["Songkhla","Pattani","Yala","Narathiwat"]:
        ws2.append([y,pe,EN[pe],cnt[(y,pe)],acts[(y,pe)]])
ws3=wb.create_sheet("sbp_special_measures")
head(ws3,["year_BE","period","measure","projects","investment_MB","jobs","activities"])
for m in meas: ws3.append(list(m))
ws4=wb.create_sheet("notes")
for line in [
"Source: BOI Investment Economic Center Region 5 (Songkhla) — hatyai.boi.go.th, 'สรุปภาวะการลงทุนใน 7 จังหวัดภาคใต้ตอนล่าง', downloaded 2026-10-09.",
"Unit: projects APPROVED for promotion (การอนุมัติให้การส่งเสริม), calendar year Jan–Dec. Approved ≠ actually built/operating.",
"Province figures = section 4 'แยกตามที่ตั้งโครงการ' (by project location). Parsed by script; all 24 rows cross-checked with the narrative text.",
"2569 = Jan–Jun only (bonus, partial year) — do not compare to full years.",
"Songkhla = WHOLE province (incl. Hat Yai, Songkhla SEZ/Sadao). Our scope is 4 districts only (Chana, Thepha, Saba Yoi, Na Thawi) — use Songkhla only as regional context.",
"Model-city measure covers Nong Chik (PTN), Betong (YLA), Sungai Kolok (NWT), Chana (SKA) — the only BOI line that touches our Songkhla districts.",
"Special-measure rows overlap province rows (same projects, different view) — never add them together.",
"Measure names change between reports (2564–2566 'policy', 2567+ 'measure'); kept as reported.",
"Caution: large 2568 values in Pattani (3,182 MB) and Narathiwat (2,290 MB) are mostly solar power plants with very few jobs (100 and 22).",
"2568 SBP-policy jobs not stated in the report → blank, not zero.",
]: ws4.append([line])
for w in wb:
    for col in w.columns: w.column_dimensions[col[0].column_letter].width=min(60,max(10,max(len(str(c.value or "")) for c in col)*0.9))
out=f"{R}/02_clean/BOI_4prov_2564-2568_clean.xlsx"; wb.save(out); print("saved",out)
print(os.listdir(D))
