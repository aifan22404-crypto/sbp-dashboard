import os, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
R=os.path.expanduser("~/mnt/SBP_Dashboard/05_red_flags")
cols=["ID","Found","Level","Type","Dataset","Province","Period","What we see (numbers)","Why it looks wrong","Innocent explanations","Not-innocent possibility","How to confirm / refute","Impact on our analysis (rule for now)","Status","Say it in public as"]
F=[
["RF01","2026-10-09","🔴 High","A. Method change","NSO LFS","All 4 (biggest NWT, YLA)","2564Q4→2565Q1",
 "Pop 15+ Narathiwat 523,022 → 580,949 (+11.1%) in ONE quarter; employed 308,203 → 348,369 (+40,166). Yala +11.7%, Pattani +7.9%, Songkhla +4.7%. All provinces jump in the same quarter.",
 "Real population cannot grow 11% in 3 months.",
 "NSO switched to new population projections for survey weights from 2565 (very likely). It's a reweighting, not real people.",
 "None likely — pattern is simultaneous everywhere.",
 "Find the NSO note on the weight change (LFS 2565 methodology page). Compare 2564Q4 published under old vs new weights if NSO re-released.",
 "Compare LEVELS (persons employed) only from 2565 onward, or rescale 2564. RATES (unemployment %, LFP %) are OK across the break. Never say 'Narathiwat gained 40,000 jobs in 2565'.",
 "Explained (likely) — confirm NSO note","'Survey weights changed in 2022, so head-counts jump; we use rates.'"],
["RF02","2026-10-09","🔴 High","C. Outlier","DBD new juristic persons","Narathiwat (+12 other provinces)","April 2566",
 "Narathiwat: 119 new firms in Apr-2566 vs a normal month of ~13; registered capital only 13.5 MB → 0.11 MB per firm (normal ~1.7). Same month, 13 of 77 provinces had ≥2× their normal: Kalasin 511 (normal 25.5), Surin 338 (31.5), Chumphon 101 (18.5), Trat 48, Satun 33, Ranong 35. All with tiny capital per firm.",
 "A burst of very small firms in one month, in a cluster of mostly rural provinces at once. This drives Narathiwat's whole '2566 = 270 new firms' spike.",
 "A government programme that required groups to register as juristic persons; a DBD batch upload/back-log entered in one month; registration campaign.",
 "Mass-formed small partnerships created for a purpose (e.g. to qualify for public contracts or programme money) — would show as same addresses, same partners, same business type.",
 "DBD DataWarehouse: list firms registered in Narathiwat Apr-2566 → business type (TSIC), capital, address, partners. If many share address/partners → strong signal. Then match names against public procurement winners (ACT Ai / e-GP).",
 "Do NOT call 2566 a Narathiwat business boom. Show 2566 with and without April. Use 2564–2568 average for trend.",
 "Open — needs firm-level list","'An unusual one-month burst of tiny firms, seen in 13 provinces at once; we're checking what it was.'"],
["RF03","2026-10-09","🟡 Medium","D. Money vs outcome","BOI approvals","Pattani, Narathiwat","2568",
 "Pattani 3 projects 3,182.35 MB → 100 jobs; Narathiwat 3 projects 2,289.62 MB → 22 jobs (~104 MB per job). SBP-policy measure 2568: 8 projects 5,883.96 MB, ALL solar, jobs not stated in report.",
 "Largest 'investment' year for the 3 provinces, but almost no local employment. Tax privileges are granted for it.",
 "Solar farms are capital-heavy and need few staff — normal for the sector. Approval ≠ construction.",
 "Projects approved for privileges but never built, or land/benefit flowing out of the area while headline 'investment' is used as a success claim.",
 "Get company names from BOI promoted-project search; check ERC/PEA power purchase licences; check if built (satellite imagery / news). Compare approved vs 'certificate issued' vs 'started operation'.",
 "Always show jobs next to money. Report 2568 with and without solar.",
 "Open","'Big money on paper, very few jobs — mostly solar farms. Approved doesn't mean built.'"],
["RF04","2026-10-09","🟡 Medium","A. Definition","BOI approvals","All","2564–2569",
 "All BOI numbers are APPROVALS for promotion, not money actually spent.",
 "Readers will read 'investment' as real spending.",
 "Standard BOI reporting.","—",
 "Collect BOI 'issued certificates' and 'started operations' for the same provinces.",
 "Label every chart 'approved'.","Explained","'Approved investment (not yet spent).'"],
["RF05","2026-10-09","🟢 Low","B. Source inconsistency","BOI 2569 H1 report","Trang (out of scope)","2569 Jan–Jun",
 "Trang shown as 525.44 MB in the table but 524.44 MB in the text of the same report.",
 "Same report, two numbers.","Typo.","—","Use table value (it sums to the total 1,337.17).",
 "Shows BOI summaries contain typos → always cross-check table vs text (done for our 4 provinces: all match).","Explained","—"],
["RF06","2026-10-09","🟡 Medium","C. Outlier (base effect)","NESDC GPP","Yala","2565–2566 (2022–2023)",
 "Yala real growth +12.4% in 2023 — but agriculture fell -22.0% in 2022 then +44.9% in 2023. 2021→2023 total = +1.7% only. Accommodation & food +93% (2022) and +92% (2023).",
 "Headline growth looks like a boom.",
 "Rubber/fruit output swing (weather, prices) + tourism rebound after COVID.",
 "—",
 "Compare with OAE rubber/fruit output for Yala 2565–2566 (we have it).",
 "Never quote one-year growth alone; show 2-year or 5-year average. 2024 is preliminary (p).",
 "Explained (mostly)","'2023 jump was a rebound from a bad 2022, not a boom.'"],
["RF07","2026-10-09","🟡 Medium","D. Watch (budget check later)","NESDC GPP","Yala, Narathiwat","2566 (2023)",
 "Construction output (real) fell -22.9% in Yala and -20.9% in Narathiwat in 2023, while Pattani rose +27.1%.",
 "Construction in the deep south is driven heavily by public budgets — a sharp fall needs a reason.",
 "Budget delay (FY2567 budget was approved late, in 2024), end of big projects, election year 2023.",
 "If budgets were disbursed but construction output fell, money may not be reaching real works.",
 "When SBPAC / provincial budget data comes in: compare disbursed capital budget vs construction GPP by province and year.",
 "Hold — revisit with budget data (P2).","Open — waiting for budget data","—"],
["RF08","2026-10-09","🟢 Low","C. Outlier","NESDC GPP","Yala","2565–2567",
 "'Administrative & support services' 301 → 113 → 78 MB (-62.5%, -31.3%).",
 "Sector shrinks by 3/4 in two years.","Reclassification of a firm/agency; small base.","—",
 "Check previous NESDC edition (in _superseded) for the same rows.","Ignore for headline analysis.","Open (low priority)","—"],
["RF09","2026-10-09","🟡 Medium","C. Outlier","MOTS tourism","Yala","2566–2568",
 "Revenue 5,848 → 8,314 MB; hotel occupancy 80.93% in 2568 — higher than Songkhla (69.42%). Rooms ~3,500.",
 "Occupancy above major tourist provinces is unusual for a conflict-affected province.",
 "Betong boom (airport, Malaysian visitors, small room stock) can push occupancy high.",
 "MOTS provincial numbers are survey-based estimates; could be over-estimated.",
 "Cross-check: Malaysian arrivals at Betong immigration checkpoint, Betong airport passengers, provincial hotel registry (room count).",
 "Use as indicative; show it's an estimate.","Open","'Official estimates show Yala hotels ~81% full — we're cross-checking.'"],
["RF10","2026-10-09","🔴 High","E. Preliminary data","OAE agriculture","All 4","2568",
 "Longkong 2568 = 0 in every province; durian & other fruit collapse in 2568.",
 "Impossible values.","2568 data still preliminary/incomplete.","—",
 "Re-download from OAE after final release.","Use fruit data only to 2567.","Explained","—"],
["RF11","2026-10-09","🟢 Low","C. Outlier","FMO fishery ports","Narathiwat","FY2568",
 "88.5 tonnes, 11.9 MB → 134.69 baht/kg (other years 38–57).",
 "Price per kg more than doubles.","Tiny volume + high-value species; or unit typo.","—",
 "Read FY2568 quarterly tables for Narathiwat port.","Don't use Narathiwat FMO port for price trend.","Open (low)","—"],
["RF12","2026-10-09","🟢 Low","C. Outlier","FMO fishery ports","Pattani","FY2567",
 "Pattani 46.15 baht/kg vs 57.44 (FY65) and 58.55 (FY68).","Dip in value per kg.","Species mix / price cycle.","—",
 "Quarterly species breakdown if available.","Report volume and value separately.","Open (low)","—"],
["RF13","2026-10-09","🟢 Low","B. Revision","MOC border trade","Songkhla (Ban Prakob)","2566",
 "2566 total 1,915.24 MB in Dec-2566 report, revised to 1,891.59 MB in Dec-2567 report (-23.65 MB).",
 "Number changed after publication.","Normal revision.","—","—","Use latest report.","Explained","—"],
["RF14","2026-10-09","🟢 Low","E. Missing","MOC border trade","Yala (Betong)","Nov 2564",
 "Reported as 0.00 in two reports.","Zero trade for a whole month is implausible.","COVID border closure / data gap.","—",
 "Check Betong customs house.","Treated as missing; 2564 marked incomplete.","Explained","—"],
["RF15","2026-10-09","🟡 Medium","C. Outlier","NSO LFS","Narathiwat","2564–2569",
 "Unemployment 9.36% (2564Q1) → 1.90% (2569Q2), steady fall; was ~3× other provinces.",
 "Very large improvement.","Post-COVID recovery; Malaysia border reopening (cross-border workers count as employed).",
 "Small provincial sample → noisy.",
 "Compare with Dept. of Employment registered job-seekers and SSO insured persons for Narathiwat.",
 "Use 4-quarter averages, not single quarters.","Open","—"],
]
wb=openpyxl.Workbook(); ws=wb.active; ws.title="red_flags"
ws.append(cols)
for c in ws[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="7B1E1E"); c.alignment=Alignment(wrap_text=True,vertical="center")
fill={"🔴":"F8CBAD","🟡":"FFE699","🟢":"C6E0B4"}
for f in F:
    ws.append(f)
    for c in ws[ws.max_row]: c.alignment=Alignment(wrap_text=True,vertical="top")
    ws.cell(ws.max_row,3).fill=PatternFill("solid",fgColor=fill[f[2][0]])
w=[6,11,11,18,16,16,13,55,32,38,38,45,42,20,38]
for i,x in enumerate(w): ws.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width=x
ws.freeze_panes="C2"
dv=DataValidation(type="list",formula1='"🔴 High,🟡 Medium,🟢 Low"'); ws.add_data_validation(dv); dv.add("C2:C500")
r=wb.create_sheet("rules")
for line in [
["กติกาของแฟ้มนี้ (Red Flags)"],
["1. Red flag = 'คำถาม' ไม่ใช่ 'คำตัดสิน' ข้อมูลไม่โกหก แต่ข้อมูลหลอกเราได้ถ้าอ่านผิด (เปลี่ยนวิธีเก็บ, ข้อมูลเบื้องต้น, พิมพ์ผิด, ฐานต่ำ)"],
["2. ทุกแถวต้องมี: ตัวเลขจริง + คำอธิบายแบบบริสุทธิ์ใจ + วิธีพิสูจน์ ถ้ายังไม่ได้พิสูจน์ ห้ามสรุปว่าโกง"],
["3. 🔴 = ถ้าไม่จัดการจะทำให้ผลวิเคราะห์ผิด หรือมีสัญญาณที่ต้องหาเอกสารเพิ่ม / 🟡 = ต้องเช็คก่อนใช้ / 🟢 = จดไว้ ไม่กระทบภาพใหญ่"],
["4. Type: A = วิธีการเปลี่ยน, B = ตัวเลขไม่ตรงกันในแหล่งเดียวกัน, C = ค่าผิดปกติ, D = เงินกับผลลัพธ์ไม่สอดคล้อง (หมวดจับตา), E = ข้อมูลหาย/เบื้องต้น"],
["5. ห้ามเอ่ยชื่อบุคคล/บริษัท/หน่วยงานว่า 'โกง' ในที่สาธารณะ (TikTok/รายงาน) จนกว่าจะมีเอกสารยืนยัน — กฎหมายหมิ่นประมาทและ พ.ร.บ.คอมพิวเตอร์ของไทยแรงมาก ใช้คอลัมน์ 'Say it in public as'"],
["6. หลักฐาน (screenshot, PDF, ลิงก์) เก็บใน 05_red_flags/evidence/ ตั้งชื่อตาม ID เช่น RF02_DBD_Narathiwat_Apr2566_firmlist.pdf"],
["7. แหล่งที่ใช้ตรวจเชิงลึกในอนาคต: DBD DataWarehouse (รายชื่อบริษัท), ACT Ai / e-GP (การจัดซื้อจัดจ้างภาครัฐ), BOI promoted-project search, ERC (ใบอนุญาตโรงไฟฟ้า), งบ ศอ.บต./จังหวัด"],
]: r.append(line)
r.column_dimensions["A"].width=150
wb.save(f"{R}/red_flags_log.xlsx")
open(f"{R}/README.md","w",encoding="utf-8").write("""# 05_red_flags — ข้อมูลที่น่าสงสัย

- `red_flags_log.xlsx` — ทุกจุดที่ตัวเลขดูผิดปกติ (ชีต red_flags) + กติกา (ชีต rules)
- `evidence/` — หลักฐานประกอบ ตั้งชื่อไฟล์ขึ้นต้นด้วย ID เช่น `RF02_...`

Red flag = คำถาม ไม่ใช่คำตัดสิน ต้องพิสูจน์ก่อนสรุป และห้ามกล่าวหาใครในที่สาธารณะโดยไม่มีเอกสาร
""")
print("ok",len(F))
