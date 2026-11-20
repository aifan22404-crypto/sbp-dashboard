"""Build 3 supplementary clean datasets (20 Nov 2026):
 1. District population (DOPA registration via provincial GD Catalogs) 2562-2568
 2. Tourism 2562 pre-COVID baseline (MOTS Domestic Tourism Statistics Q1-Q4 2019, South)
 3. Southern-border integrated-plan budget by agency (Rocket Media Lab compilation of Budget Acts)
"""
import pandas as pd
OUT = 'data/clean/'

# 1. population -------------------------------------------------------------
pop = pd.read_csv('/tmp/district_population_long.csv')
notes = pd.DataFrame({'หมายเหตุ': [
    'ที่มา: ทะเบียนราษฎร (กรมการปกครอง) เผยแพร่ผ่านระบบบัญชีข้อมูลจังหวัด (GD Catalog) ของแต่ละจังหวัด ดาวน์โหลด 2026-11-20',
    'นราธิวาส: narathiwat.gdcatalog.go.th/dataset/12_01 (2563–2568) — สุไหงโก-ลก ปี 2563–2566 ≈ 40,000 แต่ 2567–2568 ≈ 79,000 → ปี 2563–2566 น่าจะนับไม่ครบ (RF25)',
    'ปัตตานี: pattani.gdcatalog.go.th/dataset/population (2564–2566) — ตารางแยก "อำเภอ" (นอกเขตเทศบาล) กับ "ท้องถิ่นเทศบาล" ; รวมเทศบาลเข้าอำเภอที่ตั้งแล้ว (ยอดรวม 2566 = 737,077 ตรงกับยอดจังหวัด)',
    'ปัตตานี: เทศบาลตำบลตันหยง (2,898 คน) จัดเข้าอำเภอยะหริ่งตามข้อมูลเขตเทศบาล — ยังไม่ได้ยืนยันกับกรมการปกครอง',
    'ยะลา: yala.gdcatalog.go.th/dataset/population_01 (2566–2568) ; สงขลา: songkhla.gdcatalog.go.th/dataset/dataset_03_46 (2562–2568)',
    'ตัวหารที่ใช้ในการวิเคราะห์ต่อหัว: ปี 2566 (ปีกลางของช่วง) ยกเว้นนราธิวาสใช้ 2567 เพราะ RF25',
]})

# 2. tourism 2562 baseline ---------------------------------------------------
t = [  # province, year_BE, visitors, thai, foreign, revenue_mb, occupancy_pct, guest_arrivals
    ('Pattani', 2561, 281988, 279643, 2345, 1063.16, 52.37, 246620), ('Pattani', 2562, 282059, 279674, 2385, 1046.10, 51.89, 246854),
    ('Yala', 2561, 680018, 146222, 533796, 3333.56, 51.51, 633926), ('Yala', 2562, 689938, 145191, 544747, 3340.55, 51.93, 640793),
    ('Narathiwat', 2561, 692057, 259853, 432204, 3143.44, 56.66, 626831), ('Narathiwat', 2562, 703810, 262309, 441501, 3170.56, 56.16, 630859),
    ('Songkhla', 2561, 7628813, 4614546, 3014267, 68252.64, 73.54, 5230352), ('Songkhla', 2562, 7503461, 4420558, 3082903, 66365.52, 71.48, 5175422),
]
tour = pd.DataFrame(t, columns=['province', 'year_BE', 'visitors', 'visitors_thai', 'visitors_foreign', 'revenue_mb', 'occupancy_pct', 'guest_arrivals'])
tour['source'] = 'MOTS สถิตินักท่องเที่ยวภายในประเทศ Q1-Q4 ปี 2562 (ภาคใต้) mots.go.th/download/article/article_20201105143022.xlsx, block January-December'
panel = pd.read_csv('data/master_panel/SBP_master_panel_long.csv', encoding='utf-8-sig')
v = panel[(panel.indicator_code == 'tour_visitors') & (panel.year_BE == 2568)].set_index('area')['value']
cmp_ = tour[tour.year_BE == 2562][['province', 'visitors']].copy()
cmp_['visitors_2568'] = cmp_.province.map(v)
cmp_['ratio_2568_to_2562'] = (cmp_.visitors_2568 / cmp_.visitors).round(2)
tnotes = pd.DataFrame({'หมายเหตุ': [
    'ใช้เป็นฐานก่อนโควิดเท่านั้น ไม่รวมใน Master Panel (ปีหลักของโครงการคือ 2564–2568)',
    'ข้อควรระวัง (RF26): ปี 2568 ปัตตานี/ยะลา/นราธิวาส สูงกว่า 2562 ถึง 1.8–2.9 เท่า แต่สงขลาต่ำกว่า 2562 — อาจมีการเปลี่ยนวิธีประมาณการของ กก. ระหว่างช่วง ยังยืนยันไม่ได้',
]})

# 3. budget by agency --------------------------------------------------------
bud = pd.DataFrame([
    ('สำนักนายกรัฐมนตรีและหน่วยงานในสังกัด (รวม กอ.รมน.)', 83402701700, 31.16),
    ('  ในนั้น: กอ.รมน.', 81650975600, 30.50),
    ('กระทรวงมหาดไทย', 49673165300, 18.56),
    ('  ในนั้น: กรมการปกครอง', 26854493000, 10.03),
    ('  ในนั้น: กรมโยธาธิการและผังเมือง', 17338202600, 6.48),
    ('กระทรวงกลาโหม', 42721519700, 15.96),
    ('  ในนั้น: กองทัพบก', 27553455000, 10.29),
    ('  ในนั้น: กองทัพเรือ', 11027655500, 4.12),
    ('  ในนั้น: กองทัพอากาศ', 3049163000, 1.14),
    ('หน่วยงานไม่สังกัดสำนักนายกฯ กระทรวง หรือทบวง', 41806290700, 15.62),
    ('  ในนั้น: ศอ.บต.', 22827233900, 8.53),
    ('  ในนั้น: สำนักงานตำรวจแห่งชาติ', 16724270800, 6.25),
    ('กระทรวงศึกษาธิการ', 26326105900, 9.84),
], columns=['หน่วยงาน', 'งบรวม_2551_2568_บาท', 'สัดส่วน_%'])
security = 30.50 + 15.96 + 6.25
bnotes = pd.DataFrame({'หมายเหตุ': [
    'ที่มา (แหล่งรอง): Rocket Media Lab, "17 ปีงบแก้ปัญหาไฟใต้ถูกใช้ผ่านหน่วยงานไหน เท่าไรบ้าง" rocketmedialab.co/deep-south-budget/ — รวบรวมจาก พ.ร.บ.งบประมาณ ปี 2551–2568 (4 แผนงานที่เปลี่ยนชื่อตามช่วงเวลา) งบรวม 267,673,113,800 บาท',
    f'หน่วยงานด้านความมั่นคง (กอ.รมน. + กลาโหม + ตำรวจ) รวม ≈ {security:.1f}% ของงบสะสม 2551–2568',
    'ข้อจำกัด: เป็นการแยกตาม "หน่วยงานที่รับงบ" ไม่ใช่ "ภารกิจ" — กอ.รมน. ทำโครงการพัฒนาด้วย; เป็นงบสะสม ไม่ได้แยกรายปี; ยังไม่ได้ตรวจกับเอกสารงบประมาณต้นฉบับ',
    'ปีงบ 2569 (ร่าง): แผนบูรณาการเหลือ 1,475,003,600 บาท ลดลง ~74% จาก 5,737,960,100 บาท; กลาโหม 12.9 ล้าน, สำนักนายกฯ 16.5 ล้าน (แหล่งเดียวกัน)',
]})

with pd.ExcelWriter(OUT + 'Population_district_DOPA_2562-2568_clean.xlsx') as w:
    pop.to_excel(w, 'population_long', index=False)
    pop.pivot_table(index=['province', 'district'], columns='year_BE', values='population').reset_index().to_excel(w, 'wide', index=False)
    notes.to_excel(w, 'notes', index=False)
with pd.ExcelWriter(OUT + 'Tourism_baseline_2561-2562_clean.xlsx') as w:
    tour.to_excel(w, 'tourism_2561_2562', index=False)
    cmp_.to_excel(w, 'compare_2568', index=False)
    tnotes.to_excel(w, 'notes', index=False)
with pd.ExcelWriter(OUT + 'Budget_southern_by_agency_2551-2568_clean.xlsx') as w:
    bud.to_excel(w, 'by_agency_cumulative', index=False)
    bnotes.to_excel(w, 'notes', index=False)
print(cmp_.to_string()); print('security share', security)
