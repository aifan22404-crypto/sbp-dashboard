# คู่มือประกอบ Dashboard ใน Power BI — SBP Economic Dashboard

ทำตามทีละขั้น ใช้เวลาประมาณ 2–3 ชั่วโมงสำหรับคนที่เพิ่งเริ่ม ไฟล์ทั้งหมดอยู่ในโฟลเดอร์นี้ (`04_output/powerbi`)

## ไฟล์ในชุดนี้

| ไฟล์ | แถว | คืออะไร |
|---|---|---|
| fact_panel.csv | 2,147 | ตารางหลัก: ค่าของตัวชี้วัดทุกตัว ทุกพื้นที่ ทุกปี |
| dim_indicator.csv | 113 | รายชื่อตัวชี้วัด หมวด หน่วย และใช้ใน P ไหน |
| dim_area.csv | 8 | พื้นที่ (3 จังหวัด, สงขลา, กลุ่มเทียบ) |
| dim_year.csv | 7 | ปี พ.ศ./ค.ศ. |
| fact_lq.csv | 76 | Location Quotient รายสาขา รายจังหวัด (P1) |
| fact_district.csv | 37 | ธุรกิจตั้งใหม่/เลิก และเหตุรุนแรง รายอำเภอ (P2, P3) |
| fact_fish.csv | 12 | สัตว์น้ำขึ้นท่า (P1, P4) |
| fact_sections.csv | 16 | ธุรกิจตั้งใหม่/เลิก ตามหมวด TSIC (P4) |
| fact_disbursement.csv | 22 | การเบิกจ่ายงบ นราธิวาส รายอำเภอ (P2) |
| fact_incident_types.csv | 28 | ประเภทเหตุการณ์ นราธิวาส (P3) |
| fact_price_monthly.csv | 229 | ราคายาง RSS3 และปาล์ม รายเดือน (P4) |
| fact_hypothesis.csv | 8 | สมมติฐาน 8 ข้อเทียบกับข้อมูล |
| measures.dax | — | สูตร DAX ทั้งหมด |
| sbp_theme.json | — | ธีมสี |

## ขั้นที่ 1 — นำเข้าข้อมูล (15 นาที)

1. เปิด Power BI Desktop → **Home > Get data > Text/CSV**
2. เลือก `fact_panel.csv` → ในหน้าต่าง preview ตั้ง **File Origin = 65001: Unicode (UTF-8)** (ถ้าภาษาไทยเป็นตัวประหลาด) → กด **Load**
3. ทำซ้ำกับไฟล์ CSV อีก 11 ไฟล์
4. **View > Themes > Browse for themes** → เลือก `sbp_theme.json`
5. กด **Save** ตั้งชื่อ `SBP_Dashboard.pbix` เก็บใน `04_output/powerbi`

ตรวจ: ที่แถบ Data ด้านขวาต้องเห็น 12 ตาราง และ `fact_panel` มี 2,147 แถว (ดูที่ Table view มุมล่างซ้าย)

## ขั้นที่ 2 — เชื่อมตาราง (15 นาที)

ไปที่ **Model view** (ไอคอนที่สามด้านซ้าย) แล้วลากคอลัมน์ทับกันตามตาราง ถ้า Power BI สร้างเส้นให้อัตโนมัติแล้ว ให้ตรวจว่าตรงตามนี้

| ลากจาก (ฝั่ง many) | ไปยัง (ฝั่ง one) |
|---|---|
| fact_panel[indicator_code] | dim_indicator[indicator_code] |
| fact_panel[area] | dim_area[area] |
| fact_panel[year_BE] | dim_year[year_BE] |
| fact_lq[province] | dim_area[area] |
| fact_district[province] | dim_area[area] |
| fact_fish[province] | dim_area[area] |
| fact_fish[year] | dim_year[year_BE] |
| fact_sections[area] | dim_area[area] |
| fact_price_monthly[year_be] | dim_year[year_BE] |

ทุกเส้นต้องเป็น **Many to one (\*:1)** และ Cross filter direction = **Single**
`fact_disbursement`, `fact_incident_types`, `fact_hypothesis` ไม่ต้องเชื่อม (เป็นข้อมูลนราธิวาส/ตารางอ้างอิง)

จากนั้นตั้งค่าเรียงลำดับ: คลิกคอลัมน์ `dim_area[area_th]` → **Column tools > Sort by column > sort**

## ขั้นที่ 3 — ใส่สูตร DAX (20 นาที)

1. **Home > Enter data** → ตั้งชื่อตาราง `_Measures` → Load
2. คลิกขวาที่ `_Measures` → **New measure** → เปิด `measures.dax` ด้วย Notepad แล้วคัดลอกมาวางทีละบล็อก (หนึ่งบล็อกคือชื่อ measure จนถึงบรรทัดว่าง)
3. จัดรูปแบบ: `YoY %`, `Change since 2564 %`, `Disbursed %` → Format = **Percentage**; `LQ`, `Spearman rho...` → ทศนิยม 2 ตำแหน่ง

สำคัญ: measure `Value` จะแสดงค่าเมื่อเลือกตัวชี้วัด **ตัวเดียว** เท่านั้น เพราะตัวชี้วัดต่างหน่วยบวกกันไม่ได้ ทุกกราฟที่ใช้ `Value` ต้องกรองตัวชี้วัดไว้ที่ **Filters on this visual**

## ขั้นที่ 4 — Slicer ที่ใช้ร่วมทุกหน้า

สร้างในหน้าแรกก่อน แล้วคัดลอกไปหน้าอื่น

| Slicer | ฟิลด์ | รูปแบบ |
|---|---|---|
| จังหวัด | dim_area[area_th] | **Tile** (ปุ่ม) วางบนสุด |
| ปี | dim_year[year_BE] | **Between** (แถบเลื่อน) |

เลือก slicer → **View > Sync slicers** → ติ๊กทุกหน้า เพื่อให้การเลือกตามไปทุกหน้า

## ขั้นที่ 5 — สร้างหน้า (ประมาณ 20 นาทีต่อหน้า)

ในตาราง "ตัวกรองของกราฟ" หมายถึงลาก `dim_indicator[indicator_code]` ไปที่ Filters on this visual แล้วติ๊กรหัสนั้น

### หน้า 1 ภาพรวม

| กราฟ | ชนิด | แกน / ค่า | ตัวกรองของกราฟ |
|---|---|---|---|
| GPP ต่อหัว | Card | Value Latest | gpp_per_capita_baht |
| อัตราความยากจน | Card | Value Latest | poverty_rate_pct |
| ผู้เยี่ยมเยือน | Card | Value Latest | tour_visitors |
| ธุรกิจตั้งใหม่ | Card | Value Latest | dbd_new_firms |
| ความยากจนเทียบประเทศ | Line chart | X: dim_year[year_BE] · Y: Value · Legend: dim_area[area_th] | poverty_rate_pct |
| GPP ต่อหัวรายจังหวัด | Clustered bar | Y: dim_area[area_th] · X: Value Latest | gpp_per_capita_baht |

สำหรับกราฟเส้นความยากจน: เลือกกราฟ → **Format > Edit interactions** → ตั้ง slicer จังหวัดเป็น **None** เพื่อให้เส้น "ทั้งประเทศ" ยังอยู่ให้เทียบ

### หน้า 2 P1 โอกาส

| กราฟ | ชนิด | แกน / ค่า | ตัวกรอง |
|---|---|---|---|
| LQ รายสาขา | Clustered bar | Y: fact_lq[sector_th] · X: LQ · เรียงมากไปน้อย | — |
| เส้นอ้างอิง LQ = 1 | Analytics pane > Constant line = 1 | | |
| สัดส่วนสาขาเทียบประเทศ | Table | sector_th, Sector Share %, National Share %, LQ, real_cagr_2021_2024_pct | — |
| สัตว์น้ำขึ้นท่า | Clustered column | X: fact_fish[year] · Y: Fish Landings (t) · Legend: province | — |
| ผู้เยี่ยมเยือน | Line chart | X: year_BE · Y: Value · Legend: area_th | tour_visitors |
| แขกชาวมาเลเซีย | Line chart | X: year_BE · Y: Value | tour_hotel_guests_malaysian |

ใส่สีตามเงื่อนไขแท่ง LQ: **Format > Bars > Color > fx** → Rules: ค่า LQ ≥ 1 เป็นสีเขียว, < 1 เป็นสีเทา

### หน้า 3 P2 สิ่งที่ควรพัฒนา

| กราฟ | ชนิด | แกน / ค่า | ตัวกรอง |
|---|---|---|---|
| รายได้ครัวเรือนต่อเดือน | Clustered bar | Y: area_th · X: Value Latest | hh_income_month_baht |
| หนี้นอกระบบ | Clustered bar | Y: area_th · X: Value Latest | hh_informal_debt_baht |
| อัตราว่างงาน | Line chart | X: year_BE · Y: Value · Legend: area_th | unemployment_rate_pct |
| ทุนจดทะเบียนธุรกิจใหม่ | Clustered column | X: year_BE · Y: Value · Legend: area_th | dbd_new_capital_mb |
| การเบิกจ่ายงบ นราธิวาส | Clustered bar | Y: fact_disbursement[district] · X: Disbursed % · Legend: expense_type | — |

ใส่ข้อความใต้กราฟเบิกจ่าย: "ข้อมูลเฉพาะจังหวัดนราธิวาส ปีงบประมาณ 2568"

### หน้า 4 P3 ความเสี่ยง

| กราฟ | ชนิด | แกน / ค่า | ตัวกรอง |
|---|---|---|---|
| เหตุการณ์ (สถิติทางการ) | Line chart | X: year_BE · Y: Value · Legend: area_th | sec_official_incidents |
| เหตุที่มีผู้เสียชีวิต (UCDP) | Line chart | X: year_BE · Y: Value · Legend: area_th | sec_ucdp_lethal_events |
| เหตุรุนแรงกับธุรกิจใหม่ รายอำเภอ | Scatter | X: Lethal Events (district) · Y: New Firms (district) · Values: fact_district[district] | — |
| ค่าสหสัมพันธ์ | Card | Spearman rho (events vs new firms) | — |
| ประเภทเหตุการณ์ นราธิวาส | Stacked column | X: fact_incident_types[year] · Y: Incidents by Type · Legend: type | — |

ใส่กล่องข้อความ: "ค่าสหสัมพันธ์บอกแค่ว่าตัวเลขสองชุดขึ้นลงไปด้วยกันหรือไม่ ไม่ได้บอกว่าอะไรเป็นสาเหตุ และข้อมูลไม่ได้ระบุว่าใครเป็นผู้ก่อเหตุ"

### หน้า 5 P4 กลุ่มเปราะบาง

| กราฟ | ชนิด | แกน / ค่า | ตัวกรอง |
|---|---|---|---|
| ราคายาง RSS3 รายเดือน | Line chart | X: fact_price_monthly[date] · Y: Rubber Price (baht/kg) | — |
| เส้นอ้างอิง 100 บาท | Analytics pane > Constant line = 100 | | |
| เรือประมงที่มีใบอนุญาต | Line chart | X: year_BE · Y: Value · Legend: area_th | fish_vessels_licensed |
| ธุรกิจเลิกต่อตั้งใหม่ 100 ราย ตามหมวด | Clustered bar | Y: fact_sections[section_th] · X: Dissolved per 100 New | — |
| สัดส่วนรายได้จากเกษตรของครัวเรือน | Clustered bar | Y: area_th · X: Value Latest | hh_farm_income_share_pct |

สำหรับกราฟราคายาง: ตั้ง slicer ปีเป็น **None** ใน Edit interactions เพื่อให้เห็นย้อนถึงปี 2554

### หน้า 6 สมมติฐาน

Table จาก `fact_hypothesis`: no, P, hypothesis, result, evidence
ใส่สีพื้นคอลัมน์ result: **Format > Cell elements > Background color > fx > Rules** ตามข้อความ (ทำได้โดยเพิ่มคอลัมน์ตัวเลข หรือใช้ Field value)

## ขั้นที่ 6 — ตรวจตัวเลข (ห้ามข้าม)

เลือกจังหวัดและปีตามตาราง แล้วเทียบกับค่าที่ต้องได้ ถ้าไม่ตรงแปลว่าเส้นเชื่อมตารางหรือตัวกรองผิด

| ตรวจอะไร | ต้องได้ |
|---|---|
| LQ เกษตร ป่าไม้ ประมง — ปัตตานี / ยะลา / นราธิวาส | 2.69 / 4.14 / 3.29 |
| Spearman rho (ไม่เลือกจังหวัด) | −0.06 (−0.055) |
| ผู้เยี่ยมเยือน นราธิวาส 2565 | 431,003 |
| ผู้เยี่ยมเยือน ปัตตานี 2565 → 2568 | 255,452 → 643,995 |
| ราคายาง: จำนวนเดือนที่ ≥ 100 บาท | 14 เดือน (เดือนสุดท้าย พ.ค. 2555) |
| ราคายางสูงสุดตั้งแต่ปี 2564 | 81.61 บาท (มิ.ย. 2569) |
| ตากใบ ธุรกิจตั้งใหม่ 2565–2568 | 52 |

## ข้อควรระวังเรื่องข้อมูล

- คอลัมน์ `usable = 0` (39 แถว) คือค่าที่ไม่ควรใช้: ข้อมูลเบื้องต้นปี 2568, ข้อมูลไม่ครบปี และผลไม้ปี 2565 (RF23) measure `Value` ตัดออกให้แล้ว
- BOI เป็นยอด **อนุมัติ** ไม่ใช่ลงทุนจริง
- `tpmap_target_poor_rate_pct` ไม่ใช่อัตราความยากจน ห้ามเทียบข้ามปี
- ตัวเลข GPP ปี 2567 เป็นค่าเบื้องต้น
- อย่าเขียนข้อความที่ระบุว่าใครก่อเหตุหรือกล่าวหาหน่วยงานใด ข้อมูลชุดนี้พิสูจน์เรื่องนั้นไม่ได้

## เผยแพร่

**File > Publish > Publish to Power BI** ต้องใช้บัญชีที่มีสิทธิ์ (อีเมลมหาวิทยาลัยมักใช้ได้) ถ้าเผยแพร่ไม่ได้ ให้ **File > Export > Export to PDF** หรืออัดวิดีโอหน้าจอสำหรับ TikTok
