"""Spot-check numbers quoted in the article drafts against the master panel and the Power BI tables.
Writes notes/spot_check_2026-11-20.md. Each claim: (label, article value, computed value)."""
import pandas as pd, math
P = pd.read_csv('data/master_panel/SBP_master_panel_long.csv', encoding='utf-8-sig')
def v(area, code, year):
    s = P[(P.area == area) & (P.indicator_code == code) & (P.year_BE == year)]['value']
    return None if s.empty else float(s.iloc[0])
lq = pd.read_csv('powerbi/csv/fact_lq.csv', encoding='utf-8-sig')
def L(prov, sector_th): return float(lq[(lq.province == prov) & (lq.sector_th == sector_th)].lq.iloc[0])
sec = pd.read_csv('powerbi/csv/fact_sections.csv', encoding='utf-8-sig')
s3 = sec[sec.area.str.startswith('3 provinces')].set_index('section_th')
pm = pd.read_csv('powerbi/csv/fact_price_monthly.csv', encoding='utf-8-sig')
rss = pm[pm['product'] == 'ยางแผ่นดิบ ชั้น 3']
fd = pd.read_csv('powerbi/csv/fact_district.csv', encoding='utf-8-sig')
claims = [
 ('Poverty Pattani 2568 (%)', 21.21, v('Pattani', 'poverty_rate_pct', 2568)),
 ('Poverty Whole Kingdom 2568 (%)', 3.27, v('Whole Kingdom', 'poverty_rate_pct', 2568)),
 ('Poverty Pattani 2564 (%)', 30.43, v('Pattani', 'poverty_rate_pct', 2564)),
 ('Poverty Yala 2566 (%)', 9.04, v('Yala', 'poverty_rate_pct', 2566)),
 ('Poverty Yala 2567 (%)', 25.54, v('Yala', 'poverty_rate_pct', 2567)),
 ('LQ agriculture Pattani', 2.69, L('Pattani', 'เกษตร ป่าไม้ ประมง')),
 ('LQ agriculture Yala', 4.14, L('Yala', 'เกษตร ป่าไม้ ประมง')),
 ('LQ manufacturing Yala', 0.29, L('Yala', 'อุตสาหกรรมการผลิต')),
 ('LQ manufacturing Pattani', 0.36, L('Pattani', 'อุตสาหกรรมการผลิต')),
 ('LQ education Narathiwat', 4.00, L('Narathiwat', 'การศึกษา')),
 ('Fish landings Pattani 2567 (t)', 84694, v('Pattani', 'fish_landings_t', 2567)),
 ('Fish landings value Pattani 2567 (mb)', 3359, v('Pattani', 'fish_landings_value_mb', 2567)),
 ('Manufacturing share of GPP Pattani 2567 (%)', 8.72, v('Pattani', 'gpp_share_manuf_pct', 2567)),
 ('Visitors Narathiwat 2565', 431003, v('Narathiwat', 'tour_visitors', 2565)),
 ('Visitors Yala 2568', 2033274, v('Yala', 'tour_visitors', 2568)),
 ('Malaysian hotel guests Pattani 2568', 54926, v('Pattani', 'tour_hotel_guests_malaysian', 2568)),
 ('Occupancy Yala 2568 (%)', 80.9, v('Yala', 'tour_occupancy_pct', 2568)),
 ('GPP per capita Narathiwat 2567 (baht)', 66785, v('Narathiwat', 'gpp_per_capita_baht', 2567)),
 ('GPP per capita Songkhla 2567 (baht)', 156755, v('Songkhla', 'gpp_per_capita_baht', 2567)),
 ('HH income national 2566 (baht/month)', 29030, v('Whole Kingdom', 'hh_income_month_baht', 2566)),
 ('Informal debt Pattani 2566 (baht)', 11847, v('Pattani', 'hh_informal_debt_baht', 2566)),
 ('Unemployment Narathiwat 2564 (%)', 8.47, v('Narathiwat', 'unemployment_rate_pct', 2564)),
 ('Unemployment Narathiwat 2568 (%)', 2.37, v('Narathiwat', 'unemployment_rate_pct', 2568)),
 ('Official incidents Narathiwat 2568', 103, v('Narathiwat', 'sec_official_incidents', 2568)),
 ('UCDP lethal events Narathiwat 2568', 13, v('Narathiwat', 'sec_ucdp_lethal_events', 2568)),
 ('Vessels Pattani 2564', 2214, v('Pattani', 'fish_vessels_licensed', 2564)),
 ('Vessels Pattani 2568', 1624, v('Pattani', 'fish_vessels_licensed', 2568)),
 ('Border trade Narathiwat 2564 (mb)', 3405, v('Narathiwat', 'border_trade_mb', 2564)),
 ('Manufacturing dissolved per 100 new (3 prov)', 50.4, float(s3.loc['การผลิต (อุตสาหกรรม)', 'dissolved_per_100_new'])),
 ('RSS3 annual average 2554 (baht/kg)', 127.96, round(rss[rss.year_be == 2554].price_baht_per_kg.mean(), 2)),
 ('RSS3 annual average 2561 (baht/kg)', 41.48, round(rss[rss.year_be == 2561].price_baht_per_kg.mean(), 2)),
 ('RSS3 months at or above 100 baht', 14, int((rss.price_baht_per_kg >= 100).sum())),
 ('Chanae lethal events 2564-2568', 13, int(fd[fd.district == 'จะแนะ'].ucdp_lethal_events_2564_2568.iloc[0])),
 ('Nong Chik new firms 2565-2568', 64, int(fd[fd.district == 'หนองจิก'].new_firms_2565_2568.iloc[0])),
]
rows = []
for lab, art, calc in claims:
    if calc is None: st = 'ไม่พบในข้อมูล'
    else:
        tol = max(abs(art) * 0.005, 0.011)
        st = 'ตรง' if abs(calc - art) <= tol else 'ไม่ตรง'
    rows.append((lab, art, calc, st))
df = pd.DataFrame(rows, columns=['ข้อความในบทความ', 'ค่าในบทความ', 'ค่าจากข้อมูล', 'ผล'])
print(df.to_string(index=False)); print(df['ผล'].value_counts().to_dict())
with open('notes/spot_check_2026-11-20.md', 'w', encoding='utf-8') as f:
    f.write('# สุ่มตรวจตัวเลขในบทความ (20 พ.ย. 2569)\n\nเทียบตัวเลขที่อ้างในบทความกับ Master Panel และตาราง Power BI ด้วยสคริปต์ `scripts/spot_check_article_2026-11-20.py` (ยอมคลาดเคลื่อนจากการปัดเศษ 0.5%)\n\n')
    f.write(df.to_markdown(index=False))
    f.write(f"\n\nสรุป: {df['ผล'].value_counts().to_dict()}\n\nหมายเหตุ: นี่คือการตรวจว่าบทความคัดลอกตัวเลขจากชุดข้อมูลถูกต้อง ไม่ใช่การตรวจชุดข้อมูลกับไฟล์ต้นฉบับของหน่วยงาน (ส่วนนั้นตรวจตอนทำ clean และบันทึกใน red flags)\n")
