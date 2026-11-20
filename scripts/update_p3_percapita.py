"""20 Nov 2026: add district population + per-capita measures; update beliefs 3, 5, 6.
Writes powerbi/csv/fact_district.csv, fact_hypothesis.csv, docs/data.json and (if given) an SBP_PowerBI.xlsx path.
Usage: python update_p3_percapita.py [path/to/SBP_PowerBI.xlsx ...]"""
import sys, json, math
import pandas as pd, openpyxl
from scipy import stats

pop = pd.read_excel('data/clean/Population_district_DOPA_2562-2568_clean.xlsx', 'population_long')
den = pd.concat([pop[(pop.province != 'Narathiwat') & (pop.year_BE == 2566)], pop[(pop.province == 'Narathiwat') & (pop.year_BE == 2567)]])
den = den[['province', 'district', 'year_BE', 'population']].rename(columns={'year_BE': 'pop_year'})
fd = pd.read_csv('powerbi/csv/fact_district.csv', encoding='utf-8-sig')
fd = fd[[c for c in fd.columns if c not in ('population', 'pop_year', 'lethal_per_100k', 'new_firms_per_10k', 'rank_lethal_per_100k', 'rank_new_firms_per_10k')]]
fd = fd.merge(den, on=['province', 'district'], how='left')
assert fd.population.notna().all()
fd['lethal_per_100k'] = (fd.ucdp_lethal_events_2564_2568 / fd.population * 1e5).round(2)
fd['new_firms_per_10k'] = (fd.new_firms_2565_2568 / fd.population * 1e4).round(2)
fd['rank_lethal_per_100k'] = fd.lethal_per_100k.rank()
fd['rank_new_firms_per_10k'] = fd.new_firms_per_10k.rank()
fd.to_csv('powerbi/csv/fact_district.csv', index=False, encoding='utf-8-sig')

def ci(r, n, k=3):
    z, se = math.atanh(r), 1 / math.sqrt(n - k)
    return math.tanh(z - 1.96 * se), math.tanh(z + 1.96 * se)
r_raw = stats.spearmanr(fd.ucdp_lethal_events_2564_2568, fd.new_firms_2565_2568)[0]
r_pc = stats.spearmanr(fd.lethal_per_100k, fd.new_firms_per_10k)[0]
lo, hi = ci(r_pc, len(fd))
print(f'raw {r_raw:.3f}  per-capita {r_pc:.3f} CI {lo:.2f} {hi:.2f}')

hy = pd.read_csv('powerbi/csv/fact_hypothesis.csv', encoding='utf-8-sig')
hy.loc[hy.no == 3, 'evidence'] = 'ผู้เยี่ยมเยือนเพิ่มทุกปีหลังโควิด: นราธิวาส 431,003 → 1.24 ล้านคน (2565→2568) และสูงกว่าปี 2562 ก่อนโควิด (703,810 คน) — เทียบกับ 2562 แบบระวัง เพราะวิธีประมาณอาจเปลี่ยน (RF26)'
hy.loc[hy.no == 5, ['result', 'evidence']] = ['ตรงบางส่วน', 'กอ.รมน. + กลาโหม + ตำรวจ ≈ 53% ของงบแก้ปัญหาชายแดนใต้สะสม 2551–2568 (แหล่งรอง: Rocket Media Lab) — แยกตามหน่วยงานที่รับงบ ไม่ใช่ภารกิจ; ปีงบ 2569 งบแผนบูรณาการลดลง ~74%']
hy.loc[hy.no == 6, ['result', 'evidence']] = ['ตรงบางส่วน', f'นับดิบ ρ = {r_raw:.2f}; ปรับตามประชากร ρ = {r_pc:.2f} (95% CI {lo:.2f} ถึง {hi:.2f}, 37 อำเภอ); คุมประชากรแบบ partial ρ = −0.31 (CI −0.58 ถึง 0.02) — ทิศทางลบทุกจังหวัด แต่เป็นความสัมพันธ์ ไม่ใช่เหตุและผล'.replace('-', '−')]
hy.to_csv('powerbi/csv/fact_hypothesis.csv', index=False, encoding='utf-8-sig')
print(hy.result.value_counts().to_dict())

d = json.load(open('docs/data.json', encoding='utf-8'))
extra = fd.set_index(['province', 'district'])[['population', 'pop_year', 'lethal_per_100k', 'new_firms_per_10k']].to_dict('index')
for r in d['districts']:
    e = extra[(r['province'], r['district'])]
    r.update({k: (int(v) if k in ('population', 'pop_year') else float(v)) for k, v in e.items()})
d['hyp'] = [{k: str(v) for k, v in row.items()} for row in hy.to_dict('records')]
d['tour_baseline_2562'] = {'Pattani': 282059, 'Yala': 689938, 'Narathiwat': 703810, 'Songkhla': 7503461}
json.dump(d, open('docs/data.json', 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))

for x in sys.argv[1:]:
    wb = openpyxl.load_workbook(x)
    for name, df in (('fact_district', fd), ('fact_hypothesis', hy)):
        idx = wb.sheetnames.index(name); wb.remove(wb[name]); ws = wb.create_sheet(name, idx)
        ws.append(list(df.columns))
        for row in df.itertuples(index=False):
            ws.append([None if (isinstance(v, float) and math.isnan(v)) else (v.item() if hasattr(v, 'item') else v) for v in row])
    wb.save(x); print('updated', x)
