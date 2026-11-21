"""21 Nov 2026: indicator-count reconciliation (113 in master panel vs 72 on the site) and data fixes.
- 40 missing codes = DBD new/dissolved firms split by 20 business sections (A–S + unspecified); the site shows them
  summarised as 'dissolved per 100 new firms by section' (data.json 'sections'), not as 40 explorer items.
- 1 missing code = agri_output_coconut_t: the OAE source reports coconut production in nuts (ผล), not tonnes,
  so it cannot sit next to the other crops; unit corrected and logged as RF27.
- Site units/domains get Thai labels (unit_th, domain_th) so the explorer no longer shows 'tonnes (coconut: as source)'.
Run from repo root."""
import json, openpyxl, pandas as pd

# ---- master panel: coconut unit
csv = 'data/master_panel/SBP_master_panel_long.csv'
m = pd.read_csv(csv)
fix = m.indicator_code.str.startswith('agri_output_')
m.loc[fix, 'unit'] = 'tonnes'
m.loc[m.indicator_code == 'agri_output_coconut_t', 'unit'] = 'nuts (ผล) — not tonnes despite code suffix (RF27)'
m.to_csv(csv, index=False, encoding='utf-8-sig')
wb = openpyxl.load_workbook('data/master_panel/SBP_master_panel_2564-2568.xlsx')
for name in ('panel_long', 'indicator_dictionary'):
    ws = wb[name]; hdr = [c.value for c in ws[1]]
    ci = hdr.index('indicator_code') if 'indicator_code' in hdr else hdr.index('code')
    ui = hdr.index('unit')
    for row in ws.iter_rows(min_row=2):
        code = str(row[ci].value or '')
        if code.startswith('agri_output_'):
            row[ui].value = 'nuts (ผล) — not tonnes despite code suffix (RF27)' if code == 'agri_output_coconut_t' else 'tonnes'
wb.save('data/master_panel/SBP_master_panel_2564-2568.xlsx')

# ---- red flag RF27
wb = openpyxl.load_workbook('data/reference/red_flags_log.xlsx'); ws = wb['red_flags']
if not any(c.value == 'RF27' for c in ws['A']):
    ws.append(['RF27', '2026-11-21', '🟢 Low', 'B. Units', 'OAE crop production (Agriculture_4prov_2563-2568_clean.xlsx)', 'PTN, YLA, NWT, SKA', '2563–2568',
               'Coconut production 20–34 million per province per year with yield in "ผลต่อไร่" (nuts per rai)',
               'Panel code agri_output_coconut_t and unit "tonnes" implied tonnes; the source counts nuts',
               'Source convention: coconut is reported in nuts, other crops in tonnes', '—',
               'Yield column says ผลต่อไร่; production = area × yield in nuts',
               'Unit corrected to nuts; kept out of the site explorer and out of any cross-crop comparison', 'Resolved', '—'])
    wb.save('data/reference/red_flags_log.xlsx')

# ---- site data: Thai unit/domain labels
UNIT = {'%': '%', '% of GPP': '% ของ GPP', 'million baht': 'ล้านบาท', 'tonnes (coconut: as source)': 'ตัน', 'tonnes': 'ตัน', 'persons': 'คน',
        'baht': 'บาท', 'firms': 'ราย', 'baht/kg': 'บาท/กก.', 'baht/person/year': 'บาท/คน/ปี', '1,000 persons': 'พันคน', 'vessels': 'ลำ',
        'projects': 'โครงการ', '0–1': 'ดัชนี 0–1', 'rank': 'อันดับ', '% of population': '% ของประชากร', '% of จปฐ. persons': '% ของคนที่สำรวจ จปฐ.',
        'index': 'ดัชนี', 'incidents': 'ครั้ง', 'events': 'ครั้ง'}
DOM = {'Agriculture': 'เกษตร', 'Business': 'ธุรกิจ', 'Economy': 'เศรษฐกิจมหภาค', 'Fisheries': 'ประมง', 'Investment': 'การลงทุน', 'Labour': 'แรงงาน',
       'Poverty': 'ความยากจนและรายได้', 'Public budget': 'งบประมาณรัฐ', 'Region': 'พื้นที่', 'Security': 'ความมั่นคง', 'Tourism': 'ท่องเที่ยว', 'Trade': 'การค้า'}
d = json.load(open('docs/data.json', encoding='utf-8'))
for i in d['indicators']:
    if i['unit'].startswith('tonnes'): i['unit'] = 'tonnes'
    i['unit_th'] = UNIT[i['unit']]; i['domain_th'] = DOM[i['domain']]
d['counts'] = {'master_indicators': int(m.indicator_code.nunique()), 'site_indicators': len(d['indicators']),
               'dbd_by_section': 40, 'coconut_units': 1, 'master_values': int(len(m)), 'site_values': len(d['panel'])}
json.dump(d, open('docs/data.json', 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print(d['counts'])
