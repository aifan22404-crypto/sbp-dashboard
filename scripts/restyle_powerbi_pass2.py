"""Second polish pass on the SBP PBIR report. Usage: python restyle_pass2.py <pbip_root>"""
import json, os, sys, glob

ROOT = sys.argv[1]
PAGES = os.path.join(ROOT, 'SBP_Dashboard.Report', 'definition', 'pages')

def lit(v): return {"expr": {"Literal": {"Value": v}}}
def s(t): return lit("'" + t.replace("'", "''") + "'")
def col(h): return {"solid": {"color": s(h)}}

RENAME = {
    'Poverty %': 'พื้นที่ที่เลือก', 'Poverty % (Kingdom)': 'ทั้งประเทศ',
    'HH Income (baht/month)': 'พื้นที่ที่เลือก', 'HH Income (Kingdom)': 'ทั้งประเทศ',
    'no': 'ข้อ', 'P': 'ส่วน', 'hypothesis': 'ความเชื่อก่อนดูข้อมูล', 'result': 'ผล', 'evidence': 'หลักฐานจากข้อมูล',
    'section_th': 'สาขา', 'area': 'พื้นที่', 'district': 'อำเภอ', 'year_label': 'ปี', 'sector_th': 'สาขา',
}
BELIEF_COLOR = {'สอดคล้อง': '#1A7F5E', 'ตรงบางส่วน': '#A9741A', 'ไม่ตรง': '#B0332A', 'ทดสอบไม่ได้': '#7A7F88'}
ZERO_BASE = {'Official Incidents', 'Lethal Events (UCDP)', 'Fishing Vessels', 'Visitors'}

n = 0
for fp in glob.glob(os.path.join(PAGES, '*', 'visuals', '*', 'visual.json')):
    with open(fp, encoding='utf-8') as f:
        d = json.load(f)
    v = d['visual']; vt = v['visualType']
    q = v.get('query', {}).get('queryState', {})
    names = []
    for role in q.values():
        for pr in role.get('projections', []):
            nm = pr.get('nativeQueryRef')
            names.append(nm)
            if nm in RENAME and vt != 'card':
                pr['displayName'] = RENAME[nm]
    o = v.setdefault('objects', {})
    if vt == 'card':
        lab = o.setdefault('labels', [{"properties": {}}])[0]['properties']
        lab['labelDisplayUnits'] = lit("1D")
        if names and names[0] in BELIEF_COLOR:
            lab['color'] = col(BELIEF_COLOR[names[0]])
            lab['fontSize'] = lit("36D")
        o['categoryLabels'] = [{"properties": {"show": lit("true"), "fontSize": lit("12D"), "color": col('#5E6673')}}]
    if vt == 'scatterChart':
        o['categoryAxis'] = [{"properties": {"showAxisTitle": lit("true"), "titleText": s("ธุรกิจตั้งใหม่ 2565–2568 (ราย)")}}]
        o['valueAxis'] = [{"properties": {"showAxisTitle": lit("true"), "titleText": s("เหตุที่มีผู้เสียชีวิต 2564–2568 (ครั้ง)")}}]
    if vt == 'lineChart' and set(names) & ZERO_BASE:
        o['valueAxis'] = [{"properties": {"start": lit("0D")}}]
    if vt == 'clusteredBarChart' and 'Dissolved per 100 New' in names:
        o['legend'] = [{"properties": {"show": lit("true"), "position": s("Top")}}]
        v['query']['sortDefinition'] = {"sort": [{"field": {"Measure": {"Expression": {"SourceRef": {"Entity": "_Measures"}}, "Property": "Dissolved per 100 New"}}, "direction": "Descending"}]}
    with open(fp, 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    n += 1
print('pass2 updated', n, 'visuals')
