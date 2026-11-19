"""Restyle the SBP Power BI report (PBIR + TMDL) in place.
Usage: python restyle.py <pbip_root> <theme_v2.json>
"""
import json, os, sys, shutil, hashlib, re

ROOT, THEME = sys.argv[1], sys.argv[2]
REP = os.path.join(ROOT, 'SBP_Dashboard.Report')
DEF = os.path.join(REP, 'definition')
PAGES = os.path.join(DEF, 'pages')
TMDL = os.path.join(ROOT, 'SBP_Dashboard.SemanticModel', 'definition', 'tables', '_Measures.tmdl')
VSCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.14.0/schema.json"

def lit(v):
    return {"expr": {"Literal": {"Value": v}}}
def s(t):  # string literal
    return lit("'" + t.replace("'", "''") + "'")
def col(hexc):
    return {"solid": {"color": s(hexc)}}
def uid(seed):
    return hashlib.md5(seed.encode()).hexdigest()[:20]

NAVY, GOLD, INK, GREY, MUTED = '#0F2236', '#C9A24B', '#14202E', '#9A9488', '#7A7F88'

# ---------------------------------------------------------------- theme
with open(THEME, encoding='utf-8') as f:
    theme = f.read()
for fn in os.listdir(os.path.join(REP, 'StaticResources', 'RegisteredResources')):
    if fn.endswith('.json'):
        with open(os.path.join(REP, 'StaticResources', 'RegisteredResources', fn), 'w', encoding='utf-8') as f:
            f.write(theme)

# ---------------------------------------------------------------- measures
FMT = {
    'Visitors': '#,0', 'Poverty %': '0.0', 'New Capital (mb)': '#,0', 'Official Incidents': '#,0',
    'Lethal Events (UCDP)': '#,0', 'Fishing Vessels': '#,0', 'HH Income (baht/month)': '#,0',
    'GPP per Capita': '#,0', 'Rubber Price (baht/kg)': '0.00', 'Dissolved per 100 New': '0.0',
    'Poverty % (Kingdom)': '0.0', 'HH Income (Kingdom)': '#,0', 'Fish Landings (t)': '#,0',
    'Spearman rho': '0.00', 'Unemployment %': '0.0', 'Lethal Events (district)': '#,0', 'New Firms (district)': '#,0',
}

def latest(base, code):
    return (f'VAR y = CALCULATE ( MAX ( fact_panel[year_BE] ), dim_indicator[indicator_code] = "{code}", '
            f'fact_panel[usable] = 1, NOT ISBLANK ( fact_panel[value] ) ) '
            f'RETURN CALCULATE ( [{base}], dim_year[year_BE] = y )')

NEW = [
    # P1
    ('ผู้เยี่ยมเยือน ปีล่าสุด (คน)', latest('Visitors', 'tour_visitors'), '#,0'),
    ('สัตว์น้ำขึ้นท่า ปีล่าสุด (ตัน)', 'VAR y = CALCULATE ( MAX ( fact_fish[year] ), NOT ISBLANK ( fact_fish[tonnes] ) ) RETURN CALCULATE ( [Fish Landings (t)], fact_fish[year] = y )', '#,0'),
    ('สาขาที่ LQ สูงสุด', 'VAR t = TOPN ( 1, VALUES ( fact_lq[sector_th] ), [LQ], DESC ) RETURN CONCATENATEX ( t, fact_lq[sector_th], ", " )', None),
    ('LQ สูงสุด', 'MAXX ( VALUES ( fact_lq[sector_th] ), [LQ] )', '0.00'),
    # P2
    ('สัดส่วนคนจน ปีล่าสุด (%)', latest('Poverty %', 'poverty_rate_pct'), '0.0'),
    ('ห่างจากทั้งประเทศ (จุด %)', 'VAR y = CALCULATE ( MAX ( fact_panel[year_BE] ), dim_indicator[indicator_code] = "poverty_rate_pct", fact_panel[usable] = 1, NOT ISBLANK ( fact_panel[value] ) ) RETURN CALCULATE ( [Poverty %] - [Poverty % (Kingdom)], dim_year[year_BE] = y )', '+0.0;-0.0;0.0'),
    ('รายได้ครัวเรือน ปีล่าสุด (บาท/เดือน)', latest('HH Income (baht/month)', 'hh_income_month_baht'), '#,0'),
    ('GPP ต่อหัว ปีล่าสุด (บาท/ปี)', latest('GPP per Capita', 'gpp_per_capita_baht'), '#,0'),
    # P3
    ('เหตุการณ์ทางการ ปีล่าสุด (ครั้ง)', latest('Official Incidents', 'sec_official_incidents'), '#,0'),
    ('เหตุมีผู้เสียชีวิต UCDP ปีล่าสุด', latest('Lethal Events (UCDP)', 'sec_ucdp_lethal_events'), '#,0'),
    ('Spearman ρ ความรุนแรง vs ธุรกิจใหม่', 'CALCULATE ( [Spearman rho], REMOVEFILTERS ( dim_area ), REMOVEFILTERS ( fact_district ) )', '0.00'),
    ('จำนวนอำเภอที่วิเคราะห์', 'CALCULATE ( COUNTROWS ( fact_district ), REMOVEFILTERS ( dim_area ), REMOVEFILTERS ( fact_district ) )', '#,0'),
    # P4
    ('ราคายางเฉลี่ยปีล่าสุด (บาท/กก.)', 'VAR y = CALCULATE ( MAX ( fact_price_monthly[year_be] ), fact_price_monthly[product] = "ยางแผ่นดิบ ชั้น 3" ) RETURN CALCULATE ( [Rubber Price (baht/kg)], fact_price_monthly[year_be] = y )', '0.00'),
    ('เรือประมงที่ได้รับอนุญาต ปีล่าสุด (ลำ)', latest('Fishing Vessels', 'fish_vessels_licensed'), '#,0'),
    ('เรือประมง เปลี่ยนแปลงจากปี 2564', 'VAR y = CALCULATE ( MAX ( fact_panel[year_BE] ), dim_indicator[indicator_code] = "fish_vessels_licensed", fact_panel[usable] = 1, NOT ISBLANK ( fact_panel[value] ) ) VAR last = CALCULATE ( [Fishing Vessels], dim_year[year_BE] = y ) VAR base = CALCULATE ( [Fishing Vessels], dim_year[year_BE] = 2564 ) RETURN DIVIDE ( last - base, base )', '+0.0%;-0.0%;0.0%'),
    ('เลิกกิจการต่อตั้งใหม่ 100 ราย (3 จังหวัด)', 'CALCULATE ( [Dissolved per 100 New], REMOVEFILTERS ( dim_area ), CONTAINSSTRING ( fact_sections[area], "3 provinces" ) )', '0.0'),
    # beliefs page
    ('สอดคล้อง', 'CALCULATE ( COUNTROWS ( fact_hypothesis ), fact_hypothesis[result] = "สอดคล้อง" ) + 0', '0'),
    ('ตรงบางส่วน', 'CALCULATE ( COUNTROWS ( fact_hypothesis ), fact_hypothesis[result] = "ตรงบางส่วน" ) + 0', '0'),
    ('ไม่ตรง', 'CALCULATE ( COUNTROWS ( fact_hypothesis ), fact_hypothesis[result] = "ไม่ตรง" ) + 0', '0'),
    ('ทดสอบไม่ได้', 'CALCULATE ( COUNTROWS ( fact_hypothesis ), fact_hypothesis[result] = "ทดสอบไม่ได้" ) + 0', '0'),
]

with open(TMDL, encoding='utf-8') as f:
    tm = f.read()

# format strings for existing measures: drop the general-number hint, add formatString
blocks = re.split(r'(?m)^(?=\tmeasure )', tm)
out = []
for b in blocks:
    m = re.match(r"\tmeasure ('([^']|'')+'|[^ =]+) =", b)
    if m:
        name = m.group(1).strip("'").replace("''", "'")
        if name in FMT and 'formatString:' not in b:
            b = b.replace('\t\tannotation PBI_FormatHint = {"isGeneralNumber":true}\n\n', '')
            b = b.replace('\t\tannotation PBI_FormatHint = {"isGeneralNumber":true}\n', '')
            b = re.sub(r'(\t\tlineageTag: [^\n]+\n)', lambda mm: f'\t\tformatString: {FMT[name]}\n' + mm.group(1), b, count=1)
    out.append(b)
tm = ''.join(out)
existing = set(re.findall(r"\tmeasure '((?:[^']|'')+)'", tm))
add = ''
for name, expr, fmt in NEW:
    if name in existing:
        continue
    add += f"\n\tmeasure '{name}' = {expr}\n"
    if fmt:
        add += f"\t\tformatString: {fmt}\n"
# insert new measures before the first non-measure child (partition/column) or at end
idx = re.search(r'(?m)^\t(partition|column|annotation) ', tm)
tm = (tm[:idx.start()] + add.lstrip('\n') + '\n' + tm[idx.start():]) if idx else tm.rstrip('\n') + '\n' + add
with open(TMDL, 'w', encoding='utf-8') as f:
    f.write(tm)

# ---------------------------------------------------------------- visual builders
def write_visual(page, name, pos, visual, z):
    d = {"$schema": VSCHEMA, "name": name,
         "position": {"x": pos[0], "y": pos[1], "z": z, "width": pos[2], "height": pos[3], "tabOrder": z},
         "visual": visual}
    p = os.path.join(PAGES, page, 'visuals', name)
    os.makedirs(p, exist_ok=True)
    with open(os.path.join(p, 'visual.json'), 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)

def textbox(paragraphs, bg=None):
    v = {"visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": paragraphs}}]},
         "drillFilterOtherVisuals": True}
    vco = {"title": [{"properties": {"show": lit("false")}}],
           "dropShadow": [{"properties": {"show": lit("false")}}],
           "border": [{"properties": {"show": lit("false")}}]}
    if bg:
        vco["background"] = [{"properties": {"show": lit("true"), "color": col(bg), "transparency": lit("0D")}}]
        vco["padding"] = [{"properties": {"top": lit("14D"), "left": lit("28D"), "right": lit("28D"), "bottom": lit("8D")}}]
    else:
        vco["background"] = [{"properties": {"show": lit("false")}}]
    v["visualContainerObjects"] = vco
    return v

def run(text, size, color, bold=False, font='Segoe UI'):
    st = {"fontSize": f"{size}pt", "color": color, "fontFamily": font}
    if bold:
        st["fontWeight"] = "bold"
    return {"value": text, "textStyle": st}

def card(measure, size=28):
    return {"visualType": "card",
            "query": {"queryState": {"Values": {"projections": [{
                "field": {"Measure": {"Expression": {"SourceRef": {"Entity": "_Measures"}}, "Property": measure}},
                "queryRef": f"_Measures.{measure}", "nativeQueryRef": measure}]}}},
            "objects": {"labels": [{"properties": {"fontSize": lit(f"{size}D"), "color": col(NAVY), "fontFamily": s("Georgia")}}],
                        "categoryLabels": [{"properties": {"show": lit("true"), "fontSize": lit("11D"), "color": col(MUTED)}}]},
            "visualContainerObjects": {"title": [{"properties": {"show": lit("false")}}]},
            "drillFilterOtherVisuals": True}

M, HEADER_H, KPI_Y, KPI_H, CH_Y, FOOT_Y = 16, 88, 104, 128, 248, 1050
KPI_W = (1920 - 5 * M) / 4

PAGE_INFO = {
    'P1 Opportunity': ('P1 · Opportunity Mapping', 'สาขาไหนเป็นจุดแข็ง และสาขาไหนยังมีที่ให้โต'),
    'P2 Development Priority': ('P2 · Development Priority', 'ช่องว่างการพัฒนาเทียบกับทั้งประเทศอยู่ตรงไหน'),
    'P3 Risk Factors': ('P3 · Risk Factor Analysis', 'ปัจจัยเสี่ยงสัมพันธ์กับกิจกรรมทางเศรษฐกิจแค่ไหน'),
    'P4 Vulnerable Sectors': ('P4 · Vulnerable Sectors', 'กลุ่มไหนเปราะบางที่สุดเมื่อเจอแรงกระแทก'),
    'ความเชื่อ vs ข้อมูล': ('ความเชื่อ vs ข้อมูล', '8 ความคาดหวังที่เขียนไว้ก่อนเห็นผล เทียบกับสิ่งที่ข้อมูลบอก'),
}
KPIS = {
    'P1 Opportunity': [('ผู้เยี่ยมเยือน ปีล่าสุด (คน)', 30), ('สัตว์น้ำขึ้นท่า ปีล่าสุด (ตัน)', 30), ('สาขาที่ LQ สูงสุด', 20), ('LQ สูงสุด', 30)],
    'P2 Development Priority': [('สัดส่วนคนจน ปีล่าสุด (%)', 30), ('ห่างจากทั้งประเทศ (จุด %)', 30), ('รายได้ครัวเรือน ปีล่าสุด (บาท/เดือน)', 30), ('GPP ต่อหัว ปีล่าสุด (บาท/ปี)', 30)],
    'P3 Risk Factors': [('เหตุการณ์ทางการ ปีล่าสุด (ครั้ง)', 30), ('เหตุมีผู้เสียชีวิต UCDP ปีล่าสุด', 30), ('Spearman ρ ความรุนแรง vs ธุรกิจใหม่', 30), ('จำนวนอำเภอที่วิเคราะห์', 30)],
    'P4 Vulnerable Sectors': [('ราคายางเฉลี่ยปีล่าสุด (บาท/กก.)', 30), ('เรือประมงที่ได้รับอนุญาต ปีล่าสุด (ลำ)', 30), ('เรือประมง เปลี่ยนแปลงจากปี 2564', 30), ('เลิกกิจการต่อตั้งใหม่ 100 ราย (3 จังหวัด)', 30)],
    'ความเชื่อ vs ข้อมูล': [('สอดคล้อง', 34), ('ตรงบางส่วน', 34), ('ไม่ตรง', 34), ('ทดสอบไม่ได้', 34)],
}
L, R2 = 16, 1920 - 16
FULL_W = 1920 - 2 * M
HALF_H = (FOOT_Y - CH_Y - M) / 2
ROW2_Y = CH_Y + HALF_H + M
# measure/field signature -> (position, title, subtitle)
LAYOUT = {
    'P1 Opportunity': {
        'LQ': ((L, CH_Y, 600, FOOT_Y - CH_Y), 'Location Quotient รายสาขา (GPP 2567)', 'LQ > 1 = สัดส่วนสูงกว่าระดับประเทศ · ที่มา: สศช.'),
        'Visitors': ((L + 616, CH_Y, FULL_W - 616, HALF_H), 'ผู้เยี่ยมเยือน (คน)', 'ที่มา: กระทรวงการท่องเที่ยวและกีฬา · ยังไม่มีฐานก่อนโควิด'),
        'Fish Landings (t)': ((L + 616, ROW2_Y, FULL_W - 616, HALF_H), 'สัตว์น้ำเค็มขึ้นท่า (ตัน)', 'ที่มา: กรมประมง'),
    },
    'P2 Development Priority': {
        'Poverty %': ((L, CH_Y, 944, HALF_H), 'สัดส่วนคนจน (%) เทียบทั้งประเทศ', 'ที่มา: สศช. · สีเทา = ทั้งประเทศ'),
        'HH Income (baht/month)': ((L + 960, CH_Y, FULL_W - 960, HALF_H), 'รายได้ครัวเรือนเฉลี่ย (บาท/เดือน)', 'ที่มา: สำนักงานสถิติแห่งชาติ · สำรวจทุก 2 ปี · สีเทา = ทั้งประเทศ'),
        'New Capital (mb)': ((L, ROW2_Y, 944, HALF_H), 'ทุนจดทะเบียนธุรกิจตั้งใหม่ (ล้านบาท)', 'ที่มา: กรมพัฒนาธุรกิจการค้า'),
        'Disbursed %': ((L + 960, ROW2_Y, FULL_W - 960, HALF_H), 'อัตราเบิกจ่ายงบประมาณรายอำเภอ (นราธิวาส ปีงบ 2568)', 'ที่มา: กรมบัญชีกลาง · ไม่เปลี่ยนตามพื้นที่ที่เลือก'),
    },
    'P3 Risk Factors': {
        'Official Incidents': ((L, CH_Y, 944, HALF_H), 'เหตุการณ์ความไม่สงบ (ตัวเลขทางการ)', 'นิยามต่างจาก UCDP · ไม่ระบุผู้ก่อเหตุ'),
        'Lethal Events (UCDP)': ((L + 960, CH_Y, FULL_W - 960, HALF_H), 'เหตุการณ์ที่มีผู้เสียชีวิต (UCDP GED)', 'ที่มา: Uppsala Conflict Data Program'),
        'Lethal Events (district)': ((L, ROW2_Y, FULL_W, HALF_H), '37 อำเภอ: เหตุที่มีผู้เสียชีวิต 2564–2568 เทียบธุรกิจตั้งใหม่ 2565–2568', 'แต่ละจุด = 1 อำเภอ · ยังไม่ปรับตามประชากร · ไม่เปลี่ยนตามพื้นที่ที่เลือก'),
    },
    'P4 Vulnerable Sectors': {
        'Rubber Price (baht/kg)': ((L, CH_Y, FULL_W, HALF_H), 'ราคายางแผ่นดิบชั้น 3 รายเดือน (บาท/กก.)', 'เส้นประ = 100 บาท · ที่มา: สำนักงานเศรษฐกิจการเกษตร'),
        'Fishing Vessels': ((L, ROW2_Y, 760, HALF_H), 'เรือประมงพาณิชย์ที่ได้รับอนุญาต (ลำ)', 'ที่มา: กรมประมง'),
        'Dissolved per 100 New': ((L + 776, ROW2_Y, FULL_W - 776, HALF_H), 'ธุรกิจเลิกกิจการต่อตั้งใหม่ 100 ราย แยกสาขา', 'ที่มา: กรมพัฒนาธุรกิจการค้า · ไม่เปลี่ยนตามพื้นที่ที่เลือก'),
    },
    'ความเชื่อ vs ข้อมูล': {
        'hypothesis': ((L, CH_Y, FULL_W, FOOT_Y - CH_Y), 'ความเชื่อของผู้วิจัย เทียบกับผลจากข้อมูล', 'ผลทั้งหมดเป็นความสัมพันธ์เชิงพรรณนา ไม่ใช่เหตุและผล'),
    },
}
NOFILTER = {'Disbursed %', 'Lethal Events (district)', 'Dissolved per 100 New'}
GREY_SERIES = {'Poverty % (Kingdom)', 'HH Income (Kingdom)'}

def fields(v):
    q = v.get('query', {}).get('queryState', {})
    return [pr.get('nativeQueryRef') for val in q.values() for pr in val.get('projections', [])]

for pid in os.listdir(PAGES):
    pdir = os.path.join(PAGES, pid)
    if not os.path.isdir(pdir):
        continue
    with open(os.path.join(pdir, 'page.json'), encoding='utf-8') as f:
        page = json.load(f)
    pname = page['displayName']
    if pname not in PAGE_INFO:
        continue
    vdir = os.path.join(pdir, 'visuals')
    slicer_id, nofilter_targets = None, []
    for vid in list(os.listdir(vdir)):
        fp = os.path.join(vdir, vid, 'visual.json')
        with open(fp, encoding='utf-8') as f:
            d = json.load(f)
        v = d['visual']
        vt, fl = v['visualType'], fields(v)
        if vt in ('textbox', 'card'):            # generated by this script -> rewritten below
            continue
        if vt == 'cardVisual':                     # old rho card -> replaced by KPI row (moved out, not deleted)
            trash = os.path.join(os.path.dirname(os.path.abspath(ROOT)), '_to_delete', 'pbip_visuals')
            os.makedirs(trash, exist_ok=True)
            shutil.move(os.path.join(vdir, vid), os.path.join(trash, vid)); continue
        if vt == 'slicer':
            slicer_id = vid
            d['position'].update({"x": 1560, "y": 14, "width": 344, "height": 60, "z": 900})
            o = v.setdefault('objects', {})
            o['data'] = [{"properties": {"mode": s("Dropdown")}}]
            o['header'] = [{"properties": {"show": lit("true"), "text": s("เลือกพื้นที่"), "fontColor": col(MUTED), "textSize": lit("9D")}}]
            o['items'] = [{"properties": {"textSize": lit("12D"), "fontColor": col(INK)}}]
            v['visualContainerObjects'] = {
                "background": [{"properties": {"show": lit("true"), "color": col('#FFFFFF'), "transparency": lit("0D")}}],
                "border": [{"properties": {"show": lit("false")}}],
                "dropShadow": [{"properties": {"show": lit("false")}}],
                "padding": [{"properties": {"top": lit("4D"), "bottom": lit("4D"), "left": lit("12D"), "right": lit("12D")}}]}
        else:
            key = next((k for k in LAYOUT[pname] if k in fl), None)
            if key is None:
                print('WARN unplaced visual', pname, vt, fl); continue
            (x, y, w, h), title, sub = LAYOUT[pname][key]
            d['position'].update({"x": x, "y": y, "width": w, "height": h})
            v['visualContainerObjects'] = {
                "title": [{"properties": {"show": lit("true"), "text": s(title)}}],
                "subTitle": [{"properties": {"show": lit("true"), "text": s(sub)}}]}
            o = v.setdefault('objects', {})
            if vt == 'lineChart':
                o['lineStyles'] = [{"properties": {"lineChartType": s("linear"), "strokeWidth": lit("3D"),
                                                   "showMarker": lit("true"), "markerSize": lit("6D")}}]
                if key == 'Rubber Price (baht/kg)':
                    o['labels'] = [{"properties": {"show": lit("false")}}]
                    o['lineStyles'][0]['properties']['showMarker'] = lit("false")
            if vt in ('clusteredColumnChart', 'clusteredBarChart', 'lineChart') and key != 'Rubber Price (baht/kg)':
                o['labels'] = [{"properties": {"show": lit("true")}}]
            if 'y1AxisReferenceLine' in o:
                rl = o['y1AxisReferenceLine'][0]['properties']
                rl.update({"lineColor": col(GOLD), "style": s("dashed"), "width": lit("2D"), "transparency": lit("0D")})
            dp = []
            for meas in fl:
                if meas in GREY_SERIES:
                    dp.append({"properties": {"fill": col('#B9B2A6')}, "selector": {"metadata": f"_Measures.{meas}"}})
            if key == 'Dissolved per 100 New':
                o['legend'] = [{"properties": {"show": lit("true"), "position": s("TopLeft")}}]
            if dp:
                o['dataPoint'] = dp
            if key == 'hypothesis':
                o['grid'] = [{"properties": {"textSize": lit("15D"), "rowPadding": lit("8D")}}]
            if key == 'Lethal Events (district)':
                o['dataPoint'] = [{"properties": {"fill": col('#2D63B0')}}]
                o['markers'] = [{"properties": {"markerSize": lit("8D")}}] if False else o.get('markers', [])
                if not o['markers']:
                    o.pop('markers')
            if key in NOFILTER:
                nofilter_targets.append(vid)
        with open(fp, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=2)

    # header band
    kicker, question = PAGE_INFO[pname]
    write_visual(pid, uid(pid + 'header'), (0, 0, 1920, HEADER_H), textbox([
        {"textRuns": [run('THE SOUTHERN FRONTIER REVIEW', 10, GOLD, True), run('   ·   ปัตตานี ยะลา นราธิวาส และ 4 อำเภอของสงขลา · 2564–2568', 10, '#AEB6C2')]},
        {"textRuns": [run(kicker, 22, '#FFFFFF', True, 'Georgia'), run('    ' + question, 13, '#D5DAE1')]},
    ], bg=NAVY), 0)
    # KPI cards
    for i, (meas, size) in enumerate(KPIS[pname]):
        cid = uid(pid + 'kpi' + meas)
        write_visual(pid, cid, (M + i * (KPI_W + M), KPI_Y, KPI_W, KPI_H), card(meas, size), 10 + i)
        if pname != 'ความเชื่อ vs ข้อมูล' and meas.startswith(('Spearman', 'จำนวนอำเภอ', 'เลิกกิจการ')):
            nofilter_targets.append(cid)
    # footer
    write_visual(pid, uid(pid + 'footer'), (M, FOOT_Y + 4, FULL_W, 24), textbox([
        {"textRuns": [run('Irfan Aroma · Al Qasimia University · ข้อมูลทางการจาก 19 แหล่ง (2,147 ค่า / 113 ตัวชี้วัด) · '
                          'ผลเป็นความสัมพันธ์เชิงพรรณนา ไม่ใช่เหตุและผล · github.com/aifan22404-crypto/sbp-dashboard', 9, GREY)]},
    ]), 1)

    # interactions: slicer must not filter cross-area visuals
    if slicer_id:
        page['visualInteractions'] = [{"source": slicer_id, "target": t, "type": "NoFilter"} for t in nofilter_targets]
    with open(os.path.join(pdir, 'page.json'), 'w', encoding='utf-8') as f:
        json.dump(page, f, ensure_ascii=False, indent=2)
    print('restyled', pname, 'nofilter:', len(nofilter_targets))
print('done')
