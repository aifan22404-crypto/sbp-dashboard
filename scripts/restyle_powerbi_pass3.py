"""Pass 3 (20 Nov 2026): district population columns, per-capita measures, P3 scatter/KPIs, P1 non-public LQ KPI.
Usage: python restyle_pass3.py <pbip_root>"""
import json, os, re, sys, glob

ROOT = sys.argv[1]
SM = os.path.join(ROOT, 'SBP_Dashboard.SemanticModel', 'definition', 'tables')
PAGES = os.path.join(ROOT, 'SBP_Dashboard.Report', 'definition', 'pages')
def lit(v): return {"expr": {"Literal": {"Value": v}}}
def s(t): return lit("'" + t.replace("'", "''") + "'")

# ---- fact_district: new columns + types ------------------------------------
fp = os.path.join(SM, 'fact_district.tmdl')
t = open(fp, encoding='utf-8').read()
NEWCOLS = [('pop_year', 'int64', '0', 'none', 'Int64.Type'), ('population', 'int64', '#,0', 'sum', 'Int64.Type'),
           ('lethal_per_100k', 'double', '0.0', 'none', 'type number'), ('new_firms_per_10k', 'double', '0.0', 'none', 'type number'),
           ('rank_lethal_per_100k', 'double', '0.0', 'none', 'type number'), ('rank_new_firms_per_10k', 'double', '0.0', 'none', 'type number')]
if 'column population' not in t:
    block = ''.join(f"\tcolumn {c}\n\t\tdataType: {dt}\n\t\tformatString: {fs}\n\t\tsummarizeBy: {sb}\n\t\tsourceColumn: {c}\n\n\t\tannotation SummarizationSetBy = Automatic\n\n"
                    for c, dt, fs, sb, _ in NEWCOLS)
    i = t.index('\tpartition ')
    t = t[:i] + block + t[i:]
    types = ''.join(f', {{"{c}", {m}}}' for c, _, _, _, m in NEWCOLS)
    t = t.replace('{"dissolved_per_100_new", type number}}', '{"dissolved_per_100_new", type number}' + types + '}')
    assert '"population", Int64.Type' in t
    open(fp, 'w', encoding='utf-8').write(t)

# ---- measures -----------------------------------------------------------------
fp = os.path.join(SM, '_Measures.tmdl')
t = open(fp, encoding='utf-8').read()
RANK = ('VAR n = COUNTROWS ( fact_district ) VAR mx = AVERAGE ( fact_district[{a}] ) VAR my = AVERAGE ( fact_district[{b}] ) '
        'VAR sxy = SUMX ( fact_district, ( fact_district[{a}] - mx ) * ( fact_district[{b}] - my ) ) '
        'VAR sxx = SUMX ( fact_district, ( fact_district[{a}] - mx ) ^ 2 ) VAR syy = SUMX ( fact_district, ( fact_district[{b}] - my ) ^ 2 ) '
        'RETURN IF ( n >= 5, DIVIDE ( sxy, SQRT ( sxx * syy ) ) )')
PUBLIC = '{ "การศึกษา", "บริหารราชการและป้องกันประเทศ", "สุขภาพและสังคมสงเคราะห์" }'
# rename old KPI measures
t = t.replace("measure 'Spearman ρ ความรุนแรง vs ธุรกิจใหม่' =", "measure 'Spearman ρ นับดิบ (ไม่ปรับประชากร)' =")
NEW = [
    ('Spearman rho (per capita)', RANK.format(a='rank_lethal_per_100k', b='rank_new_firms_per_10k'), '0.00'),
    ('Spearman ρ ต่อประชากร (37 อำเภอ)', 'CALCULATE ( [Spearman rho (per capita)], REMOVEFILTERS ( dim_area ), REMOVEFILTERS ( fact_district ) )', '0.00'),
    ('เหตุมีผู้เสียชีวิตต่อแสนคน', 'DIVIDE ( SUM ( fact_district[ucdp_lethal_events_2564_2568] ), SUM ( fact_district[population] ) ) * 100000', '0.0'),
    ('ธุรกิจใหม่ต่อหมื่นคน', 'DIVIDE ( SUM ( fact_district[new_firms_2565_2568] ), SUM ( fact_district[population] ) ) * 10000', '0.0'),
    ('สาขาเอกชนที่ LQ สูงสุด', f'VAR t = TOPN ( 1, FILTER ( VALUES ( fact_lq[sector_th] ), NOT ( fact_lq[sector_th] IN {PUBLIC} ) ), [LQ], DESC ) RETURN CONCATENATEX ( t, fact_lq[sector_th], ", " )', None),
    ('LQ สูงสุด (ไม่รวมภาครัฐ)', f'MAXX ( FILTER ( VALUES ( fact_lq[sector_th] ), NOT ( fact_lq[sector_th] IN {PUBLIC} ) ), [LQ] )', '0.00'),
]
existing = set(re.findall(r"\tmeasure '((?:[^']|'')+)'", t))
add = ''
for name, expr, fmt in NEW:
    if name in existing: continue
    add += f"\tmeasure '{name}' = {expr}\n" + (f"\t\tformatString: {fmt}\n" if fmt else '') + '\n'
i = re.search(r'(?m)^\t(column|partition) ', t).start()
t = t[:i] + add + t[i:]
open(fp, 'w', encoding='utf-8').write(t)

# ---- visuals -------------------------------------------------------------------
def proj(name):
    return {"field": {"Measure": {"Expression": {"SourceRef": {"Entity": "_Measures"}}, "Property": name}},
            "queryRef": f"_Measures.{name}", "nativeQueryRef": name}
SWAP_CARD = {'สาขาที่ LQ สูงสุด': 'สาขาเอกชนที่ LQ สูงสุด', 'LQ สูงสุด': 'LQ สูงสุด (ไม่รวมภาครัฐ)',
             'Spearman ρ ความรุนแรง vs ธุรกิจใหม่': 'Spearman ρ ต่อประชากร (37 อำเภอ)',
             'จำนวนอำเภอที่วิเคราะห์': 'Spearman ρ นับดิบ (ไม่ปรับประชากร)'}
for fp in glob.glob(os.path.join(PAGES, '*', 'visuals', '*', 'visual.json')):
    d = json.load(open(fp, encoding='utf-8')); v = d['visual']; changed = False
    q = v.get('query', {}).get('queryState', {})
    if v['visualType'] == 'card':
        pr = q['Values']['projections'][0]
        if pr['nativeQueryRef'] in SWAP_CARD:
            q['Values']['projections'] = [proj(SWAP_CARD[pr['nativeQueryRef']])]; changed = True
            v.get('query', {}).pop('sortDefinition', None)
    if v['visualType'] == 'scatterChart':
        q['X']['projections'] = [dict(proj('ธุรกิจใหม่ต่อหมื่นคน'), displayName='ธุรกิจใหม่ต่อหมื่นคน')]
        q['Y']['projections'] = [dict(proj('เหตุมีผู้เสียชีวิตต่อแสนคน'), displayName='เหตุมีผู้เสียชีวิตต่อแสนคน')]
        v.get('query', {}).pop('sortDefinition', None)
        o = v.setdefault('objects', {})
        o['categoryAxis'] = [{"properties": {"showAxisTitle": lit("true"), "titleText": s("ธุรกิจตั้งใหม่ 2565–2568 ต่อประชากรหมื่นคน")}}]
        o['valueAxis'] = [{"properties": {"showAxisTitle": lit("true"), "titleText": s("เหตุที่มีผู้เสียชีวิต 2564–2568 ต่อแสนคน")}}]
        v['visualContainerObjects']['title'] = [{"properties": {"show": lit("true"), "text": s("37 อำเภอ: เหตุที่มีผู้เสียชีวิต เทียบธุรกิจตั้งใหม่ ต่อประชากร")}}]
        v['visualContainerObjects']['subTitle'] = [{"properties": {"show": lit("true"), "text": s("แต่ละจุด = 1 อำเภอ · ประชากรทะเบียนราษฎร 2566 (นราธิวาส 2567) · ρ นับดิบ −0.06 แต่ต่อประชากร −0.47 · ความสัมพันธ์ ไม่ใช่สาเหตุ")}}]
        changed = True
    if changed:
        json.dump(d, open(fp, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
        print('updated', v['visualType'], os.path.basename(os.path.dirname(fp)))
print('pass3 done')
