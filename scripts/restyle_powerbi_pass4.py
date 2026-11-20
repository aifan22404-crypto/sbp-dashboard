"""Pass 4: fixed column widths + wrap for the beliefs table. Usage: python restyle_pass4.py <pbip_root>"""
import json, os, sys, glob
ROOT = sys.argv[1]
def lit(v): return {"expr": {"Literal": {"Value": v}}}
W = {'no': 60, 'P': 70, 'hypothesis': 420, 'result': 140, 'evidence': 1140}
for fp in glob.glob(os.path.join(ROOT, 'SBP_Dashboard.Report', 'definition', 'pages', '*', 'visuals', '*', 'visual.json')):
    d = json.load(open(fp, encoding='utf-8')); v = d['visual']
    if v['visualType'] != 'tableEx': continue
    o = v.setdefault('objects', {})
    o['columnWidth'] = [{"properties": {"value": lit(f"{w}D")}, "selector": {"metadata": f"fact_hypothesis.{c}"}} for c, w in W.items()]
    o['columnHeaders'] = [{"properties": {"autoSizeColumnWidth": lit("false"), "wordWrap": lit("true")}}]
    o['values'] = [{"properties": {"wordWrap": lit("true")}}]
    json.dump(d, open(fp, 'w', encoding='utf-8'), ensure_ascii=False, indent=2); print('table updated')
