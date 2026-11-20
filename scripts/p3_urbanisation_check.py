"""P3 robustness (20 Nov 2026): does the per-capita violence–new-firm association survive controls for urbanisation?
Urbanisation proxy = population density (registered population / district land area).
Area: chingchai/opengisdata-thailand districts.geojson (area_sqkm; CC BY 4.0; from Thailand Geography JSON, MIT).
Usage: python -I scripts/p3_urbanisation_check.py <districts.geojson>"""
import sys, json, math
import numpy as np, pandas as pd
from scipy import stats

g = json.load(open(sys.argv[1], encoding='utf-8'))
a = pd.DataFrame([f['properties'] for f in g['features'] if f['properties']['pro_en'] in ('Pattani', 'Yala', 'Narathiwat', 'Songkhla')])
a['district'] = a.amp_th.replace({'กรงปีนัง': 'กรงปินัง'})
a = a.rename(columns={'pro_en': 'province'})[['province', 'district', 'amp_en', 'area_sqkm']]
fd = pd.read_csv('powerbi/csv/fact_district.csv', encoding='utf-8-sig')
m = fd.merge(a, on=['province', 'district'], how='left')
assert m.area_sqkm.notna().all()
m['density'] = m.population / m.area_sqkm

def rk(x): return pd.Series(x).rank().values
def ci(r, n, k):
    z, se = math.atanh(r), 1 / math.sqrt(n - 3 - k)
    return math.tanh(z - 1.96 * se), math.tanh(z + 1.96 * se)
def partial(x, y, Z):
    """Spearman partial correlation: residualise ranks of x and y on ranks of controls Z (OLS), correlate residuals."""
    X = np.column_stack([np.ones(len(x))] + [rk(z) for z in Z])
    ex = rk(x) - X @ np.linalg.lstsq(X, rk(x), rcond=None)[0]
    ey = rk(y) - X @ np.linalg.lstsq(X, rk(y), rcond=None)[0]
    r = np.corrcoef(ex, ey)[0, 1]
    return r, *ci(r, len(x), len(Z))
def perm_p(x, y, Z, B=10000, seed=7):
    X = np.column_stack([np.ones(len(x))] + [rk(z) for z in Z])
    ex = rk(x) - X @ np.linalg.lstsq(X, rk(x), rcond=None)[0]
    ey = rk(y) - X @ np.linalg.lstsq(X, rk(y), rcond=None)[0]
    obs = abs(np.corrcoef(ex, ey)[0, 1]); rng = np.random.default_rng(seed)
    return np.mean([abs(np.corrcoef(ex, rng.permutation(ey))[0, 1]) >= obs for _ in range(B)])

L, F = m.lethal_per_100k, m.new_firms_per_10k
res = []
def add(label, r, lo, hi, n, p=None): res.append({'model': label, 'rho': round(r, 3), 'ci_low': round(lo, 2), 'ci_high': round(hi, 2), 'n': n, 'perm_p': None if p is None else round(p, 4)})
r = stats.spearmanr(L, F)[0]; add('ต่อประชากร (ไม่คุม)', r, *ci(r, 37, 0), 37, perm_p(L, F, []))
r, lo, hi = partial(L, F, [m.density]); add('ต่อประชากร คุมความหนาแน่นประชากร', r, lo, hi, 37, perm_p(L, F, [m.density]))
r, lo, hi = partial(L, F, [m.density, m.is_muang]); add('ต่อประชากร คุมความหนาแน่น + อำเภอเมือง', r, lo, hi, 37, perm_p(L, F, [m.density, m.is_muang]))
pd_ = pd.get_dummies(m.province, drop_first=True).astype(float)
r, lo, hi = partial(L, F, [m.density] + [pd_[c] for c in pd_]); add('ต่อประชากร คุมความหนาแน่น + จังหวัด', r, lo, hi, 37, perm_p(L, F, [m.density] + [pd_[c] for c in pd_]))
nm = m[m.is_muang == 0]
r, lo, hi = partial(nm.lethal_per_100k, nm.new_firms_per_10k, [nm.density]); add('ไม่นับอำเภอเมือง คุมความหนาแน่น', r, lo, hi, len(nm), perm_p(nm.lethal_per_100k, nm.new_firms_per_10k, [nm.density]))
out = pd.DataFrame(res)
print(out.to_string(index=False))
print('density vs lethal/100k rho', round(stats.spearmanr(m.density, L)[0], 2), '| density vs firms/10k rho', round(stats.spearmanr(m.density, F)[0], 2))
m[['province', 'district', 'amp_en', 'area_sqkm', 'population', 'density', 'lethal_per_100k', 'new_firms_per_10k', 'is_muang']].round(2).to_csv('/tmp/p3_density.csv', index=False, encoding='utf-8-sig')
out.to_csv('/tmp/p3_robustness.csv', index=False, encoding='utf-8-sig')
