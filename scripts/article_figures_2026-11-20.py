"""Figures for the article drafts (EN + TH labels). Output: notes/figures/*.png"""
import pandas as pd, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
fm.fontManager.addfont('/usr/share/fonts/opentype/tlwg/Loma.otf'); fm.fontManager.addfont('/usr/share/fonts/opentype/tlwg/Loma-Bold.otf')
INK, MUTED, GRID = '#14202E', '#6B7280', '#E5E1D8'
COL = {'Pattani': '#2D63B0', 'Yala': '#1A9A74', 'Narathiwat': '#C0392B', 'Songkhla': '#A8844A', 'Whole Kingdom': '#9A9488'}
TH = {'Pattani': 'ปัตตานี', 'Yala': 'ยะลา', 'Narathiwat': 'นราธิวาส', 'Songkhla': 'สงขลา 4 อำเภอ', 'Whole Kingdom': 'ทั้งประเทศ'}
P = pd.read_csv('data/master_panel/SBP_master_panel_long.csv', encoding='utf-8-sig')
fd = pd.read_csv('powerbi/csv/fact_district.csv', encoding='utf-8-sig')

def style(ax):
    for s in ('top', 'right'): ax.spines[s].set_visible(False)
    for s in ('left', 'bottom'): ax.spines[s].set_color(MUTED)
    ax.tick_params(colors=INK, labelsize=9); ax.grid(axis='y', color=GRID, lw=0.8); ax.set_axisbelow(True)

for lang in ('en', 'th'):
    plt.rcParams['font.family'] = ['Loma', 'DejaVu Sans'] if lang == 'th' else ['DejaVu Sans']
    # Figure 1: poverty
    fig, ax = plt.subplots(figsize=(7, 3.6), dpi=200)
    pv = P[P.indicator_code == 'poverty_rate_pct']
    for a in ['Pattani', 'Yala', 'Narathiwat', 'Whole Kingdom']:
        s = pv[pv.area == a].sort_values('year_BE')
        x = s.year_BE if lang == 'th' else s.year_CE
        ax.plot(x, s.value, marker='o', ms=4, lw=2.2 if a != 'Whole Kingdom' else 2, color=COL[a], ls='-' if a != 'Whole Kingdom' else '--')
        ax.annotate(TH[a] if lang == 'th' else a, (x.iloc[-1], s.value.iloc[-1]), xytext=(6, 0), textcoords='offset points', va='center', fontsize=8.5, color=INK)
    ax.set_ylim(0, 33); ax.set_xticks(list(x)); ax.set_xlim(x.min() - 0.3, x.max() + 0.9)
    ax.set_ylabel('สัดส่วนคนจน (%)' if lang == 'th' else 'Poverty rate (%)', fontsize=9, color=INK); style(ax)
    fig.tight_layout(); fig.savefig(f'notes/figures/fig1_poverty_{lang}.png'); plt.close(fig)
    # Figure 2: per-capita scatter
    fig, ax = plt.subplots(figsize=(7, 3.9), dpi=200)
    for prov in ['Pattani', 'Yala', 'Narathiwat', 'Songkhla']:
        g = fd[fd.province == prov]
        ax.scatter(g.lethal_per_100k, g.new_firms_per_10k, s=34, color=COL[prov], edgecolor='white', lw=0.8, label=TH[prov] if lang == 'th' else ('Songkhla (4 districts)' if prov == 'Songkhla' else prov), zorder=3)
    ax.set_xlabel('เหตุที่มีผู้เสียชีวิต 2564–2568 ต่อประชากรแสนคน' if lang == 'th' else 'Lethal events 2021–2025 per 100,000 people', fontsize=9, color=INK)
    ax.set_ylabel('ธุรกิจตั้งใหม่ 2565–2568\nต่อประชากรหมื่นคน' if lang == 'th' else 'New firms 2022–2025\nper 10,000 people', fontsize=9, color=INK)
    ax.legend(frameon=False, fontsize=8.5, loc='upper right'); style(ax); ax.grid(axis='x', color=GRID, lw=0.8)
    ax.text(0.99, 0.40, 'Spearman ρ = −0.47\n(95% CI −0.69 ถึง −0.18)' if lang == 'th' else 'Spearman ρ = −0.47\n(95% CI −0.69 to −0.18)', transform=ax.transAxes, ha='right', fontsize=8.5, color=MUTED)
    fig.tight_layout(); fig.savefig(f'notes/figures/fig2_district_percapita_{lang}.png'); plt.close(fig)
print('ok')
