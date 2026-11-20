"""Append RF25-RF26 to red_flags_log and new sources to source_log. Usage: python log_updates.py <red_flags.xlsx> <source_log.xlsx>"""
import sys, openpyxl
rf, sl = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(rf); ws = wb['red_flags']
have = {r[0] for r in ws.iter_rows(min_row=2, values_only=True)}
rows = [
 ['RF25', '2026-11-20', '🟡 Medium', 'A. Inconsistency', 'DOPA district population (Narathiwat GD Catalog)', 'Narathiwat', '2563–2568',
  'Su-ngai Kolok 40,822 / 40,583 / 40,160 / 39,984 (2563–2566) → 78,853 (2567) → 79,048 (2568); also Bacho 57,188 → 54,702 and Yi-ngo 48,416 → 51,482 (2567→2568)',
  'A district does not double in one year; other districts change <1%/yr',
  '2563–2566 probably counted only the municipality (or only the non-municipal part); 2568 may reflect a boundary/registration-office reassignment',
  '—', 'Compare with stat.bora.dopa.go.th district table (blocked from abroad) or the Narathiwat provincial office',
  'Use 2567 as Narathiwat denominator; 2566 for other provinces. Do not compute population growth for Kolok.', 'Open', '—'],
 ['RF26', '2026-11-20', '🟡 Medium', 'C. Comparability', 'MOTS domestic tourism 2562 vs 2564–2568', 'PTN, YLA, NWT, SKA', '2562 vs 2568',
  'Visitors 2568/2562: Pattani 2.28×, Yala 2.95×, Narathiwat 1.76×, Songkhla 0.96×',
  'Three border provinces far above pre-COVID while the main tourism province is below',
  'Real recovery + Malaysian day-trips; or a change in MOTS estimation method between 2562 and 2565',
  '—', 'Check MOTS methodology notes for 2565+ provincial estimates; compare hotel guest arrivals (survey-based) instead of visitors',
  'Quote 2562 as context only ("เทียบแบบระวัง"); do not state growth multiples as fact.', 'Open', '—'],
]
for r in rows:
    if r[0] not in have: ws.append(r)
wb.save(rf)
wb = openpyxl.load_workbook(sl); ws = wb['source_log']
have = {r[0] for r in ws.iter_rows(min_row=2, values_only=True)}
src = [
 ['POP01', 'Population_district_DOPA_2562-2568_clean.xlsx', '02_clean', 'กรมการปกครอง (ทะเบียนราษฎร) ผ่าน GD Catalog จังหวัด', 'narathiwat.gdcatalog.go.th/dataset/12_01 ; pattani.gdcatalog.go.th/dataset/population ; yala.gdcatalog.go.th/dataset/population_01 ; songkhla.gdcatalog.go.th/dataset/dataset_03_46', '2026-11-20', '2562–2568 (varies by province)', 'District', 'PTN,YLA,NWT,SKA-4', 'csv → xlsx', 'P2,P3', 'Pass with RF25', 'Pattani municipalities added to host district; Tanyong → Yaring assumed. stat.bora.dopa.go.th blocked from outside Thailand.'],
 ['TO02', 'Tourism_baseline_2561-2562_clean.xlsx', '02_clean', 'กระทรวงการท่องเที่ยวและกีฬา', 'mots.go.th/download/article/article_20201105143022.xlsx', '2026-11-20', '2561–2562', 'Province', 'PTN,YLA,NWT,SKA', 'xlsx', 'P1,P3', 'Pass with RF26', 'Pre-COVID baseline only; not in master panel.'],
 ['BUD03', 'Budget_southern_by_agency_2551-2568_clean.xlsx', '02_clean', 'Rocket Media Lab (compiled from Budget Acts 2551–2568)', 'rocketmedialab.co/deep-south-budget/', '2026-11-20', 'FY2551–2568 cumulative; FY2569 draft', 'Region (SBP)', '—', 'web article', 'P3', 'Secondary source', 'By receiving agency, not mission; not yet checked against Budget Bureau documents.'],
]
for r in src:
    if r[0] not in have: ws.append(r)
wb.save(sl); print('ok')
