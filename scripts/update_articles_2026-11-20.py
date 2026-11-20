"""Update TH/EN article drafts (20 Nov 2026): per-capita district result, 2019 tourism baseline, budget by agency,
and replace 'สมมติฐาน' with 'ความคาดหวังก่อนดูข้อมูล' in the Thai draft. Usage: python update_articles.py th.docx en.docx out_dir"""
import sys, os, copy
from docx import Document

def set_par(p, text, bold_lead=None):
    """Rewrite a paragraph keeping the formatting of its first run(s). If bold_lead given, first run bold = lead."""
    runs = p.runs
    if bold_lead is not None:
        lead, rest = runs[0], (runs[1] if len(runs) > 1 else None)
        lead.text = bold_lead; lead.bold = True
        if rest is None:
            rest = p.add_run(); rest.bold = False
        rest.text = text
        for r in runs[2:]: r.text = ''
    else:
        runs[0].text = text
        for r in runs[1:]: r.text = ''

def find(doc, start):
    hits = [p for p in doc.paragraphs if p.text.startswith(start)]
    assert len(hits) == 1, (start, len(hits))
    return hits[0]

def set_cell(cell, text):
    p = cell.paragraphs[0]; set_par(p, text)
    for extra in cell.paragraphs[1:]:
        for r in extra.runs: r.text = ''

def table_row(doc, key):
    for t in doc.tables:
        for r in t.rows:
            if r.cells[0].text.strip() == key and len(r.cells) >= 4:
                yield r

th, en, out = sys.argv[1], sys.argv[2], sys.argv[3]
FOURTH_EN = ('Fourth, across 37 districts, raw counts of lethal violent events and new business registrations are uncorrelated (ρ = −0.06), '
             'but per head of population they are moderately negatively correlated (ρ = −0.47, 95% CI −0.69 to −0.18).')
TALLY_EN = ('three were consistent with the data, four partly consistent, and one not supported. All results are descriptive associations, not causal estimates. '
            'The main limits are a pre-COVID tourism baseline that may not be comparable, a district analysis of only 37 districts without controls for urbanisation, and no fieldwork.')

# ---------------------------------------------------------------- Thai
d = Document(th)
p = find(d, 'Four findings stand out.')
set_par(p, p.text.split('Fourth, across 37 districts')[0] + FOURTH_EN)
set_par(find(d, 'Of the eight hypotheses,'), 'Of the eight prior expectations, ' + TALLY_EN.replace('All results are descriptive associations, not causal estimates.', 'Results are correlational and not causal.'))
p = find(d, 'ข้อค้นพบหลักมี 4 ข้อ')
t = p.text.split('(3)')[0]
set_par(p, t + '(3) จำนวนผู้เยี่ยมเยือนเพิ่มขึ้นทุกปีหลังโควิด-19 และนราธิวาสปี 2568 (1.24 ล้านคน) สูงกว่าปี 2562 (703,810 คน) จึงไม่สนับสนุนความเชื่อว่าข่าวความไม่สงบทำให้คนไม่มาเที่ยว แม้การเทียบกับปี 2562 ต้องทำอย่างระวัง '
         '(4) ใน 37 อำเภอ จำนวนนับดิบของเหตุรุนแรงถึงชีวิตกับธุรกิจตั้งใหม่แทบไม่สัมพันธ์กัน (ρ = −0.06) แต่เมื่อคิดต่อประชากร อำเภอที่มีเหตุต่อหัวสูงมีธุรกิจตั้งใหม่ต่อหัวต่ำ (ρ = −0.47 ช่วงความเชื่อมั่น 95% −0.69 ถึง −0.18)')
set_par(find(d, 'จากสมมติฐาน 8 ข้อ สอดคล้อง'), 'จากความคาดหวังก่อนดูข้อมูล 8 ข้อ สอดคล้องกับข้อมูล 3 ข้อ ตรงบางส่วน 4 ข้อ และไม่ตรง 1 ข้อ ผลทั้งหมดเป็นความสัมพันธ์ ไม่ใช่ความเป็นเหตุเป็นผล')
set_par(find(d, 'เหตุการณ์ความไม่สงบเพิ่มขึ้นในปี 2567–2568'), 'เหตุการณ์ความไม่สงบเพิ่มขึ้นในปี 2567–2568 เมื่อนับจำนวนดิบ ระดับอำเภอยังไม่เห็นความสัมพันธ์กับการตั้งธุรกิจ แต่เมื่อคิดต่อประชากร อำเภอที่มีเหตุรุนแรงต่อหัวสูงมีธุรกิจตั้งใหม่ต่อหัวต่ำกว่า')
set_par(find(d, 'ความสัมพันธ์กับการตั้งธุรกิจ'), ' ค่าสหสัมพันธ์ของ Spearman ระหว่างจำนวนเหตุรุนแรงถึงชีวิตกับจำนวนนิติบุคคลตั้งใหม่ใน 37 อำเภอเท่ากับ −0.06 (ช่วงความเชื่อมั่น 95% ประมาณ −0.37 ถึง 0.27) เมื่อตัดอำเภอเมืองออกได้ 0.02 และเมื่อใช้สถิติทางการเฉพาะ 13 อำเภอของนราธิวาสได้ −0.26 แต่จำนวนนับดิบปนผลของขนาดอำเภอ เพราะจำนวนธุรกิจตั้งใหม่สัมพันธ์กับประชากรสูงมาก (ρ = 0.91) ขณะที่เหตุรุนแรงแทบไม่สัมพันธ์กับประชากร (ρ = 0.08)', bold_lead='ความสัมพันธ์กับการตั้งธุรกิจ')
set_par(find(d, 'จำนวนธุรกิจยังไม่ได้ปรับตามจำนวนประชากร'), ' เมื่อใช้ประชากรจากทะเบียนราษฎรปี 2566 (นราธิวาสใช้ปี 2567 เพราะตัวเลขสุไหงโก-ลกปีก่อนหน้านับไม่ครบ) ค่าสหสัมพันธ์ระหว่างเหตุรุนแรงถึงชีวิตต่อแสนคนกับธุรกิจตั้งใหม่ต่อหมื่นคนเท่ากับ −0.47 (ช่วงความเชื่อมั่น 95% −0.69 ถึง −0.18) เมื่อตัดอำเภอเมืองออกได้ −0.40 (−0.65 ถึง −0.07) และทิศทางเป็นลบในทุกจังหวัด อย่างไรก็ตาม อัตราสองตัวที่ใช้ตัวหารเดียวกันอาจสัมพันธ์กันเองได้บางส่วน เมื่อใช้ partial correlation ที่คุมประชากรแทนได้ −0.31 (−0.58 ถึง 0.02) ซึ่งยังเป็นลบแต่ไม่ชัดเท่า ผลนี้ยังไม่ได้คุมความเป็นเมือง ถนน หรือปัจจัยอื่น จึงเป็นความสัมพันธ์ ไม่ใช่ผลกระทบ', bold_lead='ปรับตามประชากร')
set_par(find(d, 'งบประมาณ ข้อมูลงบประมาณที่รวบรวมได้'), ' ข้อมูลงบประมาณจังหวัดที่รวบรวมได้แยกตามประเภทรายจ่าย (ประจำ ลงทุน) ไม่ได้แยกตามภารกิจ แหล่งรองจาก Rocket Media Lab ซึ่งรวบรวม พ.ร.บ.งบประมาณ ปี 2551–2568 พบว่างบแก้ปัญหาชายแดนใต้สะสม 267,673 ล้านบาท ไปที่ กอ.รมน. 30.5% กระทรวงกลาโหม 16.0% และสำนักงานตำรวจแห่งชาติ 6.3% รวมราว 53% แต่เป็นการแยกตามหน่วยงานที่รับงบ ไม่ใช่ภารกิจ เพราะ กอ.รมน. ทำโครงการพัฒนาด้วย และร่างงบปี 2569 แผนบูรณาการลดลงราว 74%', bold_lead='งบประมาณ')
set_par(find(d, 'จากสมมติฐาน 8 ข้อ ข้อมูลสอดคล้อง'), 'จากความคาดหวังก่อนดูข้อมูล 8 ข้อ ข้อมูลสอดคล้อง 3 ข้อ ตรงบางส่วน 4 ข้อ และไม่ตรง 1 ข้อ (ปรับปรุง 20 พ.ย. 2569 หลังเพิ่มประชากรรายอำเภอ ข้อมูลท่องเที่ยวปี 2562 และงบแยกตามหน่วยงาน)')
set_par(find(d, 'ข้อ 3 ท่องเที่ยว'), ' ตัวเลขไม่สนับสนุนว่าคนไม่มา ผู้เยี่ยมเยือนนราธิวาสปี 2568 (1.24 ล้านคน) สูงกว่าปี 2562 ก่อนโควิด (703,810 คน) แต่ต้องเทียบอย่างระวัง เพราะปัตตานีและยะลาก็สูงกว่าปี 2562 ถึง 2.3 และ 2.9 เท่า ขณะที่สงขลาต่ำกว่าเล็กน้อย ซึ่งอาจมาจากการเปลี่ยนวิธีประมาณการ นอกจากนี้ผู้เยี่ยมเยือนต่างชาติส่วนใหญ่เป็นชาวมาเลเซียที่คุ้นเคยกับพื้นที่ ข่าวอาจมีผลกับนักท่องเที่ยวไทยจากภาคอื่นมากกว่า ซึ่งข้อมูลชุดนี้แยกไม่ได้', bold_lead='ข้อ 3 ท่องเที่ยว')
set_par(find(d, 'ข้อ 5 งบประมาณ'), ' ตรงบางส่วน หน่วยงานด้านความมั่นคงได้รับราว 53% ของงบแก้ปัญหาชายแดนใต้สะสมปี 2551–2568 (Rocket Media Lab) ซึ่งไปทางเดียวกับ Suwanthanin และ Sirisunhirun (2026) ที่รายงานว่างบสะสมกระจุกอยู่ที่ด้านความมั่นคง แต่ตัวเลขนี้แยกตามหน่วยงาน ไม่ใช่ภารกิจ ยังไม่ได้ตรวจกับเอกสารต้นฉบับ และงบล่าสุดลดลง ข้อมูลของงานนี้เองเห็นว่าสาขาบริหารราชการและป้องกันประเทศมีสัดส่วน 17.39% ของ GPP ปัตตานี เทียบกับ 6.09% ของประเทศ', bold_lead='ข้อ 5 งบประมาณ')
set_par(find(d, 'ข้อ 6 เหตุการณ์กับธุรกิจใหม่'), ' ตรงบางส่วน เมื่อนับดิบไม่พบความสัมพันธ์ แต่เมื่อคิดต่อประชากร อำเภอที่มีเหตุรุนแรงต่อหัวสูงมีธุรกิจตั้งใหม่ต่อหัวต่ำ (ρ = −0.47) ซึ่งไปทางเดียวกับ Chongwilaikasaem และคณะ (2022) ที่พบว่าพื้นที่นี้มีธุรกิจใหม่น้อยกว่าจังหวัดเทียบในช่วงปี 2547–2560 ต่างกันที่งานนี้เทียบอำเภอภายในพื้นที่ในช่วงปี 2564–2568 ค่าที่ได้ขึ้นกับวิธี (partial −0.31) และอำเภอที่มีเหตุต่อหัวสูงมักเป็นอำเภอชนบท จึงยังแยกผลของความรุนแรงออกจากความเป็นชนบทไม่ได้', bold_lead='ข้อ 6 เหตุการณ์กับธุรกิจใหม่')
set_par(find(d, 'ไม่ปรับตามประชากรระดับอำเภอ'), ' ใช้ประชากรจากทะเบียนราษฎร ซึ่งอาจต่างจากประชากรที่อยู่จริง ยังไม่ได้คุมความเป็นเมืองหรือการเข้าถึงตลาด และมีเพียง 37 อำเภอ', bold_lead='ปรับตามประชากรแล้ว แต่ยังไม่ได้คุมปัจจัยอื่น')
set_par(find(d, 'ไม่มีฐานก่อนโควิด-19'), ' ข้อมูลท่องเที่ยวหลักเริ่มปี 2564 ฐานปี 2562 ที่เพิ่มเข้ามาแสดงว่าสามจังหวัดสูงกว่าก่อนโควิดมาก ขณะที่สงขลาต่ำกว่า จึงอาจมีการเปลี่ยนวิธีประมาณการ', bold_lead='ฐานก่อนโควิด-19 อาจเทียบกันไม่ได้')
set_par(find(d, 'งานขั้นต่อไปคือปรับข้อมูลรายอำเภอ'), 'งานขั้นต่อไปคือคุมปัจจัยความเป็นเมืองในการวิเคราะห์รายอำเภอ ตรวจวิธีประมาณการท่องเที่ยวของกระทรวงการท่องเที่ยวและกีฬา ตรวจงบแยกหน่วยงานกับเอกสารสำนักงบประมาณ และให้อาจารย์ตรวจระเบียบวิธี')
for r in table_row(d, '3'):
    set_cell(r.cells[3], 'ผู้เยี่ยมเยือนเพิ่มทุกปีหลังโควิด นราธิวาส 431,003 เป็น 1.24 ล้านคน สูงกว่าปี 2562 (703,810) แต่อาจเทียบกันไม่ได้ตรง ๆ')
for r in table_row(d, '5'):
    set_cell(r.cells[2], 'ตรงบางส่วน'); set_cell(r.cells[3], 'หน่วยงานความมั่นคงได้ราว 53% ของงบสะสม 2551–2568 (แหล่งรอง แยกตามหน่วยงาน ไม่ใช่ภารกิจ)')
for r in table_row(d, '6'):
    set_cell(r.cells[2], 'ตรงบางส่วน'); set_cell(r.cells[3], 'นับดิบ ρ = −0.06; ต่อประชากร ρ = −0.47 (37 อำเภอ)')
# terminology: สมมติฐาน -> ความคาดหวังก่อนดูข้อมูล (run level, then paragraph fallback)
n = 0
def fix(par):
    global n
    for r in par.runs:
        if 'สมมติฐาน' in r.text:
            r.text = r.text.replace('สมมติฐาน', 'ความคาดหวังก่อนดูข้อมูล'); n += 1
for par in d.paragraphs: fix(par)
for t in d.tables:
    for row in t.rows:
        for c in row.cells:
            for par in c.paragraphs: fix(par)
left = [p.text for p in d.paragraphs if 'สมมติฐาน' in p.text]
print('TH replaced', n, 'left', left)
d.save(os.path.join(out, 'SBP_บทความวิจัย_TH_ฉบับร่าง.docx'))

# ---------------------------------------------------------------- English
d = Document(en)
p = find(d, 'Four findings stand out.')
t = p.text.split('Third, visitor numbers')[0]
set_par(p, t + 'Third, visitor numbers rose every year after COVID-19 (Narathiwat: 431,003 to 1.24 million, 2022–2025, above the 703,810 recorded in 2019), which does not support the belief that news of insecurity keeps visitors away, although the 2019 comparison is tentative. ' + FOURTH_EN)
set_par(find(d, 'Of the eight expectations, three were'), 'Of the eight expectations, ' + TALLY_EN)
set_par(find(d, 'Of the eight expectations, the data are'), 'Of the eight expectations, the data are consistent with three, partly consistent with four, and do not support one (updated 20 November 2026 after adding district population, a 2019 tourism baseline and a budget breakdown by agency).')
set_par(find(d, 'Association with business formation.'), " Spearman's ρ between lethal events and new juristic persons across 37 districts is −0.06 (approximate 95% CI −0.37 to 0.27). Excluding the provincial capital districts gives 0.02; using the official count for Narathiwat's 13 districts alone gives −0.26. Raw counts, however, mix in district size: new firms track population closely (ρ = 0.91), while lethal events hardly do (ρ = 0.08).", bold_lead='Association with business formation.')
set_par(find(d, 'Business counts are not adjusted'), " Using registered population for 2023 (2024 for Narathiwat, where Su-ngai Kolok is undercounted in earlier years), ρ between lethal events per 100,000 people and new firms per 10,000 people is −0.47 (95% CI −0.69 to −0.18); excluding capital districts it is −0.40 (−0.65 to −0.07), and the sign is negative within every province. Two rates that share a denominator can be correlated partly by construction, so a partial rank correlation controlling for population was also computed: −0.31 (−0.58 to 0.02), still negative but weaker. Urbanisation, roads and other factors are not controlled for, so this is an association, not an effect.", bold_lead='Population-adjusted.')
set_par(find(d, 'Budget. The budget data collected'), ' Provincial budget data are classified by type of expenditure, not by mission. A secondary compilation of the Budget Acts for 2008–2025 by Rocket Media Lab shows that of 267.7 billion baht allocated to the southern-border plans, ISOC received 30.5%, the Ministry of Defence 16.0% and the Royal Thai Police 6.3%, about 53% in total. This is by receiving agency rather than mission (ISOC also runs development projects), and the draft 2026 integrated-plan budget is about 74% lower than the year before.', bold_lead='Budget.')
set_par(find(d, 'Expectation 3, tourism.'), " The figures do not support the claim that visitors stay away. Narathiwat's 1.24 million visitors in 2025 exceed the 703,810 recorded in 2019, but the comparison is tentative: Pattani and Yala now report 2.3 and 2.9 times their 2019 figures while Songkhla reports slightly fewer, which may reflect a change in estimation method. Foreign visitors are largely Malaysians who know the area; news may weigh more on Thai tourists from other regions, which this dataset cannot separate.", bold_lead='Expectation 3, tourism.')
set_par(find(d, 'Expectation 5, budget.'), " Partly consistent. Security agencies received about 53% of cumulative southern-border plan budgets in 2008–2025 (Rocket Media Lab), in line with Suwanthanin and Sirisunhirun (2026), who report that cumulative spending is concentrated on security. The figure is by agency rather than mission, has not been checked against the original documents, and the latest budgets point downward. This study's own data show that public administration and defence is 17.39% of Pattani's GPP, against 6.09% nationally.", bold_lead='Expectation 5, budget.')
set_par(find(d, 'Expectation 6, incidents and new businesses.'), ' Partly consistent. Raw counts show no association, but per head of population, districts with more lethal events have fewer new businesses (ρ = −0.47). This points the same way as Chongwilaikasaem et al. (2022), who found fewer new businesses here than in comparison provinces over 2004–2017; this study compares districts within the area over 2021–2025. The estimate depends on method (partial ρ = −0.31), and high-violence districts tend to be rural, so the effect of violence cannot yet be separated from rurality.', bold_lead='Expectation 6, incidents and new businesses.')
set_par(find(d, 'No district population adjustment.'), ' Rates use registered population, which can differ from residents; urbanisation and market access are not controlled for; there are only 37 districts.', bold_lead='Population adjusted, other factors not.')
set_par(find(d, 'No pre-COVID baseline.'), " Main tourism data start in 2021. The added 2019 baseline shows the three provinces far above pre-COVID levels while Songkhla is slightly below, which suggests a possible change in estimation method.", bold_lead='Pre-COVID baseline may not be comparable.')
set_par(find(d, 'Next steps are to adjust district data'), "Next steps are to control for urbanisation in the district analysis, check the tourism ministry's estimation method, verify the budget-by-agency figures against Budget Bureau documents, and have the method reviewed by a faculty member.")
for r in table_row(d, '3'):
    set_cell(r.cells[3], 'Visitors rose every year after COVID; Narathiwat 431,003 to 1.24 million, above 703,810 in 2019 (comparison tentative)')
for r in table_row(d, '5'):
    set_cell(r.cells[2], 'Partly consistent'); set_cell(r.cells[3], 'Security agencies received about 53% of cumulative 2008–2025 budgets (secondary source, by agency not mission)')
for r in table_row(d, '6'):
    set_cell(r.cells[2], 'Partly consistent'); set_cell(r.cells[3], 'Raw ρ = −0.06; per capita ρ = −0.47 across 37 districts')
d.save(os.path.join(out, 'SBP_Article_EN_draft.docx'))
print('EN saved')
# (applied afterwards) TH draft English abstract: 'eight hypotheses' -> 'eight prior expectations'
