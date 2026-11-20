"""Article revision v2 (20 Nov 2026): Irfan's own passages, international literature, P3 method + robustness table,
two figures, consistency fixes after the data update, data availability, references.
Usage: python revise_articles_v2.py th.docx en.docx figures_dir"""
import sys, copy, os
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

TH, EN, FIG = sys.argv[1], sys.argv[2], sys.argv[3]

def find(doc, start):
    hits = [p for p in doc.paragraphs if p.text.startswith(start)]
    assert len(hits) == 1, (start, len(hits))
    return hits[0]

def set_text(p, text, lead=None):
    runs = p.runs
    if lead is None:
        runs[0].text = text
        for r in runs[1:]: r.text = ''
        return p
    runs[0].text = lead; runs[0].bold = True
    if len(runs) < 2:
        r = p.add_run(); r.bold = False
    p.runs[1].text = text; p.runs[1].bold = False
    for r in p.runs[2:]: r.text = ''
    return p

def replace_in(p, old, new):
    for r in p.runs:
        if old in r.text:
            r.text = r.text.replace(old, new); return True
    raise AssertionError(('not found', old[:50]))

def insert_after(anchor, text='', like=None, lead=None):
    """Insert a paragraph after `anchor` (paragraph or table) copying paragraph properties of `like`."""
    el = anchor._p if hasattr(anchor, '_p') else anchor._tbl
    new_p = OxmlElement('w:p'); el.addnext(new_p)
    par = Paragraph(new_p, anchor._parent)
    src = like if like is not None else (anchor if hasattr(anchor, '_p') else None)
    if src is not None and src._p.pPr is not None:
        new_p.insert(0, copy.deepcopy(src._p.pPr))
    if lead:
        r = par.add_run(lead); r.bold = True
        par.add_run(' ' + text.lstrip() if text else '')
    elif text:
        par.add_run(text)
    return par

def table_after(doc, anchor, rows, like_table):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = like_table.style
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = t.cell(i, j); c.text = str(val)
            if i == 0:
                for r in c.paragraphs[0].runs: r.bold = True
    t.autofit = False
    for j, w in enumerate([2.4, 0.9, 1.5, 0.5, 1.0][:len(rows[0])]):
        for c in t.columns[j].cells: c.width = Inches(w)
        t._tbl.tblGrid.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}gridCol')[j].set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', str(int(w*1440)))
    anchor._p.addnext(t._tbl)
    return t

def picture_after(anchor, path, caption, like_caption):
    p = insert_after(anchor, '', like=like_caption)
    p.add_run().add_picture(path, width=Inches(6.2))
    cap = insert_after(p, caption, like=like_caption)
    for r in cap.runs: r.italic = True; r.font.size = Pt(9)
    return cap


W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
def relink(p, text, url):
    h = p._p.find(W_NS + 'hyperlink')
    ts = h.findall('.//' + W_NS + 't'); ts[0].text = text
    for t in ts[1:]: t.text = ''
    rid = h.get(R_NS + 'id')
    p.part.rels[rid]._target = url

def add_lit_row(doc, header_start, cells):
    for t in doc.tables:
        if t.rows[0].cells[0].text.strip() == header_start:
            row = t.add_row()
            for c, v in zip(row.cells, cells): c.text = v
            return
    raise AssertionError('lit table')

def ref_after(doc, start, text):
    p = find(doc, start)
    return insert_after(p, text)

ROB_EN = [['Specification (per-capita rates)', 'Spearman ρ', '95% CI', 'n', 'Permutation p'],
          ['No controls', '−0.47', '−0.69 to −0.18', '37', '0.003'],
          ['Control: population density', '−0.52', '−0.72 to −0.22', '37', '0.001'],
          ['Control: density + capital district', '−0.46', '−0.69 to −0.15', '37', '0.004'],
          ['Excluding capital districts, control: density', '−0.45', '−0.69 to −0.13', '34', '0.008'],
          ['Control: density + province', '−0.31', '−0.59 to 0.04', '37', '0.065'],
          ['Raw counts (no rates)', '−0.06', '−0.37 to 0.27', '37', '0.75']]
ROB_TH = [['แบบที่ใช้ (อัตราต่อประชากร)', 'Spearman ρ', 'ช่วงความเชื่อมั่น 95%', 'n', 'p (permutation)'],
          ['ไม่คุมตัวแปรอื่น', '−0.47', '−0.69 ถึง −0.18', '37', '0.003'],
          ['คุมความหนาแน่นประชากร', '−0.52', '−0.72 ถึง −0.22', '37', '0.001'],
          ['คุมความหนาแน่น + อำเภอเมือง', '−0.46', '−0.69 ถึง −0.15', '37', '0.004'],
          ['ไม่นับอำเภอเมือง คุมความหนาแน่น', '−0.45', '−0.69 ถึง −0.13', '34', '0.008'],
          ['คุมความหนาแน่น + จังหวัด', '−0.31', '−0.59 ถึง 0.04', '37', '0.065'],
          ['นับดิบ (ไม่คิดต่อประชากร)', '−0.06', '−0.37 ถึง 0.27', '37', '0.75']]

REFS_EN = [  # (insert after paragraph starting with, text)
 (None, 'Abadie, A., & Gardeazabal, J. (2003). The economic costs of conflict: A case study of the Basque Country. American Economic Review, 93(1), 113–132. https://doi.org/10.1257/000282803321455188'),
 ('Benbourenane', 'Blattman, C., & Miguel, E. (2010). Civil war. Journal of Economic Literature, 48(1), 3–57. https://doi.org/10.1257/jel.48.1.3'),
 ('Blattman', 'Camacho, A., & Rodriguez, C. (2013). Firm exit and armed conflict in Colombia. Journal of Conflict Resolution, 57(1), 89–116.'),
 ('Chongwilaikasaem', 'Enders, W., Sandler, T., & Parise, G. F. (1992). An econometric analysis of the impact of terrorism on tourism. Kyklos, 45(4), 531–554.'),
 ('Hayiyakoh', 'Kronmal, R. A. (1993). Spurious correlation and the fallacy of the ratio standard revisited. Journal of the Royal Statistical Society: Series A, 156(3), 379–392.'),
 ('Kronmal', 'Neumayer, E. (2004). The impact of political violence on tourism: Dynamic cross-national estimation. Journal of Conflict Resolution, 48(2), 259–281. https://doi.org/10.1177/0022002703262358'),
 ('Phongsuwansiri', 'Rocket Media Lab. (2025). [Seventeen years of the southern-border budget: Through which agencies, and how much]. https://rocketmedialab.co/deep-south-budget/'),
]

def common(doc, lang):
    # date line
    p = doc.paragraphs[1]
    replace_in(p, 'Nov 17, 2026', 'Revised Nov 20, 2026')
    # references
    for after, text in REFS_EN:
        if after is None:
            first = find(doc, 'Andriesse')
            new_p = OxmlElement('w:p'); first._p.addprevious(new_p)
            par = Paragraph(new_p, first._parent); new_p.insert(0, copy.deepcopy(first._p.pPr)); par.add_run(text)
        else:
            ref_after(doc, after, text)

# ===================================================================== EN
d = Document(EN)
common(d, 'en')
replace_in(find(d, 'Of the eight expectations, three were'), 'a district analysis of only 37 districts without controls for urbanisation', 'a district analysis of only 37 districts with few controls')
set_text(find(d, '[Irfan] Add two or three sentences'),
 "I grew up in Tak Bai, Narathiwat, a quiet town with both the sea and a river. Outsiders mostly know it through violence, above all the Tak Bai incident of 2004, which people have not forgotten twenty years on. "
 "When I won a scholarship to study economics in the United Arab Emirates, I saw how Dubai keeps building and draws tourists and investors as its main sources of income, and every trip home made the contrast sharper: my hometown is developing, but very slowly. "
 "I built this project to put the area's economic structure in one place, so that investors and entrepreneurs can see which sectors are already crowded and which could grow from existing strengths, and so that young people in the area can read their own economy before they invest or start a business. Many people outside imagine that bombs go off every day; they do not, but the fear alone is enough to keep visitors and money away. I also plan to invest and build a business here myself, and I want to use this evidence to do it.")
find(d, 'Fifteen studies were reviewed').runs[0].text = 'Fifteen studies on the area were reviewed and fall into four groups; five international studies are added as a fifth group so that the results can be compared with evidence from other conflict-affected places. The full list of area studies with links is in the companion document '
add_lit_row(d, 'Group', ['E International evidence', 'Political violence reduces tourist arrivals and diverts them to neighbouring places; armed attacks raise the exit of small and young firms; terrorism lowered per-capita GDP in the Basque Country by about 10% against a synthetic comparison', 'Enders et al. (1992); Neumayer (2004); Camacho & Rodriguez (2013); Abadie & Gardeazabal (2003); Blattman & Miguel (2010)'])
set_text(find(d, 'The summaries above are drawn'), 'The summaries above are drawn from abstracts, summary pages or full texts; not every study has yet been read in full by the author.')
p = find(d, 'A total of 333 raw files')
insert_after(p, 'Three supplementary datasets were added on 20 November 2026: registered district population (Department of Provincial Administration, published through provincial GD Catalogs), a 2019 tourism baseline (Ministry of Tourism and Sports) and a breakdown of the southern-border budget by receiving agency (Rocket Media Lab, a secondary compilation of the Budget Acts). District land area comes from OpenGISData-Thailand.')
set_text(find(d, "P3 uses Spearman's rank correlation"),
 "P3 uses Spearman's rank correlation between lethal violent events (UCDP, 2021–2025) and newly registered juristic persons (DBD, 2022–2025) across 37 districts; tied ranks are averaged. Because raw counts mix in district size, the main specification uses rates: lethal events per 100,000 registered residents and new firms per 10,000. Confidence intervals use the Fisher z transformation with standard error 1/√(n − 3 − k), where k is the number of controls. Robustness checks are partial rank correlations controlling for population, population density (a proxy for urbanisation), capital districts and province, with permutation p-values from 10,000 shuffles.")
replace_in(find(d, 'Raw files are never edited.'), 'a red-flag register of 24 entries', 'a red-flag register of 26 entries')
set_text(find(d, 'Unrest incidents rose in 2024–2025'), 'Unrest incidents rose in 2024–2025. Counted raw, districts with more lethal violence did not register fewer new businesses; per head of population, they did.')
p = find(d, 'Population-adjusted.')
set_text(p, " Using registered population for 2023 (2024 for Narathiwat, where Su-ngai Kolok is undercounted in earlier years), ρ between lethal events per 100,000 people and new firms per 10,000 people is −0.47 (95% CI −0.69 to −0.18; permutation p = 0.003), and the sign is negative within every province. The table below shows how the estimate changes with controls. It holds after controlling for population density (−0.52) and capital districts (−0.46), but weakens to −0.31, with an interval that includes zero, when districts are compared only within the same province, so part of the association reflects differences between provinces. Two rates that share a denominator can be correlated partly by construction (Kronmal, 1993); the partial correlation on population, −0.31 (−0.58 to 0.02), is reported for that reason.", lead='Population-adjusted.')
cap_like = find(d, 'The summaries above are drawn')
lit_table = [t for t in d.tables if t.rows[0].cells[0].text.strip() == 'District'][0]
t = table_after(d, p, ROB_EN, lit_table)
cap = picture_after(t, os.path.join(FIG, 'fig2_district_percapita_en.png'), 'Figure 2. Lethal events and new firms per head of population, 37 districts. Population: civil registration 2023 (Narathiwat 2024).', cap_like)
pov_table = [t for t in d.tables if t.rows[0].cells[0].text.strip() == 'Poverty rate (%)'][0]
picture_after(pov_table, os.path.join(FIG, 'fig1_poverty_en.png'), 'Figure 1. Poverty rate, 2021–2025. Source: NESDC.', cap_like)
replace_in(find(d, 'Expectation 3, tourism.'), 'which this dataset cannot separate.', 'which this dataset cannot separate. Cross-country studies find that political violence reduces tourist arrivals (Enders et al., 1992; Neumayer, 2004); the provincial totals here do not show this, which suggests that any effect falls on particular places or groups of visitors that provincial data cannot see.')
replace_in(find(d, 'Expectation 6, incidents and new businesses.'), 'so it cannot be read as causal.', 'so it cannot be read as causal. The direction is consistent with Camacho and Rodriguez (2013), who found in Colombia that armed attacks raise the exit of small and young plants.')
set_text(find(d, '[Irfan] Write one paragraph'),
 "The result that surprised me most was expectation 3. I had always believed that news of violence keeps people from visiting and from investing, because I saw it myself: a floating market near the Malaysian border that used to be busy went quiet after a bombing, and every report of a shooting or a bomb made people more afraid. Yet the provincial data show visitors rising every year after COVID. "
 "I think both can be true. The provincial totals are driven mostly by Malaysian visitors who know the area, while the damage from news may fall on small places, such as a market or a row of shops where an attack has just happened, which provincial figures cannot see. "
 "Expectation 6 taught me the same lesson from the other side: counted raw, violence and new businesses looked unrelated, but per head of population the districts with more violence had fewer new firms. How something is measured matters as much as the data itself. Local efforts such as halal food fairs help, but I believe the area needs structural change more than one-off events, so the next question I want to answer is how other places that lived through long conflicts rebuilt their local economies.")
set_text(find(d, 'Not causal.'), ' All results are descriptive associations. Controls are limited to population, population density, capital districts and province.', lead='Not causal.')
set_text(find(d, 'Low statistical power.'), ' With 37 districts, confidence intervals are about ±0.25 wide, and the within-province estimate cannot be distinguished from zero.', lead='Low statistical power.')
p = find(d, 'All results can be explored interactively'); p.runs[0].text = 'All results can be explored interactively in '; relink(p, 'The Southern Frontier Review', 'https://aifan22404-crypto.github.io/sbp-dashboard/'); p.runs[-1].text = ' dashboard (in Thai): https://aifan22404-crypto.github.io/sbp-dashboard/'
p = find(d, 'All results can be explored interactively')
insert_after(p, ' Cleaned datasets, the master panel, the scripts used for every figure and the red-flag register are available at https://github.com/aifan22404-crypto/sbp-dashboard.', lead='Data and code availability.')
replace_in(find(d, 'Data sources.'), 'Uppsala Conflict Data Program (UCDP GED).', 'Uppsala Conflict Data Program (UCDP GED); Department of Provincial Administration (district population, via provincial GD Catalogs); OpenGISData-Thailand (district land area, CC BY 4.0); Rocket Media Lab (budget by agency, secondary).')
d.save(EN); print('EN revised')

# ===================================================================== TH
d = Document(TH)
common(d, 'th')
replace_in(find(d, 'Of the eight prior expectations,'), 'a district analysis of only 37 districts without controls for urbanisation', 'a district analysis of only 37 districts with few controls')
set_text(find(d, '[Irfan] เพิ่ม 2–3 ประโยค'),
 'ผมเกิดและโตที่อำเภอตากใบ จังหวัดนราธิวาส เมืองเล็กที่สงบ มีทั้งทะเลและแม่น้ำ แต่คนภายนอกมักรู้จักที่นี่ผ่านความรุนแรง โดยเฉพาะเหตุการณ์ตากใบเมื่อปี 2547 ซึ่งผ่านมากว่า 20 ปีแล้วคนก็ยังไม่ลืม '
 'เมื่อได้ทุนไปเรียนเศรษฐศาสตร์ที่สหรัฐอาหรับเอมิเรตส์ ผมได้เห็นดูไบที่พัฒนาเมืองไม่หยุดและดึงนักท่องเที่ยวกับนักลงทุนมาเป็นรายได้หลัก ทุกครั้งที่กลับบ้าน ผมรู้สึกถึงความต่างชัดขึ้นเรื่อย ๆ บ้านเกิดของผมพัฒนาก็จริง แต่ช้ามาก '
 'ผมจึงทำงานนี้เพื่อรวมโครงสร้างเศรษฐกิจของพื้นที่ไว้ในที่เดียว ให้นักลงทุนและผู้ประกอบการเห็นว่าสาขาไหนแน่นแล้ว สาขาไหนยังโตได้จากจุดแข็งที่มีอยู่ และให้คนรุ่นใหม่ในพื้นที่ได้อ่านเศรษฐกิจบ้านตัวเองก่อนตัดสินใจลงทุนหรือเริ่มธุรกิจ คนภายนอกจำนวนมากคิดว่าที่นี่มีระเบิดทุกวัน ความจริงไม่ใช่อย่างนั้น แต่ความกลัวอย่างเดียวก็พอจะทำให้คนไม่มาและเงินไม่หมุน ผมเองก็ตั้งใจจะลงทุนและทำธุรกิจที่บ้าน และอยากใช้หลักฐานชุดนี้ในการตัดสินใจด้วย')
find(d, 'งานวิจัยที่ทบทวน 15 ชิ้น').runs[0].text = 'งานวิจัยเกี่ยวกับพื้นที่ที่ทบทวน 15 ชิ้นจัดได้เป็น 4 กลุ่ม และเพิ่มงานต่างประเทศ 5 ชิ้นเป็นกลุ่มที่ 5 เพื่อเทียบผลกับพื้นที่อื่นที่เคยเผชิญความขัดแย้ง รายละเอียดและรายการอ้างอิงฉบับเต็มของงานในพื้นที่อยู่ในเอกสาร '
add_lit_row(d, 'กลุ่ม', ['E หลักฐานจากต่างประเทศ', 'ความรุนแรงทางการเมืองลดจำนวนนักท่องเที่ยวและทำให้นักท่องเที่ยวไปที่อื่นแทน การโจมตีด้วยอาวุธทำให้โรงงานขนาดเล็กและเพิ่งตั้งปิดกิจการมากขึ้น และการก่อการร้ายทำให้ GDP ต่อหัวของแคว้นบาสก์ต่ำกว่ากลุ่มเทียบราว 10%', 'Enders et al. (1992); Neumayer (2004); Camacho & Rodriguez (2013); Abadie & Gardeazabal (2003); Blattman & Miguel (2010)'])
set_text(find(d, 'ข้อสรุปในตารางมาจากบทคัดย่อ'), 'ข้อสรุปในตารางมาจากบทคัดย่อ หน้าสรุป หรือฉบับเต็ม ผู้เขียนยังไม่ได้อ่านฉบับเต็มครบทุกชิ้น')
p = find(d, 'ข้อมูลดิบ 333 ไฟล์')
insert_after(p, 'เมื่อ 20 พ.ย. 2569 เพิ่มข้อมูลเสริม 3 ชุด ได้แก่ ประชากรรายอำเภอจากทะเบียนราษฎร (กรมการปกครอง เผยแพร่ผ่านระบบบัญชีข้อมูลจังหวัด) ข้อมูลท่องเที่ยวปี 2562 (กระทรวงการท่องเที่ยวและกีฬา) และงบแก้ปัญหาชายแดนใต้แยกตามหน่วยงานที่รับงบ (Rocket Media Lab ซึ่งรวบรวมจาก พ.ร.บ.งบประมาณ) พื้นที่อำเภอใช้ข้อมูลจาก OpenGISData-Thailand')
set_text(find(d, 'P3 ใช้สหสัมพันธ์อันดับของ Spearman'),
 'P3 ใช้สหสัมพันธ์อันดับของ Spearman ระหว่างเหตุรุนแรงที่มีผู้เสียชีวิต (UCDP, 2564–2568) กับนิติบุคคลตั้งใหม่ (DBD, 2565–2568) ใน 37 อำเภอ อันดับที่เท่ากันใช้ค่าเฉลี่ย เนื่องจากจำนวนนับดิบปนผลของขนาดอำเภอ แบบหลักจึงใช้อัตรา คือเหตุต่อประชากรแสนคนและธุรกิจใหม่ต่อประชากรหมื่นคน ช่วงความเชื่อมั่นคำนวณด้วย Fisher z โดยใช้ความคลาดเคลื่อนมาตรฐาน 1/√(n − 3 − k) เมื่อ k คือจำนวนตัวแปรควบคุม การตรวจความทนทานใช้ partial rank correlation ที่คุมประชากร ความหนาแน่นประชากร (แทนความเป็นเมือง) อำเภอเมือง และจังหวัด พร้อมค่า p แบบ permutation จากการสุ่ม 10,000 ครั้ง')
replace_in(find(d, 'ไฟล์ดิบไม่ถูกแก้ไข'), 'red flags 24 รายการ', 'red flags 26 รายการ')
p = find(d, 'ปรับตามประชากร เมื่อใช้ประชากร')
set_text(p, ' เมื่อใช้ประชากรจากทะเบียนราษฎรปี 2566 (นราธิวาสใช้ปี 2567 เพราะตัวเลขสุไหงโก-ลกปีก่อนหน้านับไม่ครบ) ค่าสหสัมพันธ์ระหว่างเหตุรุนแรงถึงชีวิตต่อแสนคนกับธุรกิจตั้งใหม่ต่อหมื่นคนเท่ากับ −0.47 (ช่วงความเชื่อมั่น 95% −0.69 ถึง −0.18; p แบบ permutation = 0.003) และเป็นลบในทุกจังหวัด ตารางด้านล่างแสดงว่าค่านี้เปลี่ยนอย่างไรเมื่อคุมตัวแปรอื่น ค่ายังอยู่เมื่อคุมความหนาแน่นประชากร (−0.52) และอำเภอเมือง (−0.46) แต่อ่อนลงเหลือ −0.31 และช่วงความเชื่อมั่นคร่อมศูนย์เมื่อเทียบเฉพาะอำเภอภายในจังหวัดเดียวกัน ส่วนหนึ่งของความสัมพันธ์จึงมาจากความต่างระหว่างจังหวัด อัตราสองตัวที่ใช้ตัวหารเดียวกันอาจสัมพันธ์กันเองได้บางส่วน (Kronmal, 1993) จึงรายงาน partial correlation ที่คุมประชากรไว้ด้วย ได้ −0.31 (−0.58 ถึง 0.02)', lead='ปรับตามประชากร')
cap_like = find(d, 'ข้อสรุปในตารางมาจากบทคัดย่อ')
dist_table = [t for t in d.tables if t.rows[0].cells[0].text.strip() == 'อำเภอ'][0]
t = table_after(d, p, ROB_TH, dist_table)
picture_after(t, os.path.join(FIG, 'fig2_district_percapita_th.png'), 'ภาพที่ 2 เหตุที่มีผู้เสียชีวิตกับธุรกิจตั้งใหม่ต่อประชากร 37 อำเภอ ประชากรจากทะเบียนราษฎรปี 2566 (นราธิวาส 2567)', cap_like)
pov_table = [t for t in d.tables if t.rows[0].cells[0].text.strip() == 'อัตราความยากจน (%)'][0]
picture_after(pov_table, os.path.join(FIG, 'fig1_poverty_th.png'), 'ภาพที่ 1 สัดส่วนคนจน พ.ศ. 2564–2568 ที่มา: สศช.', cap_like)
replace_in(find(d, 'ข้อ 3 ท่องเที่ยว'), 'ซึ่งข้อมูลชุดนี้แยกไม่ได้', 'ซึ่งข้อมูลชุดนี้แยกไม่ได้ งานระหว่างประเทศพบว่าความรุนแรงทางการเมืองลดจำนวนนักท่องเที่ยว (Enders et al., 1992; Neumayer, 2004) แต่ตัวเลขรวมระดับจังหวัดในงานนี้ไม่เห็นผลนั้น ผลกระทบจึงอาจไปตกที่บางจุดหรือนักท่องเที่ยวบางกลุ่มที่ข้อมูลระดับจังหวัดมองไม่เห็น')
replace_in(find(d, 'ข้อ 6 เหตุการณ์กับธุรกิจใหม่'), 'จึงยังสรุปเป็นเหตุและผลไม่ได้', 'จึงยังสรุปเป็นเหตุและผลไม่ได้ ทิศทางนี้สอดคล้องกับ Camacho และ Rodriguez (2013) ที่พบในโคลอมเบียว่าการโจมตีด้วยอาวุธทำให้โรงงานขนาดเล็กและเพิ่งตั้งปิดกิจการมากขึ้น')
set_text(find(d, '[Irfan] เขียน 1 ย่อหน้า'),
 'ผลที่ทำให้ผมแปลกใจที่สุดคือข้อ 3 ผมเชื่อมาตลอดว่าข่าวความรุนแรงทำให้คนไม่กล้ามาเที่ยวและไม่กล้าลงทุน เพราะเห็นกับตาว่าตลาดน้ำแห่งหนึ่งใกล้ชายแดนมาเลเซียที่เคยคึกคักเงียบลงหลังเกิดเหตุระเบิด และทุกครั้งที่มีข่าวยิงหรือระเบิด คนในพื้นที่ก็ยิ่งกลัว แต่ข้อมูลระดับจังหวัดกลับแสดงว่าผู้เยี่ยมเยือนเพิ่มขึ้นทุกปีหลังโควิด '
 'ผมคิดว่าทั้งสองอย่างเป็นจริงได้พร้อมกัน ตัวเลขรวมของจังหวัดส่วนใหญ่มาจากนักท่องเที่ยวมาเลเซียที่คุ้นกับพื้นที่ ขณะที่ความเสียหายจากข่าวอาจไปตกที่จุดเล็ก ๆ อย่างตลาดหรือร้านค้าที่เพิ่งเกิดเหตุ ซึ่งตัวเลขระดับจังหวัดมองไม่เห็น '
 'ผลข้อ 6 สอนผมเรื่องเดียวกันจากอีกด้าน เมื่อนับดิบ ความรุนแรงกับธุรกิจใหม่ดูไม่เกี่ยวกัน แต่เมื่อคิดต่อประชากร อำเภอที่มีความรุนแรงสูงกลับมีธุรกิจใหม่น้อยกว่า วิธีวัดจึงสำคัญพอ ๆ กับตัวข้อมูล กิจกรรมในพื้นที่อย่างงานอาหารฮาลาลก็ช่วยได้ แต่ผมเชื่อว่าพื้นที่นี้ต้องการการเปลี่ยนแปลงเชิงโครงสร้างมากกว่ากิจกรรมครั้งคราว คำถามต่อไปที่ผมอยากตอบจึงเป็นว่า พื้นที่อื่นที่ผ่านความขัดแย้งยาวนานฟื้นเศรษฐกิจท้องถิ่นด้วยวิธีใด')
set_text(find(d, 'ไม่ใช่ความเป็นเหตุเป็นผล'), ' ผลทั้งหมดเป็นการพรรณนาและความสัมพันธ์ ตัวแปรควบคุมมีเพียงประชากร ความหนาแน่นประชากร อำเภอเมือง และจังหวัด', lead='ไม่ใช่ความเป็นเหตุเป็นผล')
p = find(d, 'ปรับตามประชากรแล้ว แต่ยังไม่ได้คุมปัจจัยอื่น')
insert_after(p, ' ด้วยจำนวน 37 อำเภอ ช่วงความเชื่อมั่นกว้างราว ±0.25 และค่าที่เทียบเฉพาะภายในจังหวัดยังแยกไม่ออกจากศูนย์', lead='อำนาจทางสถิติต่ำ')
p = find(d, 'ผลทั้งหมดเปิดดูแบบโต้ตอบได้ที่'); p.runs[0].text = 'ผลทั้งหมดเปิดดูแบบโต้ตอบได้ที่แดชบอร์ด '; relink(p, 'The Southern Frontier Review', 'https://aifan22404-crypto.github.io/sbp-dashboard/'); p.add_run(' (https://aifan22404-crypto.github.io/sbp-dashboard/)')
p = find(d, 'ผลทั้งหมดเปิดดูแบบโต้ตอบได้ที่')
insert_after(p, ' ชุดข้อมูลที่ทำความสะอาดแล้ว Master Panel สคริปต์ที่ใช้สร้างทุกตัวเลข และทะเบียน red flags อยู่ที่ https://github.com/aifan22404-crypto/sbp-dashboard', lead='การเข้าถึงข้อมูลและโค้ด')
replace_in(find(d, 'รายการครบ 15 ชิ้นพร้อมลิงก์'), 'รายการครบ 15 ชิ้น', 'รายการงานในพื้นที่ครบ 15 ชิ้น')
replace_in(find(d, 'แหล่งข้อมูล '), 'Uppsala Conflict Data Program (UCDP GED)', 'สำนักงบประมาณ; Uppsala Conflict Data Program (UCDP GED); กรมการปกครอง (ประชากรรายอำเภอ ผ่านระบบบัญชีข้อมูลจังหวัด); OpenGISData-Thailand (พื้นที่อำเภอ, CC BY 4.0); Rocket Media Lab (งบแยกหน่วยงาน, แหล่งรอง)')
d.save(TH); print('TH revised')
