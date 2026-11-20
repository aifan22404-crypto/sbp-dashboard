"""Add the urbanisation (density) robustness check to TH/EN drafts. Usage: python update_articles_robustness.py th.docx en.docx"""
import sys
from docx import Document
REP = {
 'th': [
  ('จึงเป็นความสัมพันธ์ ไม่ใช่ผลกระทบ',
   'เมื่อคุมความหนาแน่นประชากร (ประชากรต่อตารางกิโลเมตร ใช้แทนความเป็นเมือง) ค่ายังอยู่ที่ −0.52 (−0.72 ถึง −0.22) และเมื่อคุมทั้งความหนาแน่นและอำเภอเมืองได้ −0.46 (−0.69 ถึง −0.15) แต่ถ้าเทียบเฉพาะอำเภอภายในจังหวัดเดียวกัน (คุมจังหวัด) ค่าอ่อนลงเหลือ −0.31 (−0.59 ถึง 0.04) ส่วนหนึ่งของความสัมพันธ์จึงมาจากความต่างระหว่างจังหวัด ทั้งหมดนี้ยังเป็นความสัมพันธ์ ไม่ใช่ผลกระทบ'),
  ('และอำเภอที่มีเหตุต่อหัวสูงมักเป็นอำเภอชนบท จึงยังแยกผลของความรุนแรงออกจากความเป็นชนบทไม่ได้',
   'ความสัมพันธ์ยังอยู่เมื่อคุมความหนาแน่นประชากร (−0.52) แต่อ่อนลงเมื่อเทียบเฉพาะภายในจังหวัด (−0.31) จึงยังสรุปเป็นเหตุและผลไม่ได้'),
  ('ยังไม่ได้คุมความเป็นเมืองหรือการเข้าถึงตลาด และมีเพียง 37 อำเภอ',
   'คุมความเป็นเมืองด้วยความหนาแน่นประชากรเท่านั้น ยังไม่ได้คุมการเข้าถึงตลาดหรือถนน และมีเพียง 37 อำเภอ'),
  ('งานขั้นต่อไปคือคุมปัจจัยความเป็นเมืองในการวิเคราะห์รายอำเภอ',
   'งานขั้นต่อไปคือคุมปัจจัยอื่นเพิ่ม เช่น ระยะทางถึงเมืองหลักหรือถนน ในการวิเคราะห์รายอำเภอ'),
 ],
 'en': [
  ('so this is an association, not an effect.',
   'so this is an association, not an effect. Controlling for population density (people per km², a proxy for urbanisation) leaves ρ at −0.52 (−0.72 to −0.22); controlling for density and capital districts gives −0.46 (−0.69 to −0.15). Comparing districts only within the same province weakens it to −0.31 (−0.59 to 0.04), so part of the association reflects differences between provinces.'),
  ('and high-violence districts tend to be rural, so the effect of violence cannot yet be separated from rurality.',
   'and it survives a control for population density (−0.52) but weakens within provinces (−0.31), so it cannot be read as causal.'),
  ('urbanisation and market access are not controlled for; there are only 37 districts.',
   'urbanisation is controlled only through population density, not market access or roads; there are only 37 districts.'),
  ('Next steps are to control for urbanisation in the district analysis,',
   'Next steps are to add further district controls such as distance to the main towns or road access,'),
 ]}
for path, key in ((sys.argv[1], 'th'), (sys.argv[2], 'en')):
    d = Document(path); done = 0
    for old, new in REP[key]:
        hit = False
        for p in d.paragraphs:
            for r in p.runs:
                if old in r.text:
                    r.text = r.text.replace(old, new); hit = True
        assert hit, (key, old[:40]); done += 1
    d.save(path); print(key, done)
# (applied afterwards) Discussion of expectation 6 reworded so the two "−0.31" values (partial on population vs within-province) are not confused.
