# สร้างชุดไฟล์สำหรับ Power BI จาก Master Panel + ตารางวิเคราะห์ (17 พ.ย. 2569)
import pandas as pd, json, os, openpyxl
B=os.path.expanduser("~/mnt/SBP_Dashboard"); O=f"{B}/04_output/powerbi"
w=lambda df,n: df.to_csv(f"{O}/{n}",index=False,encoding="utf-8-sig")
p=pd.read_csv(f"{B}/03_analysis/00_master_panel/SBP_master_panel_long.csv")
ATH={'Pattani':'ปัตตานี','Yala':'ยะลา','Narathiwat':'นราธิวาส','Songkhla':'สงขลา (ทั้งจังหวัด)','Songkhla-4districts':'สงขลา 4 อำเภอ',
 '3 provinces (PTN+YLA+NWT)':'รวม 3 จังหวัด','South region':'ภาคใต้','Whole Kingdom':'ทั้งประเทศ'}
ORD=['Pattani','Yala','Narathiwat','Songkhla-4districts','Songkhla','3 provinces (PTN+YLA+NWT)','South region','Whole Kingdom']
DTH={'Agriculture':'เกษตร','Business':'ธุรกิจ','Economy':'เศรษฐกิจ','Fisheries':'ประมง','Investment':'การลงทุน','Labour':'แรงงาน','Poverty':'ความยากจนและรายได้',
 'Public budget':'งบประมาณรัฐ','Region':'ภาพรวมภาค','Security':'ความมั่นคง','Tourism':'ท่องเที่ยว','Trade':'การค้าชายแดน'}
# ---- fact_panel
f=p[['area','year_BE','year_CE','period_type','indicator_code','value','quality_flag']].copy()
fl=f.quality_flag.fillna('').astype(str)
fruit=f.indicator_code.isin(['agri_output_longkong_t','agri_output_durian_t','agri_output_mangosteen_t','agri_output_rambutan_t'])
bad=fl.str.contains('do not use|INCOMPLETE',case=False)|(fruit&(f.year_BE==2565))
f.loc[fruit&(f.year_BE==2565),'quality_flag']='caution RF23 (source reports near-zero fruit output 2565; cause unverified)'
f['usable']=(~bad).astype(int)
f['caution']=(f.quality_flag.fillna('').str.contains('caution|RF|preliminary|≠',case=False)).astype(int)
w(f,'fact_panel.csv')
# ---- dims
di=p.drop_duplicates('indicator_code')[['indicator_code','indicator_th','indicator_en','domain','unit','source_id','used_in']].copy()
di['domain_th']=di.domain.map(DTH)
for k in ['P1','P2','P3','P4']: di[k]=di.used_in.apply(lambda s: int(k in str(s) or str(s)=='All'))
w(di.sort_values(['domain','indicator_code']),'dim_indicator.csv')
da=p.drop_duplicates('area')[['area','area_level','scope']].copy()
da['area_th']=da.area.map(ATH); da['sort']=da.area.map({a:i+1 for i,a in enumerate(ORD)})
da['is_focus_province']=da.area.isin(['Pattani','Yala','Narathiwat']).astype(int)
w(da.sort_values('sort'),'dim_area.csv')
ys=sorted(p.year_BE.unique()); w(pd.DataFrame({'year_BE':ys,'year_CE':[y-543 for y in ys],'year_label':[f"{y} ({y-543})" for y in ys],'in_scope':[int(2564<=y<=2568) for y in ys]}),'dim_year.csv')
# ---- analysis tables from dashboard data
d=json.load(open(f"{B}/04_output/dashboard/sbp_dashboard_data.json",encoding="utf-8"))
lq=pd.DataFrame(d['lq']); lq['lq_class']=lq.lq.apply(lambda x:'เชี่ยวชาญ (LQ>1.25)' if x>1.25 else ('ใกล้เคียงประเทศ' if x>=0.75 else 'ต่ำกว่าประเทศ (LQ<0.75)')); w(lq,'fact_lq.csv')
ds=pd.DataFrame(d['districts']).drop(columns=['key'])
ds['rank_new_firms']=ds.new_firms_2565_2568.rank(ascending=False,method='average'); ds['rank_lethal_events']=ds.ucdp_lethal_events_2564_2568.rank(ascending=False,method='average')
ds['is_muang']=ds.district.astype(str).str.startswith('เมือง').astype(int)
ds['dissolved_per_100_new']=(ds.dissolved_2565_2568/ds.new_firms_2565_2568*100).round(1)
w(ds,'fact_district.csv')
w(pd.DataFrame(d['fish']),'fact_fish.csv'); w(pd.DataFrame(d['sections']).drop(columns=['key']),'fact_sections.csv')
w(pd.DataFrame(d['disbursement']).drop(columns=['key']),'fact_disbursement.csv'); w(pd.DataFrame(d['incident_types']),'fact_incident_types.csv')
wb=openpyxl.load_workbook(f"{B}/02_clean/Agriculture_4prov_2563-2568_clean.xlsx",read_only=True)
r=list(wb['price_monthly_national'].iter_rows(values_only=True)); pm=pd.DataFrame(r[1:],columns=r[0]).drop(columns=['source_file'])
pm['date']=pd.to_datetime(dict(year=pm.year_ad,month=pm.month,day=1)).dt.strftime('%Y-%m-%d'); w(pm.sort_values(['product','date']),'fact_price_monthly.csv')
H=[(1,'P1','ประมงและอาหารแปรรูปคือจุดแข็งของสามจังหวัด','ตรงบางส่วน','ปัตตานีขึ้นท่า 84,694 ตัน (2567) = 7.35% ของประเทศ แต่ LQ อุตสาหกรรมการผลิตต่ำกว่า 1 ทั้งสามจังหวัด'),
(2,'P2','ขาดเงินทุน','สอดคล้อง','ทุนจดทะเบียนธุรกิจใหม่ 217–546 ล้านบาท/จังหวัด/ปี เทียบสงขลา 1,866–2,658 ล้านบาท'),
(3,'P3','ข่าวทำให้คนไม่กล้ามาเที่ยว','ไม่ตรง','ผู้เยี่ยมเยือนเพิ่มทุกปีหลังโควิด: นราธิวาส 431,003 → 1.24 ล้านคน (2565→2568); ยังไม่มีฐานก่อนโควิด'),
(4,'P2','รายได้ต่อหัวเดือนละ 7–8 พันบาท','ตรงบางส่วน','GPP ต่อหัว 2567 ÷ 12: ปัตตานี 7,067; นราธิวาส 5,565; ยะลา 10,120 บาท'),
(5,'P3','งบส่วนใหญ่เป็นงบความมั่นคง','ทดสอบไม่ได้','งบที่มีแยกตามประเภทรายจ่าย ไม่ได้แยกตามภารกิจ'),
(6,'P3','อำเภอเหตุการณ์สูงมีธุรกิจใหม่น้อย','ไม่ตรง','Spearman rho = −0.06 (37 อำเภอ); ยังไม่ปรับตามประชากร'),
(7,'P3','เหตุการณ์จำนวนมากไม่มีผู้เสียชีวิต','สอดคล้อง','นราธิวาส 2568: ทางการ 103 ครั้ง เทียบเหตุที่มีผู้เสียชีวิต (UCDP) 13 ครั้ง'),
(8,'P4','ครัวเรือนพึ่งยางเปราะบางเมื่อราคาตก','สอดคล้อง','ราคา RSS3 เฉลี่ยปี 127.96 (2554) → 41.48 (2561) → 69.87 (2567) บาท/กก.')]
w(pd.DataFrame(H,columns=['no','P','hypothesis','result','evidence']),'fact_hypothesis.csv')
for n in sorted(os.listdir(O)):
    if n.endswith('.csv'): print(n, len(pd.read_csv(f"{O}/{n}")))
print(f.usable.value_counts().to_dict())
