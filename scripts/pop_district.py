import csv, math, json
import pandas as pd
from scipy import stats
rows=[]
# Narathiwat (gdcatalog narathiwat 12_01), totals male+female
for l in open('data/reference/narathiwat_population_by_district_2563-2568.txt',encoding='utf-8'):
    y,d,v=l.strip().split('|'); d=d.replace('อำเภอ','')
    rows.append(('Narathiwat',d,int(y),int(v),'narathiwat.gdcatalog.go.th/dataset/12_01 (ที่ทำการปกครองจังหวัดนราธิวาส)'))
# Pattani 2564-2566: district (non-municipal) + municipalities mapped to their district
ptn_d={2564:dict(เมืองปัตตานี=91401,โคกโพธิ์=62353,หนองจิก=59619,ปะนาเระ=36852,มายอ=56488,ทุ่งยางแดง=25297,สายบุรี=57008,ไม้แก่น=13014,ยะหริ่ง=70825,ยะรัง=90617,กะพ้อ=19099,แม่ลาน=17604),
       2565:dict(เมืองปัตตานี=91409,โคกโพธิ์=62612,หนองจิก=60114,ปะนาเระ=37024,มายอ=56967,ทุ่งยางแดง=25661,สายบุรี=57309,ไม้แก่น=13088,ยะหริ่ง=71564,ยะรัง=91378,กะพ้อ=19234,แม่ลาน=17829),
       2566:dict(เมืองปัตตานี=91440,โคกโพธิ์=62941,หนองจิก=60769,ปะนาเระ=37246,มายอ=57647,ทุ่งยางแดง=25974,สายบุรี=57829,ไม้แก่น=13195,ยะหริ่ง=72460,ยะรัง=92340,กะพ้อ=19399,แม่ลาน=17940)}
# municipality -> district
mun_map={'ยะรัง':'ยะรัง','ยะหริ่ง':'ยะหริ่ง','บางปู':'ยะหริ่ง','ตันหยง':'ยะหริ่ง','มายอ':'มายอ','ปะนาเระ':'ปะนาเระ','หนองจิก':'หนองจิก','บ่อทอง':'หนองจิก','นาประดู่':'โคกโพธิ์','โคกโพธิ์':'โคกโพธิ์','ตะลุบัน':'สายบุรี','ปัตตานี':'เมืองปัตตานี'}
ptn_m={2564:dict(ยะรัง=5089,ยะหริ่ง=6260,บางปู=10442,ตันหยง=2820,มายอ=4907,ปะนาเระ=10681,หนองจิก=6245,บ่อทอง=16241,นาประดู่=3102,โคกโพธิ์=3359,ตะลุบัน=14570,ปัตตานี=45688),
       2565:dict(ยะรัง=5128,ยะหริ่ง=6278,บางปู=10494,ตันหยง=2866,มายอ=4929,ปะนาเระ=10772,หนองจิก=6247,บ่อทอง=16216,นาประดู่=3095,โคกโพธิ์=3337,ตะลุบัน=14551,ปัตตานี=44853),
       2566:dict(ยะรัง=5158,ยะหริ่ง=6246,บางปู=10581,ตันหยง=2898,มายอ=4974,ปะนาเระ=10881,หนองจิก=6256,บ่อทอง=15987,นาประดู่=3077,โคกโพธิ์=3383,ตะลุบัน=14541,ปัตตานี=43915)}
for y in ptn_d:
    tot=dict(ptn_d[y])
    for m,v in ptn_m[y].items(): tot[mun_map[m]]+=v
    if y==2566: assert sum(tot.values())==737077, sum(tot.values())
    for d,v in tot.items(): rows.append(('Pattani',d,y,v,'pattani.gdcatalog.go.th/dataset/population (อำเภอ + เทศบาลรวมเข้าอำเภอที่ตั้ง)'))
yala={2566:[174457,61753,64666,26104,65558,98868,26051,30881],2567:[175569,61642,65156,26215,66000,99368,26190,31218],2568:[176794,61701,65539,26286,66482,99727,26362,31581]}
yd=['เมืองยะลา','เบตง','บันนังสตา','ธารโต','ยะหา','รามัน','กาบัง','กรงปินัง']
for y,vs in yala.items():
    for d,v in zip(yd,vs): rows.append(('Yala',d,y,v,'yala.gdcatalog.go.th/dataset/population_01'))
skm={2562:{'จะนะ':(53407,54283),'เทพา':(40019,38541),'นาทวี':(35119,34310),'สะบ้าย้อย':(39968,39201)},
 2563:{'จะนะ':(53586,54414),'เทพา':(40096,38497),'นาทวี':(35143,34366),'สะบ้าย้อย':(40386,39434)},
 2564:{'จะนะ':(53671,54574),'เทพา':(40161,38621),'นาทวี':(35129,34417),'สะบ้าย้อย':(40620,39745)},
 2565:{'จะนะ':(53757,54753),'เทพา':(40196,38670),'นาทวี':(34979,34371),'สะบ้าย้อย':(40822,40020)},
 2566:{'จะนะ':(54058,54840),'เทพา':(40149,38773),'นาทวี':(34827,34269),'สะบ้าย้อย':(41048,40287)},
 2567:{'จะนะ':(54167,54876),'เทพา':(40179,38901),'นาทวี':(34730,34158),'สะบ้าย้อย':(41108,40528)},
 2568:{'จะนะ':(54235,55017),'เทพา':(40221,38906),'นาทวี':(34417,34068),'สะบ้าย้อย':(41199,40716)}}
for y,dd in skm.items():
    for d,(m,f) in dd.items(): rows.append(('Songkhla',d,y,m+f,'songkhla.gdcatalog.go.th/dataset/dataset_03_46'))
pop=pd.DataFrame(rows,columns=['province','district','year_BE','population','source'])
pop['district']=pop['district'].replace({'สุไหงโกลก':'สุไหงโก-ลก'})
pop.to_csv('/tmp/district_population_long.csv',index=False,encoding='utf-8-sig')
# denominator: 2566; Narathiwat 2567 (Su-ngai Kolok undercounted 2563-2566, RF25)
den=pd.concat([pop[(pop.province!='Narathiwat')&(pop.year_BE==2566)], pop[(pop.province=='Narathiwat')&(pop.year_BE==2567)]])
den=den[['province','district','year_BE','population']].rename(columns={'year_BE':'pop_year'})
fd=pd.read_csv('/home/claude/sbp-dashboard/powerbi/csv/fact_district.csv',encoding='utf-8-sig')
m=fd.merge(den,on=['province','district'],how='left')
print('unmatched:',m[m.population.isna()][['province','district']].values.tolist())
m['lethal_per_100k']=m.ucdp_lethal_events_2564_2568/m.population*1e5
m['new_firms_per_10k']=m.new_firms_2565_2568/m.population*1e4
def rho_ci(x,y):
    r,p=stats.spearmanr(x,y); n=len(x); z=math.atanh(r); se=1/math.sqrt(n-3)
    return round(r,3),round(p,3),round(math.tanh(z-1.96*se),2),round(math.tanh(z+1.96*se),2),n
print('raw     ',rho_ci(m.ucdp_lethal_events_2564_2568,m.new_firms_2565_2568))
print('percap  ',rho_ci(m.lethal_per_100k,m.new_firms_per_10k))
nm=m[m.is_muang==0]
print('percap non-muang',rho_ci(nm.lethal_per_100k,nm.new_firms_per_10k))
print('pop total by prov', m.groupby('province').population.sum().to_dict())
print(m.sort_values('lethal_per_100k',ascending=False)[['province','district','population','ucdp_lethal_events_2564_2568','lethal_per_100k','new_firms_2565_2568','new_firms_per_10k']].round(1).to_string())
m.to_csv('/tmp/fact_district_v2.csv',index=False,encoding='utf-8-sig')
