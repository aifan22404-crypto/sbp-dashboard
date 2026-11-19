import openpyxl,glob,json,csv,collections,os
R=os.path.expanduser('~/mnt/SBP_Dashboard')
def sheet(f,n): return [list(r) for r in openpyxl.load_workbook(f,read_only=True)[n].iter_rows(values_only=True)]
raw=glob.glob(R+'/01_raw/01_economy_gpp/NESDC_GPP_edition2024*.xlsx')[0]
wk=sheet(raw,'WK')
hdr=[i for i,r in enumerate(wk) if r[1] and str(r[1]).startswith('1995')][0]
cols=[str(c) for c in wk[hdr]]
nat={}
for r in wk[hdr+1:]:
    if r[0] and isinstance(r[-1],(int,float)) and str(r[0]).strip() not in nat: nat[str(r[0]).strip()]={cols[i]:r[i] for i in range(1,len(cols)) if isinstance(r[i],(int,float))}
print([ (k,round(v.get('2024p',0))) for k,v in nat.items()][:30])
g=R+'/02_clean/GPP_4prov_2019-2024_clean.xlsx'
sl=sheet(g,'sectors_long'); ss=sheet(g,'sector_structure_2021_vs_2024')
print(sorted(set(r[3] for r in sl[1:])))
print(ss[0]); print(set(r[2] for r in sl[1:]))

# ---------- LQ ----------
tot=nat['Gross domestic product (GDP)']['2024p']
def nk(s): return ' '.join(str(s).split())
natn={nk(k):v for k,v in nat.items()}
lq=[]
for r in ss[1:]:
    p,s,v,sh,sh21,dpp,cagr=r[:7]
    n=natn.get(nk(s))
    if not n or sh is None: continue
    nsh=n['2024p']/tot*100
    lq.append(dict(p=p,s=nk(s),v=round(v,1),sh=round(sh,2),nsh=round(nsh,2),lq=round(sh/nsh,2),cagr=cagr))
# ---------- districts ----------
S=R+'/02_clean/Security_4prov_2562-2568_clean.xlsx'; D=R+'/02_clean/DBD_firms_by_business_type_2565-2568_clean.xlsx'
ev={(r[0],r[2]):(r[3],r[4]) for r in sheet(S,'ucdp_by_district_2021-2025')[1:] if r[2]}
dist=[]
for r in sheet(D,'by_district')[1:]:
    if r[2]!='IN SCOPE': continue
    e=ev.get((r[0],r[1]),(0,0))
    dist.append(dict(p=r[0],d=r[1],new=r[3],dis=r[4],ev=e[0],deaths=e[1],ratio=round(r[4]/r[3]*100,1) if r[3] else None))
def rank(x):
    o=sorted(range(len(x)),key=lambda i:x[i]); rk=[0]*len(x); i=0
    while i<len(o):
        j=i
        while j+1<len(o) and x[o[j+1]]==x[o[i]]: j+=1
        for k in range(i,j+1): rk[o[k]]=(i+j)/2+1
        i=j+1
    return rk
def spear(a,b):
    ra,rb=rank(a),rank(b); n=len(a); ma=sum(ra)/n; mb=sum(rb)/n
    num=sum((x-ma)*(y-mb) for x,y in zip(ra,rb)); da=sum((x-ma)**2 for x in ra)**.5; db=sum((y-mb)**2 for y in rb)**.5
    return round(num/(da*db),3)
nonm=[d for d in dist if not d['d'].startswith('เมือง')]
print('n',len(dist),'rho ev~new',spear([d['ev'] for d in dist],[d['new'] for d in dist]),'excl mueang',spear([d['ev'] for d in nonm],[d['new'] for d in nonm]),
      'rho ev~ratio',spear([d['ev'] for d in dist if d['ratio'] is not None],[d['ratio'] for d in dist if d['ratio'] is not None]))
hi=[d for d in nonm if d['ev']>=5]; lo=[d for d in nonm if d['ev']<5]
import statistics as st
print('median new: hi-ev',st.median(d['new'] for d in hi),len(hi),'lo-ev',st.median(d['new'] for d in lo),len(lo))
# NWT official by district vs firms
off=collections.Counter(); typ=collections.defaultdict(dict)
for r in sheet(S,'official_NWT_district_type')[1:]:
    if r[0] and 2565<=int(r[0])<=2568: off[r[1]]+=int(r[3] or 0)
    if r[0]: typ[int(r[0])][r[2]]=typ[int(r[0])].get(r[2],0)+int(r[3] or 0)
nw=[d for d in dist if d['p']=='Narathiwat']
for d in nw: d['off']=off.get(d['d'],0)
print('NWT rho official~new',spear([d['off'] for d in nw],[d['new'] for d in nw]), 'excl mueang', spear([d['off'] for d in nw if not d['d'].startswith('เมือง')],[d['new'] for d in nw if not d['d'].startswith('เมือง')]))
# ---------- other tables ----------
fish=[dict(y=r[0],p=r[2],t=r[3],v=r[4],st=r[5],sv=r[6]) for r in sheet(R+'/02_clean/Fisheries_DoF_4prov_2562-2567_clean.xlsx','landings_province_total')[1:]]
disb=[dict(d=r[0],t=r[1],a=r[2],x=r[3],pct=r[4]) for r in sheet(R+'/02_clean/Budget_4prov_FY2564-2569_clean.xlsx','NWT_disbursement_FY2568')[1:]]
print(disb[-4:]); print(fish)
sec=[dict(a=r[0],s=r[1],th=r[2],new=r[3],dis=r[4],net=r[5],per=r[6],cap=r[7]) for r in sheet(D,'summary_2565-2568_by_section')[1:]]
ton=[dict(zip(['y','m','p','occ68','occ67','g68','g67','v68','v67'],r[0:1]+r[2:10])) for r in sheet(R+'/02_clean/Tourism_4prov_2564-2568_clean.xlsx','monthly_2568_vs_2567')[1:]]
panel=list(csv.DictReader(open(R+'/03_analysis/00_master_panel/SBP_master_panel_long.csv',encoding='utf-8-sig')))
ind={}; vals=[]
for x in panel:
    c=x['indicator_code']
    if '_sec' in c and c.startswith('dbd_'): continue
    ind[c]=dict(th=x['indicator_th'],en=x['indicator_en'],u=x['unit'],dom=x['domain'],src=x['source_id'])
    try: v=round(float(x['value']),3)
    except: continue
    vals.append([x['area'],int(x['year_BE']),c,v,'' if x['quality_flag']=='ok' else x['quality_flag']])
out=dict(ind=ind,vals=vals,lq=lq,dist=dist,fish=fish,disb=disb,sec=sec,tour_m=ton,inc_type={y:t for y,t in typ.items()},built='2026-11-16')
os.makedirs(R+'/04_output/dashboard',exist_ok=True)
json.dump(out,open(R+'/04_output/dashboard/data.json','w',encoding='utf-8'),ensure_ascii=False,separators=(',',':'))
print(len(vals),len(ind),os.path.getsize(R+'/04_output/dashboard/data.json'))
for p in ['Pattani','Yala','Narathiwat']:
    print(p,[(x['s'][:22],x['lq'],x['sh'],x['cagr']) for x in sorted([l for l in lq if l['p']==p],key=lambda z:-z['lq'])[:6]])
print(ton[0])
