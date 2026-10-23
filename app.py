#!/usr/bin/env python3
"""LNG Business Case NPV Calculator v15.
Run: python -m streamlit run lng_business_case_npv_calculator_v15.py
"""
from __future__ import annotations
import io, re, zipfile, tempfile, subprocess, sys, importlib
from datetime import datetime, timedelta
from xml.etree import ElementTree as ET
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title='LNG Business Case NPV Calculator',page_icon='◆',layout='wide')
st.markdown("""<style>.block-container{padding-top:.6rem;max-width:1580px}.hero{padding:18px 24px;border-radius:16px;background:linear-gradient(110deg,#071f25,#0c4d52 72%,#b8163c);color:#fff}.hero h1{margin:0}.hero p{margin:.3rem 0 0;color:#dff5f2}.brand{display:flex;align-items:center;gap:16px}.eq{width:58px}.sea{position:relative;height:90px;margin:8px 4px 14px;overflow:hidden;border-radius:14px;background:linear-gradient(#bce9f2 0 55%,#2f98aa 56% 73%,#176879 74%)}.ship{position:absolute;top:21px;left:3%;width:138px;height:58px;animation:sail 14s ease-in-out infinite;filter:drop-shadow(0 7px 5px #0004)}.hull{position:absolute;left:3px;bottom:0;width:132px;height:25px;background:linear-gradient(#eef4f5 0 42%,#173e4b 43%);clip-path:polygon(0 0,100% 0,80% 100%,18% 100%)}.bridge{position:absolute;left:83px;bottom:24px;width:31px;height:20px;background:#edf6f6}.stack{position:absolute;left:16px;bottom:25px;display:grid;grid-template-columns:repeat(4,19px);gap:2px}.stack i{width:19px;height:13px;background:#d53054}.stack i:nth-child(2n){background:#278e96}.wake{position:absolute;left:-23px;bottom:5px;width:35px;height:8px;border-top:3px solid #fff;border-radius:50%}@keyframes sail{0%{left:3%;transform:scaleX(1)}45%{left:88%;transform:scaleX(1)}50%{left:88%;transform:scaleX(-1)}95%{left:3%;transform:scaleX(-1)}100%{left:3%;transform:scaleX(1)}}[data-testid=stMetric]{position:relative;overflow:hidden;padding:16px!important;border:1px solid #dce7e7;border-radius:16px!important;box-shadow:0 7px 18px #03212716}[data-testid=stMetric]:before{content:'';position:absolute;left:0;top:0;bottom:0;width:5px;background:linear-gradient(#e51b46,#0d8f83)}.stButton>button{border-radius:10px;font-weight:700}</style>""",unsafe_allow_html=True)
logo='<svg class="eq" viewBox="0 0 80 80"><g fill="#e51b46"><path d="M40 4 48 27 40 36 32 27Z"/><path d="M8 20 32 29 36 40 24 42Z"/><path d="M72 20 48 29 44 40 56 42Z"/><path d="M15 62 34 43 42 46 34 70Z"/><path d="M65 62 46 43 38 46 46 70Z"/></g></svg>'
st.markdown(f'<div class="hero"><div class="brand">{logo}<div><h1>LNG Business Case NPV Calculator</h1><p>Transparent netback, break-even formula, shipping split and sensitivities.</p></div></div></div><div class="sea"><div class="ship"><div class="wake"></div><div class="hull"></div><div class="stack">'+''.join('<i></i>' for _ in range(8))+'</div><div class="bridge"></div></div></div>',unsafe_allow_html=True)

MONTHS=[datetime(2000,m,1).strftime('%b') for m in range(1,13)];NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';RN='{http://schemas.openxmlformats.org/package/2006/relationships}';DR='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
def cn(s):
 n=0
 for x in s:n=n*26+ord(x)-64
 return n
def cr(s):
 m=re.match(r'([A-Z]+)(\d+)',s.upper());return(int(m.group(2)),cn(m.group(1)))if m else(0,0)
def col(n):
 s=''
 while n:n,r=divmod(n-1,26);s=chr(65+r)+s
 return s
def xcells(raw,sheet,pred):
 z=zipfile.ZipFile(io.BytesIO(raw));ss=[]
 try:ss=[''.join(t.text or '' for t in q.iter(f'{NS}t'))for q in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(f'{NS}si')]
 except KeyError:pass
 wb=ET.fromstring(z.read('xl/workbook.xml'));rel=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'));tg={q.attrib['Id']:q.attrib['Target']for q in rel.findall(f'{RN}Relationship')};path=None
 for q in wb.find(f'{NS}sheets'):
  if q.attrib.get('name','').strip().lower()==sheet.lower():path=tg[q.attrib[f'{DR}id']];break
 if not path:raise ValueError(f'Missing sheet: {sheet}')
 path=('xl/'+path.lstrip('/')).replace('xl/xl/','xl/');out={}
 with z.open(path)as f:
  for _,c in ET.iterparse(f,events=('end',)):
   if c.tag!=f'{NS}c':continue
   ref=c.attrib.get('r','');r,k=cr(ref)
   if pred(r,k):
    typ=c.attrib.get('t');v=c.find(f'{NS}v');ins=c.find(f'{NS}is');val=None
    if ins is not None:val=''.join(t.text or ''for t in ins.iter(f'{NS}t'))
    elif v is not None and v.text is not None:
     q=v.text
     if typ=='s':val=ss[int(q)]
     else:
      try:val=float(q)
      except:val=q
    out[ref.upper()]=val
   c.clear()
 z.close();return out
def yr(v):
 if isinstance(v,(int,float)):
  if 2000<=v<=2100:return int(v)
  if 30000<=v<=80000:return(datetime(1899,12,30)+timedelta(days=v)).year
 if isinstance(v,str):
  m=re.search(r'(20\d{2})',v)
  if m:return int(m.group(1))
@st.cache_data(show_spinner=False)
def gas(raw):
 c=xcells(raw,'Gas Basecase',lambda r,k:8<=r<=15 and k<=250);h=[]
 for k in range(4,251):
  y=yr(c.get(f'{col(k)}8'));q=c.get(f'{col(k)}9')
  if y and isinstance(q,str)and re.fullmatch(r'Q[1-4]',q,re.I):h.append((k,y,int(q[-1])))
 out=[]
 for r in range(10,16):
  m=str(c.get(f'B{r}','')).strip();u=str(c.get(f'C{r}',''))
  if not m or'USD/MMBtu'not in u:continue
  for k,y,q in h:
   v=c.get(f'{col(k)}{r}')
   if isinstance(v,(int,float)):out.append({'Market':m,'Year':y,'Quarter':q,'Price':float(v),'Source':'Gas'})
 if not out:raise ValueError('No gas prices found.')
 return pd.DataFrame(out)
@st.cache_data(show_spinner=False)
def brent(raw):
 c=xcells(raw,'Brent',lambda r,k:7<=r<=8 and k<=250);out=[]
 for k in range(1,251):
  y=yr(c.get(f'{col(k)}7'));v=c.get(f'{col(k)}8')
  if y and isinstance(v,(int,float)):
   for q in range(1,5):out.append({'Market':'Brent price','Year':y,'Quarter':q,'Price':float(v),'Source':'Brent'})
 if not out:raise ValueError('No Brent prices found.')
 return pd.DataFrame(out)
@st.cache_data(show_spinner=False)
def shipping(raw):
 c=xcells(raw,'Interface',lambda r,k:(r==45 and k==13)or(79<=r<=99 and 9<=k<=37));ys={}
 for k in range(9,38):
  for r in range(79,83):
   y=yr(c.get(f'{col(k)}{r}'))
   if y:ys[k]=y;break
 out=[]
 for k,y in ys.items():out.append({'Year':y,'Shipping Opex USD/MMBtu':float(c.get(f'{col(k)}98',0)or 0),'Shipping Capex USD/MMBtu':float(c.get(f'{col(k)}97',0)or 0),'Other Shipping Cost USD/MMBtu':sum(float(c.get(f'{col(k)}{r}',0)or 0)for r in range(93,97))})
 b=float(c.get('M45'));b=b/100 if b>1 else b
 return pd.DataFrame(out).groupby('Year',as_index=False).mean(numeric_only=True),b
def cpif(years,base,rates):
 d={base:1.}
 for y in range(base+1,max(years+[base])+1):d[y]=d[y-1]*(1+rates.get(y,rates.get(y-1,.02)))
 for y in range(base-1,min(years+[base])-1,-1):d[y]=d[y+1]/(1+rates.get(y+1,.02))
 return{y:d[y]for y in years}
def seed(years,n,mode,months):
 allowed={'Quarterly':['Feb','May','Aug','Nov'],'Monthly':MONTHS}.get(mode,months or['Mar','Jun','Sep','Dec']);out=[]
 for y in years:
  r={'Year':y,**{m:0 for m in MONTHS}}
  if mode!='Manual by year and month':
   for i in range(n):r[allowed[i%len(allowed)]]+=1
  out.append(r)
 return pd.DataFrame(out)
def form(prefix,title,markets,default,comp=False):
 st.markdown('#### '+title);a,b=st.columns([1.4,1]);m=a.selectbox('EPA market',markets,index=markets.index(default)if default in markets else 0,key=prefix+'m');sl=b.number_input('Slope / multiplier',value=1.15 if prefix=='s'else .125,step=.0001,format='%.4f',key=prefix+'s');a,b,c,d=st.columns([.4,1,.4,1]);fs=a.selectbox('Sign',['+','−'],key=prefix+'fs');fx=b.number_input('Fixed, no CPI',value=0.,step=.05,key=prefix+'f');cs=c.selectbox('CPI sign',['+','−'],key=prefix+'cs');cf=d.number_input('Fixed, CPI-linked',value=0.,step=.05,key=prefix+'c');bc=st.number_input('Boil-off compensation (% of sales index)',0.,100.,8.5,.1,key=prefix+'bc')/100 if comp else 0
 return{'market':m,'slope':sl,'fixed':fx*(1 if fs=='+'else-1),'cpi':cf*(1 if cs=='+'else-1),'comp':bc}
def calculate(schedule,prices,pur,sal,cpi,base,disc,tax,ext,cargo,boil,shipdf,flat_opex,flat_capex,flat_other,shipmode,scales=(1,1,1,1)):
 ps,ss,hs,es=scales;lp={(r.Market,int(r.Year),int(r.Quarter)):float(r.Price)for _,r in prices.iterrows()};ls={int(r.Year):r for _,r in shipdf.iterrows()};rows=[]
 for _,r in schedule.iterrows():
  y=int(r.Year)
  for mi,m in enumerate(MONTHS,1):
   n=float(r[m]);q=(mi-1)//3+1
   if n<=0:continue
   try:pi=lp[(pur['market'],y,q)];si=lp[(sal['market'],y,q)]
   except KeyError as e:raise ValueError(f'Missing price {e.args[0]}')
   buy=(pi*pur['slope']+pur['fixed']+pur['cpi']*cpi[y])*ps;sbase=(si*sal['slope']+sal['fixed']+sal['cpi']*cpi[y])*ss;bc=si*sal['comp']*ss
   if shipmode=='Upload shipping Excel':
    if y not in ls:raise ValueError(f'Missing shipping {y}')
    o=float(ls[y]['Shipping Opex USD/MMBtu'])*hs;ca=float(ls[y]['Shipping Capex USD/MMBtu'])*hs;ot=float(ls[y]['Other Shipping Cost USD/MMBtu'])*hs
   else:o=flat_opex*hs;ca=flat_capex*hs;ot=flat_other*hs
   loaded=n*cargo;deliv=loaded*(1-boil);boilc=buy*boil/(1-boil)if boil<1 else 0;eu=ext*es;pre=(sbase+bc+eu-buy-boilc-o-ca-ot)*deliv;df=(1+disc)**(y-base)
   rows.append({'Year':y,'Month':m,'Quarter':q,'Cargoes':n,'Delivered MMBtu':deliv,'Sales Revenue USD':sbase*deliv,'Boil-off Compensation Revenue USD':bc*deliv,'Extrinsic Revenue USD':eu*deliv,'Purchase Base Cost USD':buy*deliv,'Boil-off Cost USD':boilc*deliv,'Shipping Opex USD':o*deliv,'Shipping Capex USD':ca*deliv,'Other Shipping Cost USD':ot*deliv,'Pre-tax Cash Flow USD':pre,'PV Sales Revenue USD':sbase*deliv/df,'PV Boil-off Compensation Revenue USD':bc*deliv/df,'PV Extrinsic Revenue USD':eu*deliv/df,'PV Purchase Base Cost USD':buy*deliv/df,'PV Boil-off Cost USD':boilc*deliv/df,'PV Shipping Opex USD':o*deliv/df,'PV Shipping Capex USD':ca*deliv/df,'PV Other Shipping Cost USD':ot*deliv/df,'PV Pre-tax Cash Flow USD':pre/df,'PV Delivered MMBtu':deliv/df})
 if not rows:raise ValueError('No cargoes scheduled.')
 d=pd.DataFrame(rows);a=d.groupby('Year',as_index=False)['Pre-tax Cash Flow USD'].sum();a['Corporate Tax USD']=a['Pre-tax Cash Flow USD'].clip(lower=0)*tax;a['After-tax Cash Flow USD']=a['Pre-tax Cash Flow USD']-a['Corporate Tax USD'];a['PV Corporate Tax USD']=a.apply(lambda r:r['Corporate Tax USD']/(1+disc)**(int(r.Year)-base),axis=1);a['PV After-tax Cash Flow USD']=a.apply(lambda r:r['After-tax Cash Flow USD']/(1+disc)**(int(r.Year)-base),axis=1);return d,a
def totals(d,a):return{'revenue':d['PV Sales Revenue USD'].sum(),'comp':d['PV Boil-off Compensation Revenue USD'].sum(),'ext':d['PV Extrinsic Revenue USD'].sum(),'purchase':d['PV Purchase Base Cost USD'].sum(),'boil':d['PV Boil-off Cost USD'].sum(),'opex':d['PV Shipping Opex USD'].sum(),'capex':d['PV Shipping Capex USD'].sum(),'other':d['PV Other Shipping Cost USD'].sum(),'tax':a['PV Corporate Tax USD'].sum(),'npv_before_tax':d['PV Pre-tax Cash Flow USD'].sum(),'npv':a['PV After-tax Cash Flow USD'].sum(),'volume':d['PV Delivered MMBtu'].sum()}
def solve_zero(inputs,side,target,cpi_share):
 base=inputs[side].copy()
 def evalx(x):
  z={k:(v.copy()if isinstance(v,dict)else v)for k,v in inputs.items()}
  if target=='Slope / multiplier':z[side]['slope']=x
  else:z[side]['fixed']=x*(1-cpi_share);z[side]['cpi']=x*cpi_share
  d,a=calculate(**z);return totals(d,a)['npv']
 lo,hi=-20.,20.;flo,fhi=evalx(lo),evalx(hi)
 for _ in range(12):
  if flo==0:return lo
  if fhi==0:return hi
  if flo*fhi<0:break
  lo*=2;hi*=2;flo,fhi=evalx(lo),evalx(hi)
 if flo*fhi>0:raise ValueError('A zero-NPV solution could not be bracketed. Review the deal inputs or allowable formula range.')
 for _ in range(100):
  mid=(lo+hi)/2;fm=evalx(mid)
  if abs(fm)<.01:return mid
  if flo*fm<=0:hi=mid;fhi=fm
  else:lo=mid;flo=fm
 return(lo+hi)/2
def formula_text(side,f):
 idx=f['market'];slope=f['slope'];fixed=f['fixed'];cpi=f['cpi'];parts=[f'{slope*100:.4f}% × {idx}']
 if abs(fixed)>1e-10:parts.append(('+'if fixed>=0 else'−')+f' {abs(fixed):.4f}')
 if abs(cpi)>1e-10:parts.append(('+'if cpi>=0 else'−')+f' {abs(cpi):.4f} (CPI)')
 if side=='sal'and abs(f.get('comp',0))>1e-10:parts.append(f"+ {f['comp']*100:.2f}% × {idx} boil-off compensation")
 return' '.join(parts)

def ensure_ppt():
 for mod,pkg in [('pptx','python-pptx'),('matplotlib','matplotlib')]:
  try:importlib.import_module(mod)
  except ImportError:subprocess.check_call([sys.executable,'-m','pip','install','--quiet',pkg])
def ppt_bytes(d,a,inp):
 ensure_ppt();from pptx import Presentation;from pptx.util import Inches,Pt;from pptx.dml.color import RGBColor;import matplotlib.pyplot as plt,os
 t=totals(d,a);prs=Presentation();prs.slide_width=Inches(13.333);prs.slide_height=Inches(7.5);files=[]
 def pic(kind):
  fig,ax=plt.subplots(figsize=(11,5.5))
  if kind=='waterfall':labs=['Revenue','Comp.','Extrinsic','Purchase','Boil-off','Opex','Capex','Other','Tax'];vs=[t['revenue'],t['comp'],t['ext'],-t['purchase'],-t['boil'],-t['opex'],-t['capex'],-t['other'],-t['tax']];starts=[];r=0
  else:labs=[];vs=[]
  if kind=='waterfall':
   for v in vs:starts.append(r);r+=v
   ax.bar(range(len(vs)),vs,bottom=starts,color=['#0d8f83'if v>=0 else'#e51b46'for v in vs]);ax.bar([len(vs)],[t['npv']],color='#123842');ax.set_xticks(range(len(vs)+1),labs+['NPV'],rotation=25,ha='right')
  elif kind=='cash':ax.plot(a.Year,a['Pre-tax Cash Flow USD'],marker='o',label='Pre-tax');ax.plot(a.Year,a['After-tax Cash Flow USD'],marker='o',label='After-tax');ax.plot(a.Year,a['PV After-tax Cash Flow USD'],marker='o',ls='--',label='PV after tax');ax.legend()
  else:
   base=t['npv'];spec=[('Sales',10,1),('Purchase',10,0),('Shipping',10,2),('Extrinsic',10,3)];rs=[]
   for n,p,i in spec:
    lo=[1]*4;hi=[1]*4;lo[i]=.9;hi[i]=1.1;dl,al=calculate(**inp,scales=tuple(lo));dh,ah=calculate(**inp,scales=tuple(hi));rs.append((n,totals(dl,al)['npv'],totals(dh,ah)['npv']))
   ax.barh(range(4),[base-min(x,y)for _,x,y in rs],left=[min(x,y)for _,x,y in rs],color='#e51b46');ax.barh(range(4),[max(x,y)-base for _,x,y in rs],left=base,color='#0d8f83');ax.set_yticks(range(4),[r[0]for r in rs]);ax.axvline(base,color='#123842')
  fig.tight_layout();f=tempfile.NamedTemporaryFile(suffix='.png',delete=False);fig.savefig(f.name,dpi=170);plt.close(fig);files.append(f.name);return f.name
 sl=prs.slides.add_slide(prs.slide_layouts[6]);sl.background.fill.solid();sl.background.fill.fore_color.rgb=RGBColor(9,36,43);tb=sl.shapes.add_textbox(Inches(.8),Inches(2.2),Inches(11.5),Inches(2));p=tb.text_frame.paragraphs[0];p.text='LNG Business Case NPV';p.font.size=Pt(34);p.font.color.rgb=RGBColor(255,255,255)
 for title,kind in [('NPV waterfall','waterfall'),('Cash flow per year','cash'),('NPV sensitivities','sens')]:
  sl=prs.slides.add_slide(prs.slide_layouts[6]);sl.shapes.add_picture(pic(kind),Inches(.65),Inches(1),width=Inches(12),height=Inches(5.8));tb=sl.shapes.add_textbox(Inches(.65),Inches(.25),Inches(12),Inches(.5));tb.text_frame.paragraphs[0].text=title
 out=io.BytesIO();prs.save(out)
 for f in files:
  try:os.remove(f)
  except:pass
 return out.getvalue()

calc_tab,price_tab,ship_tab,sens_tab=st.tabs(['Calculator','Prices','Shipping','Sensitivities'])
with calc_tab:
 with st.expander('1. EPA files, period and CPI',expanded=True):
  a,b=st.columns(2)
  with a:
   f=st.file_uploader('EPA Gas Prices',type=['xlsx'])
   if f:
    try:st.session_state.gas=gas(f.getvalue());st.success('Gas prices loaded')
    except Exception as e:st.error(str(e))
  with b:
   f=st.file_uploader('EPA Brent Prices',type=['xlsx'])
   if f:
    try:st.session_state.brent=brent(f.getvalue());st.success('Brent prices loaded')
    except Exception as e:st.error(str(e))
  frames=[st.session_state[x]for x in['gas','brent']if x in st.session_state];prices=pd.concat(frames,ignore_index=True)if frames else None;a,b,c=st.columns(3);start=int(a.number_input('First cargo year',2020,2100,2028));end=int(b.number_input('Last cargo year',2020,2100,2037));base=int(c.number_input('Base year',2020,2100,2026));years=list(range(start,end+1));mode=st.radio('Annual CPI',['Fixed annual CPI','Preset CPI curve','Specific CPI by year'],horizontal=True)
  if mode=='Fixed annual CPI':p=st.number_input('Annual CPI (%)',value=2.,step=.1)/100;rates={y:p for y in range(min(years+[base]),max(years+[base])+1)}
  elif mode=='Preset CPI curve':
   p=st.selectbox('Preset',['2.0% long-term','2.7%, 2.4%, then 2.0%','No inflation']);rates={y:(0 if p=='No inflation'else .02)for y in range(min(years+[base]),max(years+[base])+1)}
   if p.startswith('2.7'):rates[base+1]=.027;rates[base+2]=.024
  else:d=st.data_editor(pd.DataFrame({'Year':range(min(years+[base]),max(years+[base])+1),'CPI %':2.}),hide_index=True);rates={int(r.Year):float(r['CPI %'])/100 for _,r in d.iterrows()}
  cpi=cpif(years,base,rates)
 with st.expander('2. Purchase and sales formulas',expanded=True):
  markets=sorted(prices.Market.unique())if prices is not None else['Upload prices'];a,b=st.columns(2)
  with a:pur=form('p','Purchase price',markets,'Brent price')
  with b:sal=form('s','Sales price',markets,'TTF price',True)
 with st.expander('3. Volumes and schedule',expanded=True):
  a,b,c=st.columns(3);mode2=a.selectbox('Allocation',['Even through selected months','Quarterly','Monthly','Manual by year and month']);n=int(b.number_input('Cargoes/year',0,100,4));cargo=float(c.number_input('Loaded cargo MMBtu',0.,10_000_000.,3_800_000.,100_000.,format='%.0f'));months=st.multiselect('Eligible months',MONTHS,default=['Mar','Jun','Sep','Dec'])if mode2 in['Even through selected months','Manual by year and month']else[];schedule=st.data_editor(seed(years,n,mode2,months),hide_index=True,use_container_width=True,key=f's{start}{end}{mode2}')
 with st.expander('4. Shipping and boil-off',expanded=True):
  shipmode=st.radio('Shipping source',['Flat fixed assumption','Upload shipping Excel'],horizontal=True);fo=fc=ft=boil=0.;shipdf=pd.DataFrame(columns=['Year','Shipping Opex USD/MMBtu','Shipping Capex USD/MMBtu','Other Shipping Cost USD/MMBtu'])
  if shipmode=='Flat fixed assumption':
   a,b,c,d=st.columns(4);fo=float(a.number_input('Shipping Opex USD/MMBtu',0.,value=.16,step=.01));fc=float(b.number_input('Shipping Capex USD/MMBtu',0.,value=.57,step=.01));ft=float(c.number_input('Other Shipping Cost USD/MMBtu',0.,value=0.,step=.01));boil=float(d.number_input('Physical boil-off %',0.,99.,1.5,.1))/100;shipdf=pd.DataFrame({'Year':years,'Shipping Opex USD/MMBtu':fo,'Shipping Capex USD/MMBtu':fc,'Other Shipping Cost USD/MMBtu':ft})
  else:
   f=st.file_uploader('Shipping workbook',type=['xlsx','xlsm'])
   if f:
    try:st.session_state.ship=shipping(f.getvalue());st.success('Shipping loaded')
    except Exception as e:st.error(str(e))
   if 'ship'in st.session_state:shipdf,boil=st.session_state.ship
  st.metric('Sales cargo size',f'{cargo*(1-boil):,.0f} MMBtu',f'{boil:.3%} boil-off')
 with st.expander('5. Financial assumptions',expanded=True):
  a,b,c=st.columns(3);disc=float(a.number_input('Discount rate %',0.,100.,8.,.25))/100;tax=float(b.number_input('Corporate tax rate %',0.,100.,22.,.5))/100;ext=float(c.number_input('Extrinsic USD/MMBtu',0.,20.,0.,.05))
 inp={'schedule':schedule,'prices':prices,'pur':pur,'sal':sal,'cpi':cpi,'base':base,'disc':disc,'tax':tax,'ext':ext,'cargo':cargo,'boil':boil,'shipdf':shipdf,'flat_opex':fo,'flat_capex':fc,'flat_other':ft,'shipmode':shipmode};st.session_state.inp=inp
 if st.button('Calculate NPV',type='primary',use_container_width=True,disabled=prices is None):
  try:st.session_state.detail,st.session_state.annual=calculate(**inp);st.session_state.pop('ppt',None)
  except Exception as e:st.error(str(e));st.session_state.pop('detail',None)
 if 'detail'in st.session_state:
  with st.expander('Find Break Even Price',expanded=False):
   a,b,c=st.columns(3);side_label=a.selectbox('Find break-even for',['Purchase price','Sales price']);target=b.selectbox('Solve for',['Slope / multiplier','Fixed elements']);share=c.selectbox('CPI-linked share of fixed elements',list(range(0,101,5)),index=0,format_func=lambda x:f'{x}%');side='pur'if side_label.startswith('Purchase')else'sal'
   if st.button('Calculate Break-Even Price',use_container_width=True):
    try:
     root=solve_zero(inp,side,target,share/100);f=inp[side].copy()
     if target=='Slope / multiplier':f['slope']=root
     else:f['fixed']=root*(1-share/100);f['cpi']=root*(share/100)
     st.session_state.be=(side_label,formula_text(side,f),root,target,share)
    except Exception as e:st.error(str(e))
   if 'be'in st.session_state:
    lab,txt,root,target,sh=st.session_state.be;st.success(f'Break-even {lab}: {txt}');st.caption(f'Solved {target.lower()} for NPV after tax = 0. Fixed-element CPI share: {sh}%. All other assumptions are unchanged.')
  if st.button('Make a PowerPoint Presentation',use_container_width=True):
   try:st.session_state.ppt=ppt_bytes(st.session_state.detail,st.session_state.annual,st.session_state.inp)
   except Exception as e:st.error(str(e))
  if st.session_state.get('ppt'):st.download_button('Download PowerPoint Presentation',st.session_state.ppt,'LNG_Business_Case_NPV.pptx','application/vnd.openxmlformats-officedocument.presentationml.presentation',use_container_width=True)
  d,a=st.session_state.detail,st.session_state.annual;t=totals(d,a);x,y,z,w,v=st.columns(5);x.metric('NPV before tax',f"${t['npv_before_tax']/1e6:,.1f}m");y.metric('NPV before tax / MMBtu',f"${t['npv_before_tax']/t['volume']:.3f}");z.metric('NPV after tax',f"${t['npv']/1e6:,.1f}m");w.metric('NPV after tax / MMBtu',f"${t['npv']/t['volume']:.3f}");v.metric('PV Volume',f"{t['volume']/1e6:.1f} TBtu")
  unit=st.radio('Chart view',['$','$/MMBtu'],horizontal=True);div=1 if unit=='$'else t['volume'];vals=[t['revenue']/div,t['comp']/div,t['ext']/div,-t['purchase']/div,-t['boil']/div,-t['opex']/div,-t['capex']/div,-t['other']/div,-t['tax']/div,t['npv']/div];labs=['PV Sales','PV Boil-off Comp.','PV Extrinsic','PV Purchase','PV Boil-off Cost','PV Ship Opex','PV Ship Capex','PV Other Ship','PV Tax','NPV After Tax'];f=go.Figure(go.Waterfall(x=labs,y=vals,measure=['relative']*9+['total']));f.update_layout(height=470);st.plotly_chart(f,use_container_width=True)
  d['Period']=d.Year.astype(str)+' '+d.Month;mf=go.Figure();comps=[('Sales','Sales Revenue USD',1),('Boil-off comp.','Boil-off Compensation Revenue USD',1),('Extrinsic','Extrinsic Revenue USD',1),('Purchase','Purchase Base Cost USD',-1),('Boil-off cost','Boil-off Cost USD',-1),('Ship Opex','Shipping Opex USD',-1),('Ship Capex','Shipping Capex USD',-1),('Other ship','Other Shipping Cost USD',-1)]
  for name,colu,sgn in comps:mf.add_trace(go.Bar(x=d.Period,y=d[colu]*sgn if unit=='$'else d[colu]*sgn/d['Delivered MMBtu'],name=name))
  mf.update_layout(barmode='relative',height=480,legend={'orientation':'h'});st.plotly_chart(mf,use_container_width=True)
  lf=go.Figure();lf.add_trace(go.Scatter(x=a.Year,y=a['Pre-tax Cash Flow USD'],mode='lines+markers',name='Pre-tax'));lf.add_trace(go.Scatter(x=a.Year,y=a['After-tax Cash Flow USD'],mode='lines+markers',name='After-tax'));lf.add_trace(go.Scatter(x=a.Year,y=a['PV After-tax Cash Flow USD'],mode='lines+markers',name=f'PV to {base}',line={'dash':'dot'}));lf.update_layout(height=420,legend={'orientation':'h'});st.plotly_chart(lf,use_container_width=True)
with price_tab:
 st.subheader('All uploaded EPA prices');frames=[st.session_state[x]for x in['gas','brent']if x in st.session_state]
 if not frames:st.info('Upload prices first.')
 else:
  p=pd.concat(frames);sel=st.multiselect('Series',sorted(p.Market.unique()),default=sorted(p.Market.unique()));v=p[p.Market.isin(sel)].copy();v['Period']=v.Year.astype(str)+' Q'+v.Quarter.astype(str);st.dataframe(v,hide_index=True,use_container_width=True);f=go.Figure()
  for x in sel:q=v[v.Market==x];f.add_trace(go.Scatter(x=q.Period,y=q.Price,mode='lines+markers',name=x))
  st.plotly_chart(f,use_container_width=True)
with ship_tab:
 st.subheader('Shipping assumptions')
 if 'detail'not in st.session_state:st.info('Calculate first.')
 else:
  sd=st.session_state.detail.groupby('Year',as_index=False).agg({'Delivered MMBtu':'sum','Shipping Opex USD':'sum','Shipping Capex USD':'sum','Other Shipping Cost USD':'sum'})
  for c,nm in [('Shipping Opex USD','Shipping Opex USD/MMBtu'),('Shipping Capex USD','Shipping Capex USD/MMBtu'),('Other Shipping Cost USD','Other Shipping Cost USD/MMBtu')]:sd[nm]=sd[c]/sd['Delivered MMBtu']
  sd['Total USD']=sd[['Shipping Opex USD','Shipping Capex USD','Other Shipping Cost USD']].sum(axis=1);sd['Total USD/MMBtu']=sd['Total USD']/sd['Delivered MMBtu'];sd['Year label']=sd.Year.astype(str);st.dataframe(sd.drop(columns='Year label'),hide_index=True,use_container_width=True);unit=st.radio('Shipping view',['$','$/MMBtu'],horizontal=True);f=go.Figure();series=[('Shipping Opex USD','Opex'),('Shipping Capex USD','Capex'),('Other Shipping Cost USD','Other')]if unit=='$'else[('Shipping Opex USD/MMBtu','Opex'),('Shipping Capex USD/MMBtu','Capex'),('Other Shipping Cost USD/MMBtu','Other')]
  for c,n in series:f.add_trace(go.Bar(x=sd['Year label'],y=sd[c],name=n))
  f.update_layout(barmode='stack',xaxis={'type':'category'},height=450);st.plotly_chart(f,use_container_width=True);tot=sd['Total USD']if unit=='$'else sd['Total USD/MMBtu'];g=go.Figure(go.Scatter(x=sd['Year label'],y=tot,mode='lines+markers+text',text=[f'{x:,.3f}'for x in tot],name='Total shipping'));g.update_layout(xaxis={'type':'category'},height=390);st.plotly_chart(g,use_container_width=True)
with sens_tab:
 st.subheader('NPV sensitivities')
 if 'detail'not in st.session_state:st.info('Calculate first.')
 else:
  i=st.session_state.inp;bt=totals(st.session_state.detail,st.session_state.annual);bn=bt['npv'];vol=bt['volume'];unit=st.radio('View',['$','$/MMBtu'],horizontal=True);a,b,c,d=st.columns(4);ss=a.slider('Sales price change (%)',0,20,10);ps=b.slider('Purchase price change (%)',0,20,10);hs=c.slider('Shipping cost change (%)',0,20,10);es=d.slider('Extrinsic value change (%)',0,20,10)
  def run(sc):x,y=calculate(**i,scales=sc);return totals(x,y)['npv']
  rows=[]
  for n,p,ix in[('Sales price',ss,1),('Purchase price',ps,0),('Shipping cost',hs,2),('Extrinsic value',es,3)]:
   lo=[1]*4;hi=[1]*4;lo[ix]=1-p/100;hi[ix]=1+p/100;rows.append((n,p,run(tuple(lo)),run(tuple(hi))))
  q=pd.DataFrame(rows,columns=['Driver','Percent','Low','High']);div=1 if unit=='$'else vol;bd=bn/div;human=lambda x:f'${x/1e6:.1f}m'if unit=='$'else f'${x/vol:.3f}/MMBtu';f=go.Figure();f.add_trace(go.Bar(y=['Base case'],x=[bd],orientation='h',name='Base case NPV',text=[human(bn)],textposition='inside'))
  f.add_trace(go.Bar(y=q.Driver,x=(q.High-bn)/div,base=bd,orientation='h',name='Higher-input scenario',text=[f'NPV when {r.Driver.lower()} is {r.Percent}% higher: {human(r.High)}'for _,r in q.iterrows()],textposition='inside'))
  f.add_trace(go.Bar(y=q.Driver,x=(q.Low-bn)/div,base=bd,orientation='h',name='Lower-input scenario',text=[f'NPV when {r.Driver.lower()} is {r.Percent}% lower: {human(r.Low)}'for _,r in q.iterrows()],textposition='inside'));f.update_layout(barmode='overlay',height=600,legend={'orientation':'h'});st.plotly_chart(f,use_container_width=True)

