#!/usr/bin/env python3
"""LNG Business Case | redesigned single-file Streamlit application.
Run: python -m streamlit run lng_business_case_studio_v29.py
Dependencies: streamlit>=1.41, pandas, plotly, numpy, python-pptx, matplotlib.
V29: V27 baseline plus split fixed/CPI liquefaction boil-off, enhanced matched-formula display, shipping horizon capped at 2050 and deployment-safe input styling.
Separate leg CPI timing, CCT2 preset, annual-average cargo programme,
component-matched sales formula and explicit annual/PV hedge reconciliation.
Optional offline brand asset: equinor_logo.svg or equinor_logo.png beside this file.
The valuation conventions match the supplied calculator: annual discounting,
no taxation, separate sales/purchase CPI years and independent discount year.
"""
from __future__ import annotations
import base64
import copy
import io
import re
import zipfile
from pathlib import Path
from datetime import datetime, timedelta
from xml.etree import ElementTree as ET
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title='LNG Business Case | Studio',page_icon='◆',layout='wide',initial_sidebar_state='collapsed')
MONTHS=[datetime(2000,m,1).strftime('%b') for m in range(1,13)]
NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RN='{http://schemas.openxmlformats.org/package/2006/relationships}'
DR='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

# ---------------------------- Valuation engine ----------------------------
def cn(s):
 n=0
 for x in s:n=n*26+ord(x)-64
 return n
def cr(s):
 m=re.match(r'([A-Z]+)(\d+)',s.upper())
 return (int(m.group(2)),cn(m.group(1))) if m else (0,0)
def col(n):
 s=''
 while n:n,r=divmod(n-1,26);s=chr(65+r)+s
 return s
def xcells(raw,sheet,pred):
 z=zipfile.ZipFile(io.BytesIO(raw));ss=[]
 try:ss=[''.join(t.text or '' for t in q.iter(f'{NS}t')) for q in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(f'{NS}si')]
 except KeyError:pass
 wb=ET.fromstring(z.read('xl/workbook.xml'));rel=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
 tg={q.attrib['Id']:q.attrib['Target'] for q in rel.findall(f'{RN}Relationship')};path=None
 for q in wb.find(f'{NS}sheets'):
  if q.attrib.get('name','').strip().lower()==sheet.lower():path=tg[q.attrib[f'{DR}id']];break
 if not path:raise ValueError(f'Missing sheet: {sheet}')
 path=('xl/'+path.lstrip('/')).replace('xl/xl/','xl/');out={}
 with z.open(path) as f:
  for _,c in ET.iterparse(f,events=('end',)):
   if c.tag!=f'{NS}c':continue
   ref=c.attrib.get('r','');r,k=cr(ref)
   if pred(r,k):
    typ=c.attrib.get('t');v=c.find(f'{NS}v');ins=c.find(f'{NS}is');val=None
    if ins is not None:val=''.join(t.text or '' for t in ins.iter(f'{NS}t'))
    elif v is not None and v.text is not None:
     q=v.text
     if typ=='s':val=ss[int(q)]
     elif typ=='b':val=q=='1'
     else:
      try:val=float(q)
      except ValueError:val=q
    out[ref.upper()]=val
   c.clear()
 z.close();return out
def yr(v):
 if isinstance(v,(int,float)):
  if 2000<=v<=2100:return int(v)
  if 30000<=v<=80000:return (datetime(1899,12,30)+timedelta(days=v)).year
 if isinstance(v,str):
  m=re.search(r'(20\d{2})',v)
  if m:return int(m.group(1))
@st.cache_data(show_spinner=False)
def gas(raw):
 c=xcells(raw,'Gas Basecase',lambda r,k:8<=r<=15 and k<=250);h=[]
 for k in range(4,251):
  y=yr(c.get(f'{col(k)}8'));q=c.get(f'{col(k)}9')
  if y and isinstance(q,str) and re.fullmatch(r'Q[1-4]',q,re.I):h.append((k,y,int(q[-1])))
 out=[]
 for r in range(10,16):
  m=str(c.get(f'B{r}','')).strip();u=str(c.get(f'C{r}',''))
  if not m or 'USD/MMBtu' not in u:continue
  for k,y,q in h:
   v=c.get(f'{col(k)}{r}')
   if isinstance(v,(int,float)):out.append(dict(Market=m,Year=y,Quarter=q,Price=float(v),Source='Gas'))
 if not out:raise ValueError('No gas prices found.')
 return pd.DataFrame(out)
@st.cache_data(show_spinner=False)
def brent(raw):
 c=xcells(raw,'Brent',lambda r,k:7<=r<=8 and k<=250);out=[]
 for k in range(1,251):
  y=yr(c.get(f'{col(k)}7'));v=c.get(f'{col(k)}8')
  if y and isinstance(v,(int,float)):
   for q in range(1,5):out.append(dict(Market='Brent price',Year=y,Quarter=q,Price=float(v),Source='Brent'))
 if not out:raise ValueError('No Brent prices found.')
 return pd.DataFrame(out)
@st.cache_data(show_spinner=False)
def shipping(raw):
 c=xcells(raw,'Interface',lambda r,k:(r==45 and k==13) or (41<=r<=43 and 2<=k<=3) or (79<=r<=99 and 9<=k<=250));ys={}
 for k in range(9,251):
  for r in range(79,83):
   y=yr(c.get(f'{col(k)}{r}'))
   if y:
    if y<=2050:ys[k]=y
    break
 def number(ref):
  v=c.get(ref)
  if not isinstance(v,(int,float)) or isinstance(v,bool):raise ValueError(f'Interface!{ref}: missing numeric cached value. Recalculate and save the workbook in Excel.')
  return float(v)
 out=[]
 for k,y in ys.items():out.append({'Year':y,'Shipping Opex USD/MMBtu':sum(number(f'{col(k)}{r}') for r in [92,93,94,95,96,98]),'Shipping Capex USD/MMBtu':number(f'{col(k)}97')})
 if not out:raise ValueError('No shipping years found in Interface rows 79-82.')
 df=pd.DataFrame(out)
 if df.Year.duplicated().any():raise ValueError('Duplicate shipping years. Expected one annual column per year.')
 b=number('M45');b=b/100 if b>1 else b
 if not 0<=b<1:raise ValueError('Physical boil-off must be between 0% and less than 100%.')
 routes=[]
 for r,label in [(41,'Suez Channel'),(42,'Panama Channel'),(43,'Northern Sea Route Fee')]:
  v=c.get(f'C{r}')
  if isinstance(v,bool):value='TRUE' if v else 'FALSE'
  elif isinstance(v,(int,float)) and v in (0,1):value='TRUE' if v else 'FALSE'
  elif str(v).strip().upper() in ('TRUE','FALSE'):value=str(v).strip().upper()
  else:value='Not available'
  routes.append({'Route / fee':str(c.get(f'B{r}') or label),'Included in shipping workbook':value,'Source':f'Interface!B{r}:C{r}'})
 return df,b,pd.DataFrame(routes)

def cpif(years,base,rates):
 """Factor 1 in anchor year; include CPI for each following year."""
 import math
 if not years:return {}
 if any(not math.isfinite(float(v)) or float(v)<=-1 for v in rates.values()):raise ValueError('CPI rates must be finite and greater than -100%.')
 def rate(y):
  if y not in rates:raise ValueError(f'Missing CPI rate for {y}.')
  return float(rates[y])
 d={base:1.}
 for y in range(base+1,max(years+[base])+1):d[y]=d[y-1]*(1+rate(y))
 for y in range(base-1,min(years+[base])-1,-1):d[y]=d[y+1]/(1+rate(y+1))
 return {y:d[y] for y in years}

def seed(years,n,mode,months):
 allowed={'Quarterly':['Feb','May','Aug','Nov'],'Monthly':MONTHS}.get(mode,months or ['Mar','Jun','Sep','Dec']);out=[]
 for y in years:
  r={'Year':y,**{m:0 for m in MONTHS}}
  if mode!='Manual by year and month':
   for i in range(n):r[allowed[i%len(allowed)]]+=1
  out.append(r)
 return pd.DataFrame(out)
def calculate(schedule,prices,pur,sal,sales_cpi,purchase_cpi,sales_base,purchase_base,discount_year,disc,ext,cargo,boil,shipdf,flat_opex,flat_capex,shipmode,price_mode='Scheduled quarterly prices',scales=(1,1,1,1),epa_cpi=None,epa_real_base=2026):
 import math
 ps,ss,hs,es=scales
 if epa_cpi is None:raise ValueError('EPA real-to-nominal CPI factors are required. Recalculate with an EPA real-price reference year.')
 if not 0<=boil<1:raise ValueError('Physical boil-off must be between 0% and less than 100%.')
 if cargo<=0:raise ValueError('Loaded cargo size must be positive.')
 if disc<=-1:raise ValueError('Discount rate must exceed -100%.')
 if prices is None or prices.empty:raise ValueError('Upload EPA prices first.')
 if schedule.Year.duplicated().any():raise ValueError('Cargo schedule contains duplicate years.')
 if schedule[MONTHS].isna().any().any() or (schedule[MONTHS]<0).any().any():raise ValueError('Cargo counts must be non-negative numbers.')
 if prices.duplicated(['Market','Year','Quarter']).any():raise ValueError('Duplicate EPA market/year/quarter prices found.')
 if not prices['Price'].map(lambda v:math.isfinite(float(v))).all():raise ValueError('EPA prices must be finite.')
 lp={(r.Market,int(r.Year),int(r.Quarter)):float(r.Price) for _,r in prices.iterrows()}
 annual={}
 if price_mode=='Annual average prices':
  for (market,y),g in prices.groupby(['Market','Year']):
   if set(g.Quarter.astype(int))=={1,2,3,4}:annual[(market,int(y))]=float(g.Price.mean())
 ls={int(r.Year):r for _,r in shipdf.iterrows()};rows=[]
 for _,r in schedule.iterrows():
  y=int(r.Year)
  for mi,m in enumerate(MONTHS,1):
   n=float(r[m]);q=(mi-1)//3+1
   if n<=0:continue
   try:
    if price_mode=='Annual average prices':pi=annual[(pur['market'],y)];si=annual[(sal['market'],y)]
    else:pi=lp[(pur['market'],y,q)];si=lp[(sal['market'],y,q)]
   except KeyError as e:raise ValueError(f'Missing EPA price {e.args[0]}. Annual averaging requires all four quarters for each selected market/year.')
   pi_real=pi;si_real=si
   if y not in epa_cpi:raise ValueError(f'Missing EPA inflation factor for {y}.')
   factor=float(epa_cpi[y])
   if not math.isfinite(factor) or factor<=0:raise ValueError(f'Invalid EPA inflation factor for {y}.')
   pi=pi_real*factor;si=si_real*factor
   loaded=n*cargo;deliv=loaded*(1-boil);lost=loaded-deliv;ratio=1-boil
   # Decompose the full contractual formulas without changing their economics.
   # EPA base uses delivered volume on both sides. The BoG rows contain only
   # the residual slope/volume exposure, so the full totals remain:
   # sales = nominal EPA x sales slope x delivered volume;
   # purchase = nominal EPA x purchase slope x loaded volume.
   rev=si*deliv*ss
   bc=si*(sal['slope']-1)*deliv*ss
   fixed_rev=sal['fixed']*deliv*ss;cpi_rev=sal['cpi']*sales_cpi[y]*deliv*ss
   purchase=pi*deliv*ps
   bog_cost=pi*(pur['slope']*loaded-deliv)*ps
   liq_fixed=pur['fixed']*deliv*ps;liq_cpi=pur['cpi']*purchase_cpi[y]*deliv*ps
   liq_bog_fixed=(pur['fixed']/ratio-pur['fixed'])*deliv*ps
   liq_bog_cpi=(pur['cpi']*purchase_cpi[y]/ratio-pur['cpi']*purchase_cpi[y])*deliv*ps
   liq_bog=liq_bog_fixed+liq_bog_cpi
   if shipmode=='Upload shipping Excel':
    if y not in ls:raise ValueError(f'Missing shipping year {y}.')
    raw_o=float(ls[y]['Shipping Opex USD/MMBtu']);raw_ca=float(ls[y]['Shipping Capex USD/MMBtu'])
   else:raw_o=flat_opex;raw_ca=flat_capex
   o=raw_o*purchase_cpi[y]*hs*deliv;ca=raw_ca*hs*deliv;eu=ext*es*deliv
   income=rev+bc+fixed_rev+cpi_rev+eu;cost=purchase+bog_cost+liq_fixed+liq_cpi+liq_bog+o+ca
   cash=income-cost;df=(1+disc)**(y-int(discount_year))
   row={'Year':y,'Month':m,'Quarter':q,'Price Mode':price_mode,'Cargoes':n,'Loaded MMBtu':loaded,'Boiled-off MMBtu':lost,'Delivered MMBtu':deliv,'EPA Base Volume MMBtu':deliv,'EPA Purchase Base Volume MMBtu':loaded,'DTL':ratio,'Physical Boil-off Fraction':boil,'BoG Index Coefficient Sales':sal['slope']-1,'BoG Index Coefficient Purchase':pur['slope']/ratio-1,'EPA Sales Price Real':si_real,'EPA Purchase Price Real':pi_real,'EPA Inflation Factor':factor,'EPA Sales Price':si,'EPA Purchase Price':pi,'Sales CPI Factor':sales_cpi[y],'Purchase CPI Factor':purchase_cpi[y],'Shipping CPI Factor':purchase_cpi[y],'Raw Shipping Opex USD/MMBtu':raw_o,'Raw Shipping Capex USD/MMBtu':raw_ca,'Discount Factor':df,'Sales Revenue USD':rev,'Boil-off Compensation Revenue USD':bc,'Fixed Revenue USD':fixed_rev,'CPI Revenue USD':cpi_rev,'Extrinsic Revenue USD':eu,'Purchase Base Cost USD':purchase,'Boil-off Cost USD':bog_cost,'Liquefaction Fixed Cost USD':liq_fixed,'Liquefaction CPI Cost USD':liq_cpi,'Liquefaction Boil-off Cost USD':liq_bog,'Liquefaction Boil-off Fixed Cost USD':liq_bog_fixed,'Liquefaction Boil-off CPI Cost USD':liq_bog_cpi,'Shipping Opex USD':o,'Shipping Capex USD':ca,'Total Income USD':income,'Total Cost USD':cost,'Cash Flow USD':cash}
   for key,value in list(row.items()):
    if key.endswith(' USD'):row['PV '+key]=value/df
   row['PV Delivered MMBtu']=deliv/df
   if not all(math.isfinite(float(v)) for v in row.values() if isinstance(v,(int,float))):raise ValueError('Non-finite calculation. Check inputs.')
   rows.append(row)
 if not rows:raise ValueError('No cargoes scheduled.')
 d=pd.DataFrame(rows)
 a=d.groupby('Year',as_index=False)[['Cargoes','Loaded MMBtu','Boiled-off MMBtu','Delivered MMBtu','Total Income USD','Total Cost USD','Cash Flow USD','PV Total Income USD','PV Total Cost USD','PV Cash Flow USD','PV Delivered MMBtu']].sum()
 return d,a

def totals(d,a):
 keys={'revenue':'Sales Revenue USD','comp':'Boil-off Compensation Revenue USD','fixed_rev':'Fixed Revenue USD','cpi_rev':'CPI Revenue USD','ext':'Extrinsic Revenue USD','purchase':'Purchase Base Cost USD','boil':'Boil-off Cost USD','liq_fixed':'Liquefaction Fixed Cost USD','liq_cpi':'Liquefaction CPI Cost USD','liq_bog':'Liquefaction Boil-off Cost USD','liq_bog_fixed':'Liquefaction Boil-off Fixed Cost USD','liq_bog_cpi':'Liquefaction Boil-off CPI Cost USD','opex':'Shipping Opex USD','capex':'Shipping Capex USD','income':'Total Income USD','cost':'Total Cost USD'}
 t={k:d['PV '+v].sum() for k,v in keys.items()};t.update(npv=a['PV Cash Flow USD'].sum(),volume=d['PV Delivered MMBtu'].sum());return t

def solve_zero(inputs,side,target,cpi_share):
 def evalx(x):
  z={k:(v.copy() if isinstance(v,dict) else v) for k,v in inputs.items()}
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
 return (lo+hi)/2
def formula_text(side,f):
 if side=='pur':
  return f"EPA {f['market']} x delivered volume + BoG residual [EPA x ({f['slope']:.6f} x loaded - delivered)] + fixed {f['fixed']:.6f} + CPI-linked {f['cpi']:.6f}"
 return f"EPA {f['market']} x delivered volume + BoG residual [EPA x ({f['slope']:.6f} - 1) x delivered] + fixed {f['fixed']:.6f} + CPI-linked {f['cpi']:.6f}"

def esc(s):
 import html
 return html.escape(str(s),quote=True)
def money(v,unit='$'):
 if unit=='$/MMBtu':return f'${v:,.3f}'
 if abs(v)>=1e9:return f'${v/1e9:,.2f}bn'
 if abs(v)>=1e6:return f'${v/1e6:,.1f}m'
 if abs(v)>=1e3:return f'${v/1e3:,.0f}k'
 return f'${v:,.0f}'
def section(kicker,title,note=None):
 st.markdown(f'<div class="section-head"><span>{esc(kicker)}</span><h3>{esc(title)}</h3>'+ (f'<p>{esc(note)}</p>' if note else '')+'</div>',unsafe_allow_html=True)
def formula_form(prefix,title,markets,default,comp=False):
 st.markdown(f'<div class="deal-label">{esc(title)}</div>',unsafe_allow_html=True)
 a,b=st.columns([1.4,1])
 m=a.selectbox('EPA market',markets,index=markets.index(default) if default in markets else 0,key=prefix+'m')
 sl=b.number_input('Slope / multiplier',value=1.15,step=.0001,format='%.6f',key=prefix+'s',help='Contract price slope. The waterfall decomposes each full formula into a common delivered-volume EPA base plus a residual BoG component. Sales total = EPA x slope x delivered; purchase total = EPA x slope x loaded.')
 a,b=st.columns([.55,1.6]);fs=a.selectbox('Sign',['+','-'],key=prefix+'fs');fx=b.number_input('Fixed, no CPI',value=0.,step=.0001,format='%.6f',key=prefix+'f')
 a,b=st.columns([.55,1.6]);cs=a.selectbox('CPI sign',['+','-'],key=prefix+'cs');cf=b.number_input('Fixed, CPI-linked',value=0.,step=.0001,format='%.6f',key=prefix+'c')
 f=dict(market=m,slope=sl,fixed=fx*(1 if fs=='+' else -1),cpi=cf*(1 if cs=='+' else -1))
 st.markdown(f'<div class="formula-preview">{esc(formula_text("sal" if comp else "pur",f))}</div>',unsafe_allow_html=True)
 return f

def chart(fig,title,unit='$',height=480):
 fig.update_layout(title=dict(text=title,x=0,font=dict(size=18,color=TEXT)),height=height,paper_bgcolor=PANEL,plot_bgcolor=PANEL,font=dict(family='Arial, sans-serif',color=TEXT,size=12),margin=dict(l=20,r=35,t=65,b=95),hoverlabel=dict(bgcolor=PANEL,font_color=TEXT),legend=dict(orientation='h',y=-.23,x=0,font=dict(color=TEXT)),bargap=.25,hovermode='x unified',uniformtext=dict(minsize=9,mode='hide'))
 fig.update_xaxes(showgrid=False,zeroline=False,automargin=True,tickfont_color=MUTED)
 fig.update_yaxes(gridcolor=GRID,griddash='dot',zerolinecolor=GRID,tickprefix='$',title_text='USD / MMBtu' if unit=='$/MMBtu' else 'USD',automargin=True)
 return fig
def table(df):
 formats={c:('{:,.0f}' if c in ['Year','Quarter'] else '{:,.3f}' if 'MMBtu' in c and 'Delivered' not in c else '{:,.2f}') for c in df.select_dtypes(include='number').columns}
 st.dataframe(df.style.format(formats),hide_index=True,use_container_width=True)
COMPONENTS=[('EPA sales','Sales Revenue USD','revenue',1),('BoG income','Boil-off Compensation Revenue USD','comp',1),('Fixed income','Fixed Revenue USD','fixed_rev',1),('CPI income','CPI Revenue USD','cpi_rev',1),('Extrinsic','Extrinsic Revenue USD','ext',1),('EPA purchase','Purchase Base Cost USD','purchase',-1),('BoG cost','Boil-off Cost USD','boil',-1),('Liq. no CPI','Liquefaction Fixed Cost USD','liq_fixed',-1),('Liq. CPI','Liquefaction CPI Cost USD','liq_cpi',-1),('Liq. BoG no CPI','Liquefaction Boil-off Fixed Cost USD','liq_bog_fixed',-1),('Liq. BoG CPI','Liquefaction Boil-off CPI Cost USD','liq_bog_cpi',-1),('Ship opex CPI','Shipping Opex USD','opex',-1),('Ship capex','Shipping Capex USD','capex',-1)]
def waterfall(t,unit):
 div=1 if unit=='$' else t['volume'];vals=[sgn*t[key]/div for name,c,key,sgn in COMPONENTS]+[t['npv']/div];labels=[name for name,c,key,sgn in COMPONENTS]+['NPV']
 f=go.Figure(go.Waterfall(x=labels,y=vals,measure=['relative']*len(COMPONENTS)+['total'],text=[money(v,unit) for v in vals],textposition='outside',textfont_color=TEXT,cliponaxis=False,connector=dict(line=dict(color=GRID,width=1)),increasing=dict(marker=dict(color=GREEN)),decreasing=dict(marker=dict(color=RED)),totals=dict(marker=dict(color=BLUE))))
 return chart(f,'From revenue to value',unit,570)
def monthly_chart(d,unit):
 periods=d.Year.astype(str)+' annual' if (d['Price Mode']=='Annual average prices').all() else d.Year.astype(str)+' '+d.Month;f=go.Figure()
 for j,(name,c,key,sgn) in enumerate(COMPONENTS):
  vals=d[c]*sgn if unit=='$' else d[c]*sgn/d['Delivered MMBtu']
  f.add_trace(go.Bar(x=periods,y=vals,name=name,marker_color=PALETTE[j%len(PALETTE)],text=[money(v,unit) if abs(v)>1e-10 else '' for v in vals],textposition='inside',insidetextfont_color='white',hovertemplate='%{x}<br>%{text}<extra>%{fullData.name}</extra>'))
 f.update_layout(barmode='relative');return chart(f,'Annual revenue & cost' if (d['Price Mode']=='Annual average prices').all() else 'Monthly revenue & cost',unit,600)
def cash_chart(d,a,base,unit):
 f=go.Figure()
 for j,(c,name) in enumerate([('Cash Flow USD','Annual cash flow'),('PV Cash Flow USD',f'PV to {base}')]):
  denom=a['PV Delivered MMBtu'] if j==1 else a['Delivered MMBtu'];vals=a[c] if unit=='$' else a[c]/denom
  f.add_trace(go.Scatter(x=a.Year,y=vals,name=name,mode='lines+markers+text',text=[money(v,unit) for v in vals],textfont_color=TEXT,textposition=['top center','bottom center'][j],line=dict(width=3,color=LINES[j],dash='dot' if j==1 else 'solid'),marker=dict(size=8),cliponaxis=False))
 return chart(f,'Annual cash flow',unit,480)

def sensitivity_rows(inp,changes):
 out=[]
 for name,p,ix in [('Sales price',changes[0],1),('Purchase price',changes[1],0),('Shipping cost',changes[2],2),('Extrinsic value',changes[3],3)]:
  lo=[1]*4;hi=[1]*4;lo[ix]=1-p/100;hi[ix]=1+p/100
  dl,al=calculate(**inp,scales=tuple(lo));dh,ah=calculate(**inp,scales=tuple(hi))
  out.append((name,p,totals(dl,al)['npv'],totals(dh,ah)['npv']))
 return pd.DataFrame(out,columns=['Driver','Percent','Low','High'])
def sensitivity_chart(q,t,unit):
 div=1 if unit=='$' else t['volume'];bn=t['npv'];bd=bn/div
 f=go.Figure(go.Bar(y=['Base case'],x=[bd],orientation='h',name='Base case NPV',marker_color=BLUE,text=[money(bd,unit)],textposition='inside',insidetextfont_color='white'))
 for c,name,colour in [('High','Higher-input scenario',GREEN),('Low','Lower-input scenario',RED)]:
  labels=[f"{int(r.Percent)}% {'higher' if c=='High' else 'lower'}: {money(r[c]/div,unit)}" for _,r in q.iterrows()]
  f.add_trace(go.Bar(y=q.Driver,x=(q[c]-bn)/div,base=bd,orientation='h',name=name,marker_color=colour,text=labels,textposition='inside',insidetextfont_color='white',hovertemplate='%{y}<br>%{text}<extra></extra>'))
 f.update_layout(barmode='overlay',yaxis=dict(categoryorder='array',categoryarray=['Base case']+q.Driver.tolist(),autorange='reversed'))
 f=chart(f,'What moves the NPV?',unit,560)
 f.update_yaxes(title_text=None,tickprefix='');f.update_xaxes(tickprefix='$',showgrid=True,gridcolor=GRID)
 return f

def ppt_bytes(d,a,inp):
 # Local export only. No network installation or external conversion service.
 from pptx import Presentation
 from pptx.util import Inches,Pt
 from pptx.dml.color import RGBColor
 import matplotlib
 matplotlib.use('Agg')
 import matplotlib.pyplot as plt
 t=totals(d,a);prs=Presentation();prs.slide_width=Inches(13.333);prs.slide_height=Inches(7.5)
 def slide(title):
  sl=prs.slides.add_slide(prs.slide_layouts[6]);sl.background.fill.solid();sl.background.fill.fore_color.rgb=RGBColor(13,25,36)
  box=sl.shapes.add_textbox(Inches(.65),Inches(.35),Inches(12),Inches(.65));p=box.text_frame.paragraphs[0];p.text=title;p.font.size=Pt(26);p.font.color.rgb=RGBColor(239,245,249)
  return sl
 sl=slide('LNG Business Case | Investment overview')
 box=sl.shapes.add_textbox(Inches(.8),Inches(2.0),Inches(11.8),Inches(3.6));tf=box.text_frame
 lines=[f"NPV to {inp['discount_year']}  {money(t['npv'])}",f"Sales fee CPI from {inp['sales_base']+1}; purchase CPI from {inp['purchase_base']}; EPA real year {inp['epa_real_base']}",f"NPV / MMBtu  {money(t['npv']/t['volume'],'$/MMBtu')}",f"PV volume  {t['volume']/1e6:,.1f} TBtu",f"Purchase: {formula_text('pur',inp['pur'])}",f"Sales: {formula_text('sal',inp['sal'])}"]
 for j,line in enumerate(lines):
  p=tf.paragraphs[0] if j==0 else tf.add_paragraph();p.text=line;p.font.size=Pt(25 if j<2 else 17);p.font.color.rgb=RGBColor(239,245,249)
 q=sensitivity_rows(inp,[10]*4)
 for title,kind in [('NPV waterfall','waterfall'),('Cash flow per year','cash'),('NPV sensitivities','sens')]:
  sl=slide(title)
  with plt.rc_context({'figure.facecolor':'#0d1924','axes.facecolor':'#0d1924','text.color':'#edf5fa','axes.labelcolor':'#b7cbd8','xtick.color':'#b7cbd8','ytick.color':'#b7cbd8','axes.edgecolor':'#365264','font.size':11}):
   fig,ax=plt.subplots(figsize=(12,5.2))
   if kind=='waterfall':
    vals=[sgn*t[key] for name,c,key,sgn in COMPONENTS];running=0
    for j,v in enumerate(vals):ax.bar([j],[v/1e6],bottom=[running/1e6],color='#198c7c' if v>=0 else '#c14763');running+=v
    ax.bar([len(COMPONENTS)],[t['npv']/1e6],color='#426fc2');ax.set_xticks(range(len(COMPONENTS)+1),[v[0] for v in COMPONENTS]+['NPV'],rotation=25,ha='right');ax.set_ylabel('USD million')
   elif kind=='cash':
    for c,n,colour in [('Cash Flow USD','Cash flow','#36c9b5'),('PV Cash Flow USD','PV cash flow','#7fbaff')]:ax.plot(a.Year,a[c]/1e6,marker='o',label=n,color=colour)
    ax.legend(facecolor='#13232f',labelcolor='#edf5fa');ax.set_ylabel('USD million')
   else:
    bn=t['npv'];ax.axvline(bn/1e6,color='#7fbaff',lw=1)
    for j,r in q.iterrows():
     ax.barh([j],[(r.High-bn)/1e6],left=[bn/1e6],color='#198c7c');ax.barh([j],[(r.Low-bn)/1e6],left=[bn/1e6],color='#c14763')
    ax.set_yticks(range(4),q.Driver);ax.invert_yaxis();ax.set_xlabel('NPV, USD million; ±10% input change')
   ax.spines[['top','right']].set_visible(False);fig.tight_layout();buf=io.BytesIO();fig.savefig(buf,format='png',dpi=170);plt.close(fig);buf.seek(0)
  sl.shapes.add_picture(buf,Inches(.65),Inches(1.25),width=Inches(12),height=Inches(5.6))
 out=io.BytesIO();prs.save(out);return out.getvalue()


def hedge_table(d,pv=False):
 p='PV ' if pv else ''
 rows=[]
 for y,g in d.groupby('Year'):
  def total(*cols):return sum(g[p+c].sum() for c in cols)
  index_income=total('Sales Revenue USD','Boil-off Compensation Revenue USD');index_cost=total('Purchase Base Cost USD','Boil-off Cost USD')
  fixed_income=total('Fixed Revenue USD');fixed_cost=total('Liquefaction Fixed Cost USD','Liquefaction Boil-off Fixed Cost USD','Shipping Capex USD')
  cpi_income=total('CPI Revenue USD');cpi_cost=total('Liquefaction CPI Cost USD','Liquefaction Boil-off CPI Cost USD','Shipping Opex USD')
  ext=total('Extrinsic Revenue USD')
  rows.append({'Year':int(y),'EPA + BoG income USD':index_income,'EPA + BoG cost USD':index_cost,'Index residual USD':index_income-index_cost,'Fixed income USD':fixed_income,'Fixed costs incl. fee BoG USD':fixed_cost,'Fixed residual USD':fixed_income-fixed_cost,'CPI income USD':cpi_income,'CPI costs incl. fee BoG USD':cpi_cost,'CPI residual USD':cpi_income-cpi_cost,'Extrinsic USD':ext,'Net residual USD':index_income-index_cost+fixed_income-fixed_cost+cpi_income-cpi_cost+ext})
 return pd.DataFrame(rows)

def match_sales_components(inputs):
 """Match three discounted component buckets, not a claim of annual hedge."""
 z=copy.deepcopy(inputs);d,a=calculate(**z)
 volume=d['PV Delivered MMBtu'].sum()
 index_basis=(d['EPA Sales Price']*d['PV Delivered MMBtu']).sum()
 index_target=(d['PV Purchase Base Cost USD']+d['PV Boil-off Cost USD']).sum()
 cpi_basis=(d['Sales CPI Factor']*d['PV Delivered MMBtu']).sum()
 if abs(cpi_basis)<1e-12 or volume<=0:raise ValueError('Cannot match components with zero sales CPI/volume basis.')
 if abs(index_basis)<1e-12 and abs(index_target)>.01:raise ValueError('Cannot match the index bucket because the sales EPA basis is zero.')
 sales=z['sal'].copy()
 sales['slope']=index_target/index_basis if abs(index_basis)>=1e-12 else sales['slope']
 sales['fixed']=(d['PV Liquefaction Fixed Cost USD']+d['PV Liquefaction Boil-off Fixed Cost USD']+d['PV Shipping Capex USD']).sum()/volume
 sales['cpi']=(d['PV Liquefaction CPI Cost USD']+d['PV Liquefaction Boil-off CPI Cost USD']+d['PV Shipping Opex USD']).sum()/cpi_basis
 z['sal']=sales;dx,ax=calculate(**z)
 return sales,dx,ax

def compact_price_formula(f):
 def signed(v):
  return (' + ' if v>=0 else ' - ')+f'{abs(v):.4f}'.rstrip('0').rstrip('.')
 market=str(f['market'])
 slope=f"{f['slope']*100:.4f}".rstrip('0').rstrip('.')+'%'
 return f"{slope} x {market}{signed(f['fixed'])}{signed(f['cpi'])} (CPI)"
def apply_purchase_preset():
 if st.session_state.get('purchase_preset')=='CCT2':
  st.session_state['purchase_base_year']=2022
  st.session_state['ps']=1.15;st.session_state['pf']=1.93;st.session_state['pc']=.5124
  st.session_state['pfs']='+';st.session_state['pcs']='+'
  st.session_state['pending_cct2_market']=True

def apply_matched_sales():
 h=st.session_state.get('studio_hedge')
 if h:
  f=h['formula'];st.session_state['sm']=f['market'];st.session_state['ss']=float(f['slope'])
  st.session_state['sf']=abs(float(f['fixed']));st.session_state['sc']=abs(float(f['cpi']))
  st.session_state['sfs']='+' if f['fixed']>=0 else '-'
  st.session_state['scs']='+' if f['cpi']>=0 else '-'
  st.session_state['sales_preset']='Custom'

# ------------------------- Theme and application shell -------------------
if st.session_state.get('studio_engine_version')!='v29':
 for old_key in ['ship','studio_result','studio_ppt','studio_be','studio_hedge']:
  st.session_state.pop(old_key,None)
 st.session_state.studio_engine_version='v29'
with st.container(key='masthead'):
 brand_col,theme_col=st.columns([5,1],vertical_alignment='center')
 with brand_col:brand_slot=st.empty()
 with theme_col:light=st.toggle('Light mode',value=False,key='studio_light_mode',help='Appearance only. Your deal inputs and results are retained.')
DARK=not light
BG,PANEL,FIELD,TEXT,MUTED,GRID=('#0b1119','#121d29','#192737','#eef3f8','#a2b5c7','#2a3b4d') if DARK else ('#f2f5f7','#ffffff','#edf2f5','#172d40','#52697b','#d9e2e9')
GREEN,RED,BLUE=('#137f70','#b83d58','#3e65b4') if DARK else ('#087365','#ae304b','#365ba5')
PALETTE=[GREEN,'#426c9f','#826119',RED,'#854e83','#5367a2','#4c7555','#8b573c']
LINES=['#46d0b3','#ff91a7','#85b6ff'] if DARK else ['#087365','#ae304b','#365ba5']
pio.templates['studio']=go.layout.Template(layout=dict(paper_bgcolor=PANEL,plot_bgcolor=PANEL,font=dict(color=TEXT),colorway=PALETTE,xaxis=dict(gridcolor=GRID,tickfont=dict(color=MUTED)),yaxis=dict(gridcolor=GRID,tickfont=dict(color=MUTED)),legend=dict(font=dict(color=TEXT)),hoverlabel=dict(bgcolor=PANEL,font=dict(color=TEXT))))
pio.templates.default=('plotly_dark' if DARK else 'plotly_white')+'+studio'
logo='https://upload.wikimedia.org/wikipedia/commons/c/c6/Equinor.svg'
for name,mime in [('equinor_logo.svg','image/svg+xml'),('equinor_logo.png','image/png')]:
 path=Path(__file__).resolve().parent/name
 if path.is_file():logo='data:'+mime+';base64,'+base64.b64encode(path.read_bytes()).decode();break
st.markdown(f'''<style>
:root{{--studio-bg:{BG};--studio-panel:{PANEL};--studio-field:{FIELD};--studio-text:{TEXT};--studio-muted:{MUTED};--studio-grid:{GRID};--text-color:{TEXT};--background-color:{BG};--secondary-background-color:{FIELD};color-scheme:{'dark' if DARK else 'light'};}}
html,body,.stApp,[data-testid="stAppViewContainer"],[data-testid="stMain"]{{background:{BG};color:{TEXT};font-family:Inter,"Segoe UI",Arial,sans-serif}}
[data-testid="stHeader"]{{background:{BG};height:3rem}}
.block-container{{max-width:1510px;padding:4.2rem 2.8rem 3rem!important}}
.st-key-masthead{{padding:0 0 20px!important;border-bottom:1px solid {GRID};margin-bottom:6px;overflow:visible!important}}
.brand{{display:flex;align-items:center;gap:22px;min-height:68px;min-width:0}}
.brand img{{width:80px;height:64px;object-fit:contain;flex-shrink:0;{'filter:brightness(0) invert(1)' if DARK else ''}}}
.brand-rule{{height:40px;width:1px;background:{GRID};flex-shrink:0}}
.brand h1{{font-size:clamp(1.4rem,2.5vw,2.3rem)!important;letter-spacing:-.05em;line-height:1.2!important;margin:0!important;padding:6px 0!important;color:{TEXT}!important;overflow-wrap:anywhere}}
.studio-badge{{display:inline-block;font-size:10px;letter-spacing:.16em;font-weight:700;color:{'#70d7c3' if DARK else '#087365'};border:1px solid {GRID};padding:5px 8px;border-radius:5px;margin-left:14px;vertical-align:middle}}
[data-testid="stMarkdownContainer"],[data-testid="stWidgetLabel"],[data-testid="stMetricLabel"],[data-testid="stMetricValue"],label{{color:{TEXT}}}
h1,h2,h3,h4{{color:{TEXT}!important;letter-spacing:-.025em}}
[data-testid="stCaptionContainer"],small{{color:{MUTED}}}
.section-head{{margin:4px 0 20px}}.section-head>span{{color:{'#70d7c3' if DARK else '#087365'};font-size:10px;font-weight:750;letter-spacing:.18em;text-transform:uppercase}}
.section-head h3{{font-size:24px;margin:7px 0 6px;padding:0;font-weight:650;letter-spacing:-.04em}}.section-head p{{color:{MUTED};font-size:13px;line-height:1.65;margin:0}}
.status-strip{{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:6px 0 16px}}.status-chip{{font-size:11px;color:{MUTED};padding:7px 11px;border:1px solid {GRID};border-radius:7px;background:{PANEL}}}.status-chip.ready{{color:{'#70d7c3' if DARK else '#087365'}}}.status-chip b{{color:{TEXT};font-weight:600}}
[data-testid="stTabs"] [role="tablist"]{{gap:6px;border-bottom:1px solid {GRID};padding-bottom:8px}}
[data-testid="stTabs"] [role="tab"]{{height:42px;padding:0 17px;border-radius:8px;color:{MUTED};font-weight:600;font-size:13px}}
[data-testid="stTabs"] [aria-selected="true"]{{background:{FIELD};color:{TEXT}!important}}
[data-testid="stTabs"] [data-baseweb="tab-highlight"]{{background:{'#70d7c3' if DARK else '#087365'};height:2px}}
[data-testid="stTabs"] [role="tabpanel"]{{padding-top:24px}}
[data-testid="stVerticalBlockBorderWrapper"]>div{{border-color:{GRID}!important;border-radius:14px!important}}
[data-testid="stExpander"]{{background:{PANEL};border:1px solid {GRID};border-radius:12px}}
[data-testid="stExpander"] summary{{color:{TEXT}}}
.deal-label{{font-size:11px;letter-spacing:.16em;text-transform:uppercase;font-weight:750;color:{MUTED};margin-bottom:20px}}
.formula-preview{{background:{FIELD};border-left:3px solid {'#70d7c3' if DARK else '#087365'};padding:14px 16px;border-radius:0 8px 8px 0;color:{TEXT};font-size:12px;line-height:1.8;margin-top:14px;min-height:65px}}
[data-baseweb="input"],[data-baseweb="base-input"],[data-baseweb="select"]>div,[data-baseweb="textarea"]{{background:{FIELD}!important;color:{TEXT}!important;border-color:{GRID}!important;border-radius:8px}}
.matched-formula-card{{background:linear-gradient(120deg,#123b35,#16483f);border:1px solid #2b8f79;border-radius:12px;padding:18px 20px;min-height:130px;display:flex;flex-direction:column;justify-content:center}}
.matched-formula-label{{font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:#8fdcc9;font-weight:700;margin-bottom:10px}}
.matched-formula-value{{font-size:clamp(1.35rem,2.2vw,2.15rem);line-height:1.25;font-weight:700;color:#9bf0d8;overflow-wrap:anywhere}}
.matched-formula-note{{font-size:11px;color:#acd3c8;margin-top:9px}}
/* Deployment-safe BaseWeb input styling: Radix/browser defaults must not paint white fields. */
[data-testid="stNumberInput"] [data-baseweb="input"],
[data-testid="stTextInput"] [data-baseweb="input"],
[data-testid="stDateInput"] [data-baseweb="input"],
[data-testid="stTimeInput"] [data-baseweb="input"],
[data-testid="stNumberInput"] [data-baseweb="base-input"],
[data-testid="stTextInput"] [data-baseweb="base-input"]{{background:{FIELD}!important;color:{TEXT}!important;border-color:{GRID}!important;color-scheme:{'dark' if DARK else 'light'}!important}}
[data-testid="stNumberInput"] input,
[data-testid="stTextInput"] input,
[data-testid="stDateInput"] input,
[data-testid="stTimeInput"] input,
input[type="number"],input[type="text"]{{background:{FIELD}!important;background-color:{FIELD}!important;color:{TEXT}!important;-webkit-text-fill-color:{TEXT}!important;caret-color:{TEXT}!important;border-color:{GRID}!important;box-shadow:inset 0 0 0 1000px {FIELD}!important;color-scheme:{'dark' if DARK else 'light'}!important}}
[data-testid="stNumberInput"] input:focus,
[data-testid="stTextInput"] input:focus{{background:{FIELD}!important;background-color:{FIELD}!important;color:{TEXT}!important;-webkit-text-fill-color:{TEXT}!important;box-shadow:inset 0 0 0 1000px {FIELD}!important}}
input:-webkit-autofill,input:-webkit-autofill:hover,input:-webkit-autofill:focus{{-webkit-text-fill-color:{TEXT}!important;box-shadow:0 0 0 1000px {FIELD} inset!important;-webkit-box-shadow:0 0 0 1000px {FIELD} inset!important;caret-color:{TEXT}!important}}
[data-testid="stNumberInput"] button{{background:{FIELD}!important;color:{TEXT}!important;border-color:{GRID}!important}}
[data-testid="stNumberInput"] button svg{{fill:{TEXT}!important;color:{TEXT}!important}}

input,textarea{{color:{TEXT}!important;caret-color:{TEXT}!important}}input::placeholder{{color:{MUTED}!important}}
[data-baseweb="select"] span,[data-baseweb="select"] svg{{color:{TEXT}!important}}
[data-baseweb="popover"] [role="listbox"],[data-baseweb="menu"],[role="option"]{{background:{PANEL}!important;color:{TEXT}!important}}
[role="option"]:hover{{background:{FIELD}!important}}[data-baseweb="tag"]{{background:{GRID}!important;color:{TEXT}!important}}
[data-testid="stFileUploaderDropzone"]{{background:{FIELD};border:1px dashed {GRID};border-radius:10px;color:{TEXT}}}
[data-testid="stFileUploaderDropzone"] small{{color:{MUTED}}}
[data-testid="stMetric"]{{padding:20px 18px;border:1px solid {GRID};background:{PANEL};border-radius:12px;min-height:130px}}
[data-testid="stMetricLabel"]{{font-size:11px;color:{MUTED}!important;text-transform:uppercase;letter-spacing:.04em}}
[data-testid="stMetricValue"]{{font-size:clamp(1.35rem,2.1vw,2.05rem);font-weight:650;letter-spacing:-.05em}}
.st-key-value_primary [data-testid="stMetric"]{{background:{'linear-gradient(120deg,#163c3b,#142d35)' if DARK else 'linear-gradient(120deg,#e1f2ed,#f1f8f6)'};border-color:{'#2d6c61' if DARK else '#a6d4c7'}}}
.st-key-value_primary [data-testid="stMetricValue"]{{color:{'#8fe3ce' if DARK else '#075d51'}!important}}
button[kind="primary"]{{background:{'#c83557' if DARK else '#b42b4b'};border:1px solid transparent;border-radius:9px;min-height:45px;font-weight:650;color:white}}
button[kind="primary"] p,button[kind="primary"] [data-testid="stMarkdownContainer"]{{color:white!important}}
button[kind="secondary"],[data-testid="stNumberInput"] button{{background:{FIELD};color:{TEXT};border-color:{GRID};border-radius:8px}}
button:disabled{{opacity:.45}}
[data-testid="stDownloadButton"] button{{background:{FIELD};color:{TEXT};border-color:{GRID};border-radius:8px}}
[data-testid="stAlert"]{{background:{PANEL}!important;border:1px solid {GRID};border-radius:10px}}
[data-testid="stAlert"] [data-testid="stMarkdownContainer"]{{color:{TEXT}!important}}
[data-testid="stDataFrame"],[data-testid="stDataEditor"]{{border:1px solid {GRID};border-radius:10px;overflow:hidden;--gdg-bg-cell:{PANEL};--gdg-bg-header:{FIELD};--gdg-text-dark:{TEXT};--gdg-text-medium:{MUTED};--gdg-text-header:{TEXT};--gdg-border-color:{GRID}}}
[data-testid="stPlotlyChart"]{{border:1px solid {GRID};border-radius:14px;overflow:hidden;background:{PANEL};padding:8px}}
.empty-state{{padding:45px 32px;border:1px dashed {GRID};border-radius:14px;background:{PANEL};text-align:center}}
.empty-state .empty-icon{{font-size:30px;color:{'#70d7c3' if DARK else '#087365'};margin-bottom:14px}}.empty-state h2{{font-size:26px;margin:0 0 12px}}.empty-state p{{color:{MUTED};max-width:550px;margin:0 auto;line-height:1.8;font-size:14px}}
.empty-steps{{display:flex;gap:12px;justify-content:center;flex-wrap:wrap;margin-top:26px}}.empty-steps span{{border:1px solid {GRID};padding:10px 14px;border-radius:8px;font-size:12px;color:{TEXT}}}
.footer{{border-top:1px solid {GRID};padding-top:18px;margin-top:32px;display:flex;justify-content:space-between;gap:16px;color:{MUTED};font-size:11px;line-height:1.6}}
@media(max-width:750px){{.block-container{{padding:3.8rem 1rem 2rem!important}}.brand{{gap:12px}}.brand img{{width:55px;height:50px}}.brand h1{{font-size:1.5rem!important}}.studio-badge{{display:none}}[data-testid="stTabs"] [role="tab"]{{padding:0 10px}}}}
@media(prefers-reduced-motion:reduce){{*{{scroll-behavior:auto!important}}}}
</style>''',unsafe_allow_html=True)
brand_slot.markdown(f'<div class="brand"><img src="{logo}" alt="Equinor"><div class="brand-rule"></div><h1>LNG Business Case<span class="studio-badge">STUDIO</span></h1></div>',unsafe_allow_html=True)

# Decorative carrier and its existing click-to-explode easter egg.
# The animation runs only in the browser; no financial state is read or changed.
def vessel_scene():
 svg='''<svg viewBox="0 0 520 200" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="LNG carrier"><defs><linearGradient id="h" x2="0" y2="1"><stop stop-color="#3a5c70"/><stop offset=".5" stop-color="#163246"/><stop offset="1" stop-color="#081e30"/></linearGradient><radialGradient id="t" cx=".3" cy=".15" r=".95"><stop stop-color="#fff"/><stop offset=".4" stop-color="#e4edf0"/><stop offset="1" stop-color="#819dad"/></radialGradient></defs><path d="M30 142L498 134L473 169Q465 178 441 178H64Q42 170 30 142Z" fill="url(#h)"/><path d="M44 166H476L463 178H64Z" fill="#a63448"/><path d="M30 141L498 133L493 141L32 148Z" fill="#d5e2e9"/><path d="M44 137V126H476L485 136" fill="none" stroke="#c3d9e5" stroke-width="2"/>'''
 for x in [174,248,322,396]:
  svg+=f'<g stroke="#8ca8b7"><path d="M{x-34} 127V111A34 34 0 0 1 {x+34} 111V127Z" fill="url(#t)"/><path d="M{x-34} 120H{x+34}V139H{x-34}Z" fill="#9db5c3"/><path d="M{x-32} 121H{x+32}" stroke="#f0f6f8"/><path d="M{x-30} 105Q{x} 117 {x+30} 105M{x-21} 85Q{x} 94 {x+21} 85" fill="none" stroke="#b4c9d3"/><path d="M{x} 77V119" stroke="#d0dfe5"/><rect x="{x-8}" y="73" width="16" height="5" rx="2" fill="#d7e5ec"/><path d="M{x} 73V66H{x+10}" stroke="#d5e6ee" fill="none"/></g>'
 svg+='''<path d="M139 136H439M139 132H439" stroke="#d4b46b" stroke-width="2"/><path d="M62 140V76H115V140Z" fill="#dce7ec"/><path d="M57 77V62H121V81Z" fill="#f2f7f9"/><path d="M60 65H117V73H60Z" fill="#284d66"/><path d="M68 62V48H103V62Z" fill="#c8dce5"/><path d="M86 48V22M74 32H101M78 40H96" stroke="#e0edf2" stroke-width="2"/><g fill="#375e73"><path d="M70 91H78V97H70ZM86 91H94V97H86ZM103 91H110V97H103ZM70 106H78V112H70ZM86 106H94V112H86ZM103 106H110V112H103ZM70 121H78V127H70ZM86 121H94V127H86ZM103 121H110V127H103Z"/></g><path d="M120 129V91H137V129Z" fill="#c6d7e1"/><path d="M120 91H137V101H120Z" fill="#c13854"/><path d="M120 90H137" stroke="#243d4e" stroke-width="4"/><path d="M451 136V108M444 116H458M451 108V95" stroke="#dcebf2" stroke-width="2"/><path d="M468 136L482 128L488 135Z" fill="#c0d5e0"/><path d="M36 173Q20 168 9 176M476 167Q495 162 511 171" fill="none" stroke="#cceff7" stroke-width="2"/><g fill="#c9dfe8"><circle cx="473" cy="152" r="2"/><circle cx="461" cy="153" r="2"/></g></svg>'''
 html=r'''<!doctype html><html><head><style>
*{box-sizing:border-box}body{margin:0;overflow:hidden}.sea{height:154px;border-radius:12px;position:relative;overflow:hidden;background:linear-gradient(__SKY__)}
.horizon{position:absolute;top:68px;width:100%;height:1px;background:#c4e9ef55}.wave{position:absolute;left:-30%;width:160%;height:20px;border-top:1px solid #c1e7ee44;border-radius:45%;animation:water 10s linear infinite}.w1{top:102px}.w2{top:128px;animation-direction:reverse}.track{position:absolute;top:5px;left:0;right:0;height:142px}.travel{position:absolute;width:clamp(230px,25vw,340px);height:142px;animation:sail 32s linear infinite}.face{height:100%;animation:turn 32s steps(1,end) infinite}.boat{border:0;background:none;width:100%;height:100%;padding:0;cursor:pointer;animation:bob 4s ease-in-out infinite;filter:drop-shadow(0 7px 3px #001a2e66)}.boat svg{width:100%;height:100%}.boat:focus-visible{outline:2px solid #70d7c3;border-radius:12px}.wake{position:absolute;left:-16%;bottom:15px;width:35%;height:9px;border-top:2px solid #ccefff99;border-radius:50%}.sea-note{position:absolute;right:18px;top:14px;font:10px Arial,sans-serif;letter-spacing:2px;color:__NOTE__;pointer-events:none}
@keyframes sail{0%,100%{left:0;transform:translateX(0)}48%,50%{left:100%;transform:translateX(-100%)}98%{left:0;transform:translateX(0)}}@keyframes turn{0%{transform:scaleX(1)}49%{transform:scaleX(-1)}99%{transform:scaleX(1)}}@keyframes bob{0%,100%{transform:translateY(0) rotate(-.3deg)}50%{transform:translateY(2px) rotate(.3deg)}}@keyframes water{to{transform:translateX(12%)}}@media(prefers-reduced-motion:reduce){.travel,.face,.boat,.wave{animation:none}.travel{left:30%}}
</style></head><body><div class="sea"><div class="horizon"></div><div class="wave w1"></div><div class="wave w2"></div><div class="track"><div class="travel"><div class="face"><div class="wake"></div><button class="boat" title="Click for a surprise" aria-label="LNG carrier: click for a surprise">__SVG__</button></div></div></div><div class="sea-note">LNG / ORIGINATION</div></div><script>
// Canvas tables use their own renderer. Adapt only the displayed colours.
const dark=__DARK__;let tableTimer;
try{const pd=parent.document;const adapt=()=>{pd.querySelectorAll('[data-testid="stDataFrame"] canvas,[data-testid="stDataEditor"] canvas').forEach(c=>{try{const ctx=c.getContext('2d');if(!ctx||c.width<30||c.height<12)return;const p=ctx.getImageData(24,8,1,1).data;if(p[3]<200)return;const nativeDark=(p[0]*.2126+p[1]*.7152+p[2]*.0722)<128;c.style.filter=nativeDark!==dark?'invert(1) hue-rotate(180deg)':'none';}catch(e){}})};adapt();tableTimer=setInterval(adapt,1000);addEventListener('pagehide',()=>clearInterval(tableTimer),{once:true});}catch(e){}
const boat=document.querySelector('.boat'),travel=document.querySelector('.travel');let busy=false;
boat.addEventListener('click',()=>{
 if(busy||matchMedia('(prefers-reduced-motion: reduce)').matches)return;busy=true;
 let host=window,doc=document,ox=0,oy=0;
 try{if(parent.document.body){host=parent;doc=host.document;const r=frameElement.getBoundingClientRect();ox=r.left;oy=r.top;}}catch(e){}
 const r=boat.getBoundingClientRect(),sx=r.left+ox,sy=r.top+oy;
 const layer=doc.createElement('div');layer.dataset.lngEffect='true';layer.setAttribute('aria-hidden','true');layer.style.cssText='position:fixed;inset:0;z-index:2147483000;pointer-events:none;overflow:hidden';doc.body.appendChild(layer);
 let done=false;const cleanup=()=>{if(done)return;done=true;host.clearTimeout(safety);layer.remove();travel.style.visibility='visible';busy=false;};const safety=host.setTimeout(cleanup,11000);
 const flyer=doc.createElement('div');flyer.innerHTML=boat.innerHTML;flyer.style.cssText=`position:absolute;left:0;top:0;width:${r.width}px;height:${r.height}px;filter:drop-shadow(0 12px 9px #0006)`;flyer.querySelector('svg').style.cssText='width:100%;height:100%';layer.appendChild(flyer);travel.style.visibility='hidden';
 const ex=Math.max(0,Math.min(host.innerWidth-r.width,host.innerWidth*.6)),ey=Math.max(0,host.innerHeight-r.height-80);
 const a=flyer.animate([{transform:`translate(${sx}px,${sy}px) rotate(0deg)`},{transform:`translate(${Math.min(host.innerWidth-r.width,sx+220)}px,${sy+30}px) rotate(15deg)`,offset:.3},{transform:`translate(${ex}px,${ey}px) rotate(64deg)`}],{duration:3200,easing:'cubic-bezier(.35,.05,.6,1)',fill:'forwards'});
 a.onfinish=()=>{if(done)return;flyer.style.opacity=0;const cx=ex+r.width/2,cy=ey+r.height/2;
 const particle=(css,frames,duration,delay=0)=>{const e=doc.createElement('div');e.style.cssText=`position:absolute;left:${cx}px;top:${cy}px;pointer-events:none;`+css;layer.appendChild(e);e.animate(frames,{duration,delay,fill:'both',easing:'ease-out'});};
 particle('width:620px;height:620px;margin:-310px;border-radius:50%;background:radial-gradient(circle,#fffde8,#fff2ae 13%,#ffb74b99 35%,#ff572233 55%,transparent 72%);filter:blur(6px)',[{transform:'scale(.08)',opacity:0},{transform:'scale(1)',opacity:.85,offset:.18},{transform:'scale(1.65)',opacity:0}],1200);
 particle('width:160px;height:160px;margin:-80px;border:5px solid #ffe3a5aa;border-radius:50%;box-shadow:0 0 28px #ff963e66',[{transform:'scale(.2)',opacity:.8},{transform:'scale(5)',opacity:0}],1700);
 for(let i=0;i<24;i++){const size=70+Math.random()*100,dx=(Math.random()-.5)*360,dy=-100-Math.random()*280;particle(`width:${size}px;height:${size}px;margin:${-size/2}px;border-radius:46% 54% 60% 40%;background:radial-gradient(circle at 35% 30%,#887d72cc,#49484de6 48%,#20283299 72%,transparent 76%);filter:blur(7px)`,[{transform:'scale(.1)',opacity:0},{transform:`translate(${dx*.4}px,${dy*.4}px) scale(.9)`,opacity:.85,offset:.35},{transform:`translate(${dx}px,${dy}px) scale(1.8)`,opacity:0}],3400+Math.random()*700,150+Math.random()*450);}
 const palettes=['#fffde1,#ffe584 18%,#ffb328 40%,#ff6819 62%,#b9281299 78%,transparent 84%','#fff9cf,#ffd24d 20%,#ff8b1c 43%,#ee4314 65%,#7d201d88 80%,transparent 85%','#fff6b3,#ffc85b 22%,#ff9a38 42%,#db341c 66%,#6c223377 80%,transparent 85%'];
 for(let i=0;i<58;i++){const size=35+Math.random()*90,angle=Math.random()*Math.PI*2,dist=50+Math.random()*220,dx=Math.cos(angle)*dist,dy=Math.sin(angle)*dist*.65-65;particle(`width:${size}px;height:${size*1.2}px;margin:${-size/2}px;border-radius:${35+Math.random()*30}% 55% 40% 60%;background:radial-gradient(ellipse at 45% 55%,${palettes[i%3]});filter:blur(${1+Math.random()*3}px)`,[{transform:'scale(.05)',opacity:0},{transform:`translate(${dx*.45}px,${dy*.4}px) scale(.9) rotate(-15deg)`,opacity:1,offset:.22},{transform:`translate(${dx}px,${dy-70}px) scale(1.3) rotate(20deg)`,opacity:.75,offset:.55},{transform:`translate(${dx*1.25}px,${dy-160}px) scale(.6) rotate(40deg)`,opacity:0}],1900+Math.random()*1000,Math.random()*500);}
 for(let i=0;i<8;i++)particle('width:55px;height:35px;margin:-20px;border-radius:50%;background:radial-gradient(ellipse,#d9f6ff,#69bfff99 35%,#6b56c455 60%,transparent 75%);filter:blur(4px)',[{transform:'scale(.1)',opacity:0},{transform:`translate(${(i-4)*18}px,15px)`,opacity:.7,offset:.3},{transform:'translate(0,-20px) scale(1.4)',opacity:0}],900,i*30);
 for(let i=0;i<64;i++){const angle=Math.random()*Math.PI*2,speed=100+Math.random()*340,dx=Math.cos(angle)*speed,dy=Math.sin(angle)*speed-100;particle(`width:${2+Math.random()*3}px;height:${5+Math.random()*11}px;border-radius:50%;background:#fff1a8;box-shadow:0 0 7px #ffb21c,0 0 14px #ff5b17`,[{transform:'scale(.3)',opacity:0},{transform:`translate(${dx*.45}px,${dy*.55}px) rotate(${angle*180/Math.PI}deg)`,opacity:1,offset:.2},{transform:`translate(${dx}px,${dy+150}px) scale(.1)`,background:'#d63717',opacity:0}],1400+Math.random()*1500,Math.random()*200);}
 host.setTimeout(cleanup,4500);
 };
 addEventListener('pagehide',cleanup,{once:true});
});
</script></body></html>'''
 sky='#15283b 0%,#263f50 43%,#204b5c 44%,#102e42 100%' if DARK else '#d2e5ed 0%,#e7eff2 43%,#8cb5c0 44%,#427d94 100%'
 html=html.replace('__SVG__',svg).replace('__SKY__',sky).replace('__DARK__','true' if DARK else 'false').replace('__NOTE__','#9fb9c9' if DARK else '#3b6378')
 components.html(html,height=164,scrolling=False)

vessel_scene()


# ------------------------------ Interface --------------------------------
status_slot=st.empty()
summary_slot=st.container()
overview_tab,setup_tab,prices_tab,shipping_tab,sens_tab,explanation_tab=st.tabs(['Overview','Deal setup','Prices','Shipping','Sensitivities','Explanation'])

# Render the setup first so all other views use the same current input values.
with setup_tab:
 section('Build the case','Deal configuration','The same pricing, cargo, shipping and financial inputs, organised into four work areas.')
 data_step,formula_step,volume_step,finance_step=st.tabs(['01  Data & horizon','02  Pricing','03  Cargo programme','04  Shipping & finance'])
 with data_step:
  left,right=st.columns([1.15,1],gap='large')
  with left:
   with st.container(border=True):
    section('Market inputs','EPA price files')
    fg=st.file_uploader('EPA Gas Prices',type=['xlsx'],key='gas_upload',help='Gas Basecase: years in row 8, quarters in row 9, market prices in rows 10–15.')
    if fg:
     try:st.session_state.gas=gas(fg.getvalue());st.session_state.gas_name=fg.name
     except Exception as e:st.error(f'Gas workbook: {e}')
    if 'gas' in st.session_state:st.caption(f"Loaded: {st.session_state.get('gas_name','Gas prices')} · {len(st.session_state.gas):,} price points")
    fb=st.file_uploader('EPA Brent Prices',type=['xlsx'],key='brent_upload',help='Brent sheet: years in row 7 and annual prices in row 8.')
    if fb:
     try:st.session_state.brent=brent(fb.getvalue());st.session_state.brent_name=fb.name
     except Exception as e:st.error(f'Brent workbook: {e}')
    if 'brent' in st.session_state:st.caption(f"Loaded: {st.session_state.get('brent_name','Brent prices')} · {len(st.session_state.brent):,} price points")
  with right:
   with st.container(border=True):
    section('Time horizon','Period & inflation')
    a,b=st.columns(2)
    start=int(a.number_input('First cargo year',2020,2100,2028,key='start_year'))
    end=int(b.number_input('Last cargo year',2020,2100,2037,key='end_year'))
    a,b=st.columns(2)
    sales_base=int(a.number_input('Sales base year',2000,2100,2026,key='sales_base_year',help='First included CPI rate is sales base year + 1.'))
    purchase_base=int(b.number_input('Purchase base year',2000,2100,2026,key='purchase_base_year',help='First included CPI rate is the purchase base year itself. Also used for shipping opex.'))
    discount_year=int(st.number_input('Discount back to',2000,2100,2026,step=1,key='discount_year',help='All cash flows and PV volumes are discounted to this year. Does not change CPI.'))
    epa_real_base=int(st.number_input('EPA prices: real reference year',2000,2100,2026,step=1,key='epa_real_base',help='Uploaded EPA prices are real in this year. Both EPA legs and BoG decomposition use prices inflated to the cash-flow year. Independent of fee base years and Discount back to.'))
    years=list(range(start,end+1))
    if end<start:st.error('Last cargo year must be on or after first cargo year.')
    mode=st.radio('Annual CPI',['Fixed annual CPI','Preset CPI curve','Specific CPI by year'],key='cpi_mode')
    cpi_years=list(range(min(years+[sales_base,purchase_base,epa_real_base]),max(years+[sales_base+1,purchase_base,epa_real_base])+1))
    if mode=='Fixed annual CPI':
     p=st.number_input('Annual CPI (%)',value=2.,step=.1,key='cpi_fixed')/100;rates={y:p for y in cpi_years}
    elif mode=='Preset CPI curve':
     p=st.selectbox('Preset',['2.0% long-term','2.7%, 2.4%, then 2.0%','No inflation'],key='cpi_preset')
     rates={y:(0 if p=='No inflation' else .02) for y in cpi_years}
     if p.startswith('2.7'):rates[cpi_years[0]]=.027;rates[cpi_years[0]+1]=.024
    else:
     cd=st.data_editor(pd.DataFrame({'Year':cpi_years,'CPI %':2.}),hide_index=True,key=f'cpi_curve_{start}_{end}_{sales_base}_{purchase_base}_{epa_real_base}',use_container_width=True)
     rates={int(r.Year):float(r['CPI %'])/100 for _,r in cd.iterrows()}
    try:sales_cpi=cpif(years,sales_base,rates);purchase_cpi=cpif(years,purchase_base-1,rates);epa_cpi=cpif(years,epa_real_base,rates);cpi_error=None
    except Exception as e:sales_cpi={};purchase_cpi={};epa_cpi={};cpi_error=str(e);st.error(f'CPI curve: {e}')
    st.caption(f'Sales: first CPI rate {sales_base+1}. Purchase and shipping opex: first CPI rate {purchase_base}. EPA prices: real {epa_real_base} converted to nominal each year. Discount back to: {discount_year}, independent of CPI.')
 frames=[st.session_state[x] for x in ['gas','brent'] if x in st.session_state]
 prices=pd.concat(frames,ignore_index=True) if frames else None
 markets=sorted(prices.Market.unique()) if prices is not None else ['Upload prices']
 with formula_step:
  a,b=st.columns(2)
  a.selectbox('Purchase deal preset',['Custom','CCT2'],key='purchase_preset',on_change=apply_purchase_preset)
  b.selectbox('Sales deal preset',['Custom'],key='sales_preset')
  hh=next((m for m in markets if 'henry' in m.lower() or re.search(r'\bhh\b',m,re.I)),None)
  if st.session_state.get('pending_cct2_market') and hh:
   st.session_state['pm']=hh;st.session_state.pop('pending_cct2_market',None)
  if st.session_state.get('purchase_preset')=='CCT2':
   st.caption('CCT2 preset inputs: HH index; BoG slope 1.15; 1.93 fixed; 0.5124 CPI-linked; purchase base year 2022. Inputs remain editable after loading the preset.')
   if hh is None:st.warning('CCT2 requires an HH / Henry Hub EPA series. Upload that series before calculating.')
  a,b=st.columns(2,gap='large')
  with a:
   with st.container(border=True):pur=formula_form('p','Purchase / cost leg',markets,'HH price')
  with b:
   with st.container(border=True):sal=formula_form('s','Sales / revenue leg',markets,'TTF price',True)
  st.caption('EPA sales = EPA price x delivered volume. EPA purchase = EPA price x loaded/bought volume. Slopes apply only to the separate physical BoG components. Purchase fixed fees are liquefaction costs; sales fixed fees are income.')
 with volume_step:
  section('Physical programme','Cargo programme','Choose an annual-average programme or retain the monthly schedule.')
  price_mode=st.radio('Price calculation',['Scheduled quarterly prices','Annual average prices'],key='price_mode',horizontal=True)
  if price_mode=='Annual average prices':
   a,b=st.columns(2)
   n=int(a.number_input('Default cargoes per year',0,200,1,key='annual_cargo_count'))
   cargo=float(b.number_input('Loaded cargo MMBtu',0.,10_000_000.,3_800_000.,50_000.,key='annual_cargo_size'))
   annual_seed=pd.DataFrame({'Year':years,'Cargoes':n})
   annual_program=st.data_editor(annual_seed,hide_index=True,use_container_width=True,disabled=['Year'],key=f'annual_program_{start}_{end}_{n}')
   schedule=pd.DataFrame({'Year':annual_program.Year,**{m:0. for m in MONTHS}})
   schedule['Jan']=annual_program.Cargoes.astype(float)
   st.caption('Each year uses the arithmetic mean of Q1-Q4 for the selected index, irrespective of cargo month. All four quarters are required. Edit the cargo count for each year above. January is an internal annual bucket only.')
  else:
   a,b,c=st.columns([1,1.3,1.2])
   n=int(a.number_input('Cargoes / year',0,200,12,key='cargo_count'))
   mode2=b.selectbox('Distribution',['Quarterly','Monthly','Selected months','Manual by year and month'],key='cargo_distribution')
   cargo=float(c.number_input('Loaded cargo MMBtu',0.,10_000_000.,3_800_000.,50_000.,key='cargo_size'))
   months=st.multiselect('Eligible months',MONTHS,default=['Mar','Jun','Sep','Dec'],key='cargo_months') if mode2=='Selected months' else []
   sd=seed(years,n,mode2,months) if years else pd.DataFrame(columns=['Year']+MONTHS)
   schedule=st.data_editor(sd,hide_index=True,use_container_width=True,disabled=['Year'],key=f'cargo_schedule_{start}_{end}_{n}_{mode2}_{"_".join(months)}')
  total_cargoes=float(schedule[MONTHS].sum().sum()) if len(schedule) else 0
  st.caption(f'{total_cargoes:,.1f} cargoes scheduled | Loaded volume {total_cargoes*cargo/1e6:,.1f} TBtu')
 with finance_step:
  left,right=st.columns([1.25,1],gap='large')
  with left:
   with st.container(border=True):
    section('Logistics','Shipping & boil-off')
    shipmode=st.radio('Shipping source',['Flat fixed assumption','Upload shipping Excel'],horizontal=True,key='shipping_mode')
    fo=fc=boil=0.
    routes=pd.DataFrame(columns=['Route / fee','Included in shipping workbook','Source'])
    shipdf=pd.DataFrame(columns=['Year','Shipping Opex USD/MMBtu','Shipping Capex USD/MMBtu'])
    if shipmode=='Flat fixed assumption':
     a,b=st.columns(2)
     fo=float(a.number_input('Shipping Opex USD/MMBtu',0.,value=.16,step=.01,key='flat_opex'))
     fc=float(b.number_input('Shipping Capex USD/MMBtu',0.,value=.57,step=.01,key='flat_capex'))
     boil=float(st.number_input('Physical boil-off %',0.,99.,1.5,.1,key='physical_boil'))/100
     shipdf=pd.DataFrame({'Year':years,'Shipping Opex USD/MMBtu':fo,'Shipping Capex USD/MMBtu':fc})
    else:
     fs=st.file_uploader('Shipping workbook',type=['xlsx','xlsm'],key='shipping_upload')
     if fs:
      try:st.session_state.ship=shipping(fs.getvalue());st.session_state.ship_name=fs.name
      except Exception as e:st.error(f'Shipping workbook: {e}')
     if 'ship' in st.session_state:
      shipdf,boil,routes=st.session_state.ship;st.caption(f"Loaded: {st.session_state.get('ship_name','Shipping workbook')}")
     st.caption(f'Interface: capex row 97; opex rows 92-96 + 98; years through 2050 only. First shipping CPI rate: {purchase_base}. Boil-off: M45.')
     if not routes.empty:st.dataframe(routes,hide_index=True,use_container_width=True)
     st.caption('Route flags are informational only. Uploaded opex is treated as a pre-CPI source value.')
    st.metric('Sales cargo size',f'{cargo*(1-boil):,.0f} MMBtu',f'{boil:.3%} physical boil-off',delta_color='off')
  with right:
   with st.container(border=True):
    section('Valuation','Financial assumptions')
    disc=float(st.number_input('Discount rate %',0.,100.,8.,.25,key='discount_rate'))/100
    ext=float(st.number_input('Extrinsic USD/MMBtu',0.,20.,0.,.05,key='extrinsic'))
    st.caption('Annual income minus annual cost, discounted to the Discount back to year. No tax is applied.')
 inp=dict(schedule=schedule,prices=prices,pur=pur,sal=sal,sales_cpi=sales_cpi,purchase_cpi=purchase_cpi,sales_base=sales_base,purchase_base=purchase_base,discount_year=discount_year,disc=disc,ext=ext,cargo=cargo,boil=boil,shipdf=shipdf,flat_opex=fo,flat_capex=fc,shipmode=shipmode,price_mode=price_mode,epa_cpi=epa_cpi,epa_real_base=epa_real_base)
 ready=prices is not None and bool(years) and cpi_error is None and not (st.session_state.get('purchase_preset')=='CCT2' and hh is None)
 a,b=st.columns([1,2],vertical_alignment='center')
 with a:calculate_clicked=st.button('Calculate NPV',type='primary',use_container_width=True,disabled=not ready,key='calculate_setup')
 with b:st.caption('Results, sensitivities and exports use a consistent snapshot of the calculated case. Recalculate after changing assumptions.')

# A stable signature makes it explicit when displayed results need refreshing.
def case_signature(values):
 import hashlib,json
 def normal(v):
  if isinstance(v,pd.DataFrame):return v.to_json(orient='split',double_precision=15)
  if isinstance(v,dict):return {str(k):normal(x) for k,x in v.items()}
  return v
 return hashlib.sha256(json.dumps(normal(values),sort_keys=True,default=str).encode()).hexdigest()
current_sig=case_signature({'valuation':inp,'routes':routes})

def execute_case():
 try:
  d,a=calculate(**inp)
  st.session_state.studio_result={'detail':d,'annual':a,'inputs':copy.deepcopy(inp),'signature':current_sig,'routes':routes.copy()}
  st.session_state.pop('studio_ppt',None);st.session_state.pop('studio_be',None);st.session_state.pop('studio_hedge',None)
  st.rerun()
 except Exception as e:st.error(f'Could not calculate: {e}')
if calculate_clicked:execute_case()
result=st.session_state.get('studio_result')
stale=bool(result and result['signature']!=current_sig)
chips=[]
for key,label in [('gas','Gas EPA'),('brent','Brent EPA')]:
 loaded=key in st.session_state;chips.append(f'<span class="status-chip {"ready" if loaded else ""}">{label} · {"loaded" if loaded else "not loaded"}</span>')
chips.append(f'<span class="status-chip"><b>{start}–{end}</b> cargo horizon</span>')
chips.append(f'<span class="status-chip"><b>{disc:.1%}</b> discount rate</span>')
chips.append(f'<span class="status-chip {"ready" if result and not stale else ""}">{"Recalculation needed" if stale else "Calculated case" if result else "Case not calculated"}</span>')
status_slot.markdown('<div class="status-strip">'+''.join(chips)+'</div>',unsafe_allow_html=True)
if result:
 d,a=result['detail'],result['annual'];t=totals(d,a);ci=result['inputs']
 with summary_slot:
  cols=st.columns([1.2,1,1,1,1])
  with cols[0]:
   with st.container(key='value_primary'):st.metric('NPV',money(t['npv']))
  cols[1].metric('NPV / MMBtu',money(t['npv']/t['volume'],'$/MMBtu'))
  cols[2].metric('PV total income',money(t['income']))
  cols[3].metric('PV total cost',money(t['cost']))
  cols[4].metric('PV volume',f"{t['volume']/1e6:,.1f} TBtu")
  if stale:st.warning('Inputs have changed. The figures shown are the last calculated case. Select Calculate NPV or Recalculate NPV to update them.')

with overview_tab:
 if not result:
  st.markdown('<div class="empty-state"><div class="empty-icon">◆</div><h2>A clearer view of the deal.</h2><p>Build the purchase and sales legs, allocate your cargoes and set the shipping assumptions. Your value bridge, cash flows and break-even analysis will appear here.</p><div class="empty-steps"><span>01 · Upload EPA prices</span><span>02 · Configure the deal</span><span>03 · Calculate NPV</span></div></div>',unsafe_allow_html=True)
 else:
  left,right=st.columns([2,1],vertical_alignment='center')
  with left:section('Calculated case','Investment overview',f"{int(ci['schedule'].Year.min())}–{int(ci['schedule'].Year.max())} · PV to {ci['discount_year']} · {ci['disc']:.1%} discount rate")
  with right:
   if st.button('Recalculate NPV',use_container_width=True,disabled=not ready,key='calculate_overview'):execute_case()
  unit=st.radio('Chart view',['$','$/MMBtu'],horizontal=True,key='overview_unit')
  st.plotly_chart(waterfall(t,unit),use_container_width=True,theme=None)
  st.plotly_chart(cash_chart(d,a,ci['discount_year'],unit),use_container_width=True,theme=None)
  st.plotly_chart(monthly_chart(d,unit),use_container_width=True,theme=None)
  left,right=st.columns([1.4,1],gap='large')
  with left:
   with st.container(border=True):
    section('Price discovery','Find break-even price')
    st.caption('Solves NPV = 0 for the calculated case. All other assumptions remain unchanged.')
    a1,b1=st.columns(2)
    side_label=a1.selectbox('Find break-even for',['Purchase price','Sales price'],key='be_side')
    target=b1.selectbox('Solve for',['Slope / multiplier','Fixed elements'],key='be_target')
    share=st.selectbox('CPI-linked share of fixed elements',list(range(0,101,5)),format_func=lambda x:f'{x}%',key='be_share')
    if st.button('Calculate Break-Even Price',use_container_width=True,key='be_run'):
     try:
      side='pur' if side_label.startswith('Purchase') else 'sal';root=solve_zero(ci,side,target,share/100);bf=ci[side].copy()
      if target=='Slope / multiplier':bf['slope']=root
      else:bf['fixed']=root*(1-share/100);bf['cpi']=root*(share/100)
      st.session_state.studio_be=(side_label,formula_text(side,bf),target,share)
     except Exception as e:st.error(str(e))
    if 'studio_be' in st.session_state:
     lab,txt,targ,sh=st.session_state.studio_be;st.success(f'{lab}: {txt}');st.caption(f'Solved {targ.lower()}. Fixed-element CPI share: {sh}%.')
  with right:
   with st.container(border=True):
    section('Share the case','Presentation export')
    st.caption('Generate the overview, NPV waterfall, annual cash flows and ±10% sensitivities from the calculated case.')
    if st.button('Make a PowerPoint Presentation',use_container_width=True,key='ppt_run'):
     try:st.session_state.studio_ppt=ppt_bytes(d,a,ci)
     except ImportError:st.error('PowerPoint export needs python-pptx and matplotlib. Install them in the same Python environment as Streamlit.')
     except Exception as e:st.error(f'Export failed: {e}')
    if st.session_state.get('studio_ppt'):st.download_button('Download PowerPoint Presentation',st.session_state.studio_ppt,'LNG_Business_Case_NPV.pptx','application/vnd.openxmlformats-officedocument.presentationml.presentation',use_container_width=True,key='ppt_download')

  with st.container(border=True):
   section('Hedge reconciliation','Component-matched sales formula','Match index, fixed and CPI buckets separately in present value. Annual residuals remain visible.')
   st.caption('Fixed costs include shipping capex, Liq. no CPI and Liq. BoG no CPI. CPI costs include Liq. CPI, Liq. BoG CPI and shipping opex CPI.')
   if st.button('Calculate component-matched sales formula',key='hedge_run',use_container_width=True):
    try:
     hf,hd,ha=match_sales_components(ci)
     st.session_state.studio_hedge={'formula':hf,'detail':hd,'annual':ha}
    except Exception as e:st.error(str(e))
   if 'studio_hedge' in st.session_state:
    h=st.session_state.studio_hedge
    ht=totals(h['detail'],h['annual']);matched_formula=compact_price_formula(h['formula'])
    result_left,result_right=st.columns([1,2])
    with result_left:st.metric('NPV with matched sales formula',money(ht['npv']))
    with result_right:
     st.markdown(f'<div class="matched-formula-card"><div class="matched-formula-label">Matched sales price formula</div><div class="matched-formula-value">{esc(matched_formula)}</div><div class="matched-formula-note">Slope x selected price index + fixed no-CPI + CPI-linked element</div></div>',unsafe_allow_html=True)
    st.caption('Preview only until applied. Apply the full-precision coefficients to the sales leg, then recalculate the main case.')
    st.button('Apply matched formula to sales leg',key='hedge_apply',on_click=apply_matched_sales,use_container_width=True)
    st.markdown('**Annual residuals for the matched formula (USD)**');table(hedge_table(h['detail']))
    hp=hedge_table(h['detail'],pv=True)
    st.markdown('**Discounted residual totals (USD)**');table(hp[['Index residual USD','Fixed residual USD','CPI residual USD','Extrinsic USD','Net residual USD']].sum().to_frame('PV USD').reset_index(names='Component'))
    if ci['ext']!=0:st.warning('Extrinsic income is non-zero. The three matched buckets net to zero in PV, but total NPV retains extrinsic income.')
   st.markdown('**Current calculated case: annual hedge residuals (USD)**');table(hedge_table(d))

with prices_tab:
 section('Market data','EPA price library','Every uploaded gas and Brent series, with quarterly values and a selectable curve view.')
 if prices is None:st.info('Upload EPA workbooks in Deal setup → Data & horizon.')
 else:
  sel=st.multiselect('Series',sorted(prices.Market.unique()),default=sorted(prices.Market.unique()),key='price_series')
  v=prices[prices.Market.isin(sel)].copy();v['Period']=v.Year.astype(str)+' Q'+v.Quarter.astype(str)
  f=go.Figure()
  for j,name in enumerate(sel):
   q=v[v.Market==name].sort_values(['Year','Quarter']);f.add_trace(go.Scatter(x=q.Period,y=q.Price,mode='lines+markers',name=name,line=dict(width=2.5,color=(LINES+['#d5ad63','#b59aeb','#77bd83'])[j%6])))
  f=chart(f,'Uploaded forward assumptions',height=450);f.update_yaxes(title_text='Source price units',tickprefix='')
  st.plotly_chart(f,use_container_width=True,theme=None)
  st.caption('This library displays real uploaded prices in source units. Valuation inflates both EPA legs and the BoG decomposition from the EPA real reference year. No commodity-unit conversion is applied.')
  table(v)

with shipping_tab:
 section('Logistics economics','Shipping cost structure','CPI-adjusted opex and non-CPI capex, reconciled to delivered volume.')
 if not result:st.info('Calculate a case to view shipping results.')
 else:
  sd=d.groupby('Year',as_index=False).agg({'Delivered MMBtu':'sum','Shipping Opex USD':'sum','Shipping Capex USD':'sum'})
  for c,nm in [('Shipping Opex USD','Shipping Opex USD/MMBtu'),('Shipping Capex USD','Shipping Capex USD/MMBtu')]:sd[nm]=sd[c]/sd['Delivered MMBtu']
  sd['Total USD']=sd[['Shipping Opex USD','Shipping Capex USD']].sum(axis=1)
  sd['Total USD/MMBtu']=sd['Total USD']/sd['Delivered MMBtu']
  a1,b1,c1=st.columns(3)
  a1.metric('Physical boil-off',f"{ci['boil']:.3%}")
  b1.metric('Delivered / cargo',f"{ci['cargo']*(1-ci['boil']):,.0f} MMBtu")
  c1.metric('Total shipping',money(sd['Total USD'].sum()))
  unit=st.radio('Shipping view',['$','$/MMBtu'],horizontal=True,key='shipping_unit')
  f=go.Figure();xx=sd.Year.astype(str)
  series=[('Shipping Opex USD','Opex'),('Shipping Capex USD','Capex')] if unit=='$' else [('Shipping Opex USD/MMBtu','Opex'),('Shipping Capex USD/MMBtu','Capex')]
  for j,(c,name) in enumerate(series):f.add_trace(go.Bar(x=xx,y=sd[c],name=name,marker_color=PALETTE[j],text=[money(x,unit) if abs(x)>1e-10 else '' for x in sd[c]],textposition='inside',insidetextfont_color='white',hovertemplate='%{x}<br>%{text}<extra>%{fullData.name}</extra>'))
  f.update_layout(barmode='stack',xaxis_type='category');st.plotly_chart(chart(f,'Annual shipping cost breakdown',unit),use_container_width=True,theme=None)
  vals=sd['Total USD'] if unit=='$' else sd['Total USD/MMBtu']
  f=go.Figure(go.Scatter(x=xx,y=vals,mode='lines+markers+text',name='Total shipping',text=[money(v,unit) for v in vals],textfont_color=TEXT,textposition='top center',line=dict(width=3,color=LINES[0]),marker=dict(size=9),cliponaxis=False))
  f.update_xaxes(type='category');st.plotly_chart(chart(f,'Total annual shipping cost',unit,410),use_container_width=True,theme=None)
  if not result.get('routes',pd.DataFrame()).empty:
   st.caption('Calculated-case route flags (informational only; not used in valuation).')
   st.dataframe(result['routes'],hide_index=True,use_container_width=True)
  section('Underlying assumptions','Annual shipping table')
  table(sd)

with sens_tab:
 section('Scenario analysis','Sensitivity to key drivers','Change one driver at a time around the calculated case. The base case remains first.')
 if not result:st.info('Calculate a case to run sensitivities.')
 else:
  unit=st.radio('View',['$','$/MMBtu'],horizontal=True,key='sensitivity_unit')
  with st.container(border=True):
   cols=st.columns(4)
   changes=[cols[j].slider(label,0,20,10,key=f'sensitivity_{j}') for j,label in enumerate(['Sales price change (%)','Purchase price change (%)','Shipping cost change (%)','Extrinsic value change (%)'])]
  q=sensitivity_rows(ci,changes)
  st.plotly_chart(sensitivity_chart(q,t,unit),use_container_width=True,theme=None)
  st.caption('Sales changes scale EPA sales, BoG income and fixed income. Purchase changes scale EPA purchase, BoG cost and all liquefaction fees. Shipping changes scale CPI-adjusted opex and capex. Every scenario uses the calculated snapshot and its Discount back to year.')

with explanation_tab:
 section('Calculation audit trail','Explanation','The charts, break-even solver and presentation use the same calculated snapshot.')
 st.markdown("""
**Independent years.** Discount back to affects only PV calculations. Sales base year and Purchase base year retain the existing fee/shipping CPI logic. EPA prices have a separate real reference year and are inflated to nominal before the commodity decomposition.

**CPI timing.** Sales: the first included annual CPI rate is sales base year + 1. Purchase and shipping opex: the first included annual CPI rate is purchase base year itself. For base year 2022, sales uses the product of (1+CPI) for 2023 through the cash-flow year; costs use 2022 through the cash-flow year. A 2022 purchase fee is therefore escalated by the 2022 CPI rate in 2022. Years before the reference point are deflated. This explicitly replaces the previous +2/+1 anchors.

**Volumes and non-duplicative commodity decomposition.** Loaded/bought = cargo count x loaded cargo size; delivered = loaded x (1-b). Uploaded real EPA prices are first inflated to nominal using the separate EPA real reference year. The waterfall then splits each full contractual formula into two rows without changing total economics. EPA sales = nominal sales EPA x delivered volume. BoG income = nominal sales EPA x (sales slope - 1) x delivered volume. EPA purchase = nominal purchase EPA x delivered volume. BoG cost = nominal purchase EPA x (purchase slope x loaded volume - delivered volume). Therefore EPA sales and EPA purchase match when their nominal indices match, while BoG income and BoG cost match when the slopes compensate correctly for volume loss. Combined sales still equal nominal EPA x sales slope x delivered volume; combined purchase still equals nominal EPA x purchase slope x loaded volume. Liq. boiled off remains unchanged and covers only liquefaction fees on lost volume.

**Fixed fees.** Fixed sales income = sales fixed fee x delivered volume. Purchase no-CPI fee = purchase fixed fee x delivered volume. CPI income = sales CPI fee x sales CPI factor x delivered volume. CPI purchase fee = purchase CPI fee x purchase CPI factor x delivered volume.

**Liquefaction boiled off.** The total calculation remains [(fixed fee + CPI-adjusted fee)/(1-b) - (fixed fee + CPI-adjusted fee)] x delivered volume. V29 displays the exact same total as two components: Liq. BoG no CPI = fixed liquefaction fee x boiled-off volume; Liq. BoG CPI = CPI-adjusted liquefaction fee x boiled-off volume. Their sum equals the former single Liq. boiled off element, so total cost and NPV are unchanged.

**Shipping.** Capex = Interface row 97 x delivered volume, without CPI. Opex = (rows 92-96 + row 98) x purchase CPI factor x delivered volume. Flat shipping assumptions follow the same CPI rules. Route flags B41:C43 are informational only. Source shipping opex must not already include the same CPI escalation.

**Annual-average mode.** Each selected index/year uses the arithmetic mean of all four uploaded quarterly prices, (Q1+Q2+Q3+Q4)/4. Missing quarters raise an error rather than averaging a partial year. Enter cargo counts separately for each year. The internal January row is only an annual bucket: it does not select the Q1 price. Discounting and CPI remain annual. The existing monthly/quarterly schedule is retained as the alternative mode.

**NPV.** Annual cash flow = total income - total costs. PV = annual cash flow / (1+discount rate)^(cash-flow year - Discount back to year). NPV is the sum of annual PVs. NPV/MMBtu divides by discounted delivered volume. No tax is applied. Extrinsic value remains additional income if non-zero.

**Component-matched sales formula.** The matching tool solves the separate sales BoG slope after subtracting EPA base sales from the index cost target; EPA base sales remains EPA price x delivered volume. Fixed fee and CPI fee matching are unchanged. This is not a guarantee of an annual hedge. Different index curves, changing shipping assumptions or differing CPI patterns can leave annual residuals even when discounted totals match. The annual residual table makes those differences visible. Non-zero extrinsic value is shown separately and is not silently absorbed into the three matched buckets.
""")
 if not result:st.info('Calculate a case to see the numerical audit trail.')
 else:
  st.info(f"Sales CPI first included year: {ci['sales_base']+1}. Purchase/shipping CPI first included year: {ci['purchase_base']}. EPA real reference year: {ci['epa_real_base']}. Discount back to: {ci['discount_year']}. Price mode: {ci['price_mode']}.")
  y=st.selectbox('Inspect year',sorted(d.Year.unique()),key='explain_year');yd=d[d.Year==y].copy();v=yd['Delivered MMBtu'].sum()
  st.markdown('#### Annual component reconciliation')
  rows=[{'Component':name,'Nominal USD':sgn*yd[c].sum(),'USD / delivered MMBtu':sgn*yd[c].sum()/v,'PV USD':sgn*yd['PV '+c].sum()} for name,c,key,sgn in COMPONENTS]
  rows.append({'Component':'Net cash flow / PV','Nominal USD':yd['Cash Flow USD'].sum(),'USD / delivered MMBtu':yd['Cash Flow USD'].sum()/v,'PV USD':yd['PV Cash Flow USD'].sum()});table(pd.DataFrame(rows))
  st.markdown('#### Monthly inputs and intermediate factors')
  table(yd[['Year','Month','Cargoes','Loaded MMBtu','Boiled-off MMBtu','Delivered MMBtu','EPA Purchase Base Volume MMBtu','EPA Sales Price Real','EPA Purchase Price Real','EPA Inflation Factor','EPA Sales Price','EPA Purchase Price','BoG Index Coefficient Sales','BoG Index Coefficient Purchase','Sales CPI Factor','Purchase CPI Factor','Shipping CPI Factor','Raw Shipping Opex USD/MMBtu','Raw Shipping Capex USD/MMBtu','Discount Factor']])
  st.markdown('#### Annual cash flows and discounting')
  audit=a.copy();audit['Discount Factor']=(1+ci['disc'])**(audit.Year-ci['discount_year']);table(audit)
  st.markdown('#### CPI factors by year')
  table(pd.DataFrame({'Year':sorted(ci['sales_cpi']),'Sales CPI factor':[ci['sales_cpi'][k] for k in sorted(ci['sales_cpi'])],'Purchase / shipping CPI factor':[ci['purchase_cpi'][k] for k in sorted(ci['sales_cpi'])]}))
  st.markdown('#### Annual hedge residuals (nominal USD)');table(hedge_table(d))
  st.download_button('Download calculation detail (CSV)',d.to_csv(index=False).encode('utf-8'),'LNG_case_calculation_detail.csv','text/csv',key='audit_csv')

st.markdown('<div class="footer"><span>LNG BUSINESS CASE · STUDIO</span><span>Annual discounting · Split liquefaction BoG fixed/CPI · Deployment-safe theme</span></div>',unsafe_allow_html=True)

