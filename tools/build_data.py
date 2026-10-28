import openpyxl, json, collections, datetime, warnings, sys
warnings.filterwarnings('ignore')
import sys
S=sys.argv[1] if len(sys.argv)>1 else './'
if not S.endswith('/'): S+='/'
wh=openpyxl.load_workbook(S+'src/wh.xlsx',data_only=True)
pb=openpyxl.load_workbook(S+'src/pba.xlsx',data_only=True)

def plant_of_wh(w):
    if w is None: return None
    if w.endswith('-IT') or w.endswith('-QA'): return None
    if w.startswith('AC'): return 'AC'
    if w.startswith('PFMIS'): return 'PFMIS'
    if w.startswith('BUKID') or w.startswith('ALAE'): return 'BUKID'
    if w.startswith('SOUTH'): return 'SOUTH'
    if w=='BAC' or w.startswith('HOREB') or w.startswith('HBEXT'): return 'HOREB'
    if w.startswith('ARGAO'): return 'ARGAO'
    return None
PL=['AC','PFMIS','BUKID','SOUTH','HOREB','ARGAO']
def code(x):
    try: return str(int(float(x)))
    except: return str(x)

ASOF_DT=datetime.datetime(2026,10,20)
def build(CLS,DIV):
    global items
    items={}
    def item(c,name):
        if c not in items:
            items[c]={'code':c,'name':name,'cons':dict.fromkeys(PL,0.0),'stock':dict.fromkeys(PL,0.0),'pend':dict.fromkeys(PL,0.0),'pos':[],'wh':{}}
        return items[c]

    # consumption: average weekly forecast consumption x 4, from each plant's MIN_ sheet
    mins={'MIN_BULACAN':'AC','MIN_ISA':'PFMIS','MIN_HOREB':'HOREB','MIN_ARGAO':'ARGAO','MIN_BUKID':'BUKID','MIN_SOUTH':'SOUTH'}
    for sh,pl in mins.items():
        ws=wh[sh]; h=None
        for n,r in enumerate(ws.iter_rows(min_row=1,max_row=12,values_only=True),1):
            if r[0]=='INV_GROUP': h=list(r);hdr=n;break
        ix=h.index('AV W Fcasts consump')
        for r in ws.iter_rows(min_row=hdr+1,values_only=True):
            if r[0]==CLS and r[1] is not None:
                w=r[ix]
                it=item(code(r[1]),r[2]); it['cons'][pl]+=(w if isinstance(w,(int,float)) else 0)*4
    # stock (kg)
    ws=wh['INV']
    for r in ws.iter_rows(min_row=4,values_only=True):
        if r[2]==CLS and r[3] is not None:
            p=plant_of_wh(r[1])
            if p:
                it=item(code(r[3]),r[4]); it['stock'][p]+=(r[5] or 0)
                it['wh'][r[1]]=it['wh'].get(r[1],0)+(r[5] or 0)
    # pending PBA + PO
    end={}
    for r in wh['PBA'].iter_rows(min_row=4,values_only=True):
        if isinstance(r[1],(int,float)): end[(int(r[1]),code(r[6]),r[8])]=r[3]
    for r in pb['OPEN PBA + PO'].iter_rows(min_row=6,values_only=True):
        if not r[1] or r[5] is None: continue
        c=code(r[5])
        if c not in items: continue
        p=plant_of_wh(r[7])
        if not p: continue
        bal=r[9] or 0
        if bal<=0: continue
        e=end.get((int(r[1]),c,r[7]))
        items[c]['pend'][p]+=bal
        items[c]['pos'].append({'pba':int(r[1]),'vendor':r[4],'wh':r[7],'plant':p,'created':r[2].strftime('%Y-%m-%d'),
            'end':e.strftime('%Y-%m-%d') if e else None,'t':round(bal/DIV,3),'price':r[11],'expired':bool(e and e<ASOF_DT)})
    out=[]
    for c,it in sorted(items.items(),key=lambda x:(0,int(x[0]),'') if x[0].isdigit() else (1,0,x[0])):
        t=lambda d:{k:round(v/DIV,3) for k,v in d.items()}
        row={'code':c,'name':it['name'],'cons':t(it['cons']),'stock':t(it['stock']),'pend':t(it['pend']),'stockWh':{k:round(v/DIV,3) for k,v in it['wh'].items() if v},'pos':sorted(it['pos'],key=lambda x:-x['t'])}
        row['code']=c
        if sum(row['cons'].values())+sum(row['stock'].values())+sum(row['pend'].values())>0: out.append(row)
    return out

import re as _re
def _d(txt):
    m=_re.search(r'(\d{2})/(\d{2})/(\d{4})',str(txt)); return datetime.datetime(int(m.group(3)),int(m.group(1)),int(m.group(2))) if m else None
inv_sync=wh['INV']['A2'].value; pba_sync=pb['OPEN PBA + PO']['A3'].value
ASOF=_d(inv_sync) or datetime.datetime.today()
ASOF_DT=ASOF
print('inv',inv_sync,'| pba',pba_sync)
res={'asof':ASOF.strftime('%Y-%m-%d'),'invSync':ASOF.strftime('%b %d'),'invSyncFull':str(inv_sync),'pbaSyncFull':str(pba_sync),'pbaSync':(_d(pba_sync) or ASOF).strftime('%b %d'),'plants':PL,'groups':{
 'MACRO':{'label':'Macro','unit':'t','dec':0,'stockDec':1,'items':build('MACRO',1000)},
 'PACKAGING':{'label':'Packaging','unit':'pcs','dec':0,'stockDec':0,'items':build('PACKAGING',1)},
 'MICRO':{'label':'Micro','unit':'kg','dec':0,'stockDec':0,'items':build('MICRO',1)},
 'FEEDTAGS':{'label':'Feedtags','unit':'pcs','dec':0,'stockDec':0,'items':build('FEEDTAGS',1)}}}
json.dump(res,open(S+'data.json','w'),separators=(',',':'))
for k,g in res['groups'].items(): print(k,len(g['items']),sum(sum(i['stock'].values()) for i in g['items']),sum(sum(i['pend'].values()) for i in g['items']))
