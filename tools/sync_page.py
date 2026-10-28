import json, sys, os
W=sys.argv[1]; OUT=sys.argv[2]; TAKEN=sys.argv[3]   # work dir, output dir (repo root), snapshot date YYYY-MM-DD
HIST=OUT+'history.json'
nav_pages=[('index.html','Macro coverage'),('packaging.html','Packaging coverage'),('micro.html','Micro coverage'),('feedtags.html','Feedtags coverage')]

def snap(d,taken,inv,pba,cost=None):
    out={'taken':taken,'inv':inv,'pba':pba,'groups':{},'detail':{}}
    if cost: out['cost']=cost
    for g,v in d['groups'].items():
        its=v['items']; det={}; crit=0; wc=0
        for i in its:
            c=sum(i['cons'].values()); s=sum(i['stock'].values()); p=sum(i['pend'].values())
            det[i['code']]=[i['name'],round(s,3),round(p,3),round(c,3)]
            if round(c)>0:
                wc+=1
                if s/(c/30)<30: crit+=1
        out['groups'][g]={'label':v['label'],'unit':v['unit'],'items':len(its),'withCons':wc,'critical':crit,
            'stock':round(sum(x[1] for x in det.values()),3),'pend':round(sum(x[2] for x in det.values()),3),'cons':round(sum(x[3] for x in det.values()),3)}
        out['detail'][g]=det
    return out

hist=json.load(open(HIST)) if os.path.exists(HIST) else []
cur=json.load(open(W+'data.json'))
new=snap(cur,TAKEN,cur['invSyncFull'],cur['pbaSyncFull'],cur.get('costSyncFull'))
hist=[h for h in hist if h['taken']!=TAKEN]+[new]
for h in hist:   # recompute with the current rule: under half a unit of consumption a month = no usage
    for g,v in h['groups'].items():
        det=h['detail'][g].values()
        v['withCons']=sum(1 for x in det if round(x[3])>0)
        v['critical']=sum(1 for x in det if round(x[3])>0 and x[1]/(x[3]/30)<30)
json.dump(hist,open(HIST,'w'),separators=(',',':'))

HTML='''<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Data Sync</title>
<style>
:root{--bg:#eef2ef;--panel:#fff;--ink:#14221c;--muted:#5a6a63;--line:#d6dfda;--line-strong:#9fb2a9;--accent:#1d6a4c;--accent-soft:#e3f0ea;--crit:#b42318;--crit-soft:#fdecea;--warn:#a86f12;--warn-soft:#fbf1dc;--ok:#167a4a;--ok-soft:#e2f4ea;--font:system-ui,-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]){--bg:#0e1512;--panel:#16201c;--ink:#e4ece8;--muted:#94a59d;--line:#26342e;--line-strong:#4a5f56;--accent:#4cbc8b;--accent-soft:#16302a;--crit:#f07b70;--crit-soft:#341a18;--warn:#e3b25e;--warn-soft:#2c2414;--ok:#55c98f;--ok-soft:#14301f;color-scheme:dark}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 var(--font)}
.wrap{max-width:1100px;margin:0 auto;padding:16px}h1{font-size:20px;margin:0}h2{font-size:14px;margin:0 0 8px}
.sub,.note{color:var(--muted);font-size:12.5px}nav{margin:4px 0 14px;font-size:12.5px}nav a{color:var(--accent)}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:12px 14px;margin-bottom:12px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:12px;margin-bottom:12px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:12px 14px}
.card .t{font-weight:650}.card .v{font-size:15px;margin:6px 0 2px;font-variant-numeric:tabular-nums}
.pill{display:inline-block;padding:1px 8px;border-radius:999px;font-size:11.5px;font-weight:600}
.p-ok{background:var(--ok-soft);color:var(--ok)}.p-warn{background:var(--warn-soft);color:var(--warn)}.p-crit{background:var(--crit-soft);color:var(--crit)}
table{border-collapse:collapse;width:100%;font-variant-numeric:tabular-nums;font-size:13px}
th,td{padding:6px 9px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}th{font-size:11.5px;color:var(--muted)}
th.l,td.l{text-align:left}.scroll{overflow:auto}.up{color:var(--ok)}.down{color:var(--crit)}
select{font:inherit;color:var(--ink);background:var(--panel);border:1px solid var(--line-strong);border-radius:6px;padding:5px 8px}
ol{margin:6px 0 0 18px;padding:0}
</style></head><body><div class="wrap">
<h1>Data sync</h1><div class="sub" id="sub"></div><nav id="nav"></nav>
<div class="panel" style="display:flex;gap:12px;align-items:center;flex-wrap:wrap"><button id="chk" style="font:inherit;border:1px solid var(--line-strong);background:var(--panel);color:var(--ink);border-radius:6px;padding:5px 12px;cursor:pointer">⟳ Check sources now</button><span class="note" id="chkMsg">Asks the Google Sheets when they last synced from SAP.</span></div>
<div class="cards" id="cards"></div>
<div class="panel"><h2>How to refresh</h2>
<ol class="note"><li>Each dashboard is a snapshot of the sheets. The <b>Sync now</b> button at the top of a dashboard re-reads the sheets through the Apps Script and refreshes that page in your browser (needs the sync key; see SETUP-LIVE-SYNC.md).</li>
<li>Both source sheets sync from SAP on their own schedule. Check the "Source synced" times above first, since a snapshot can only be as fresh as the sheets.</li>
<li>Sync now does not change the saved snapshot or this history. To save a new snapshot for everyone, ask Claude to <b>SYNC NOW</b>.</li></ol></div>
<div class="panel"><h2>Snapshot history</h2><div class="scroll"><table id="hist"></table></div></div>
<div class="panel"><h2>Changes since previous snapshot</h2>
<div style="margin-bottom:8px"><select id="grp"></select></div><div id="chg"></div></div>
</div>
<script src="config.js"></script>
<script>
const HIST=/*HIST*/;const NAV=/*NAV*/;
const $=id=>document.getElementById(id);
const fmt=(n,d=0)=>n.toLocaleString('en-US',{maximumFractionDigits:d,minimumFractionDigits:d});
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const fd=s=>new Date(s+'T00:00:00').toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
const cur=HIST[HIST.length-1], prev=HIST.length>1?HIST[HIST.length-2]:null;
$('nav').innerHTML=NAV.map(([f,n])=>`<a href="${f}">${n}</a>`).join(' · ')+' · <b>Sync</b>';
$('sub').textContent='Latest snapshot taken '+fd(cur.taken)+' · '+HIST.length+' snapshot'+(HIST.length>1?'s':'')+' kept';
function age(str,ref){const m=/(\\d{2})\\/(\\d{2})\\/(\\d{4})/.exec(str);if(!m)return null;const d=new Date(+m[3],+m[1]-1,+m[2]);const t=ref||new Date(cur.taken+'T00:00:00');return Math.max(0,Math.round((t-d)/864e5));}
function pill(a){if(a==null)return '';return a<=1?'<span class="pill p-ok">Fresh</span>':a<=3?`<span class="pill p-warn">${a} days old</span>`:`<span class="pill p-crit">${a} days old</span>`;}
const clean=s=>s.replace(/^(INV sync|Last sync):\\s*/,'');
function cards(inv,pba,ref,live,cost){ $('cards').innerHTML=[['Inventory and consumption','RM_WH_Mgt › INV, MIN_ sheets',inv],['Pending PBA + PO','SBO Purchase Blanket Agreement › OPEN PBA + PO',pba]].concat(cost?[['Average cost','Inventory Price › Sheet1 (Inventory Ageing import)',cost]]:[]).map(([t,src,s])=>
 `<div class="card"><div class="t">${t} ${pill(age(s,ref))}</div><div class="note">${src}</div><div class="v">${esc(clean(s))}</div><div class="note">${live?'Source synced (read live from the sheet just now)':'Source synced at the time of the latest snapshot'}</div></div>`).join(''); }
cards(cur.inv,cur.pba,new Date(cur.taken+'T00:00:00'),false,cur.cost);
(function(){ const m=/[#&]key=([^&]+)/.exec(location.hash); if(m){ try{ sessionStorage.setItem('rmcov_key',decodeURIComponent(m[1])); }catch(e){} history.replaceState(null,'',location.pathname+location.search); } })();
$('chk').onclick=async()=>{ const msg=$('chkMsg'), btn=$('chk');
  if(!window.SYNC_URL){ msg.textContent='Live check is not set up yet. Put the Apps Script URL in config.js (see SETUP-LIVE-SYNC.md).'; return; }
  let key=''; try{ key=sessionStorage.getItem('rmcov_key')||''; }catch(e){}
  if(!key){ key=(prompt('Enter the sync key')||'').trim(); if(!key){ msg.textContent='Cancelled.'; return; } try{ sessionStorage.setItem('rmcov_key',key); }catch(e){} }
  btn.disabled=true; msg.textContent='Asking the sheets…';
  try{ const j=await (await fetch(window.SYNC_URL+'?action=status&key='+encodeURIComponent(key))).json();
    if(!j.ok){ if(j.error==='unauthorized'){ try{ sessionStorage.removeItem('rmcov_key'); }catch(e){} msg.textContent='Wrong or missing sync key.'; } else msg.textContent='Check failed: '+j.error; return; }
    cards(j.status.invSyncFull,j.status.pbaSyncFull,new Date(),true,j.status.costSyncFull);
    const newer = j.status.invSyncFull!==cur.inv || j.status.pbaSyncFull!==cur.pba || (cur.cost && j.status.costSyncFull!==cur.cost);
    msg.textContent=newer?'The sheets have newer data than the latest snapshot. Open a dashboard and click Sync now.':'The sheets match the latest snapshot.';
  }catch(e){ msg.textContent='Check failed: '+(e.message||e); } finally{ btn.disabled=false; } };
const G=Object.keys(cur.groups);
$('hist').innerHTML='<thead><tr><th class="l">Snapshot</th><th class="l">Inventory synced</th><th class="l">Pending synced</th>'+G.map(g=>`<th>${cur.groups[g].label} stock</th><th>${cur.groups[g].label} pending</th><th>${cur.groups[g].label} critical</th>`).join('')+'</tr></thead><tbody>'+
 HIST.slice().reverse().map(h=>`<tr><td class="l">${fd(h.taken)}</td><td class="l">${esc(clean(h.inv).replace(/ by .*/,''))}</td><td class="l">${esc(clean(h.pba).replace(/ by .*/,''))}</td>`+
  G.map(g=>{const x=h.groups[g];return x?`<td>${fmt(x.stock)} ${x.unit}</td><td>${fmt(x.pend)} ${x.unit}</td><td>${x.critical}</td>`:'<td></td><td></td><td></td>'}).join('')+'</tr>').join('')+'</tbody>';
$('grp').innerHTML=G.map(g=>`<option value="${g}">${cur.groups[g].label}</option>`).join('');
function changes(){
  const g=$('grp').value;
  if(!prev||!prev.detail[g]){$('chg').innerHTML='<div class="note">No earlier snapshot to compare with.</div>';return;}
  const a=prev.detail[g], b=cur.detail[g], u=cur.groups[g].unit;
  const rows=[];
  new Set([...Object.keys(a),...Object.keys(b)]).forEach(c=>{
    const x=a[c]||[ (b[c]||[''])[0],0,0,0], y=b[c]||[x[0],0,0,0];
    rows.push({c,name:y[0]||x[0],ds:y[1]-x[1],dp:y[2]-x[2],dc:y[3]-x[3],s:y[1],p:y[2],added:!a[c],removed:!b[c]});
  });
  const top=rows.filter(r=>Math.abs(r.ds)+Math.abs(r.dp)>0).sort((p,q)=>(Math.abs(q.ds)+Math.abs(q.dp))-(Math.abs(p.ds)+Math.abs(p.dp))).slice(0,15);
  const cl=v=>v>0?'up':v<0?'down':'';const sg=v=>(v>0?'+':'')+fmt(v,u==='t'?1:0);
  const unchanged=rows.length-rows.filter(r=>Math.abs(r.ds)+Math.abs(r.dp)>0).length;
  $('chg').innerHTML=`<div class="note" style="margin-bottom:6px">Compared with ${fd(prev.taken)}. Showing the 15 largest movements (${u}); ${unchanged} items unchanged.</div>`+
   '<div class="scroll"><table><thead><tr><th class="l">Item</th><th>Stock now</th><th>Stock change</th><th>Pending now</th><th>Pending change</th><th>Consumption change</th></tr></thead><tbody>'+
   top.map(r=>`<tr><td class="l">${esc(r.name)} <span class="note">${esc(r.c)}</span>${r.added?' <span class="pill p-ok">new</span>':''}${r.removed?' <span class="pill p-warn">gone</span>':''}</td><td>${fmt(r.s,u==='t'?1:0)}</td><td class="${cl(r.ds)}">${sg(r.ds)}</td><td>${fmt(r.p,u==='t'?1:0)}</td><td class="${cl(r.dp)}">${sg(r.dp)}</td><td>${sg(r.dc)}</td></tr>`).join('')+'</tbody></table></div>';
}
$('grp').onchange=changes;changes();
</script></body></html>'''
open(OUT+'sync.html','w').write(HTML.replace('/*HIST*/',json.dumps(hist,separators=(',',':'))).replace('/*NAV*/',json.dumps(nav_pages)))
print('snapshots',len(hist))
