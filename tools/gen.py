import json,sys
W=sys.argv[1]; OUT=sys.argv[2]
import os
HERE=os.path.dirname(os.path.abspath(__file__))
t=open(HERE+'/template.html').read(); d=json.load(open(W+'data.json'))
pages=[('MACRO','index.html','Macro','Macro Coverage Dashboard','MACRO class, kg shown as tonnes',''),
('PACKAGING','packaging.html','Packaging','Packaging Coverage Dashboard','PACKAGING class, in pieces',"<li><b>Packaging</b>: class PACKAGING; quantities are pieces. Warehouse BAC (Visayas) is counted under HOREB. Pending lines with a zero or negative balance (over-received) are ignored.</li>\n    "),
('MICRO','micro.html','Micro','Micro Coverage Dashboard','MICRO class, in kg',"<li><b>Micro</b>: class MICRO; quantities are kg (shown in kg, not tonnes). Warehouse BAC (Visayas) is counted under HOREB. Pending lines with a zero or negative balance are ignored.</li>\n    "),
('FEEDTAGS','feedtags.html','Feedtags','Feedtags Coverage Dashboard','FEEDTAGS class, in pieces',"<li><b>Feedtags</b>: class FEEDTAGS; quantities are pieces. Pending lines with a zero or negative balance are ignored.</li>\n    ")]
nav=lambda cur:' · '.join(('<b>%s</b>'%n if f==cur else '<a href="%s" style="color:var(--accent)">%s</a>'%(f,n)) for f,n in [(p[1],p[2]+' coverage') for p in pages]+[('sync.html','Sync')])
for g,f,n,title,sn,extra in pages:
    dd={k:d[k] for k in ('asof','invSync','pbaSync','costSyncFull','plants') if k in d}; dd['groups']={g:d['groups'][g]}
    open(OUT+f,'w').write(t.replace('/*TITLE*/',title).replace('/*NAV*/',nav(f)).replace('/*PKGNOTE*/',extra).replace('/*STOCKNOTE*/',sn).replace('/*DATA*/',json.dumps(dd,separators=(',',':'))))
