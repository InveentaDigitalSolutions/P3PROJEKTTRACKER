import zipfile, re, json, datetime, xml.etree.ElementTree as ET
from collections import Counter, OrderedDict
NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
PR='{http://schemas.openxmlformats.org/package/2006/relationships}'
FN='P3 SlpP & TlP Projects Database concentrate V1.xlsx'

def clean(s): return re.sub(r'[ \t]+',' ',(s or '').replace('\t',' ')).strip()
def s2iso(n):
    try: return (datetime.date(1899,12,30)+datetime.timedelta(days=int(float(n)))).isoformat()
    except Exception: return None
def pdate(v):
    v=clean(v)
    if re.fullmatch(r'\d{5,6}',v):
        n=int(v)
        if 36500<=n<=73000: return s2iso(v)
        return None
    m=re.fullmatch(r'(\d{1,2})[./-](\d{1,2})[./-](\d{4})',v)
    if m: return '%s-%02d-%02d'%(m.group(3),int(m.group(2)),int(m.group(1)))
    return None
def cl(r): return re.match(r'([A-Z]+)',r).group(1)
def cn(s):
    n=0
    for ch in s: n=n*26+(ord(ch)-64)
    return n

z=zipfile.ZipFile(FN)
shared=[''.join(t.text or '' for t in si.iter(NS+'t')) for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(NS+'si')]
rels={r.get('Id'):r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels')).iter(PR+'Relationship')}
wb=ET.fromstring(z.read('xl/workbook.xml'))
sheet_path=None
for s in wb.iter(NS+'sheet'):
    if 'conglomerate' in s.get('name').lower(): sheet_path=rels[s.get(RNS+'id')]
ws=ET.fromstring(z.read('xl/'+sheet_path)); cells={}
for c in ws.iter(NS+'c'):
    ref=c.get('r')
    if not ref: continue
    row=int(re.match(r'[A-Z]+(\d+)',ref).group(1)); col=cn(cl(ref))
    t=c.get('t'); v=c.find(NS+'v'); isn=c.find(NS+'is'); val=''
    if t=='s' and v is not None: val=shared[int(v.text)]
    elif isn is not None: val=''.join(x.text or '' for x in isn.iter(NS+'t'))
    elif v is not None: val=v.text or ''
    cells[(row,col)]=val
maxr=max(r for r,_ in cells)
# columns: 1 Area,2 Location,3 Category,4 Project Title,5 Task no,6 Task,7 Dept,8 Start,9 Finish,10 Status
def g(r,c): return clean(cells.get((r,c),''))

projects=OrderedDict()
for r in range(2,maxr+1):
    title=g(r,4); task=g(r,6)
    if not title and not task: continue
    if not task: continue
    key=(g(r,1),g(r,2),g(r,3),title)  # area,loc,cat,title
    if key not in projects:
        projects[key]={'area':g(r,1),'location':g(r,2),'category':g(r,3),'title':title,'tasks':[]}
    projects[key]['tasks'].append({
        'no': g(r,5),'task':task,'dept':g(r,7),
        'start':pdate(g(r,8)),'finish':pdate(g(r,9)),'status':g(r,10),
    })

plist=list(projects.values())
json.dump(plist, open('_concentrate.json','w'), indent=1, ensure_ascii=False)

L=[]
L.append('PROJECTS: %d   TASKS: %d'%(len(plist), sum(len(p['tasks']) for p in plist)))
L.append('By area: '+', '.join('%s=%d'%(k,v) for k,v in sorted(Counter(p['area'] for p in plist).items())))
L.append('By location: '+', '.join('%s=%d'%(k,v) for k,v in sorted(Counter(p['location'] for p in plist).items())))
L.append('By category: '+', '.join('%s=%d'%(k,v) for k,v in sorted(Counter(p['category'] for p in plist).items())))
allstatus=Counter(t['status'] for p in plist for t in p['tasks'])
L.append('Task status values: '+', '.join('%s=%d'%(k or '(blank)',v) for k,v in allstatus.most_common()))
L.append('')
for p in plist:
    nd=sum(1 for t in p['tasks'] if t['status'].lower()=='done')
    L.append('  [%s/%s] %-14s %2dt (%d done)  %s'%(p['area'],p['location'],p['category'],len(p['tasks']),nd,p['title'][:46]))
open('_concentrate_scope.txt','w').write('\n'.join(L))
print('OK')
