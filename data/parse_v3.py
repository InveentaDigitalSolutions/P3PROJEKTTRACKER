import zipfile, re, json, datetime, xml.etree.ElementTree as ET
from collections import OrderedDict
NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
PR='{http://schemas.openxmlformats.org/package/2006/relationships}'
FN='P3 SlpP  TlP Projects Database concentrate V3.xlsx'

def clean(s): return re.sub(r'[ \t]+',' ',(s or '').replace('\t',' ')).strip()
def s2iso(n):
    try: return (datetime.date(1899,12,30)+datetime.timedelta(days=int(float(n)))).isoformat()
    except Exception: return None
def pdate(v):
    v=clean(v)
    if re.fullmatch(r'\d{5,6}',v):
        n=int(v)
        return s2iso(v) if 36500<=n<=73000 else None
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
sp=[rels[s.get(RNS+'id')] for s in wb.iter(NS+'sheet') if 'conglomerate' in s.get('name').lower()][0]
ws=ET.fromstring(z.read('xl/'+sp)); cells={}
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
def g(r,c): return clean(cells.get((r,c),''))
# V2 columns: 1 Area,2 Location,3 Category,4 Title,5 ExecSponsor,6 ProjResponsible,
#             7 StatusOverview,8 Task no,9 Task,10 Dept,11 Start,12 Finish,13 Status
projects=OrderedDict()
for r in range(2,maxr+1):
    title=g(r,4); task=g(r,9)
    if not title: continue
    # Group rows into ONE project per ECR number (e.g. 3E10177261). Some titles
    # carry a per-task trailing component code (…6002KS0844/0845/…) that would
    # otherwise split one project into many 1-task projects. Fall back to the
    # full title when no ECR number is present.
    ecr=re.search(r'3E\d{6,}', title)
    gkey=(g(r,1),g(r,2),g(r,3), ecr.group(0) if ecr else title.lower())
    if gkey not in projects:
        # Canonical title: strip a trailing " <CODE>" component suffix if present.
        ctitle=re.sub(r'\s+\d*[A-Z]{2}\w*\d+$','',title).strip() if ecr else title
        projects[gkey]={'area':g(r,1),'location':g(r,2),'category':g(r,3),'title':ctitle,
                       'sponsor':g(r,5),'responsible':g(r,6),'statusOverview':g(r,7),'tasks':[]}
    if task:
        projects[gkey]['tasks'].append({'no':g(r,8),'task':task,'dept':g(r,10),
            'start':pdate(g(r,11)),'finish':pdate(g(r,12)),'status':g(r,13)})
plist=[p for p in projects.values()]
json.dump(plist, open('_v3.json','w'), indent=1, ensure_ascii=False)
print('projects:',len(plist),'tasks:',sum(len(p['tasks']) for p in plist))
