import zipfile, re, json, datetime, glob, xml.etree.ElementTree as ET
from collections import Counter
NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
PR='{http://schemas.openxmlformats.org/package/2006/relationships}'
EXCLUDE=('resources database','timing templates','pjm projects baseline')

def clean(s): return re.sub(r'[ \t]+',' ',(s or '').replace('\t',' ')).strip()
def s2iso(n):
    try: return (datetime.date(1899,12,30)+datetime.timedelta(days=int(float(n)))).isoformat()
    except Exception: return None
def pdate(v):
    v=clean(v)
    # Excel date serials for 2000-2100 are 5 digits (~36500-73000).
    # A bare 4-digit value (e.g. a year "2025") is NOT a date serial.
    if re.fullmatch(r'\d{5,6}',v):
        n=int(v)
        if 36500 <= n <= 73000: return s2iso(v)
        return None
    m=re.fullmatch(r'(\d{1,2})[./-](\d{1,2})[./-](\d{4})',v)
    if m: return '%s-%02d-%02d'%(m.group(3),int(m.group(2)),int(m.group(1)))
    return None
def pint(v):
    v=clean(v); m=re.search(r'\d+',v)
    return int(m.group()) if (m and v.upper()!='NA') else None
def floc(fn):
    b=fn.rsplit('.',1)[0]
    loc='SLP' if re.search(r'slpp|slp',b,re.I) else ('TLP' if re.search(r'tlp|tl p',b,re.I) else '?')
    m=re.search(r'(CTG|ENG|LOG|MFE|PPS|PUQ|QMM)',b,re.I)
    return (m.group(1).upper() if m else b.split()[0].upper()), loc
def ptype(s):
    s=s.lower()
    if s.startswith('cip'): return 'CIP'
    if 'manager decision' in s: return 'ECR Manager Decision'
    if s.startswith('co ecr'): return 'CO ECR'
    if 'global ecr' in s: return 'Global ECR'
    if 'ecr' in s: return 'ECR - Standard'
    if s.startswith('new program'): return 'New Program'
    return 'Other'
def cl(r): return re.match(r'([A-Z]+)',r).group(1)
def cn(s):
    n=0
    for ch in s: n=n*26+(ord(ch)-64)
    return n
def grid(z,path,sh):
    ws=ET.fromstring(z.read('xl/'+path)); c={}
    for cc in ws.iter(NS+'c'):
        ref=cc.get('r')
        if not ref: continue
        row=int(re.match(r'[A-Z]+(\d+)',ref).group(1)); col=cn(cl(ref))
        t=cc.get('t'); v=cc.find(NS+'v'); isn=cc.find(NS+'is'); val=''
        if t=='s' and v is not None: val=sh[int(v.text)]
        elif isn is not None: val=''.join(x.text or '' for x in isn.iter(NS+'t'))
        elif v is not None: val=v.text or ''
        c[(row,col)]=val
    return c

ALL=[]; files=[]
for fn in sorted(glob.glob('*.xlsx')):
    if any(x in fn.lower() for x in EXCLUDE) or fn.startswith('_') or '_old' in fn.lower(): continue
    files.append(fn)
    z=zipfile.ZipFile(fn)
    sh=[''.join(t.text or '' for t in si.iter(NS+'t')) for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(NS+'si')] if 'xl/sharedStrings.xml' in z.namelist() else []
    rels={r.get('Id'):r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels')).iter(PR+'Relationship')}
    wb=ET.fromstring(z.read('xl/workbook.xml'))
    sheets=[(s.get('name'), rels[s.get(RNS+'id')]) for s in wb.iter(NS+'sheet')]
    area,loc=floc(fn)
    for name,path in sheets:
        try: cells=grid(z,path,sh)
        except Exception: continue
        if not cells: continue
        maxr=max(r for r,_ in cells)
        title=clean(cells.get((1,2)) or cells.get((1,1)) or '')
        title=re.sub(r'^Project ?/? ?(ECR)? ?:','',title,flags=re.I).strip()
        tasks=[]
        for r in range(3,maxr+1):
            tn=clean(cells.get((r,2),''))
            if not tn: continue
            resp=clean(cells.get((r,3),'')).strip('()')
            sd=pdate(cells.get((r,4),'')); fd=pdate(cells.get((r,6),'')); wl=pint(cells.get((r,8),''))
            if not resp and not sd and not fd and wl is None: continue
            tasks.append({'task':tn,'responsible':resp,'start':sd,'finish':fd,'leadtimeWeeks':pint(cells.get((r,5),'')),'inputs':clean(cells.get((r,7),'')).replace('\n',' / '),'workloadPct':wl})
        if not tasks: continue
        ALL.append({'file':fn,'area':area,'location':loc,'projectType':ptype(name),'name':title or name,'taskCount':len(tasks),'tasks':tasks})

json.dump(ALL, open('_projects.json','w'), indent=1, ensure_ascii=False)
al=Counter(p['area']+'/'+p['location'] for p in ALL)
ty=Counter(p['projectType'] for p in ALL)
L=[]
L.append('FILES PARSED (%d): %s'%(len(files), ', '.join(files)))
L.append('TOTAL PROJECTS: %d    TOTAL TASKS: %d'%(len(ALL), sum(p['taskCount'] for p in ALL)))
L.append('')
L.append('By area/location: '+', '.join('%s=%d'%(k,v) for k,v in sorted(al.items())))
L.append('By project type: '+', '.join('%s=%d'%(k,v) for k,v in sorted(ty.items())))
L.append('')
for p in ALL:
    L.append('  [%s/%s] %-22s %2dt  %s'%(p['area'],p['location'],p['projectType'],p['taskCount'],p['name'][:55]))
open('_scope.txt','w').write('\n'.join(L))
print('OK')
