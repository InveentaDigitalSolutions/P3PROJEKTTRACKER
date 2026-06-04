#!/usr/bin/env python3
"""Full clean reload from V3: clear Projects+Tasks, recreate all from _v3.json.
Resources + Task Templates untouched. File is source of truth."""
import json, urllib.request, urllib.error, os, time, urllib.parse
BASE='https://org91869c7b.crm4.dynamics.com/api/data/v9.2'
ROOT=os.path.dirname(os.path.abspath(__file__))
TOKEN=open(os.path.join(ROOT,'..','.env.local')).read().split('VITE_DATAVERSE_TOKEN=')[1].split('\n')[0].strip()
H={'Authorization':'Bearer '+TOKEN,'Accept':'application/json','OData-MaxVersion':'4.0','OData-Version':'4.0','Content-Type':'application/json; charset=utf-8'}
def req(method,path,body=None,prefer=None):
    url=path if path.startswith('http') else BASE+'/'+urllib.parse.quote(path,safe="?$&='()/,")
    data=json.dumps(body).encode() if body is not None else None
    h=dict(H)
    if prefer: h['Prefer']=prefer
    for _ in range(4):
        try:
            r=urllib.request.Request(url,data=data,headers=h,method=method)
            with urllib.request.urlopen(r,timeout=30) as resp:
                t=resp.read().decode(); return resp.status,(json.loads(t) if t else None)
        except urllib.error.HTTPError as e: return e.code,e.read().decode()[:160]
        except Exception: time.sleep(2)
    return 0,'neterr'

CATEGORY={'ECR':100000000,'Path Forward':100000001,'CIP':100000002,'New Programs':100000003}
SITE={'SlpP':100000000,'TlP':100000001}
ACT={'NOT_STARTED':100000000,'IN_PROGRESS':100000001,'BLOCKED':100000002,'CLOSED':100000003,'NA':100000004}
SOV={'green':100000000,'yellow':100000001,'red':100000002}
RYG_GREEN=100000002
def cat_for(c):
    c=(c or '').lower()
    if 'cip' in c: return 'CIP'
    if 'ecr' in c: return 'ECR'
    return 'New Programs'
def state_for(s):
    s=(s or '').strip().lower()
    if s=='done': return 'CLOSED'
    if s in ('on going','ongoing','in progress'): return 'IN_PROGRESS'
    if s in ('n/a','na','n.a.',''): return 'NA'
    return 'NOT_STARTED'

def clear():
    for setname,key in [('pth_activities','pth_activityid'),('pth_projects','pth_projectid')]:
        ids=[]; url=f"{setname}?$select={key}&$top=5000"
        while url:
            st,j=req('GET',url)
            if not isinstance(j,dict): break
            ids+=[r[key] for r in j['value']]
            nl=j.get('@odata.nextLink'); url=nl.replace(BASE+'/','') if nl else None
        print(f"clearing {setname}: {len(ids)}")
        for i,_id in enumerate(ids):
            req('DELETE',f"{setname}({_id})")
            if (i+1)%50==0: print(f"  {i+1}")

def main():
    clear()
    projects=json.load(open(os.path.join(ROOT,'_v3.json')))
    np=nt=0; log=[]
    for idx,p in enumerate(projects):
        starts=sorted([t['start'] for t in p['tasks'] if t['start']])
        finishes=sorted([t['finish'] for t in p['tasks'] if t['finish']])
        body={
            'pth_projectidexternal': f"P3V3-{p['area']}{p['location']}-{idx+1:03d}",
            'pth_projectname': p['title'][:200],
            'pth_category': CATEGORY[cat_for(p['category'])],
            'pth_sitelocation': SITE.get(p['location'],SITE['SlpP']),
            'pth_projectobjective': p['title'][:2000],
            'pth_projectmanagername':'Unassigned',
            'pth_sponsorexecutive':(p['sponsor'] or 'Unassigned')[:150],
            'pth_projectresponsible':(p['responsible'] or None),
            'pth_timestatus':RYG_GREEN,'pth_criticalpathchangedflag':False,
            'pth_projecttype':p['category'],
            'pth_plannedstartdate': starts[0] if starts else '2026-01-01',
            'pth_plannedenddate': finishes[-1] if finishes else (starts[-1] if starts else '2026-12-31'),
        }
        sov=SOV.get((p['statusOverview'] or '').strip().lower())
        if sov is not None: body['pth_statusoverview']=sov
        st,proj=req('POST','pth_projects',body,prefer='return=representation')
        if st not in (200,201): log.append(f"PROJ FAIL {p['title'][:30]} {st}: {proj}"); continue
        pid=proj['pth_projectid']; np+=1
        for t in p['tasks']:
            stt=state_for(t['status']); ab={
                'pth_activityidexternal':f"P3V3-{pid[:8]}-{nt}",
                'pth_name':t['task'][:200],'pth_owner':'Unassigned',
                'pth_status':ACT[stt],'pth_ryg':RYG_GREEN,'pth_iscriticalpath':False,
                'pth_responsible':(t['dept'] or None),
                'pth_Project@odata.bind':f"/pth_projects({pid})",
            }
            if t['start']: ab['pth_startdate']=t['start']
            end=t['finish'] or t['start']
            if end: ab['pth_enddate']=end; ab['pth_plannedenddate']=end
            if stt=='CLOSED' and end: ab['pth_closeddate']=end
            st2,_=req('POST','pth_activities',ab)
            if st2 in (200,201,204): nt+=1
        log.append(f"[{p['statusOverview']:6}] {p['title'][:42]} +{len(p['tasks'])}t")
    print('\n'.join(log))
    print(f"\nCREATED projects={np} tasks={nt}")

if __name__=='__main__': main()
