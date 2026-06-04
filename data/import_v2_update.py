#!/usr/bin/env python3
"""Update-in-place import of V2 fields.
- Existing project (matched by name): PATCH sponsor/responsible/statusOverview. Tasks untouched.
- New project: create project + its tasks.
Reads token from ../.env.local.
"""
import json, urllib.request, urllib.error, os, time, re
BASE='https://org91869c7b.crm4.dynamics.com/api/data/v9.2'
ROOT=os.path.dirname(os.path.abspath(__file__))
TOKEN=open(os.path.join(ROOT,'..','.env.local')).read().split('VITE_DATAVERSE_TOKEN=')[1].split('\n')[0].strip()
H={'Authorization':'Bearer '+TOKEN,'Accept':'application/json','OData-MaxVersion':'4.0','OData-Version':'4.0','Content-Type':'application/json; charset=utf-8'}
def req(method,path,body=None,prefer=None):
    url=path if path.startswith('http') else BASE+'/'+path
    data=json.dumps(body).encode() if body is not None else None
    h=dict(H)
    if prefer: h['Prefer']=prefer
    for _ in range(4):
        try:
            r=urllib.request.Request(url,data=data,headers=h,method=method)
            with urllib.request.urlopen(r,timeout=30) as resp:
                t=resp.read().decode(); return resp.status,(json.loads(t) if t else None)
        except urllib.error.HTTPError as e: return e.code, e.read().decode()[:200]
        except Exception as e: time.sleep(2)
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
def norm(s): return re.sub(r'\s+',' ',(s or '').strip().lower())

def main():
    projects=json.load(open(os.path.join(ROOT,'_v2.json')))
    # index existing projects by normalized name
    _,existing=req('GET','pth_projects?$select=pth_projectid,pth_projectname&$top=500')
    byname={norm(p['pth_projectname']):p['pth_projectid'] for p in existing['value']}
    patched=created=ntasks=0; log=[]
    for idx,p in enumerate(projects):
        title=p['title']
        fields={
            'pth_sponsorexecutive': (p['sponsor'] or 'Unassigned')[:150],
            'pth_projectresponsible': (p['responsible'] or None),
        }
        sov=SOV.get((p['statusOverview'] or '').strip().lower())
        if sov is not None: fields['pth_statusoverview']=sov
        pid=byname.get(norm(title))
        if pid:
            st,_=req('PATCH',f'pth_projects({pid})',fields)
            if st in (200,204): patched+=1; log.append(f"PATCH  {title[:45]}  status={p['statusOverview']}")
            else: log.append(f"PATCH FAIL {title[:30]} {st}: {_}")
        else:
            # new project
            starts=sorted([t['start'] for t in p['tasks'] if t['start']])
            finishes=sorted([t['finish'] for t in p['tasks'] if t['finish']])
            body={
                'pth_projectidexternal': f"P3V2-{p['area']}{p['location']}-{idx+1:03d}",
                'pth_projectname': title[:200],
                'pth_category': CATEGORY[cat_for(p['category'])],
                'pth_sitelocation': SITE.get(p['location'],SITE['SlpP']),
                'pth_projectobjective': title[:2000],
                'pth_projectmanagername':'Unassigned',
                'pth_timestatus':RYG_GREEN,'pth_criticalpathchangedflag':False,
                'pth_projecttype':p['category'],
                'pth_plannedstartdate': starts[0] if starts else '2026-01-01',
                'pth_plannedenddate': finishes[-1] if finishes else (starts[-1] if starts else '2026-12-31'),
            }
            body.update(fields)
            st,proj=req('POST','pth_projects',body,prefer='return=representation')
            if st not in (200,201): log.append(f"CREATE FAIL {title[:30]} {st}: {proj}"); continue
            npid=proj['pth_projectid']; created+=1
            for t in p['tasks']:
                stt=state_for(t['status']); ab={
                    'pth_activityidexternal': f"P3V2-{npid[:8]}-{ntasks}",
                    'pth_name':t['task'][:200],'pth_owner':'Unassigned',
                    'pth_status':ACT[stt],'pth_ryg':RYG_GREEN,'pth_iscriticalpath':False,
                    'pth_responsible':(t['dept'] or None),
                    'pth_Project@odata.bind':f"/pth_projects({npid})",
                }
                if t['start']: ab['pth_startdate']=t['start']
                end=t['finish'] or t['start']
                if end: ab['pth_enddate']=end; ab['pth_plannedenddate']=end
                if stt=='CLOSED' and end: ab['pth_closeddate']=end
                st2,_=req('POST','pth_activities',ab)
                if st2 in (200,201,204): ntasks+=1
            log.append(f"CREATE {title[:42]} +{len(p['tasks'])} tasks")
    print('\n'.join(log))
    print(f"\nPATCHED={patched} CREATED={created} new_tasks={ntasks}")

if __name__=='__main__': main()
