#!/usr/bin/env python3
"""Import projects+tasks from _concentrate.json into Dataverse via Web API (urllib).
Reads token from ../.env.local. Idempotent only if tables are empty.
"""
import json, urllib.request, urllib.error, sys, os

BASE='https://org91869c7b.crm4.dynamics.com/api/data/v9.2'
ROOT=os.path.dirname(os.path.abspath(__file__))
TOKEN=open(os.path.join(ROOT,'..','.env.local')).read().split('VITE_DATAVERSE_TOKEN=')[1].split('\n')[0].strip()
H={'Authorization':'Bearer '+TOKEN,'Accept':'application/json','OData-MaxVersion':'4.0','OData-Version':'4.0','Content-Type':'application/json; charset=utf-8'}

import time
def req(method, path, body=None, prefer=None, retries=5):
    url=path if path.startswith('http') else BASE+'/'+path
    data=json.dumps(body).encode() if body is not None else None
    h=dict(H)
    if prefer: h['Prefer']=prefer
    last=None
    for attempt in range(retries):
        r=urllib.request.Request(url, data=data, headers=h, method=method)
        try:
            with urllib.request.urlopen(r, timeout=30) as resp:
                t=resp.read().decode()
                return resp.status, (json.loads(t) if t else None)
        except urllib.error.HTTPError as e:
            return e.code, e.read().decode()[:300]
        except Exception as e:
            last=str(e); time.sleep(2*(attempt+1))   # backoff on network reset
    return 0, 'NETERR: '+str(last)

# choice maps
CATEGORY={'ECR':100000000,'Path Forward':100000001,'CIP':100000002,'New Programs':100000003}
SITE={'SlpP':100000000,'TlP':100000001}   # values relabeled to SlpP/TlP
# task status (Dataverse pth_status on pth_activities)
ACT={'NOT_STARTED':100000000,'IN_PROGRESS':100000001,'BLOCKED':100000002,'CLOSED':100000003,'NA':100000004}
RYG_GREEN=100000002

def cat_for(c):
    c=(c or '').lower()
    if 'cip' in c: return 'CIP'
    if 'ecr' in c: return 'ECR'
    return 'New Programs'   # New Program, Relocation, MSE Line, Subassembly, etc.

def state_for(status):
    s=(status or '').strip().lower()
    if s=='done': return 'CLOSED'
    if s in ('on going','ongoing','in progress'): return 'IN_PROGRESS'
    if s in ('n/a','na','n.a.',''): return 'NA'   # not applicable — won't be done
    return 'NOT_STARTED'   # plan

def main():
    projects=json.load(open(os.path.join(ROOT,'_concentrate.json')))
    np=nt=0
    log=[]
    for idx,p in enumerate(projects):
        title=p['title'] or 'Untitled'
        cat=cat_for(p['category'])
        # planned dates from task min/max
        starts=sorted([t['start'] for t in p['tasks'] if t['start']])
        finishes=sorted([t['finish'] for t in p['tasks'] if t['finish']])
        pstart=starts[0] if starts else '2026-01-01'
        pend=finishes[-1] if finishes else (starts[-1] if starts else '2026-12-31')
        body={
            'pth_projectidexternal': f"P3-{p['area']}{p['location']}-{idx+1:03d}",
            'pth_projectname': title[:200],
            'pth_category': CATEGORY[cat],
            'pth_sitelocation': SITE.get(p['location'], SITE['SlpP']),
            'pth_projectobjective': title[:2000],
            'pth_projectmanagername': 'Unassigned',
            'pth_timestatus': RYG_GREEN,
            'pth_criticalpathchangedflag': False,
            'pth_projecttype': p['category'],
            'pth_plannedstartdate': pstart,
            'pth_plannedenddate': pend,
        }
        st,proj=req('POST','pth_projects',body,prefer='return=representation')
        if st not in (200,201):
            log.append(f"PROJECT FAIL [{title[:30]}] {st}: {proj}"); continue
        pid=proj['pth_projectid']; np+=1
        for t in p['tasks']:
            stt=state_for(t['status'])
            # REAL dates only — never fabricate. N/A (and any dateless) tasks carry no schedule.
            start=t['start']
            end=t['finish'] or t['start']   # if only a start exists, use it as end
            ab={
                'pth_activityidexternal': f"P3-{pid[:8]}-{nt}",
                'pth_name': t['task'][:200],
                'pth_owner': 'Unassigned',
                'pth_status': ACT[stt],
                'pth_ryg': RYG_GREEN,
                'pth_iscriticalpath': False,
                'pth_responsible': (t['dept'] or None),
                'pth_Project@odata.bind': f"/pth_projects({pid})",
            }
            if start:
                ab['pth_startdate']=start
            if end:
                ab['pth_enddate']=end
                ab['pth_plannedenddate']=end
            if stt=='CLOSED' and end:
                ab['pth_closeddate']=end
            st2,_=req('POST','pth_activities',ab)
            if st2 in (200,201,204): nt+=1
            else: log.append(f"  TASK FAIL [{t['task'][:25]}] {st2}: {_}")
        log.append(f"[{p['area']}/{p['location']}] {title[:42]} — {len(p['tasks'])} tasks")
    print('\n'.join(log))
    print(f"\nIMPORTED projects={np} tasks={nt}")

if __name__=='__main__':
    main()
