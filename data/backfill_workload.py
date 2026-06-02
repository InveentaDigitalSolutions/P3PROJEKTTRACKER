#!/usr/bin/env python3
"""Backfill pth_workloadpct on activities from matching task templates.
Match = normalized (projectType, task name) EXACT. Only sets where template
has a workload value and the activity currently has none.
"""
import json, urllib.request, urllib.error, re, os, time
BASE='https://org91869c7b.crm4.dynamics.com/api/data/v9.2'
ROOT=os.path.dirname(os.path.abspath(__file__))
TOKEN=open(os.path.join(ROOT,'..','.env.local')).read().split('VITE_DATAVERSE_TOKEN=')[1].split('\n')[0].strip()
H={'Authorization':'Bearer '+TOKEN,'Accept':'application/json','OData-MaxVersion':'4.0','OData-Version':'4.0','Content-Type':'application/json; charset=utf-8'}
def norm(s): return re.sub(r'[^a-z0-9]','',(s or '').lower())
def get(path):
    for _ in range(5):
        try:
            with urllib.request.urlopen(urllib.request.Request(BASE+'/'+path,headers=H),timeout=30) as r:
                return json.loads(r.read().decode())
        except Exception as e: time.sleep(2)
    raise RuntimeError('GET failed '+path)
def patch(entity,_id,body):
    for _ in range(5):
        try:
            r=urllib.request.Request(f"{BASE}/{entity}({_id})",data=json.dumps(body).encode(),headers=H,method='PATCH')
            with urllib.request.urlopen(r,timeout=30) as resp: return resp.status
        except urllib.error.HTTPError as e: return e.code
        except Exception: time.sleep(2)
    return 0

tmpl=get("pth_tasktemplates?$select=pth_projecttype,pth_name,pth_workloadpct&$top=500")['value']
tindex={}
for t in tmpl:
    w=t.get('pth_workloadpct')
    if w is not None:
        tindex[(norm(t['pth_projecttype']),norm(t['pth_name']))]=w
projs={p['pth_projectid']:p for p in get("pth_projects?$select=pth_projectid,pth_projecttype&$top=200")['value']}
acts=get("pth_activities?$select=pth_activityid,pth_name,pth_workloadpct,_pth_project_value&$top=1000")['value']

updated=skipped=nomatch=0
for a in acts:
    if a.get('pth_workloadpct') is not None: skipped+=1; continue
    ptype=projs.get(a.get('_pth_project_value'),{}).get('pth_projecttype','')
    w=tindex.get((norm(ptype),norm(a['pth_name'])))
    if w is None: nomatch+=1; continue
    st=patch('pth_activities',a['pth_activityid'],{'pth_workloadpct':w})
    if st in (200,204): updated+=1
    else: print('FAIL',a['pth_name'][:30],st)
print(f"updated={updated} already_had={skipped} no_match={nomatch} total={len(acts)}")
