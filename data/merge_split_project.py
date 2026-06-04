#!/usr/bin/env python3
"""Merge the 6 wrongly-split 3E10177261 projects into one.
Keep the earliest-created as canonical, re-point other shells' activities to it,
delete the empty shells, and set the canonical title to the clean name.
"""
import json, urllib.request, urllib.error, os, time, urllib.parse
BASE='https://org91869c7b.crm4.dynamics.com/api/data/v9.2'
ROOT=os.path.dirname(os.path.abspath(__file__))
TOKEN=open(os.path.join(ROOT,'..','.env.local')).read().split('VITE_DATAVERSE_TOKEN=')[1].split('\n')[0].strip()
H={'Authorization':'Bearer '+TOKEN,'Accept':'application/json','OData-MaxVersion':'4.0','OData-Version':'4.0','Content-Type':'application/json; charset=utf-8'}
def req(method,path,body=None):
    url=path if path.startswith('http') else BASE+'/'+urllib.parse.quote(path,safe="?$&='()/,")
    data=json.dumps(body).encode() if body is not None else None
    for _ in range(4):
        try:
            r=urllib.request.Request(url,data=data,headers=H,method=method)
            with urllib.request.urlopen(r,timeout=30) as resp:
                t=resp.read().decode(); return resp.status,(json.loads(t) if t else None)
        except urllib.error.HTTPError as e: return e.code,e.read().decode()[:200]
        except Exception: time.sleep(2)
    return 0,'neterr'

CLEAN_TITLE='ECR: 3E10177261: PS-CC/NA: Reactivation of component'
st,res=req('GET',"pth_projects?$select=pth_projectid,pth_projectname&$filter=startswith(pth_projectname,'ECR: 3E10177261')&$top=50")
rows=res['value']
print('found',len(rows),'shells')
if len(rows)<2:
    print('nothing to merge'); raise SystemExit
canon=rows[0]['pth_projectid']
print('canonical:',canon)
# rename canonical
req('PATCH',f'pth_projects({canon})',{'pth_projectname':CLEAN_TITLE})
moved=deleted=0
for r in rows[1:]:
    pid=r['pth_projectid']
    # get its activities
    _,acts=req('GET',f"pth_activities?$select=pth_activityid&$filter=_pth_project_value eq {pid}&$top=200")
    for a in acts['value']:
        st2,_=req('PATCH',f"pth_activities({a['pth_activityid']})",{'pth_Project@odata.bind':f'/pth_projects({canon})'})
        if st2 in (200,204): moved+=1
    # delete the now-empty shell
    st3,_=req('DELETE',f'pth_projects({pid})')
    if st3 in (200,204): deleted+=1
print(f'moved {moved} tasks, deleted {deleted} shells')
# verify
_,v=req('GET',f"pth_activities?$select=pth_startdate,pth_enddate&$filter=_pth_project_value eq {canon}&$top=200")
ds=sorted([a['pth_startdate'][:10] for a in v['value'] if a.get('pth_startdate')])
de=sorted([a['pth_enddate'][:10] for a in v['value'] if a.get('pth_enddate')])
print('canonical now has',len(v['value']),'tasks, span',ds[0] if ds else '?','->',de[-1] if de else '?')
