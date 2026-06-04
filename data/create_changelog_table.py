#!/usr/bin/env python3
"""Create pth_changelog table + columns for the project/task History log."""
import json, urllib.request, urllib.error, os, time
B='https://org91869c7b.crm4.dynamics.com/api/data/v9.2'
ROOT=os.path.dirname(os.path.abspath(__file__))
TOKEN=open(os.path.join(ROOT,'..','.env.local')).read().split('VITE_DATAVERSE_TOKEN=')[1].split('\n')[0].strip()
H={'Authorization':'Bearer '+TOKEN,'Content-Type':'application/json','Accept':'application/json','OData-MaxVersion':'4.0','OData-Version':'4.0'}
def req(method,path,body=None):
    data=json.dumps(body).encode() if body is not None else None
    r=urllib.request.Request(B+'/'+path,data=data,headers=H,method=method)
    try:
        with urllib.request.urlopen(r,timeout=60) as resp:
            t=resp.read().decode(); return resp.status,(json.loads(t) if t else None)
    except urllib.error.HTTPError as e: return e.code,e.read().decode()[:300]
def lbl(s): return {"@odata.type":"Microsoft.Dynamics.CRM.Label","LocalizedLabels":[{"@odata.type":"Microsoft.Dynamics.CRM.LocalizedLabel","Label":s,"LanguageCode":1031}]}
def exists(p):
    r=urllib.request.Request(B+'/'+p,headers=H);
    try:
        urllib.request.urlopen(r,timeout=30); return True
    except urllib.error.HTTPError: return False

# 1. Create table (with primary name attribute)
if not exists("EntityDefinitions(LogicalName='pth_changelog')?$select=LogicalName"):
    payload={
        "@odata.type":"Microsoft.Dynamics.CRM.EntityMetadata",
        "SchemaName":"pth_ChangeLog","DisplayName":lbl("Change Log"),
        "DisplayCollectionName":lbl("Change Logs"),"Description":lbl("Audit trail of project & task changes"),
        "OwnershipType":"UserOwned","IsActivity":False,"HasActivities":False,"HasNotes":False,
        "PrimaryNameAttribute":"pth_name",
        "Attributes":[{
            "@odata.type":"Microsoft.Dynamics.CRM.StringAttributeMetadata",
            "SchemaName":"pth_Name","LogicalName":"pth_name","IsPrimaryName":True,
            "DisplayName":lbl("Summary"),"MaxLength":300,"FormatName":{"Value":"Text"},
            "RequiredLevel":{"Value":"ApplicationRequired"},
        }],
    }
    st,res=req('POST','EntityDefinitions',payload)
    print("create table:",st, res if st>=400 else "")
    time.sleep(8)
else:
    print("table already exists")

E="EntityDefinitions(LogicalName='pth_changelog')/Attributes"
def addstr(schema,label,ml=4000):
    log=schema.lower()
    if exists(f"EntityDefinitions(LogicalName='pth_changelog')/Attributes(LogicalName='{log}')?$select=LogicalName"):
        print('  ·',log,'exists'); return
    fmt={"Value":"TextArea"} if ml>850 else {"Value":"Text"}
    st,res=req('POST',E,{"@odata.type":"Microsoft.Dynamics.CRM.StringAttributeMetadata","SchemaName":schema,"DisplayName":lbl(label),"MaxLength":ml,"FormatName":fmt,"RequiredLevel":{"Value":"None"}})
    print('  +',log,st, res if st>=400 else '')
def adddt(schema,label):
    log=schema.lower()
    if exists(f"EntityDefinitions(LogicalName='pth_changelog')/Attributes(LogicalName='{log}')?$select=LogicalName"):
        print('  ·',log,'exists'); return
    st,res=req('POST',E,{"@odata.type":"Microsoft.Dynamics.CRM.DateTimeAttributeMetadata","SchemaName":schema,"DisplayName":lbl(label),"Format":"DateAndTime","DateTimeBehavior":{"Value":"UserLocal"},"RequiredLevel":{"Value":"None"}})
    print('  +',log,st, res if st>=400 else '')
def addlookup(schema,label,target):
    log=schema.lower()
    if exists(f"EntityDefinitions(LogicalName='pth_changelog')/Attributes(LogicalName='{log}')?$select=LogicalName"):
        print('  ·',log,'exists'); return
    rel={"@odata.type":"Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata",
         "SchemaName":f"{schema}_changelog_{target}",
         "ReferencedEntity":target,"ReferencingEntity":"pth_changelog",
         "CascadeConfiguration":{"Assign":"NoCascade","Delete":"RemoveLink","Merge":"NoCascade","Reparent":"NoCascade","Share":"NoCascade","Unshare":"NoCascade"},
         "Lookup":{"SchemaName":schema,"DisplayName":lbl(label),"RequiredLevel":{"Value":"None"}}}
    st,res=req('POST','RelationshipDefinitions',rel)
    print('  + lookup',log,st, res if st>=400 else '')

print("columns:")
addstr("pth_ChangeType","Change Type",100)
addstr("pth_Field","Field",200)
addstr("pth_OldValue","Old Value",4000)
addstr("pth_NewValue","New Value",4000)
addstr("pth_ChangedBy","Changed By",200)
addstr("pth_EntityKind","Entity Kind",50)   # 'Project' | 'Task'
adddt("pth_ChangedOn","Changed On")
addlookup("pth_Project","Project","pth_project")
addlookup("pth_Activity","Activity","pth_activity")
print("publishing...")
req('POST','PublishAllXml',{})
print("done")
