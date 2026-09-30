from __future__ import annotations
from collections import defaultdict
from datetime import datetime, date
from io import BytesIO
import math, re, unicodedata
import openpyxl
import pandas as pd

MAX_DISTANCE_MILES = 130.0
ACTIVE_REGISTRATION_STATUSES = {'enrolled','registered','in progress','not started','active'}

ALIASES = {
    'employee id':['employee id','employee_id','employee number'],
    'wid':['wid','worker wid','employee wid'],
    'worker':['worker','worker name','learning participant'],
    'hire date':['hire date','most recent hire date'],
    'candidate position id':['candidate position id'],
    'new hire position title':['new hire position title','position title'],
    'candidate cost center id':['candidate cost center id','cost center id'],
    'candidate department name':['candidate department name','cost center'],
    'supervisory organization id':['supervisory organization id','sup org id'],
    'worker physical location':['worker physical location','location'],
    'worker type':['worker type'],
    'email - primary work':['email - primary work','work email'],
    'email - primary home':['email - primary home','home email'],
    'hiring manager ad':['hiring manager ad (username)','hiring manager ad','hiring manager'],
    'effective date':['effective date','position effective date','staffing event effective date'],
    'job code - proposed':['job code - proposed','proposed job code'],
    'job title - proposed':['job title - proposed','proposed job title'],
    'cost center id - requested':['cost center id - requested','requested cost center id'],
    'cost center - proposed':['cost center - proposed','proposed cost center'],
    'sup org id - requested':['sup org id - requested','requested sup org id'],
    'location - proposed':['location - proposed','proposed location'],
    'candidate email':['candidate email','email - primary work'],
    'record learning content':['record learning content','learning content','learning content title'],
    'record completion status':['record completion status','completion status'],
    'record completion date':['record completion date','completion date','completed date'],
    'registration status':['registration status','learner registration status'],
    "learner's registration date":["learner's registration date",'learners registration date','registration date'],
    'enrolled course offering':['enrolled course offering','course offering'],
    'start date':['start date','session start date','offering start date'],
    'end date':['end date','session end date','offering end date'],
    'title':['title','learning content title'],
    'reference id':['reference id','reference_id'],
    'available seats':['available seats','seats available'],
    'availability status':['availability status','status'],
    'locations':['locations','location'],
}

def norm(v):
    if v is None:return ''
    s=unicodedata.normalize('NFKC',str(v)).replace('\u00a0',' ').replace('\r',' ').replace('\n',' ')
    s=s.replace('’',"'").replace('‘',"'")
    return re.sub(r'\s+',' ',s.strip().lower())

def clean(v): return '' if v is None else str(v).strip()

def id_norm(v):
    if v is None:return ''
    if isinstance(v,float) and v.is_integer():return str(int(v))
    return str(v).strip()

def dt(v):
    if isinstance(v,datetime):return v
    if isinstance(v,date):return datetime(v.year,v.month,v.day)
    if isinstance(v,str) and v.strip():
        s=v.strip()
        for f in ('%m/%d/%Y %I:%M %p','%m/%d/%Y %H:%M:%S','%Y-%m-%d %H:%M:%S','%m/%d/%Y','%Y-%m-%d'):
            try:return datetime.strptime(s,f)
            except:pass
    return None

def _match_header(v, canonical):
    nv=norm(v)
    return nv in ({norm(canonical)}|{norm(x) for x in ALIASES.get(canonical,[])})

def read_sheet(source, required_headers, optional_headers=None, scan_rows=150):
    if isinstance(source,(bytes,bytearray)):raw=bytes(source)
    elif hasattr(source,'getvalue'):raw=bytes(source.getvalue())
    elif isinstance(source,(str,bytes)): raw=open(source,'rb').read()
    else: source.seek(0);raw=source.read()
    if not raw:raise ValueError('Uploaded workbook is empty.')
    wb=openpyxl.load_workbook(BytesIO(raw),read_only=False,data_only=True)
    found=None
    for ws in wb.worksheets:
        for ridx,row in enumerate(ws.iter_rows(values_only=True),1):
            if ridx>scan_rows:break
            vals=list(row)
            matched=sum(1 for h in required_headers if any(_match_header(v,h) for v in vals))
            if matched==len(required_headers):found=(ws,ridx,vals);break
        if found:break
    if not found:raise ValueError(f'Could not recognize workbook. Expected fields: {required_headers}. Worksheets: {wb.sheetnames}')
    ws,hr,heads=found; lookup={}
    for c,v in enumerate(heads):
        for h in required_headers+(optional_headers or []):
            if h not in lookup and _match_header(v,h):lookup[h]=c;break
    out=[]
    for row in ws.iter_rows(min_row=hr+1,values_only=True):
        if not any(v not in (None,'') for v in row):continue
        out.append({h:(row[i] if i<len(row) else None) for h,i in lookup.items()})
    return out

def haversine_miles(lat1,lon1,lat2,lon2):
    R=3958.7613;p1=math.radians(lat1);p2=math.radians(lat2);dlat=math.radians(lat2-lat1);dlon=math.radians(lon2-lon1)
    a=math.sin(dlat/2)**2+math.cos(p1)*math.cos(p2)*math.sin(dlon/2)**2
    return 2*R*math.asin(math.sqrt(a))

def interval_overlap(a,b,c,d):return a<d and b>c

def person_key(wid,eid):return f'WID:{id_norm(wid)}' if id_norm(wid) else f'EID:{id_norm(eid)}'

def run_scheduler(new_hire_file, job_change_file, history_file, sessions_file, rules_df, locations_df, equivalencies_df, orientation_file):
    nh=read_sheet(new_hire_file,['employee id','wid','hire date','candidate cost center id','candidate position id','worker'],['worker type','new hire position title','candidate department name','supervisory organization id','worker physical location','hiring manager ad','email - primary work','email - primary home']) if new_hire_file else []
    jc=read_sheet(job_change_file,['employee id','wid','effective date','job code - proposed','cost center id - requested','worker'],['job title - proposed','cost center - proposed','sup org id - requested','location - proposed','worker type','hiring manager ad','candidate email']) if job_change_file else []
    hist=read_sheet(history_file,['employee id','record learning content','record completion status'],['wid','record completion date','registration status',"learner's registration date"]) 
    sessions=read_sheet(sessions_file,['title','reference id','start date','end date','available seats','wid'],['availability status','locations'])
    orientation=read_sheet(orientation_file,['employee id','enrolled course offering','registration status','start date','end date'],['wid','locations'])

    locmap={}
    for _,r in locations_df.iterrows():
        name=clean(r.get('location'))
        if not name:continue
        try:lat=float(r.get('latitude'));lon=float(r.get('longitude'))
        except:lat=lon=None
        locmap[norm(name)]={'name':name,'lat':lat,'lon':lon}
    def distance(a,b):
        x=locmap.get(norm(a));y=locmap.get(norm(b))
        if not x or not y or x['lat'] is None or y['lat'] is None:return None
        return haversine_miles(x['lat'],x['lon'],y['lat'],y['lon'])

    # Directional equivalencies. Blank/missing direction is ONE_WAY for compatibility.
    eq=defaultdict(set); eq_meta={}
    for _,r in equivalencies_df.iterrows():
        required=clean(r.get('required_training_title')); equivalent=clean(r.get('equivalent_training_title'))
        a=norm(required);b=norm(equivalent)
        if not a or not b:continue
        direction=norm(r.get('relationship_direction') or 'ONE_WAY').replace('-','_').replace(' ','_')
        two=direction in ('two_way','2_way','2')
        eq[a].add(b);eq_meta[(a,b)]=(required,equivalent,'TWO_WAY' if two else 'ONE_WAY','FORWARD')
        if two:eq[b].add(a);eq_meta[(b,a)]=(required,equivalent,'TWO_WAY','REVERSE')
    def satisfying_titles(title):return {norm(title)}|set(eq.get(norm(title),set()))

    events=[]
    for i,r in enumerate(nh,1):
        eid=id_norm(r.get('employee id'));wid=id_norm(r.get('wid'))
        events.append({'Event_Key':f'NH-{i}-{eid}','Event_Type':'NEW_HIRE','Employee_ID':eid,'WID':wid,'Person_Key':person_key(wid,eid),'Worker_Name':clean(r.get('worker')),'Worker_Type':clean(r.get('worker type')),'Anchor_Date':dt(r.get('hire date')),'Job_Code':clean(r.get('candidate position id')),'Position_Title':clean(r.get('new hire position title')),'Cost_Center_ID':clean(r.get('candidate cost center id')),'Cost_Center_Title':clean(r.get('candidate department name')),'Sup_Org_ID':clean(r.get('supervisory organization id')),'Physical_Location':clean(r.get('worker physical location')),'Hiring_Manager_AD':clean(r.get('hiring manager ad')),'Work_Email':clean(r.get('email - primary work')),'Home_Email':clean(r.get('email - primary home'))})
    for i,r in enumerate(jc,1):
        eid=id_norm(r.get('employee id'));wid=id_norm(r.get('wid'))
        events.append({'Event_Key':f'JC-{i}-{eid}','Event_Type':'JOB_CHANGE','Employee_ID':eid,'WID':wid,'Person_Key':person_key(wid,eid),'Worker_Name':clean(r.get('worker')),'Worker_Type':clean(r.get('worker type')),'Anchor_Date':dt(r.get('effective date')),'Job_Code':clean(r.get('job code - proposed')),'Position_Title':clean(r.get('job title - proposed')),'Cost_Center_ID':clean(r.get('cost center id - requested')),'Cost_Center_Title':clean(r.get('cost center - proposed')),'Sup_Org_ID':clean(r.get('sup org id - requested')),'Physical_Location':clean(r.get('location - proposed')),'Hiring_Manager_AD':clean(r.get('hiring manager ad')),'Work_Email':clean(r.get('candidate email')),'Home_Email':''})

    # Collapse exact duplicate staffing events BEFORE requirement generation.
    canonical=[];dups=[];seen_events={}
    for ev in events:
        sig=(ev['Person_Key'],norm(ev['Event_Type']),ev['Anchor_Date'],norm(ev['Job_Code']),norm(ev['Position_Title']),norm(ev['Cost_Center_ID']),norm(ev['Sup_Org_ID']),norm(ev['Physical_Location']))
        if sig in seen_events:
            d=dict(ev);d['Canonical_Event_Key']=seen_events[sig]['Event_Key'];dups.append(d)
        else:
            ev['Canonical_Event_Key']=ev['Event_Key'];seen_events[sig]=ev;canonical.append(ev)
    events=canonical

    rules=[]
    for _,r in rules_df.iterrows():
        title=clean(r.get('training_title'))
        if not title:continue
        rules.append({'Rule_ID':r.get('id'),'Job_Code':clean(r.get('job_code')),'Cost_Center':clean(r.get('cost_center')),'Sup_Org':clean(r.get('supervisory_org')),'Training_Title':title,'Prerequisite':clean(r.get('prerequisite')),'Scheduling_Policy':clean(r.get('scheduling_policy')),'Timing_Modifier':clean(r.get('timing_modifier')),'Timing_Min':r.get('timing_min'),'Timing_Max':r.get('timing_max'),'Priority':r.get('priority')})

    reqs=[];docs_found=0;docs_missing=0
    for ev in events:
        matches=[r for r in rules if norm(r['Job_Code'])==norm(ev['Job_Code']) and norm(r['Cost_Center'])==norm(ev['Cost_Center_ID']) and (not r['Sup_Org'] or norm(r['Sup_Org'])==norm(ev['Sup_Org_ID']))]
        if matches:
            docs_found+=1
            reqs.extend({**ev,**m,'Training_Documentation_Status':'TRAINING_DOCUMENTATION_FOUND'} for m in matches)
        else:
            docs_missing+=1
            reqs.append({**ev,'Rule_ID':'','Training_Title':'','Prerequisite':'','Scheduling_Policy':'','Timing_Modifier':'','Timing_Min':'','Timing_Max':'','Priority':'','Training_Documentation_Status':'NO_TRAINING_DOCUMENTATION_FOUND'})

    # Group Available Sessions rows sharing WID into one offering.
    groups=defaultdict(list)
    for s in sessions:
        key=id_norm(s.get('wid')) or f"{clean(s.get('reference id'))}|{clean(s.get('start date'))}|{norm(s.get('title'))}"
        groups[key].append(s)
    offerings=[];by_title=defaultdict(list)
    for key,rows in groups.items():
        rows=sorted(rows,key=lambda x:dt(x.get('start date')) or datetime.max)
        intervals=[(dt(r.get('start date')),dt(r.get('end date'))) for r in rows if dt(r.get('start date')) and dt(r.get('end date')) and dt(r.get('end date'))>dt(r.get('start date'))]
        if not intervals:continue
        title=clean(rows[0].get('title'));location=next((clean(r.get('locations')) for r in rows if clean(r.get('locations'))),'')
        refs=[clean(r.get('reference id')) for r in rows if clean(r.get('reference id'))]
        seats=max([0]+[int(float(r.get('available seats') or 0)) for r in rows])
        statuses={norm(r.get('availability status')) for r in rows if clean(r.get('availability status'))}
        off={'Session_Key':key,'WID':id_norm(rows[0].get('wid')),'Title':title,'Reference_ID':refs[0] if refs else '', 'Location':location,'Intervals':intervals,'First_Start':min(a for a,b in intervals),'Last_End':max(b for a,b in intervals),'Day_Count':len(intervals),'Multi_Day':len({a.date() for a,b in intervals})>1,'Seats':seats,'Status':'open' if not statuses or statuses=={'open'} else next(iter(statuses))}
        offerings.append(off);by_title[norm(title)].append(off)
    opening={o['Session_Key']:o['Seats'] for o in offerings};remaining=dict(opening)

    # Existing Workday sessions: display every active enrollment and block its intervals.
    eid_to_wid={id_norm(h.get('employee id')):id_norm(h.get('wid')) for h in hist if id_norm(h.get('employee id')) and id_norm(h.get('wid'))}
    for ev in events:
        if ev['Employee_ID'] and ev['WID']:eid_to_wid[ev['Employee_ID']]=ev['WID']
    existing=[];busy=defaultdict(list);existing_seen=set()
    for o in orientation:
        if norm(o.get('registration status')) not in ACTIVE_REGISTRATION_STATUSES:continue
        st=dt(o.get('start date'));en=dt(o.get('end date'));eid=id_norm(o.get('employee id'))
        if not st or not en or en<=st:continue
        pk=person_key(eid_to_wid.get(eid,''),eid);title=clean(o.get('enrolled course offering'));key=(pk,norm(title),st,en)
        if key not in existing_seen:
            existing_seen.add(key);existing.append({'Person_Key':pk,'WID':eid_to_wid.get(eid,''),'Employee_ID':eid,'Training_Title':title,'Registration_Status':clean(o.get('registration status')),'Start_Date':st,'End_Date':en,'Location':clean(o.get('locations'))})
        busy[pk].append({'start':st,'end':en,'title':title,'source':'EXISTING'})
    person_req_titles=defaultdict(set)
    for r in reqs:
        if r.get('Training_Title'):person_req_titles[r['Person_Key']] |= satisfying_titles(r['Training_Title'])
    for e in existing:
        e['Required_For_Current_Role']='Yes' if norm(e['Training_Title']) in person_req_titles[e['Person_Key']] else 'No'
        e['Scheduling_Impact']='BLOCKS_OVERLAP_CHECK'

    # Current identity/history helper.
    hist_by_person=defaultdict(list)
    for h in hist:
        hist_by_person[person_key(h.get('wid') or eid_to_wid.get(id_norm(h.get('employee id')),''),h.get('employee id'))].append(h)
    for r in reqs:
        defaults={'Completion_Date':'','Completion_Training':'','Existing_Enrollment_Training':'','Existing_Registration_Date':'','Existing_Session_Start':'','Equivalency_Used':'','Equivalency_Direction':'','Equivalency_Match_Direction':'','Selected_Session_WID':'','Selected_Reference_ID':'','Selected_Start':'','Selected_End':'','Selected_Location':'','Distance_Miles':'','Selected_Day_Count':'','Selected_Multi_Day':'','Selected_Session_Days':'','Original_Available_Seats':'','Remaining_Seats_After_Reservation':'','Selection_Order':'','Selection_Phase':'','Conflict_Detail':'','Disposition':'PENDING_SCHEDULING','Explanation':'','Workday_Ready':'No','Satisfied_By_Event_Key':'','Satisfied_By_Session_WID':'','Prerequisite_Status':''}
        r.update(defaults)
        if r['Training_Documentation_Status']=='NO_TRAINING_DOCUMENTATION_FOUND':
            r['Disposition']='NO_TRAINING_DOCUMENTATION_FOUND';r['Explanation']=f"No training documentation was found for Job Code '{r['Job_Code']}', Cost Center '{r['Cost_Center_ID']}'"+(f", Supervisory Organization '{r['Sup_Org_ID']}'" if r['Sup_Org_ID'] else '')+'. No training was scheduled.';continue
        titles=satisfying_titles(r['Training_Title']);ph=hist_by_person[r['Person_Key']]
        completions=[h for h in ph if norm(h.get('record learning content')) in titles and norm(h.get('record completion status'))=='completed']
        if completions:
            h=max(completions,key=lambda x:dt(x.get('record completion date')) or datetime.min);obs=clean(h.get('record learning content'));r['Completion_Date']=h.get('record completion date');r['Completion_Training']=obs
            em=eq_meta.get((norm(r['Training_Title']),norm(obs)));r['Disposition']='PREVIOUSLY_COMPLETED' if norm(obs)==norm(r['Training_Title']) else 'EQUIVALENT_COMPLETION'
            if em:r['Equivalency_Used']=f"{em[1]} -> {em[0]}" if em[3]=='FORWARD' else f"{em[0]} <-> {em[1]}";r['Equivalency_Direction']=em[2];r['Equivalency_Match_Direction']=em[3]
            r['Explanation']= 'Historical completion satisfies requirement. No reassignment required.' if norm(obs)==norm(r['Training_Title']) else f"Historical completion of equivalent training '{obs}' satisfies '{r['Training_Title']}'. No reassignment required.";continue
        enroll=[h for h in ph if norm(h.get('record learning content')) in titles and norm(h.get('registration status')) in ACTIVE_REGISTRATION_STATUSES]
        ext=[e for e in existing if e['Person_Key']==r['Person_Key'] and norm(e['Training_Title']) in titles]
        if enroll or ext:
            if ext:
                e=min(ext,key=lambda x:x['Start_Date']);obs=e['Training_Title'];r['Existing_Enrollment_Training']=obs;r['Existing_Session_Start']=e['Start_Date']
            else:
                h=max(enroll,key=lambda x:dt(x.get("learner's registration date")) or datetime.min);obs=clean(h.get('record learning content'));r['Existing_Enrollment_Training']=obs;r['Existing_Registration_Date']=h.get("learner's registration date")
            em=eq_meta.get((norm(r['Training_Title']),norm(obs)));r['Disposition']='ALREADY_ENROLLED'
            if em:r['Equivalency_Used']=f"{em[1]} -> {em[0]}" if em[3]=='FORWARD' else f"{em[0]} <-> {em[1]}";r['Equivalency_Direction']=em[2];r['Equivalency_Match_Direction']=em[3]
            r['Explanation']='Existing active registration found. Duplicate enrollment suppressed.' if norm(obs)==norm(r['Training_Title']) else f"Existing active enrollment in equivalent training '{obs}' satisfies '{r['Training_Title']}'. Duplicate enrollment suppressed.";continue
        if norm(r['Scheduling_Policy'])=='flag_manual':r['Disposition']='MANUAL_SCHEDULING_REQUIRED';r['Explanation']='Training rule uses FLAG_MANUAL.'

    # Global person/course assignment cache ensures the same course is not enrolled twice in one run.
    same_run={};selection_order=0
    by_event=defaultdict(dict)
    for r in reqs:by_event[r['Event_Key']][norm(r['Training_Title'])]=r

    def priority_val(r):
        try:return float(r.get('Priority')) if r.get('Priority') not in (None,'') else float('inf')
        except:return float('inf')
    def ready_requirements(event_key):
        out=[]
        for r in reqs:
            if r['Event_Key']!=event_key or r['Disposition']!='PENDING_SCHEDULING':continue
            bound=None;blocked=False;pr=norm(r.get('Prerequisite'))
            if pr:
                p=by_event[event_key].get(pr)
                if p:
                    if p['Disposition']=='PROPOSED_SCHEDULE':bound=dt(p.get('Selected_Start'));r['Prerequisite_Status']='SCHEDULED_EARLIER'
                    elif p['Disposition'] in ('PREVIOUSLY_COMPLETED','EQUIVALENT_COMPLETION'):r['Prerequisite_Status']='SATISFIED_BY_COMPLETION'
                    elif p['Disposition']=='ALREADY_ENROLLED' and p.get('Existing_Session_Start'):bound=dt(p.get('Existing_Session_Start'));r['Prerequisite_Status']='SATISFIED_BY_EXISTING_ENROLLMENT'
                    elif p['Disposition']=='SATISFIED_BY_SAME_RUN_ASSIGNMENT' and p.get('Selected_Start'):bound=dt(p.get('Selected_Start'));r['Prerequisite_Status']='SCHEDULED_EARLIER'
                    else:blocked=True
                else:blocked=True
            if not blocked:out.append((r,bound))
        return out

    for ev in sorted(events,key=lambda e:(e.get('Anchor_Date') or datetime.max,e['Event_Key'])):
        while True:
            pending=[r for r in reqs if r['Event_Key']==ev['Event_Key'] and r['Disposition']=='PENDING_SCHEDULING']
            if not pending:break
            ready=ready_requirements(ev['Event_Key'])
            if not ready:
                for r in pending:r['Disposition']='REVIEW_REQUIRED';r['Explanation']=f"Prerequisite '{r.get('Prerequisite')}' could not be resolved before scheduling." if r.get('Prerequisite') else 'No schedulable session could be resolved.'
                break
            evaluated=[]
            for r,bound in ready:
                course=(r['Person_Key'],norm(r['Training_Title']))
                if course in same_run:
                    p=same_run[course];ps=dt(p.get('Selected_Start'));valid=bool(ps and timing_ok(r,ps) and (not bound or ps>bound))
                    if valid:
                        for f in ('Selected_Session_WID','Selected_Reference_ID','Selected_Start','Selected_End','Selected_Location','Distance_Miles','Selected_Day_Count','Selected_Multi_Day','Selected_Session_Days','Selection_Order','Selection_Phase'):r[f]=p.get(f,'')
                        r['Disposition']='SATISFIED_BY_SAME_RUN_ASSIGNMENT';r['Satisfied_By_Event_Key']=p['Event_Key'];r['Satisfied_By_Session_WID']=p['Selected_Session_WID'];r['Explanation']=f"Same person/course was already assigned under staffing event {p['Event_Key']}; one enrollment satisfies both requirements."
                    else:r['Disposition']='REVIEW_REQUIRED';r['Explanation']='A prior same-run assignment exists, but its session does not meet this requirement timing/prerequisite rule. A duplicate enrollment was not created.'
                    continue
                offs=[]
                for o in by_title.get(norm(r['Training_Title']),[]):
                    if o['Status']!='open' or remaining.get(o['Session_Key'],0)<=0:continue
                    if r.get('Anchor_Date') and o['First_Start']<r['Anchor_Date']:continue
                    if bound and o['First_Start']<=bound:continue
                    if not timing_ok(r,o['First_Start']):continue
                    cf=any(interval_overlap(st,en,b['start'],b['end']) for st,en in o['Intervals'] for b in busy[r['Person_Key']])
                    if cf:continue
                    offs.append(o)
                if not offs:
                    # Diagnose conflict for audit.
                    all_possible=[]
                    for o in by_title.get(norm(r['Training_Title']),[]):
                        if o['Status']=='open' and remaining.get(o['Session_Key'],0)>0 and (not r.get('Anchor_Date') or o['First_Start']>=r['Anchor_Date']) and (not bound or o['First_Start']>bound) and timing_ok(r,o['First_Start']):
                            all_possible.append(o)
                    if all_possible:
                        labels=[]
                        for o in all_possible:
                            for b in busy[r['Person_Key']]:
                                if any(interval_overlap(st,en,b['start'],b['end']) for st,en in o['Intervals']):
                                    labels.append(f"{b['title']} ({b['start']:%m/%d/%Y %I:%M %p}-{b['end']:%I:%M %p})")
                        if labels:r['Conflict_Detail']='; '.join(dict.fromkeys(labels));r['Explanation']='All otherwise eligible sessions overlap an existing or newly selected class for this employee.'
                        else:r['Explanation']='No open session with remaining seats meets date/policy/prerequisite criteria.'
                    else:r['Explanation']='No open session with remaining seats meets date/policy/prerequisite criteria.'
                    r['Disposition']='REVIEW_REQUIRED';continue
                # Location precedence. Nearest physical site is considered before date/time.
                phys=[];virt=[]
                for o in offs:
                    if o['Location']:
                        d=distance(r['Physical_Location'],o['Location'])
                        if d is not None and d<=MAX_DISTANCE_MILES:phys.append((d,o))
                    else:virt.append(o)
                if phys:
                    nearest=min(d for d,o in phys);cand=[o for d,o in phys if abs(d-nearest)<0.01]
                    best=min(cand,key=lambda o:o['First_Start']);dist_m=nearest;kind='PHYSICAL'
                elif virt:
                    best=min(virt,key=lambda o:o['First_Start']);dist_m=None;kind='VIRTUAL'
                else:
                    r['Disposition']='REVIEW_REQUIRED';r['Explanation']=f'No physical session is available within {MAX_DISTANCE_MILES:.0f} miles and no eligible virtual session is available.';continue
                evaluated.append((r,bound,best,dist_m,kind))
            if not evaluated:continue
            # This is the critical global phase rule: any ready TARGET_RANGE requirement is selected before FIRST_AVAILABLE.
            targets=[x for x in evaluated if norm(x[0]['Scheduling_Policy'])=='target_range']
            pool=targets if targets else [x for x in evaluated if norm(x[0]['Scheduling_Policy'])=='first_available']
            if not pool:pool=evaluated
            def rank(x):
                r,b,o,d,k=x
                # For both policies earliest eligible session is primary; Priority resolves same start time.
                return (o['First_Start'],priority_val(r),norm(r['Training_Title']))
            r,b,o,d,kind=min(pool,key=rank)
            before=remaining[o['Session_Key']];remaining[o['Session_Key']]=before-1;selection_order+=1
            phase='TARGET_RANGE' if norm(r['Scheduling_Policy'])=='target_range' else 'FIRST_AVAILABLE'
            r.update({'Disposition':'PROPOSED_SCHEDULE','Selected_Session_WID':o['WID'],'Selected_Reference_ID':o['Reference_ID'],'Selected_Start':o['First_Start'],'Selected_End':o['Last_End'],'Selected_Location':o['Location'] or 'Virtual','Distance_Miles':round(d,1) if d is not None else '','Selected_Day_Count':o['Day_Count'],'Selected_Multi_Day':'Yes' if o['Multi_Day'] else 'No','Selected_Session_Days':' | '.join(f"{st:%m/%d/%Y %I:%M %p}-{en:%I:%M %p}" for st,en in o['Intervals']),'Original_Available_Seats':opening[o['Session_Key']],'Remaining_Seats_After_Reservation':remaining[o['Session_Key']],'Selection_Order':selection_order,'Selection_Phase':phase,'Workday_Ready':'Yes','Explanation':f"Selected {phase.replace('_',' ')} offering at nearest eligible physical location." if kind=='PHYSICAL' else 'Selected eligible virtual offering.'})
            if b:r['Explanation']+=f" Prerequisite scheduled earlier ({b:%m/%d/%Y %I:%M %p})."
            for st,en in o['Intervals']:busy[r['Person_Key']].append({'start':st,'end':en,'title':r['Training_Title'],'source':'SELECTED'})
            same_run[(r['Person_Key'],norm(r['Training_Title']))]=r

    cols=['Event_Key','Event_Type','Employee_ID','WID','Person_Key','Worker_Name','Worker_Type','Anchor_Date','Job_Code','Position_Title','Cost_Center_ID','Cost_Center_Title','Sup_Org_ID','Physical_Location','Hiring_Manager_AD','Work_Email','Home_Email','Training_Title','Training_Documentation_Status','Priority','Prerequisite','Prerequisite_Status','Scheduling_Policy','Timing_Modifier','Timing_Min','Timing_Max','Completion_Date','Completion_Training','Existing_Enrollment_Training','Existing_Registration_Date','Existing_Session_Start','Equivalency_Used','Equivalency_Direction','Equivalency_Match_Direction','Selected_Session_WID','Selected_Reference_ID','Selected_Start','Selected_End','Selected_Location','Distance_Miles','Selected_Day_Count','Selected_Multi_Day','Selected_Session_Days','Original_Available_Seats','Remaining_Seats_After_Reservation','Selection_Order','Selection_Phase','Conflict_Detail','Satisfied_By_Event_Key','Satisfied_By_Session_WID','Disposition','Explanation','Workday_Ready']
    reqdf=pd.DataFrame(reqs)
    for c in cols:
        if c not in reqdf.columns:reqdf[c]=''
    return {'events':pd.DataFrame(events),'duplicate_events':pd.DataFrame(dups),'requirements':reqdf[cols],'existing_workday':pd.DataFrame(existing),'seat_reservations':pd.DataFrame(),'summary':reqdf['Disposition'].value_counts().to_dict(),'source_counts':{'new_hires':len(nh),'job_changes':len(jc),'staffing_events_raw':len(nh)+len(jc),'staffing_events_canonical':len(events),'duplicate_source_events':len(dups),'documentation_found_events':docs_found,'documentation_not_found_events':docs_missing,'history':len(hist),'sessions':len(sessions),'existing_orientation':len(orientation)}}

def timing_ok(rec,start):
    a=rec.get('Anchor_Date')
    if not a:return True
    pol=norm(rec.get('Scheduling_Policy')); delta=(start.date()-a.date()).days
    try:mn=float(rec.get('Timing_Min')) if rec.get('Timing_Min') not in (None,'') else None
    except:mn=None
    try:mx=float(rec.get('Timing_Max')) if rec.get('Timing_Max') not in (None,'') else None
    except:mx=None
    if pol=='target_range':return (mn is None or delta>=mn) and (mx is None or delta<=mx)
    if pol=='first_available':return delta>=0
    return (mn is None or delta>=mn) and (mx is None or delta<=mx)
