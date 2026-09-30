from __future__ import annotations
from collections import defaultdict
from datetime import datetime, date
from io import BytesIO
from pathlib import Path
import math, re, unicodedata
import openpyxl
import pandas as pd

MAX_DISTANCE_MILES=130.0
ACTIVE_REGISTRATION_STATUSES={'enrolled','registered','in progress','not started','active'}

ALIASES={
'employee id':['employee id','employee_id','employee number','worker id'],
'wid':['wid','worker wid','employee wid'],
'worker':['worker','worker name','learning participant'],
'hire date':['hire date','most recent hire date'],
'candidate position id':['candidate position id','candidate position id - proposed'],
'new hire position title':['new hire position title','position title'],
'candidate cost center id':['candidate cost center id','cost center id'],
'candidate department name':['candidate department name','cost center'],
'supervisory organization id':['supervisory organization id','sup org id','sup org id - requested'],
'worker physical location':['worker physical location','location'],
'worker type':['worker type'],
'email - primary work':['email - primary work','work email'],
'email - primary home':['email - primary home','home email'],
'hiring manager ad':['hiring manager ad (username)','hiring manager ad (username)',"hiring manager ad (username\r\n)"],
'effective date':['effective date','position effective date','staffing event effective date'],
'job code - proposed':['job code - proposed','proposed job code'],
'job title - proposed':['job title - proposed','proposed job title'],
'cost center id - requested':['cost center id - requested','requested cost center id'],
'cost center - proposed':['cost center - proposed','proposed cost center'],
'sup org id - requested':['sup org id - requested','requested sup org id'],
'location - proposed':['location - proposed','proposed location'],
'candidate email':['candidate email','email - primary work'],
'learning participant':['learning participant','learner'],
'record learning content':['record learning content','learning content','learning content title'],
'record completion status':['record completion status','completion status'],
'record completion date':['record completion date','completion date','completed date'],
'registration status':['registration status','learner registration status'],
"learner's registration date":["learner's registration date",'learners registration date','registration date'],
'enrolled course offering':['enrolled course offering','course offering'],
'start date':['start date','session start date','offering start date'],
'end date':['end date','session end date','offering end date'],
'learning content type':['learning content type','content type'],
'title':['title','learning content title'],
'reference id':['reference id','reference_id'],
'available seats':['available seats','seats available','number of available seats'],
'availability status':['availability status','status'],
'locations':['locations','location'],
}

def norm(v):
    if v is None:return ''
    s=unicodedata.normalize('NFKC',str(v)).replace('\u00a0',' ').replace('\r',' ').replace('\n',' ')
    s=s.replace('’',"'").replace('‘',"'")
    return re.sub(r'\s+',' ',s.strip().lower())

def clean(v):
    return '' if v is None else str(v).strip()

def id_norm(v):
    if v is None:return ''
    if isinstance(v,float) and v.is_integer(): return str(int(v))
    return str(v).strip()

def dt(v):
    if isinstance(v,datetime):return v
    if isinstance(v,date):return datetime(v.year,v.month,v.day)
    if isinstance(v,str) and v.strip():
        s=v.strip()
        for fmt in ('%m/%d/%Y','%Y-%m-%d','%m/%d/%y','%m/%d/%Y %I:%M %p','%m/%d/%Y %H:%M:%S','%Y-%m-%d %H:%M:%S'):
            try:return datetime.strptime(s,fmt)
            except:pass
    return None

def _header_norm(v): return norm(v)

def _matches(h, canonical):
    h=_header_norm(h); aliases={_header_norm(canonical)}|{_header_norm(x) for x in ALIASES.get(canonical,[])}
    return h in aliases

def read_sheet(source, required_headers, optional_headers=None, scan_rows=150):
    if isinstance(source,(bytes,bytearray)): raw=bytes(source)
    elif hasattr(source,'getvalue'): raw=bytes(source.getvalue())
    elif hasattr(source,'read'):
        try: source.seek(0)
        except Exception: pass
        raw=source.read()
    else:
        raw=Path(source).read_bytes() if isinstance(source,(str,Path)) else bytes(source)
    if not raw: raise ValueError('Uploaded Excel file is empty.')
    wb=openpyxl.load_workbook(BytesIO(raw),read_only=False,data_only=True)
    best=None
    for ws in wb.worksheets:
        for r,row in enumerate(ws.iter_rows(values_only=True),1):
            if r>scan_rows: break
            vals=list(row)
            matched=[x for x in required_headers if any(_matches(v,x) for v in vals)]
            if len(matched)==len(required_headers):
                best=(ws,r,vals); break
        if best: break
    if not best:
        raise ValueError(f'Could not recognize workbook. Expected headers: {required_headers}. Worksheets: {wb.sheetnames}')
    ws,hr,raw_headers=best
    lookup={}
    for c,v in enumerate(raw_headers):
        for canonical in required_headers+(optional_headers or []):
            if canonical not in lookup and _matches(v,canonical): lookup[canonical]=c; break
    out=[]
    for row in ws.iter_rows(min_row=hr+1,values_only=True):
        if not any(v not in (None,'') for v in row): continue
        d={}
        for canonical,idx in lookup.items(): d[canonical]=row[idx] if idx<len(row) else None
        out.append(d)
    return out

def haversine_miles(a,b,c,d):
    R=3958.7613; p1=math.radians(a); p2=math.radians(c); dp=math.radians(c-a); dl=math.radians(d-b)
    x=math.sin(dp/2)**2+math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 2*R*math.asin(math.sqrt(x))

def overlap(a,b,c,d): return a<d and b>c

def _person_key(wid,eid): return f'WID:{id_norm(wid)}' if id_norm(wid) else f'EID:{id_norm(eid)}'

def timing_ok(rec, start):
    anchor=rec.get('Anchor_Date')
    if not anchor or not start:return True
    pol=norm(rec.get('Scheduling_Policy'))
    if pol!='target_range':return start>=anchor
    try: mn=float(rec.get('Timing_Min')) if rec.get('Timing_Min') not in (None,'') else None
    except: mn=None
    try: mx=float(rec.get('Timing_Max')) if rec.get('Timing_Max') not in (None,'') else None
    except: mx=None
    mod=norm(rec.get('Timing_Modifier')); delta=(start.date()-anchor.date()).days
    if mod=='between': return (mn is None or delta>=mn) and (mx is None or delta<=mx)
    if mod=='after': return mn is None or delta>=mn
    if mod=='before': return mx is None or delta<=mx
    return (mn is None or delta>=mn) and (mx is None or delta<=mx)

def run_scheduler(new_hire, job_change, history, sessions, rules_df, locations_df, equivalencies_df, orientation):
    nh=read_sheet(new_hire,['employee id','wid','hire date','candidate cost center id','candidate position id','worker'], ['worker type','new hire position title','candidate department name','supervisory organization id','worker physical location','hiring manager ad','email - primary work','email - primary home']) if new_hire else []
    jc=read_sheet(job_change,['employee id','wid','effective date','job code - proposed','cost center id - requested','worker'], ['job title - proposed','cost center - proposed','sup org id - requested','location - proposed','worker type','hiring manager ad','candidate email']) if job_change else []
    hist=read_sheet(history,['employee id','record learning content','record completion status'],['wid','learning participant','record completion date','registration status',"learner's registration date"]) 
    ses=read_sheet(sessions,['title','reference id','start date','end date','available seats','wid'],['availability status','locations'])
    ori=read_sheet(orientation,['employee id','enrolled course offering','registration status','start date','end date'],['wid','locations'])

    locs={}
    for _,r in locations_df.iterrows():
        name=clean(r.get('location'))
        if not name:continue
        try:lat=float(r.get('latitude'));lon=float(r.get('longitude'))
        except:lat=lon=None
        locs[norm(name)]={'name':name,'lat':lat,'lon':lon}
    def dist(emp,loc):
        a=locs.get(norm(emp));b=locs.get(norm(loc))
        if not a or not b or a['lat'] is None or b['lat'] is None:return None
        return haversine_miles(a['lat'],a['lon'],b['lat'],b['lon'])

    eq_one=defaultdict(set); eq_meta={}
    for _,r in equivalencies_df.iterrows():
        a=norm(r.get('required_training_title')); b=norm(r.get('equivalent_training_title')); direction=norm(r.get('relationship_direction') or 'one_way').replace('-','_').replace(' ','_')
        if not a or not b:continue
        eq_one[a].add(b); eq_meta[(a,b)]=(clean(r.get('required_training_title')),clean(r.get('equivalent_training_title')),'TWO_WAY' if direction in ('two_way','2_way','2') else 'ONE_WAY','FORWARD')
        if direction in ('two_way','2_way','2'):
            eq_one[b].add(a); eq_meta[(b,a)]=(clean(r.get('required_training_title')),clean(r.get('equivalent_training_title')),'TWO_WAY','REVERSE')
    def sat_titles(t):
        return {norm(t)}|set(eq_one.get(norm(t),set()))
    def eq_match(required, observed): return eq_meta.get((norm(required),norm(observed)))

    events=[]; raw_event_count=0
    for i,r in enumerate(nh,1):
        raw_event_count+=1; eid=id_norm(r.get('employee id')); wid=id_norm(r.get('wid'))
        events.append({'Event_Key':f'NH-{i}-{eid}','Event_Type':'NEW_HIRE','Employee_ID':eid,'WID':wid,'Person_Key':_person_key(wid,eid),'Worker_Name':clean(r.get('worker')),'Worker_Type':clean(r.get('worker type')),'Anchor_Date':dt(r.get('hire date')),'Job_Code':clean(r.get('candidate position id')),'Position_Title':clean(r.get('new hire position title')),'Cost_Center_ID':clean(r.get('candidate cost center id')),'Cost_Center_Title':clean(r.get('candidate department name')),'Sup_Org_ID':clean(r.get('supervisory organization id')),'Physical_Location':clean(r.get('worker physical location')),'Hiring_Manager_AD':clean(r.get('hiring manager ad')),'Work_Email':clean(r.get('email - primary work')),'Home_Email':clean(r.get('email - primary home'))})
    for i,r in enumerate(jc,1):
        raw_event_count+=1; eid=id_norm(r.get('employee id'));wid=id_norm(r.get('wid'))
        events.append({'Event_Key':f'JC-{i}-{eid}','Event_Type':'JOB_CHANGE','Employee_ID':eid,'WID':wid,'Person_Key':_person_key(wid,eid),'Worker_Name':clean(r.get('worker')),'Worker_Type':clean(r.get('worker type')),'Anchor_Date':dt(r.get('effective date')),'Job_Code':clean(r.get('job code - proposed')),'Position_Title':clean(r.get('job title - proposed')),'Cost_Center_ID':clean(r.get('cost center id - requested')),'Cost_Center_Title':clean(r.get('cost center - proposed')),'Sup_Org_ID':clean(r.get('sup org id - requested')),'Physical_Location':clean(r.get('location - proposed')),'Hiring_Manager_AD':clean(r.get('hiring manager ad')),'Work_Email':clean(r.get('candidate email')),'Home_Email':''})

    # Exact duplicate source events collapse before requirements are generated.
    seen_events={}; duplicate_events=[]; canonical=[]
    for ev in events:
        sig=(ev['Person_Key'],norm(ev['Event_Type']),ev['Anchor_Date'],norm(ev['Job_Code']),norm(ev['Position_Title']),norm(ev['Cost_Center_ID']),norm(ev['Sup_Org_ID']),norm(ev['Physical_Location']))
        if sig in seen_events:
            d=dict(ev);d['Canonical_Event_Key']=seen_events[sig]['Event_Key'];duplicate_events.append(d)
        else:
            ev['Canonical_Event_Key']=ev['Event_Key'];seen_events[sig]=ev;canonical.append(ev)
    events=canonical

    # Restore the established WID-first identity behavior for Workday output.
    # WID remains the durable identity; when multiple staffing events for the same
    # WID have different Employee IDs, use the most recent event's Employee ID as
    # the current Workday learner identifier.
    latest_employee_id={}
    for ev in events:
        pk=ev['Person_Key']
        stamp=ev.get('Anchor_Date') or datetime.min
        prior=latest_employee_id.get(pk)
        if prior is None or stamp >= prior[0]:
            latest_employee_id[pk]=(stamp,ev.get('Employee_ID',''))
    for ev in events:
        ev['Current_Employee_ID']=latest_employee_id.get(ev['Person_Key'],(None,ev.get('Employee_ID','')))[1] or ev.get('Employee_ID','')

    rules=[]
    for _,r in rules_df.iterrows():
        if not clean(r.get('training_title')):continue
        rules.append({'Rule_ID':r.get('id'),'Job_Code':clean(r.get('job_code')),'Cost_Center':clean(r.get('cost_center')),'Sup_Org':clean(r.get('supervisory_org')),'Training_Title':clean(r.get('training_title')),'Prerequisite':clean(r.get('prerequisite')),'Scheduling_Policy':clean(r.get('scheduling_policy')),'Timing_Modifier':clean(r.get('timing_modifier')),'Timing_Min':r.get('timing_min'),'Timing_Max':r.get('timing_max'),'Priority':r.get('priority')})

    reqs=[]
    docs_found=0; docs_not_found=0
    for ev in events:
        matched=[]
        for rr in rules:
            if norm(rr['Job_Code'])!=norm(ev['Job_Code']) or norm(rr['Cost_Center'])!=norm(ev['Cost_Center_ID']):continue
            if rr['Sup_Org'] and norm(rr['Sup_Org'])!=norm(ev['Sup_Org_ID']):continue
            matched.append(rr)
        if matched:
            docs_found += 1
            for rr in matched:
                reqs.append({**ev,**rr,'Training_Documentation_Status':'TRAINING_DOCUMENTATION_FOUND'})
        else:
            docs_not_found += 1
            reqs.append({**ev,'Rule_ID':'','Training_Title':'','Prerequisite':'','Scheduling_Policy':'','Timing_Modifier':'','Timing_Min':'','Timing_Max':'','Priority':'','Training_Documentation_Status':'NO_TRAINING_DOCUMENTATION_FOUND'})

    # Session offerings grouped by WID; every day/lesson interval is part of the same offering.
    grouped=defaultdict(list)
    for s in ses:
        w=id_norm(s.get('wid')) or f"{clean(s.get('reference id'))}|{clean(s.get('start date'))}|{norm(s.get('title'))}"
        grouped[w].append(s)
    offerings=[]
    for key,rows in grouped.items():
        intervals=[]
        for r in sorted(rows,key=lambda x:dt(x.get('start date')) or datetime.max):
            st=dt(r.get('start date'));en=dt(r.get('end date'))
            if st and en and en>st:intervals.append((st,en))
        if not intervals:continue
        title=clean(rows[0].get('title'));loc=next((clean(r.get('locations')) for r in rows if clean(r.get('locations'))),'')
        refs=[clean(r.get('reference id')) for r in rows if clean(r.get('reference id'))]
        seats=max([0]+[int(float(r.get('available seats') or 0)) for r in rows])
        statuses={norm(r.get('availability status')) for r in rows if clean(r.get('availability status'))}
        offerings.append({'Session_Key':key,'WID':id_norm(rows[0].get('wid')),'Title':title,'Reference_ID':refs[0] if refs else '','Location':loc,'Intervals':intervals,'First_Start':min(x[0] for x in intervals),'Last_End':max(x[1] for x in intervals),'Day_Count':len(intervals),'Multi_Day':len({x[0].date() for x in intervals})>1,'Seats':seats,'Status':'open' if not statuses or statuses=={'open'} else norm(next(iter(statuses)))})
    by_title=defaultdict(list)
    for o in offerings:by_title[norm(o['Title'])].append(o)
    remaining={o['Session_Key']:o['Seats'] for o in offerings}; opening=dict(remaining)

    # Existing Workday training and blocked calendar. Use Employee ID from the source report as
    # fallback; where the training history has a WID crosswalk, retain it.
    eid_to_wid={id_norm(r.get('employee id')):id_norm(r.get('wid')) for r in hist if id_norm(r.get('employee id')) and id_norm(r.get('wid'))}
    for ev in events:
        if ev['Employee_ID'] and ev['WID']:eid_to_wid[ev['Employee_ID']]=ev['WID']
    existing=[]; existing_seen=set(); busy=defaultdict(list)
    for o in ori:
        status=norm(o.get('registration status')); st=dt(o.get('start date'));en=dt(o.get('end date'));eid=id_norm(o.get('employee id'))
        if status not in ACTIVE_REGISTRATION_STATUSES or not st or not en or en<=st:continue
        pk=_person_key(eid_to_wid.get(eid,''),eid); title=clean(o.get('enrolled course offering')); key=(pk,norm(title),st,en)
        if key not in existing_seen:
            existing_seen.add(key); existing.append({'Person_Key':pk,'WID':eid_to_wid.get(eid,''),'Employee_ID':eid,'Training_Title':title,'Registration_Status':clean(o.get('registration status')),'Start_Date':st,'End_Date':en,'Location':clean(o.get('locations'))})
        busy[pk].append({'start':st,'end':en,'title':title,'source':'EXISTING'})

    person_req_titles=defaultdict(set)
    for r in reqs:person_req_titles[r['Person_Key']] |= sat_titles(r['Training_Title'])
    for e in existing:
        e['Required_For_Current_Role']='Yes' if norm(e['Training_Title']) in person_req_titles[e['Person_Key']] else 'No';e['Scheduling_Impact']='BLOCKS_OVERLAP_CHECK'

    # Prepopulate requirement dispositions using completion/current enrollment checks.
    for r in reqs:
        r.update({'Completion_Date':'','Completion_Training':'','Existing_Enrollment_Training':'','Existing_Registration_Date':'','Existing_Session_Start':'','Equivalency_Used':'','Equivalency_Direction':'','Equivalency_Match_Direction':'','Selected_Session_WID':'','Selected_Reference_ID':'','Selected_Start':'','Selected_End':'','Selected_Location':'','Distance_Miles':'','Selected_Day_Count':'','Selected_Multi_Day':'','Selected_Session_Days':'','Original_Available_Seats':'','Remaining_Seats_After_Reservation':'','Selection_Order':'','Selection_Phase':'','Conflict_Detail':'','Disposition':'PENDING_SCHEDULING','Explanation':'','Workday_Ready':'No','Satisfied_By_Event_Key':'','Satisfied_By_Session_WID':''})
        if r.get('Training_Documentation_Status')=='NO_TRAINING_DOCUMENTATION_FOUND':
            r['Disposition']='NO_TRAINING_DOCUMENTATION_FOUND'
            r['Explanation']=f"No training documentation was found for Job Code '{r.get('Job_Code','')}', Cost Center '{r.get('Cost_Center_ID','')}'" + (f", Supervisory Organization '{r.get('Sup_Org_ID','')}'" if r.get('Sup_Org_ID') else '') + '. No training was scheduled.'
            continue
        pt=r['Person_Key']; titles=sat_titles(r['Training_Title']); person_hist=[h for h in hist if _person_key(h.get('wid') or eid_to_wid.get(id_norm(h.get('employee id')),''),h.get('employee id'))==pt]
        completed=[h for h in person_hist if norm(h.get('record learning content')) in titles and norm(h.get('record completion status'))=='completed']
        if completed:
            h=max(completed,key=lambda x:dt(x.get('record completion date')) or datetime.min); obs=clean(h.get('record learning content'));r['Completion_Date']=h.get('record completion date');r['Completion_Training']=obs
            em=eq_match(r['Training_Title'],obs)
            r['Disposition']='PREVIOUSLY_COMPLETED' if norm(obs)==norm(r['Training_Title']) else 'EQUIVALENT_COMPLETION'
            if em:r['Equivalency_Used']=f"{em[1]} -> {em[0]}" if em[3]=='FORWARD' else f"{em[0]} <-> {em[1]}";r['Equivalency_Direction']=em[2];r['Equivalency_Match_Direction']=em[3]
            r['Explanation']='Historical completion satisfies requirement. No reassignment required.' if norm(obs)==norm(r['Training_Title']) else f"Historical completion of equivalent training '{obs}' satisfies '{r['Training_Title']}'. No reassignment required."
            continue
        enr=[h for h in person_hist if norm(h.get('record learning content')) in titles and norm(h.get('registration status')) in ACTIVE_REGISTRATION_STATUSES]
        ext=[e for e in existing if e['Person_Key']==pt and norm(e['Training_Title']) in titles]
        if enr or ext:
            if ext:
                x=min(ext,key=lambda y:y['Start_Date']); obs=x['Training_Title'];r['Existing_Enrollment_Training']=obs;r['Existing_Session_Start']=x['Start_Date']
            else:
                x=max(enr,key=lambda y:dt(y.get("learner's registration date")) or datetime.min);obs=clean(x.get('record learning content'));r['Existing_Enrollment_Training']=obs;r['Existing_Registration_Date']=x.get("learner's registration date")
            em=eq_match(r['Training_Title'],obs);r['Disposition']='ALREADY_ENROLLED'
            if em:r['Equivalency_Used']=f"{em[1]} -> {em[0]}" if em[3]=='FORWARD' else f"{em[0]} <-> {em[1]}";r['Equivalency_Direction']=em[2];r['Equivalency_Match_Direction']=em[3]
            r['Explanation']='Existing active registration found. Duplicate enrollment suppressed.' if norm(obs)==norm(r['Training_Title']) else f"Existing active enrollment in equivalent training '{obs}' satisfies '{r['Training_Title']}'. Duplicate enrollment suppressed."
            continue
        if norm(r['Scheduling_Policy'])=='flag_manual':r['Disposition']='MANUAL_SCHEDULING_REQUIRED';r['Explanation']='Training rule uses FLAG_MANUAL.'

    # Dependency-aware global policy scheduler. Every canonical event is handled as a single
    # planning problem, but TARGET_RANGE is a higher scheduling phase than FIRST_AVAILABLE.
    selection_order=0
    for ev in events:
        er=[r for r in reqs if r['Event_Key']==ev['Event_Key'] and r['Disposition']=='PENDING_SCHEDULING']
        if not er:continue
        byname={norm(r['Training_Title']):r for r in reqs if r['Event_Key']==ev['Event_Key']}
        assigned_course={}
        # Loop until all schedulable requirements are resolved.
        while True:
            remaining_reqs=[r for r in er if r['Disposition']=='PENDING_SCHEDULING']
            if not remaining_reqs:break
            ready=[]
            for r in remaining_reqs:
                prereq=norm(r.get('Prerequisite')); bound=None; blocked=False
                if prereq:
                    p=byname.get(prereq)
                    if p:
                        if p['Disposition']=='PROPOSED_SCHEDULE':bound=dt(p.get('Selected_Start'));r['Prerequisite_Status']='SCHEDULED_EARLIER'
                        elif p['Disposition'] in ('PREVIOUSLY_COMPLETED','EQUIVALENT_COMPLETION'):r['Prerequisite_Status']='SATISFIED_BY_COMPLETION'
                        elif p['Disposition']=='ALREADY_ENROLLED' and dt(p.get('Existing_Session_Start')):bound=dt(p.get('Existing_Session_Start'));r['Prerequisite_Status']='SATISFIED_BY_EXISTING_ENROLLMENT'
                        elif p['Disposition']=='SATISFIED_BY_SAME_RUN_ASSIGNMENT' and dt(p.get('Selected_Start')):bound=dt(p.get('Selected_Start'));r['Prerequisite_Status']='SCHEDULED_EARLIER'
                        else: blocked=True
                if not blocked:ready.append((r,bound))
            if not ready:
                for r in remaining_reqs:r['Disposition']='REVIEW_REQUIRED';r['Explanation']=f"Prerequisite '{r.get('Prerequisite')}' could not be resolved before scheduling." if r.get('Prerequisite') else 'No schedulable session could be resolved.'
                break
            candidates=[]
            for r,bound in ready:
                course=(r['Person_Key'],norm(r['Training_Title']))
                if course in assigned_course:
                    p=assigned_course[course]; ps=dt(p['Selected_Start']); valid=bool(ps and timing_ok(r,ps) and (not bound or ps>bound))
                    if valid:
                        for k in ('Selected_Session_WID','Selected_Reference_ID','Selected_Start','Selected_End','Selected_Location','Distance_Miles','Selected_Day_Count','Selected_Multi_Day','Selected_Session_Days','Selection_Order','Selection_Phase'):r[k]=p.get(k,'')
                        r['Disposition']='SATISFIED_BY_SAME_RUN_ASSIGNMENT';r['Satisfied_By_Event_Key']=p['Event_Key'];r['Satisfied_By_Session_WID']=p['Selected_Session_WID'];r['Explanation']=f"Same person/course was already assigned under staffing event {p['Event_Key']}; one enrollment satisfies both requirements.";continue
                    r['Disposition']='REVIEW_REQUIRED';r['Explanation']='A prior same-run assignment exists, but its session does not meet this requirement\'s timing/prerequisite constraints. A duplicate enrollment was not created.';continue
                offs=[]; conflicts=[]
                for o in by_title.get(norm(r['Training_Title']),[]):
                    if o['Status']!='open' or remaining.get(o['Session_Key'],0)<=0:continue
                    if r['Anchor_Date'] and o['First_Start']<r['Anchor_Date']:continue
                    if bound and o['First_Start']<=bound:continue
                    if not timing_ok(r,o['First_Start']):continue
                    cf=[]
                    for st,en in o['Intervals']:
                        cf += [b for b in busy[r['Person_Key']] if overlap(st,en,b['start'],b['end'])]
                    if cf:conflicts.append((o,cf));continue
                    offs.append(o)
                if not offs:
                    if conflicts:r['Disposition']='REVIEW_REQUIRED';r['Conflict_Detail']='; '.join(sorted({f"{b['title']} ({b['start']:%m/%d/%Y %I:%M %p}-{b['end']:%I:%M %p})" for o,cs in conflicts for b in cs}));r['Explanation']='All otherwise eligible sessions overlap an existing or newly selected class for this employee.'
                    else:r['Disposition']='REVIEW_REQUIRED';r['Explanation']='No open session with remaining seats meets date/policy/prerequisite/non-overlap criteria.'
                    continue
                # Best candidate by location first, then earliest date/time. Physical wins.
                physical=[];virtual=[]
                for o in offs:
                    if o['Location']:
                        d=dist(r['Physical_Location'],o['Location'])
                        if d is not None and d<=MAX_DISTANCE_MILES:physical.append((d,o))
                    else:virtual.append(o)
                if physical:
                    nd=min(d for d,o in physical); best=min((o for d,o in physical if abs(d-nd)<0.01),key=lambda x:x['First_Start']); score=(best['First_Start'],nd)
                    candidates.append((r,bound,best,score,'PHYSICAL',nd))
                elif virtual:
                    best=min(virtual,key=lambda o:o['First_Start']);candidates.append((r,bound,best,(best['First_Start'],float('inf')),'VIRTUAL',None))
                else:
                    r['Disposition']='REVIEW_REQUIRED';r['Explanation']=f"No physical session within {MAX_DISTANCE_MILES:.0f} miles and no virtual session is available.";continue
            if not candidates:continue
            # GLOBAL phase precedence: any ready TARGET_RANGE beats every FIRST_AVAILABLE.
            target=[x for x in candidates if norm(x[0]['Scheduling_Policy'])=='target_range']
            pool=target if target else [x for x in candidates if norm(x[0]['Scheduling_Policy'])=='first_available'] or candidates
            # Among competing FIRST_AVAILABLE selections at identical date/time, lower priority wins.
            # For TARGET_RANGE, earliest feasible session remains primary; priority only breaks exact ties.
            def rank(x):
                r,b,o,score,typ,d=x
                try:pval=float(r.get('Priority')) if r.get('Priority') not in (None,'') else float('inf')
                except:pval=float('inf')
                return (score[0],pval,norm(r['Training_Title']))
            r,b,o,score,typ,d=min(pool,key=rank)
            if remaining[o['Session_Key']]<=0:continue
            remaining[o['Session_Key']]-=1;selection_order+=1
            phase='TARGET_RANGE' if norm(r['Scheduling_Policy'])=='target_range' else 'FIRST_AVAILABLE'
            r.update({'Disposition':'PROPOSED_SCHEDULE','Selected_Session_WID':o['WID'],'Selected_Reference_ID':o['Reference_ID'],'Selected_Start':o['First_Start'],'Selected_End':o['Last_End'],'Selected_Location':o['Location'] or 'Virtual','Distance_Miles':round(d,1) if d is not None else '','Selected_Day_Count':o['Day_Count'],'Selected_Multi_Day':'Yes' if o['Multi_Day'] else 'No','Selected_Session_Days':' | '.join(f"{st:%m/%d/%Y %I:%M %p}-{en:%I:%M %p}" for st,en in o['Intervals']),'Original_Available_Seats':opening[o['Session_Key']],'Remaining_Seats_After_Reservation':remaining[o['Session_Key']],'Selection_Order':selection_order,'Selection_Phase':phase,'Workday_Ready':'Yes','Explanation':f"Selected {phase.replace('_',' ')} offering at nearest eligible physical location." if typ=='PHYSICAL' else 'Selected eligible virtual offering.'})
            if b:r['Explanation'] += f" Prerequisite scheduled earlier ({b:%m/%d/%Y %I:%M %p})."
            for st,en in o['Intervals']:busy[r['Person_Key']].append({'start':st,'end':en,'title':r['Training_Title'],'source':'SELECTED'})
            assigned_course[(r['Person_Key'],norm(r['Training_Title']))]=r

    req_cols=['Event_Key','Event_Type','Employee_ID','Current_Employee_ID','WID','Person_Key','Worker_Name','Worker_Type','Anchor_Date','Job_Code','Position_Title','Cost_Center_ID','Cost_Center_Title','Sup_Org_ID','Physical_Location','Hiring_Manager_AD','Work_Email','Home_Email','Training_Title','Training_Documentation_Status','Priority','Prerequisite','Scheduling_Policy','Timing_Modifier','Timing_Min','Timing_Max','Completion_Date','Completion_Training','Existing_Registration_Date','Existing_Session_Start','Existing_Enrollment_Training','Equivalency_Used','Equivalency_Direction','Equivalency_Match_Direction','Prerequisite_Status','Selected_Session_WID','Selected_Reference_ID','Selected_Start','Selected_End','Selected_Location','Distance_Miles','Selected_Day_Count','Selected_Multi_Day','Selected_Session_Days','Original_Available_Seats','Remaining_Seats_After_Reservation','Selection_Order','Selection_Phase','Conflict_Detail','Satisfied_By_Event_Key','Satisfied_By_Session_WID','Disposition','Explanation','Workday_Ready']
    requirements=pd.DataFrame(reqs)
    for c in req_cols:
        if c not in requirements.columns:requirements[c]=''
    requirements=requirements[req_cols]
    return {'events':pd.DataFrame(events),'duplicate_events':pd.DataFrame(duplicate_events),'requirements':requirements,'existing_workday':pd.DataFrame(existing),'seat_reservations':pd.DataFrame([{'Session_Key':o['Session_Key'],'Training_Title':r['Training_Title'],'Start_Date':r['Selected_Start'],'End_Date':r['Selected_End'],'Employee_ID':r['Employee_ID'],'WID':r['WID'],'Event_Key':r['Event_Key'],'Seats_Before':r['Original_Available_Seats'],'Seats_After':r['Remaining_Seats_After_Reservation']} for r in reqs for o in offerings if r.get('Selected_Session_WID')==o['WID'] and r.get('Disposition')=='PROPOSED_SCHEDULE']),'summary':requirements['Disposition'].value_counts().to_dict(),'source_counts':{'new_hires':len(nh),'job_changes':len(jc),'staffing_events_raw':raw_event_count,'staffing_events_canonical':len(events),'duplicate_source_events':len(duplicate_events),'documentation_found_events':docs_found,'documentation_not_found_events':docs_not_found,'history':len(hist),'sessions':len(ses),'existing_orientation':len(ori)}}
