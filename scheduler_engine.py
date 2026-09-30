from __future__ import annotations
from collections import defaultdict, Counter
from datetime import datetime, date
from io import BytesIO
import math, re, unicodedata
import openpyxl
import pandas as pd

ACTIVE_REGISTRATION_STATUSES={'enrolled','registered','in progress','not started'}
MAX_DISTANCE_MILES=130.0


def clean(v):
    if v is None: return ''
    try:
        if pd.isna(v): return ''
    except Exception:
        pass
    return str(v).strip() if not isinstance(v,(datetime,date)) else v

def norm(v):
    if v is None: return ''
    return re.sub(r'\s+',' ',str(v).strip().lower())

def header_norm(v):
    if v is None: return ''
    s=unicodedata.normalize('NFKC',str(v)).replace('\u00a0',' ').replace('\u200b',' ')
    s=s.replace('\r',' ').replace('\n',' ').replace('’',"'").replace('‘',"'")
    return re.sub(r'\s+',' ',s).strip().lower().rstrip(':*').strip()

def id_norm(v):
    if v is None: return ''
    if isinstance(v,float) and v.is_integer(): return str(int(v))
    return str(v).strip()

def dt(v):
    if isinstance(v,datetime): return v
    if isinstance(v,date): return datetime(v.year,v.month,v.day)
    if isinstance(v,str) and v.strip():
        s=v.strip()
        for fmt in ('%m/%d/%Y','%Y-%m-%d','%m/%d/%y','%m/%d/%Y %I:%M %p','%Y-%m-%d %H:%M:%S'):
            try: return datetime.strptime(s,fmt)
            except ValueError: pass
        try: return pd.to_datetime(s).to_pydatetime()
        except Exception: return None
    return None

def _open_wb(source):
    if isinstance(source,(bytes,bytearray)):
        raw=bytes(source)
    elif hasattr(source,'getvalue'):
        raw=bytes(source.getvalue())
    else:
        try: source.seek(0)
        except Exception: pass
        raw=source.read()
    if not raw:
        raise ValueError('The uploaded Excel file is empty or could not be read.')
    # Normal mode is intentional: some Workday exports have incorrect A1:A1
    # worksheet dimension metadata that breaks read_only mode.
    return openpyxl.load_workbook(BytesIO(raw),read_only=False,data_only=True)

HEADER_ALIASES={
'employee id':{'employee id','employee_id','employee id number','worker id','employee number'},
'wid':{'wid','worker wid','employee wid'},
'learning participant':{'learning participant','learner','learner name','learning participant name'},
'record learning content':{'record learning content','learning content','learning content title','record learning content title'},
'record completion status':{'record completion status','completion status','learning completion status'},
'record completion date':{'record completion date','completion date','completed date','completed on'},
'registration status':{'registration status','learner registration status'},
"learner's registration date":{'learner\'s registration date','learners registration date','registration date'},
'course offering':{'course offering','learning enrollment'},
'enrolled course offering':{'enrolled course offering','course offering','learning enrollment'},
'start date':{'start date','session start date','offering start date'},
'end date':{'end date','session end date','offering end date'},
'learning content type':{'learning content type','content type'},
'title':{'title','learning content title','course title'},
'reference id':{'reference id','reference_id'},
'available seats':{'available seats','seats available','number of available seats'},
'location':{'location','locations'},
'hire date':{'hire date','most recent hire date'},
'candidate cost center id':{'candidate cost center id','candidate cost center','cost center id'},
'candidate position id':{'candidate position id'},
'effective date':{'effective date','position effective date','staffing event effective date'},
'job code - proposed':{'job code - proposed','proposed job code','job code proposed'},
}

def _header_matches(v,canonical):
    hv=header_norm(v); ck=header_norm(canonical)
    return hv==ck or hv in (HEADER_ALIASES.get(ck,set())|{ck})

def read_sheet(source,required_headers,optional_headers=None,sheet=None,scan_rows=150):
    wb=_open_wb(source); required_headers=list(required_headers); optional_headers=list(optional_headers or [])
    sheets=[wb[sheet]] if sheet else list(wb.worksheets)
    best=(0,None,None,[],[]); selected=None
    for ws in sheets:
        for i,row in enumerate(ws.iter_rows(values_only=True),1):
            if i>scan_rows: break
            vals=list(row)
            missing=[h for h in required_headers if not any(_header_matches(v,h) for v in vals)]
            matched=[h for h in required_headers if h not in missing]
            if len(matched)>best[0]: best=(len(matched),ws.title,i,vals,missing)
            if not missing:
                selected=(ws,i,vals); break
        if selected: break
    if selected is None:
        n,sh,row,vals,missing=best
        samples=[]
        for ws in sheets:
            for rr in ws.iter_rows(min_row=1,max_row=min(ws.max_row,10),values_only=True):
                vv=[str(v)[:80] for v in rr if v not in (None,'')]
                if vv: samples.append(f'{ws.title}: {vv[:8]}')
        raise ValueError(f'Could not recognize report. Expected {required_headers}; best match worksheet={sh}, row={row}, matched {n}/{len(required_headers)}, missing={missing}. Samples: {samples[:12]}')
    ws,hr,hv=selected; raw_headers=list(hv); canon={}
    for h in required_headers+optional_headers:
        exact=next((i for i,v in enumerate(raw_headers) if header_norm(v)==header_norm(h)),None)
        alias=next((i for i,v in enumerate(raw_headers) if _header_matches(v,h)),None)
        j=exact if exact is not None else alias
        if j is not None: canon[j]=h
    headers=[]; seen=Counter()
    for i,v in enumerate(raw_headers):
        h=canon.get(i,str(clean(v)) if clean(v) else f'_blank_{i+1}')
        seen[h]+=1; headers.append(h if seen[h]==1 else f'{h}_{seen[h]}')
    out=[]
    for row in ws.iter_rows(min_row=hr+1,values_only=True):
        if any(v not in (None,'') for v in row):
            vals=list(row)+[None]*(len(headers)-len(row)); out.append(dict(zip(headers,vals[:len(headers)])))
    return out

def haversine_miles(a,b,c,d):
    R=3958.7613; p1=math.radians(a); p2=math.radians(c); dp=math.radians(c-a); dl=math.radians(d-b)
    x=math.sin(dp/2)**2+math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 2*R*math.asin(math.sqrt(x))

def overlap(a_start,a_end,b_start,b_end):
    return a_start < b_end and a_end > b_start

def person_key(wid,eid):
    wid=id_norm(wid); eid=id_norm(eid)
    return f'WID:{wid}' if wid else f'EID:{eid}'

def _priority(v):
    try:
        if v in (None,'','nan'): return 10**9
        return float(v)
    except Exception: return 10**9

def _policy_rank(policy):
    p=norm(policy)
    return {'target_range':0,'first_available':1,'flag_manual':2}.get(p,9)

def timing_ok(rec,start):
    anchor=rec.get('Anchor_Date')
    if not anchor: return True
    p=norm(rec.get('Scheduling_Policy'))
    if p!='target_range': return start>=anchor
    mn=rec.get('Timing_Min'); mx=rec.get('Timing_Max'); mod=norm(rec.get('Timing_Modifier'))
    try: mn=float(mn) if mn not in (None,'') else None
    except Exception: mn=None
    try: mx=float(mx) if mx not in (None,'') else None
    except Exception: mx=None
    delta=(start.date()-anchor.date()).days
    if mod=='between': return (mn is None or delta>=mn) and (mx is None or delta<=mx)
    if mod=='after': return delta >= (mn if mn is not None else (mx if mx is not None else -10**9))
    if mod=='before': return delta <= (mx if mx is not None else (mn if mn is not None else 10**9))
    return (mn is None or delta>=mn) and (mx is None or delta<=mx)

def _session_key(s):
    return id_norm(s.get('WID')) or f"{clean(s.get('Reference ID'))}|{norm(s.get('Title'))}|{clean(s.get('Start Date'))}|{clean(s.get('End Date'))}"

def _group_sessions(raw):
    groups=defaultdict(list)
    for s in raw:
        if not clean(s.get('Title')) or not id_norm(s.get('WID')): continue
        groups[id_norm(s.get('WID'))].append(s)
    out=[]
    for wid,rows in groups.items():
        rows=sorted(rows,key=lambda r:dt(r.get('Start Date')) or datetime.max)
        days=[]
        seen_days=set()
        for r in rows:
            st=dt(r.get('Start Date')); en=dt(r.get('End Date'))
            if not st or not en or en<=st: continue
            day=(st,en)
            if day in seen_days: continue
            seen_days.add(day); days.append({'start':st,'end':en})
        if not days: continue
        seats=[]
        for r in rows:
            try: seats.append(int(float(r.get('Available Seats') or 0)))
            except Exception: pass
        locs=[clean(r.get('Locations') or r.get('Location')) for r in rows if clean(r.get('Locations') or r.get('Location'))]
        title=clean(rows[0].get('Title')); ref=clean(rows[0].get('Reference ID')); status='open' if any(norm(r.get('Availability Status'))=='open' for r in rows) else clean(rows[0].get('Availability Status'))
        out.append({'WID':wid,'Title':title,'Reference ID':ref,'Availability Status':status,'Available Seats':min(seats) if seats else 0,'Locations':Counter(locs).most_common(1)[0][0] if locs else '','days':days,'start':min(d['start'] for d in days),'end':max(d['end'] for d in days),'day_count':len(days)})
    return out

def run_scheduler(new_hire_file,job_change_file,history_file,sessions_file,rules_df,locations_df,equivalencies_df=None,orientation_file=None):
    if rules_df is None or rules_df.empty: raise ValueError('Training configuration is empty.')
    nh=read_sheet(new_hire_file,['Employee ID','WID','Hire Date','Candidate Cost Center ID'],['Candidate Position ID','Worker','Worker Type','Traveler Designation','New Hire Position Title','Candidate Department Name','Supervisory Organization ID','Worker Physical Location','Hiring Manager AD (Username)','Email - Primary Work','Email - Primary Home']) if new_hire_file else []
    jc=read_sheet(job_change_file,['Employee ID','WID','Effective Date','Job Code - Proposed'],['Worker','Worker Type','Job Title - Proposed','Cost Center ID - Requested','Cost Center - Proposed','Sup Org ID - Requested','Location - Proposed','Hiring Manager AD (Username)','Candidate Email']) if job_change_file else []
    history=read_sheet(history_file,['Employee ID','Record Learning Content','Record Completion Status'],['Learning Participant','WID','Registration Status',"Learner's Registration Date",'Record Completion Date'])
    sessions_raw=read_sheet(sessions_file,['Learning Content Type','Title','Reference ID','Start Date','End Date','Available Seats','WID'],['Availability Status','Locations'])
    orientation=read_sheet(orientation_file,['Employee ID','Enrolled Course Offering','Registration Status','Start Date','End Date'],['WID','Locations','Reference ID','Course Offering']) if orientation_file else []
    sessions=_group_sessions(sessions_raw)
    sessions_by_title=defaultdict(list)
    for s in sessions: sessions_by_title[norm(s['Title'])].append(s)
    locs={}
    for _,r in locations_df.iterrows():
        n=clean(r.get('location'))
        if not n: continue
        try: lat=float(r.get('latitude')); lon=float(r.get('longitude'))
        except Exception: lat=lon=None
        locs[norm(n)]=(lat,lon)
    def distance(a,b):
        x=locs.get(norm(a)); y=locs.get(norm(b))
        if not x or not y or x[0] is None or y[0] is None: return None
        return haversine_miles(x[0],x[1],y[0],y[1])

    emp_hist_wid={id_norm(r.get('Employee ID')):id_norm(r.get('WID')) for r in history if id_norm(r.get('Employee ID')) and id_norm(r.get('WID'))}
    events=[]
    for i,r in enumerate(nh,1):
        eid=id_norm(r.get('Employee ID')); sw=id_norm(r.get('WID'))
        events.append({'Event_Key':f'NH-{i}-{eid}','Event_Type':'NEW_HIRE','Employee_ID':eid,'WID':emp_hist_wid.get(eid,sw),'Source_WID':sw,'Worker_Name':clean(r.get('Worker')),'Worker_Type':clean(r.get('Worker Type')),'Traveler_Designation':clean(r.get('Traveler Designation')),'Anchor_Date':dt(r.get('Hire Date')),'Job_Code':clean(r.get('Candidate Position ID')),'Position_Title':clean(r.get('New Hire Position Title')),'Cost_Center_ID':clean(r.get('Candidate Cost Center ID')),'Cost_Center_Title':clean(r.get('Candidate Department Name')),'Sup_Org_ID':clean(r.get('Supervisory Organization ID')),'Physical_Location':clean(r.get('Worker Physical Location')),'Hiring_Manager_AD':clean(r.get('Hiring Manager AD (Username)')),'Work_Email':clean(r.get('Email - Primary Work')),'Home_Email':clean(r.get('Email - Primary Home'))})
    for i,r in enumerate(jc,1):
        eid=id_norm(r.get('Employee ID')); sw=id_norm(r.get('WID'))
        events.append({'Event_Key':f'JC-{i}-{eid}','Event_Type':'JOB_CHANGE','Employee_ID':eid,'WID':emp_hist_wid.get(eid,sw),'Source_WID':sw,'Worker_Name':clean(r.get('Worker')),'Worker_Type':clean(r.get('Worker Type')),'Traveler_Designation':'','Anchor_Date':dt(r.get('Effective Date')),'Job_Code':clean(r.get('Job Code - Proposed')),'Position_Title':clean(r.get('Job Title - Proposed')),'Cost_Center_ID':clean(r.get('Cost Center ID - Requested')),'Cost_Center_Title':clean(r.get('Cost Center - Proposed')),'Sup_Org_ID':clean(r.get('Sup Org ID - Requested')),'Physical_Location':clean(r.get('Location - Proposed')),'Hiring_Manager_AD':clean(r.get('Hiring Manager AD (Username)')),'Work_Email':clean(r.get('Candidate Email')),'Home_Email':''})
    for e in events:
        e['Person_Key']=person_key(e['WID'],e['Employee_ID'])
        e['Contingent_Action']='STANDARD SCHEDULE' if norm(e['Worker_Type'])!='contingent worker' else ('STANDARD SCHEDULE' if norm(e['Traveler_Designation'])=='trav i' else ('ESCALATE' if norm(e['Traveler_Designation'])=='trav d' else 'NO SCHEDULE'))
    # Duplicate source-event collapse retained.
    sigmap={}; canonical=[]; duplicate=[]
    for e in events:
        sig=(e['Person_Key'],norm(e['Event_Type']),e['Anchor_Date'].isoformat() if e.get('Anchor_Date') else '',norm(e['Job_Code']),norm(e['Position_Title']),norm(e['Cost_Center_ID']),norm(e['Sup_Org_ID']),norm(e['Physical_Location']),norm(e['Worker_Type']),norm(e['Traveler_Designation']))
        if sig in sigmap:
            d=dict(e); d['Canonical_Event_Key']=sigmap[sig]['Event_Key']; d['Event_Status']='DUPLICATE_SOURCE_EVENT'; duplicate.append(d)
        else:
            e['Canonical_Event_Key']=e['Event_Key']; e['Event_Status']='CANONICAL'; sigmap[sig]=e; canonical.append(e)
    events=canonical

    # Directional equivalencies.
    eq=defaultdict(set); eq_meta={}
    if equivalencies_df is not None and not equivalencies_df.empty:
        for _,r in equivalencies_df.iterrows():
            req=norm(r.get('required_training_title')); equiv=norm(r.get('equivalent_training_title'))
            if not req or not equiv: continue
            d=norm(r.get('relationship_direction') or 'ONE_WAY').replace('-','_').replace(' ','_')
            if d in ('2_way','twoway'): d='two_way'
            else: d='one_way'
            eq[req].add(equiv); eq_meta[(req,equiv)]={'direction':'TWO_WAY' if d=='two_way' else 'ONE_WAY','forward':True,'required':clean(r.get('required_training_title')),'equivalent':clean(r.get('equivalent_training_title'))}
            if d=='two_way': eq[equiv].add(req); eq_meta[(equiv,req)]={'direction':'TWO_WAY','forward':False,'required':clean(r.get('required_training_title')),'equivalent':clean(r.get('equivalent_training_title'))}
    def satisfying(title): return {norm(title)}|set(eq.get(norm(title),set()))

    hist_by_person=defaultdict(list)
    for h in history: hist_by_person[person_key(emp_hist_wid.get(id_norm(h.get('Employee ID')),id_norm(h.get('WID'))),id_norm(h.get('Employee ID')))].append(h)
    existing_by_person=defaultdict(list); existing_unique=set()
    busy=defaultdict(list)
    for o in orientation:
        if norm(o.get('Registration Status')) not in ACTIVE_REGISTRATION_STATUSES: continue
        st=dt(o.get('Start Date')); en=dt(o.get('End Date'))
        if not st or not en or en<=st: continue
        eid=id_norm(o.get('Employee ID')); wid=emp_hist_wid.get(eid,id_norm(o.get('WID'))); pk=person_key(wid,eid); title=clean(o.get('Enrolled Course Offering'))
        ok=(pk,norm(title),st,en)
        if ok in existing_unique: continue
        existing_unique.add(ok)
        existing_by_person[pk].append({'Person_Key':pk,'WID':wid,'Employee_ID':eid,'Training_Title':title,'Registration_Status':clean(o.get('Registration Status')),'Start_Date':st,'End_Date':en,'Location':clean(o.get('Locations') or o.get('Location'))})
        busy[pk].append({'start':st,'end':en,'title':title,'source':'EXISTING_SCHEDULE'})
    # Index rules once by Job Code + Cost Center; this preserves matching behavior while
    # avoiding repeated scans of the full 18k-row configuration for every employee.
    rule_index=defaultdict(list)
    for _,rr in rules_df.iterrows():
        title=clean(rr.get('training_title'))
        if not title: continue
        rule_index[(norm(rr.get('job_code')),norm(rr.get('cost_center')))].append(rr)
    reqs=[]; seen=set()
    for e in events:
        if norm(e['Contingent_Action'])=='no schedule': continue
        candidates=rule_index.get((norm(e['Job_Code']),norm(e['Cost_Center_ID'])),[])
        for rr in candidates:
            title=clean(rr.get('training_title'))
            if clean(rr.get('supervisory_org')) and norm(rr.get('supervisory_org'))!=norm(e['Sup_Org_ID']): continue
            key=(e['Event_Key'],norm(title))
            if key in seen: continue
            seen.add(key)
            reqs.append({**e,'Rule_ID':rr.get('id'),'Training_Title':title,'Prerequisite':clean(rr.get('prerequisite')),'Topic':clean(rr.get('topic')),'Scheduling_Policy':clean(rr.get('scheduling_policy')),'Timing_Modifier':clean(rr.get('timing_modifier')),'Timing_Min':rr.get('timing_min'),'Timing_Max':rr.get('timing_max'),'Priority':rr.get('priority')})

    # Existing Workday training is visible for the employee whether or not it is a requirement.
    person_req_titles=defaultdict(set)
    for r in reqs: person_req_titles[r['Person_Key']] |= satisfying(r['Training_Title'])
    for row in sum(existing_by_person.values(),[]):
        row['Required_For_Current_Role']='Yes' if norm(row['Training_Title']) in person_req_titles.get(row['Person_Key'],set()) else 'No'
        row['Scheduling_Impact']='BLOCKS_OVERLAP_CHECK'

    # Pre-populate dispositions / suppression checks.
    for r in reqs:
        r.update({'Completion_Date':'','Completion_Training':'','Existing_Registration_Date':'','Existing_Session_Start':'','Existing_Enrollment_Training':'','Equivalency_Used':'','Equivalency_Direction':'','Equivalency_Match_Direction':'','Selected_Session_WID':'','Selected_Reference_ID':'','Selected_Start':'','Selected_End':'','Selected_Location':'','Selected_Day_Count':'','Distance_Miles':'','Original_Available_Seats':'','Remaining_Seats_After_Reservation':'','Prerequisite_Status':'','Prerequisite_Scheduled_Start':'','Conflict_Detail':'','Satisfied_By_Event_Key':'','Satisfied_By_Session_WID':'','Disposition':'','Explanation':'','Workday_Ready':'No','Override':'No','Override_Reason':''})
        satisfying_titles=satisfying(r['Training_Title']); person_hist=hist_by_person.get(r['Person_Key'],[])
        completed=[h for h in person_hist if norm(h.get('Record Learning Content')) in satisfying_titles and norm(h.get('Record Completion Status'))=='completed']
        if completed:
            h=max(completed,key=lambda x:dt(x.get('Record Completion Date')) or datetime.min); observed=clean(h.get('Record Learning Content')); r['Disposition']='PREVIOUSLY_COMPLETED' if norm(observed)==norm(r['Training_Title']) else 'EQUIVALENT_COMPLETION'; r['Completion_Date']=h.get('Record Completion Date'); r['Completion_Training']=observed
            em=eq_meta.get((norm(r['Training_Title']),norm(observed)))
            if em:
                r['Equivalency_Used']=f"{em['required']} <-> {em['equivalent']}" if em['direction']=='TWO_WAY' else f"{em['equivalent']} -> {em['required']}"; r['Equivalency_Direction']=em['direction']; r['Equivalency_Match_Direction']='FORWARD' if em['forward'] else 'REVERSE'; r['Explanation']=f"Historical completion of '{observed}' satisfies '{r['Training_Title']}' via {em['direction']} equivalency." 
            else: r['Explanation']='Historical completion satisfies requirement. No reassignment required.'
            continue
        enr=[h for h in person_hist if norm(h.get('Record Learning Content')) in satisfying_titles and norm(h.get('Registration Status')) in ACTIVE_REGISTRATION_STATUSES]
        enr2=[o for o in existing_by_person.get(r['Person_Key'],[]) if norm(o['Training_Title']) in satisfying_titles]
        if enr or enr2:
            observed=''; start=None
            if enr2:
                x=min(enr2,key=lambda o:o['Start_Date']); observed=x['Training_Title']; start=x['Start_Date']
            elif enr:
                x=max(enr,key=lambda h:dt(h.get("Learner's Registration Date")) or datetime.min); observed=clean(x.get('Record Learning Content')); start=dt(x.get('Learner Start Date'))
            r['Existing_Enrollment_Training']=observed; r['Existing_Session_Start']=start; r['Disposition']='ALREADY_ENROLLED'
            em=eq_meta.get((norm(r['Training_Title']),norm(observed)))
            if em:
                r['Equivalency_Used']=f"{em['required']} <-> {em['equivalent']}" if em['direction']=='TWO_WAY' else f"{em['equivalent']} -> {em['required']}"; r['Equivalency_Direction']=em['direction']; r['Equivalency_Match_Direction']='FORWARD' if em['forward'] else 'REVERSE'; r['Explanation']=f"Existing active enrollment in '{observed}' satisfies '{r['Training_Title']}' via {em['direction']} equivalency. Duplicate enrollment suppressed."
            else: r['Explanation']='Existing active registration found. Duplicate enrollment suppressed.'
            continue
        p=norm(r['Scheduling_Policy'])
        if p=='flag_manual': r['Disposition']='MANUAL_SCHEDULING_REQUIRED'; r['Explanation']='Training rule uses FLAG_MANUAL.'; continue
        if not p: r['Disposition']='REVIEW_REQUIRED'; r['Explanation']='Training rule has no scheduling policy.'; continue
        if norm(r['Contingent_Action'])=='escalate': r['Disposition']='REVIEW_REQUIRED'; r['Explanation']='Contingent worker policy requires escalation.'; continue
        r['Disposition']='PENDING_SCHEDULING'

    req_event=defaultdict(dict)
    for r in reqs: req_event[r['Event_Key']][norm(r['Training_Title'])]=r
    same_run={}; session_capacity={s['WID']:s['Available Seats'] for s in sessions}; seat_used=Counter(); session_map={s['WID']:s for s in sessions}
    # Scheduling order: dependency depth first (hard prerequisite), then TARGET_RANGE, then FIRST_AVAILABLE,
    # then class priority for FIRST_AVAILABLE, then anchor date, then event key/title. This preserves the
    # requested policy precedence without breaking prerequisite sequencing.
    def depth(r,memo,stack=None):
        k=norm(r['Training_Title'])
        if k in memo: return memo[k]
        stack=set() if stack is None else set(stack)
        if k in stack: return 999
        p=norm(r.get('Prerequisite'))
        if not p or p not in req_event[r['Event_Key']]: memo[k]=0; return 0
        stack.add(k); memo[k]=1+depth(req_event[r['Event_Key']][p],memo,stack); return memo[k]

    def first_available_hint(r):
        # Earliest potential FIRST_AVAILABLE selection time used only to make
        # priority a tie-breaker when FIRST_AVAILABLE classes collide in time.
        starts=[]
        for s in sessions_by_title.get(norm(r['Training_Title']),[]):
            if norm(s['Availability Status'])!='open' or session_capacity.get(s['WID'],0)<=0: continue
            if r.get('Anchor_Date') and s['start']<r['Anchor_Date']: continue
            if not timing_ok(r,s['start']): continue
            if s['Locations']:
                d=distance(r.get('Physical_Location',''),s['Locations'])
                if d is not None and d<=MAX_DISTANCE_MILES: starts.append(s['start'])
            else:
                starts.append(s['start'])
        return min(starts) if starts else datetime.max

    all_pending=[]
    for e in sorted(events,key=lambda x:(x.get('Anchor_Date') or datetime.max,x['Event_Key'])):
        memo={}; ers=[x for x in req_event[e['Event_Key']].values() if x['Disposition']=='PENDING_SCHEDULING']
        # Hard prerequisite depth is first. Then requested scheduling-policy precedence.
        # Among FIRST_AVAILABLE requirements, lower numeric Priority wins; blank priority is last.
        ers.sort(key=lambda r:(
            depth(r,memo),
            _policy_rank(r['Scheduling_Policy']),
            first_available_hint(r) if _policy_rank(r['Scheduling_Policy'])==1 else datetime.max,
            _priority(r['Priority']) if _policy_rank(r['Scheduling_Policy'])==1 else 10**9,
            norm(r['Training_Title']),r['Event_Key']))
        all_pending.extend(ers)

    reservations=[]
    for r in all_pending:
        if r['Disposition']!='PENDING_SCHEDULING': continue
        prereq=norm(r.get('Prerequisite')); prereq_end=None
        if prereq:
            pr=req_event[r['Event_Key']].get(prereq)
            if pr:
                if pr['Disposition'] in ('PREVIOUSLY_COMPLETED','EQUIVALENT_COMPLETION'): r['Prerequisite_Status']='SATISFIED_BY_COMPLETION'
                elif pr['Disposition']=='PROPOSED_SCHEDULE': prereq_end=dt(pr.get('Selected_End')); r['Prerequisite_Status']='PROPOSED_SCHEDULE'; r['Prerequisite_Scheduled_Start']=pr.get('Selected_Start')
                elif pr['Disposition']=='ALREADY_ENROLLED':
                    # Use any dated existing session as the prerequisite bound.
                    pks=existing_by_person.get(r['Person_Key'],[]); matches=[x for x in pks if norm(x['Training_Title']) in satisfying(pr['Training_Title']) and x['Start_Date']]
                    if matches: prereq_end=max(x['End_Date'] for x in matches); r['Prerequisite_Status']='ALREADY_ENROLLED'; r['Prerequisite_Scheduled_Start']=min(x['Start_Date'] for x in matches)
                    else: r['Disposition']='REVIEW_REQUIRED'; r['Explanation']='Prerequisite is already enrolled but has no usable schedule interval.'; continue
                elif pr['Disposition'] in ('MANUAL_SCHEDULING_REQUIRED','REVIEW_REQUIRED'): r['Disposition']='REVIEW_REQUIRED'; r['Explanation']=f"Prerequisite '{pr['Training_Title']}' is not in a schedulable state ({pr['Disposition']})."; continue
            else:
                ph=hist_by_person.get(r['Person_Key'],[]); titles=satisfying(prereq)
                if any(norm(h.get('Record Learning Content')) in titles and norm(h.get('Record Completion Status'))=='completed' for h in ph): r['Prerequisite_Status']='SATISFIED_BY_PRIOR_COMPLETION'
                else:
                    pexisting=[x for x in existing_by_person.get(r['Person_Key'],[]) if norm(x['Training_Title']) in titles]
                    if pexisting: prereq_end=max(x['End_Date'] for x in pexisting); r['Prerequisite_Status']='SATISFIED_BY_EXISTING_ENROLLMENT'; r['Prerequisite_Scheduled_Start']=min(x['Start_Date'] for x in pexisting)
                    else: r['Disposition']='REVIEW_REQUIRED'; r['Explanation']=f"Prerequisite '{r.get('Prerequisite')}' has no prior completion or existing enrollment."; continue

        skey=(r['Person_Key'],norm(r['Training_Title']))
        if skey in same_run:
            prior=same_run[skey]; pstart=dt(prior['Selected_Start']); pend=dt(prior['Selected_End'])
            if pstart and (not r['Anchor_Date'] or pstart>=r['Anchor_Date']) and timing_ok(r,pstart) and (not prereq_end or pstart>prereq_end):
                for k in ('Selected_Session_WID','Selected_Reference_ID','Selected_Start','Selected_End','Selected_Location','Selected_Day_Count','Distance_Miles'): r[k]=prior.get(k,'')
                r['Disposition']='SATISFIED_BY_SAME_RUN_ASSIGNMENT'; r['Satisfied_By_Event_Key']=prior['Event_Key']; r['Satisfied_By_Session_WID']=prior['Selected_Session_WID']; r['Explanation']=f"Same person/course was already assigned under {prior['Event_Key']}; this requirement is satisfied by that single enrollment."; continue
            r['Disposition']='REVIEW_REQUIRED'; r['Explanation']='The person already has the same course assigned in this run, but that assignment does not satisfy this staffing event\'s timing/prerequisite rule. No duplicate enrollment created.'; continue

        candidate=[]; overlap_count=0; all_struct=[]
        for s in sessions_by_title.get(norm(r['Training_Title']),[]):
            if norm(s['Availability Status'])!='open': continue
            if session_capacity.get(s['WID'],0)-seat_used[s['WID']]<=0: continue
            first=s['start']; last=s['end']
            if r['Anchor_Date'] and first<r['Anchor_Date']: continue
            if prereq_end and first<=prereq_end: continue
            if not timing_ok(r,first): continue
            all_struct.append(s)
            conflicts=[]
            for d in s['days']:
                for b in busy.get(r['Person_Key'],[]):
                    if overlap(d['start'],d['end'],b['start'],b['end']): conflicts.append(b)
            if conflicts:
                overlap_count+=1; continue
            candidate.append(s)
        if not candidate:
            r['Disposition']='REVIEW_REQUIRED'
            if overlap_count and all_struct:
                # Include up to 5 blocking classes, de-duplicated by title/time.
                details=[]
                for s in all_struct:
                    for d in s['days']:
                        for b in busy.get(r['Person_Key'],[]):
                            if overlap(d['start'],d['end'],b['start'],b['end']):
                                x=f"{b['title']} ({b['start']:%m/%d/%Y %I:%M %p}-{b['end']:%I:%M %p})"
                                if x not in details: details.append(x)
                r['Conflict_Detail']='; '.join(details[:5]); r['Explanation']='All otherwise eligible sessions overlap an existing or newly selected class for this employee.'
            else: r['Explanation']='No open session with remaining seats meets timing, prerequisite, location, capacity, and non-overlap criteria.'
            continue

        physical=[]; virtual=[]
        for s in candidate:
            if s['Locations']:
                d=distance(r['Physical_Location'],s['Locations'])
                if d is not None and d<=MAX_DISTANCE_MILES: physical.append((d,s))
            else: virtual.append(s)
        chosen=None; chosen_dist=''; explanation=''
        if physical:
            nearest=min(d for d,_ in physical); near=[(d,s) for d,s in physical if abs(d-nearest)<0.01]
            # Location first, then earliest qualifying offering at that location.
            chosen_dist,chosen=min(near,key=lambda x:(x[1]['start'],x[1]['end'],x[1]['WID'])); explanation=f"Selected nearest eligible physical location ({chosen_dist:.1f} miles)."
        elif virtual:
            chosen=min(virtual,key=lambda s:(s['start'],s['end'],s['WID'])); explanation=f"No eligible physical session within {MAX_DISTANCE_MILES:.0f} miles; selected virtual session."
        else:
            mapped=[]
            for s in candidate:
                if s['Locations']:
                    d=distance(r['Physical_Location'],s['Locations'])
                    if d is not None: mapped.append(d)
            if mapped:
                r['Disposition']='REVIEW_REQUIRED'; r['Explanation']=f"Nearest eligible physical session is {min(mapped):.1f} miles away, over the {MAX_DISTANCE_MILES:.0f}-mile limit, and no virtual session exists."; continue
            r['Disposition']='REVIEW_REQUIRED'; r['Explanation']='Eligible physical sessions exist but locations cannot be mapped and no virtual session exists.'; continue
        seat_used[chosen['WID']]+=1
        r.update({'Disposition':'PROPOSED_SCHEDULE','Selected_Session_WID':chosen['WID'],'Selected_Reference_ID':chosen['Reference ID'],'Selected_Start':chosen['start'],'Selected_End':chosen['end'],'Selected_Location':chosen['Locations'] or 'Virtual','Selected_Day_Count':chosen['day_count'],'Distance_Miles':round(chosen_dist,1) if chosen_dist!='' else '', 'Original_Available_Seats':chosen['Available Seats'],'Remaining_Seats_After_Reservation':session_capacity[chosen['WID']]-seat_used[chosen['WID']], 'Workday_Ready':'Yes','Explanation':explanation})
        if prereq_end: r['Explanation']+=f" Prerequisite scheduled earlier (through {prereq_end:%m/%d/%Y %I:%M %p})."
        same_run[skey]=r
        for d in chosen['days']: busy[r['Person_Key']].append({'start':d['start'],'end':d['end'],'title':r['Training_Title'],'source':'SELECTED'})
        reservations.append({'Session_WID':chosen['WID'],'Training_Title':chosen['Title'],'Start_Date':chosen['start'],'End_Date':chosen['end'],'Day_Count':chosen['day_count'],'Location':chosen['Locations'] or 'Virtual','Employee_ID':r['Employee_ID'],'WID':r['WID'],'Person_Key':r['Person_Key'],'Event_Key':r['Event_Key'],'Seats_Before':session_capacity[chosen['WID']]-seat_used[chosen['WID']]+1,'Seats_After':session_capacity[chosen['WID']]-seat_used[chosen['WID']]})
    for r in reqs:
        if r['Disposition']=='PENDING_SCHEDULING': r['Disposition']='REVIEW_REQUIRED'; r['Explanation']='Scheduling engine could not resolve this requirement.'
    req_cols=['Event_Key','Event_Type','Employee_ID','WID','Person_Key','Worker_Name','Worker_Type','Traveler_Designation','Anchor_Date','Job_Code','Position_Title','Cost_Center_ID','Cost_Center_Title','Sup_Org_ID','Physical_Location','Hiring_Manager_AD','Work_Email','Home_Email','Rule_ID','Training_Title','Prerequisite','Topic','Scheduling_Policy','Timing_Modifier','Timing_Min','Timing_Max','Priority','Completion_Date','Completion_Training','Existing_Registration_Date','Existing_Session_Start','Existing_Enrollment_Training','Equivalency_Used','Equivalency_Direction','Equivalency_Match_Direction','Selected_Session_WID','Selected_Reference_ID','Selected_Start','Selected_End','Selected_Location','Selected_Day_Count','Distance_Miles','Original_Available_Seats','Remaining_Seats_After_Reservation','Prerequisite_Status','Prerequisite_Scheduled_Start','Conflict_Detail','Satisfied_By_Event_Key','Satisfied_By_Session_WID','Disposition','Explanation','Workday_Ready','Override','Override_Reason']
    reqdf=pd.DataFrame(reqs,columns=req_cols)
    evdf=pd.DataFrame(events)
    exdf=pd.DataFrame(existing_by_person and sum(existing_by_person.values(),[]) or [],columns=['Person_Key','WID','Employee_ID','Training_Title','Registration_Status','Start_Date','End_Date','Location','Required_For_Current_Role','Scheduling_Impact'])
    rdf=pd.DataFrame(reservations)
    return {'events':evdf,'duplicate_events':pd.DataFrame(duplicate),'requirements':reqdf,'existing_workday':exdf,'sessions':pd.DataFrame(sessions),'seat_reservations':rdf,'summary':dict(Counter(reqdf['Disposition'])) if not reqdf.empty else {},'source_counts':{'new_hires':len(nh),'job_changes':len(jc),'staffing_events_raw':len(nh)+len(jc),'staffing_events_canonical':len(events),'duplicate_source_events':len(duplicate),'history':len(history),'sessions':len(sessions),'session_day_rows':len(sessions_raw)}}
