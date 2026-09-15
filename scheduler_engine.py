from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from collections import defaultdict, Counter
from datetime import datetime, date
from io import BytesIO
import math, re
import pandas as pd
import openpyxl

ACTIVE_REGISTRATION_STATUSES={'enrolled','registered','in progress','not started'}
MAX_DISTANCE_MILES=130.0


def clean(v):
    if v is None: return ''
    return str(v).strip() if not isinstance(v,(datetime,date)) else v

def norm(v):
    if v is None: return ''
    return re.sub(r'\s+',' ',str(v).strip().lower())

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
            except Exception: pass
    return None

def _open_wb(source):
    if hasattr(source,'seek'):
        source.seek(0)
        return openpyxl.load_workbook(source,read_only=True,data_only=True)
    return openpyxl.load_workbook(source,read_only=True,data_only=True)

def read_sheet(source, sheet=None, header_row=None, required_headers=None):
    wb=_open_wb(source)
    ws=wb[sheet] if sheet else wb.worksheets[0]
    rows_iter=ws.iter_rows(values_only=True)
    if header_row is not None:
        header_vals=None
        for i,row in enumerate(rows_iter,1):
            if i==header_row:
                header_vals=row; break
    else:
        req={norm(x) for x in (required_headers or [])}
        header_vals=None
        for i,row in enumerate(rows_iter,1):
            if req.issubset({norm(v) for v in row}):
                header_vals=row; header_row=i; break
    if header_vals is None:
        raise ValueError(f'Could not find expected header row: {required_headers}')
    headers=[]; seen=Counter()
    for c,v in enumerate(header_vals,1):
        h=str(clean(v)) if clean(v)!='' else f'_blank_{c}'
        seen[h]+=1
        if seen[h]>1: h=f'{h}_{seen[h]}'
        headers.append(h)
    out=[]
    for row in rows_iter:
        if not any(v not in (None,'') for v in row): continue
        vals=list(row)+[None]*(len(headers)-len(row))
        out.append(dict(zip(headers,vals[:len(headers)])))
    return out

def haversine_miles(lat1,lon1,lat2,lon2):
    R=3958.7613
    p1=math.radians(lat1); p2=math.radians(lat2)
    dphi=math.radians(lat2-lat1); dl=math.radians(lon2-lon1)
    a=math.sin(dphi/2)**2+math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 2*R*math.asin(math.sqrt(a))

def contingent_action(worker_type, traveler):
    if norm(worker_type)!='contingent worker': return 'STANDARD SCHEDULE'
    tr=norm(traveler)
    if tr=='trav i': return 'STANDARD SCHEDULE'
    if tr=='trav d': return 'ESCALATE'
    return 'NO SCHEDULE'

def timing_ok(rec, start_dt):
    anchor=rec.get('Anchor_Date')
    if not anchor or not start_dt: return True
    pol=norm(rec.get('Scheduling_Policy'))
    if pol!='target_range': return start_dt >= anchor
    mn=rec.get('Timing_Min'); mx=rec.get('Timing_Max'); mod=norm(rec.get('Timing_Modifier'))
    try: mn=float(mn) if mn not in (None,'') else None
    except Exception: mn=None
    try: mx=float(mx) if mx not in (None,'') else None
    except Exception: mx=None
    delta=(start_dt.date()-anchor.date()).days
    if mod=='between': return (mn is None or delta>=mn) and (mx is None or delta<=mx)
    if mod=='after':
        t=mn if mn is not None else mx
        return t is None or delta>=t
    if mod=='before':
        t=mx if mx is not None else mn
        return t is None or delta<=t
    return (mn is None or delta>=mn) and (mx is None or delta<=mx)

def run_scheduler(new_hire_file, job_change_file, history_file, sessions_file, rules_df, locations_df, equivalencies_df=None, orientation_file=None):
    new_hires=read_sheet(new_hire_file,required_headers=['Employee ID','WID','Hire Date','Candidate Cost Center ID']) if new_hire_file else []
    job_changes=read_sheet(job_change_file,required_headers=['Employee ID','WID','Effective Date','Job Code - Proposed']) if job_change_file else []
    history=read_sheet(history_file,required_headers=['Learning Participant','Employee ID','WID','Record Learning Content','Record Completion Status'])
    sessions=read_sheet(sessions_file,required_headers=['Learning Content Type','Title','Reference ID','Start Date','Available Seats','WID'])
    orientation=read_sheet(orientation_file,required_headers=['Employee ID','Enrolled Course Offering','Registration Status','Start Date']) if orientation_file else []

    # location master
    loc_by_name={}
    for _,r in locations_df.iterrows():
        name=clean(r.get('location'))
        if not name: continue
        try: lat=float(r.get('latitude')); lon=float(r.get('longitude'))
        except Exception: lat=lon=None
        loc_by_name[norm(name)]={'name':name,'lat':lat,'lon':lon,'inactive':clean(r.get('inactive'))}
    def dist(a,b):
        x=loc_by_name.get(norm(a)); y=loc_by_name.get(norm(b))
        if not x or not y or x['lat'] is None or y['lat'] is None: return None
        return haversine_miles(x['lat'],x['lon'],y['lat'],y['lon'])

    # WID identity from history
    emp_to_hist_wid={}
    for h in history:
        if id_norm(h.get('Employee ID')) and id_norm(h.get('WID')):
            emp_to_hist_wid[id_norm(h.get('Employee ID'))]=id_norm(h.get('WID'))

    events=[]
    for i,r in enumerate(new_hires,start=1):
        eid=id_norm(r.get('Employee ID')); source_wid=id_norm(r.get('WID'))
        events.append({'Event_Key':f'NH-{i}-{eid}','Event_Type':'NEW_HIRE','Employee_ID':eid,'WID':emp_to_hist_wid.get(eid,source_wid),'Source_WID':source_wid,
            'Worker_Name':clean(r.get('Worker')),'Worker_Type':clean(r.get('Worker Type')),'Traveler_Designation':clean(r.get('Traveler Designation')),
            'Anchor_Date':dt(r.get('Hire Date')),'Job_Code':clean(r.get('Candidate Position ID')),'Position_Title':clean(r.get('New Hire Position Title')),
            'Cost_Center_ID':clean(r.get('Candidate Cost Center ID')),'Cost_Center_Title':clean(r.get('Candidate Department Name')),
            'Sup_Org_ID':clean(r.get('Supervisory Organization ID')),'Physical_Location':clean(r.get('Worker Physical Location')),
            'Hiring_Manager_AD':clean(r.get('Hiring Manager AD (Username)')),'Work_Email':clean(r.get('Email - Primary Work')),'Home_Email':clean(r.get('Email - Primary Home'))})
    for i,r in enumerate(job_changes,start=1):
        eid=id_norm(r.get('Employee ID')); source_wid=id_norm(r.get('WID'))
        events.append({'Event_Key':f'JC-{i}-{eid}','Event_Type':'JOB_CHANGE','Employee_ID':eid,'WID':emp_to_hist_wid.get(eid,source_wid),'Source_WID':source_wid,
            'Worker_Name':clean(r.get('Worker')),'Worker_Type':clean(r.get('Worker Type')),'Traveler_Designation':'',
            'Anchor_Date':dt(r.get('Effective Date')),'Job_Code':clean(r.get('Job Code - Proposed')),'Position_Title':clean(r.get('Job Title - Proposed')),
            'Cost_Center_ID':clean(r.get('Cost Center ID - Requested')),'Cost_Center_Title':clean(r.get('Cost Center - Proposed')),
            'Sup_Org_ID':clean(r.get('Sup Org ID - Requested')),'Physical_Location':clean(r.get('Location - Proposed')),
            'Hiring_Manager_AD':clean(r.get('Hiring Manager AD (Username')) or clean(r.get('Hiring Manager AD (Username)')),
            'Work_Email':clean(r.get('Candidate Email')),'Home_Email':''})
    for ev in events: ev['Contingent_Action']=contingent_action(ev['Worker_Type'],ev['Traveler_Designation'])

    # rule list
    rule_rows=[]
    for _,r in rules_df.iterrows():
        title=clean(r.get('training_title'))
        if not title: continue
        rule_rows.append({'Rule_ID':r.get('id'),'Job_Code':clean(r.get('job_code')),'Cost_Center':clean(r.get('cost_center')),'Sup_Org':clean(r.get('supervisory_org')),
            'Training_Title':title,'Prerequisite':clean(r.get('prerequisite')),'Topic':clean(r.get('topic')),'Scheduling_Policy':clean(r.get('scheduling_policy')),
            'Timing_Modifier':clean(r.get('timing_modifier')),'Timing_Min':r.get('timing_min'),'Timing_Max':r.get('timing_max')})

    # equivalency mapping required -> set equivalent titles
    eq=defaultdict(set)
    if equivalencies_df is not None and not equivalencies_df.empty:
        for _,r in equivalencies_df.iterrows():
            req=norm(r.get('required_training_title')); equiv=norm(r.get('equivalent_training_title'))
            if req and equiv: eq[req].add(equiv)

    hist_by_wid=defaultdict(list); hist_by_eid=defaultdict(list)
    for h in history:
        if id_norm(h.get('WID')): hist_by_wid[id_norm(h.get('WID'))].append(h)
        if id_norm(h.get('Employee ID')): hist_by_eid[id_norm(h.get('Employee ID'))].append(h)
    orient_by_eid=defaultdict(list)
    for o in orientation:
        if id_norm(o.get('Employee ID')): orient_by_eid[id_norm(o.get('Employee ID'))].append(o)
    sessions_by_title=defaultdict(list)
    for s in sessions:
        if clean(s.get('Title')): sessions_by_title[norm(s.get('Title'))].append(s)

    # requirements additive with same-title collapse per event
    reqs=[]; seen=set()
    for ev in events:
        if norm(ev['Contingent_Action'])=='no schedule': continue
        for rr in rule_rows:
            if norm(rr['Job_Code'])!=norm(ev['Job_Code']) or norm(rr['Cost_Center'])!=norm(ev['Cost_Center_ID']): continue
            if rr['Sup_Org'] and norm(rr['Sup_Org'])!=norm(ev['Sup_Org_ID']): continue
            key=(ev['Event_Key'],norm(rr['Training_Title']))
            if key in seen: continue
            seen.add(key); reqs.append({**ev,**rr})

    def skey(s):
        return id_norm(s.get('WID')) or f"{clean(s.get('Reference ID'))}|{clean(s.get('Start Date'))}|{norm(s.get('Title'))}"
    opening={}; remaining={}
    for s in sessions:
        try: seats=max(0,int(float(s.get('Available Seats') or 0)))
        except Exception: seats=0
        k=skey(s); opening[k]=max(opening.get(k,0),seats); remaining[k]=opening[k]
    reservations=[]

    # base dispositions
    for rec in reqs:
        rec.update({'Completion_Date':'','Completion_Training':'','Existing_Registration_Date':'','Existing_Session_Start':'','Selected_Session_WID':'','Selected_Reference_ID':'','Selected_Start':'','Selected_End':'','Selected_Location':'','Distance_Miles':'','Original_Available_Seats':'','Remaining_Seats_After_Reservation':'','Prerequisite_Status':'','Prerequisite_Scheduled_Start':'','Disposition':'','Explanation':'','Workday_Ready':'No','Override':'No','Override_Reason':''})
        title_n=norm(rec['Training_Title']); person_hist=hist_by_wid.get(rec['WID'],[]) or hist_by_eid.get(rec['Employee_ID'],[])
        satisfy_titles={title_n}|eq.get(title_n,set())
        completed=[h for h in person_hist if norm(h.get('Record Learning Content')) in satisfy_titles and norm(h.get('Record Completion Status'))=='completed']
        if completed:
            latest=max(completed,key=lambda h:dt(h.get('Record Completion Date')) or datetime.min)
            rec['Disposition']='PREVIOUSLY_COMPLETED' if norm(latest.get('Record Learning Content'))==title_n else 'EQUIVALENT_COMPLETION'
            rec['Completion_Date']=latest.get('Record Completion Date'); rec['Completion_Training']=clean(latest.get('Record Learning Content'))
            rec['Explanation']='Historical completion satisfies requirement. No reassignment required.'
            continue
        existing=[h for h in person_hist if norm(h.get('Record Learning Content'))==title_n and norm(h.get('Registration Status')) in ACTIVE_REGISTRATION_STATUSES]
        orient_existing=[o for o in orient_by_eid.get(rec['Employee_ID'],[]) if norm(o.get('Enrolled Course Offering'))==title_n and norm(o.get('Registration Status')) in ACTIVE_REGISTRATION_STATUSES]
        if existing or orient_existing:
            starts=[dt(o.get('Start Date')) for o in orient_existing if dt(o.get('Start Date'))]
            if starts: rec['Existing_Session_Start']=min(starts)
            dates=[dt(h.get("Learner's Registration Date")) for h in existing if dt(h.get("Learner's Registration Date"))]
            if dates: rec['Existing_Registration_Date']=max(dates)
            rec['Disposition']='ALREADY_ENROLLED'; rec['Explanation']='Existing active registration found. Duplicate enrollment suppressed.'; continue
        if norm(rec['Contingent_Action'])=='escalate':
            rec['Disposition']='REVIEW_REQUIRED'; rec['Explanation']='Contingent worker policy requires escalation.'; continue
        pol=norm(rec['Scheduling_Policy'])
        if pol=='flag_manual': rec['Disposition']='MANUAL_SCHEDULING_REQUIRED'; rec['Explanation']='Training rule uses FLAG_MANUAL.'; continue
        if not pol: rec['Disposition']='REVIEW_REQUIRED'; rec['Explanation']='Training rule has no scheduling policy.'; continue
        rec['Disposition']='PENDING_SCHEDULING'

    req_event=defaultdict(dict)
    for r in reqs: req_event[r['Event_Key']][norm(r['Training_Title'])]=r
    def depth(rec,memo,active=None):
        k=norm(rec['Training_Title'])
        if k in memo: return memo[k]
        active=set() if active is None else set(active)
        if k in active: return 999
        p=norm(rec.get('Prerequisite'))
        if not p or p not in req_event[rec['Event_Key']]: memo[k]=0; return 0
        active.add(k); memo[k]=1+depth(req_event[rec['Event_Key']][p],memo,active); return memo[k]

    for ev in sorted(events,key=lambda e:(e.get('Anchor_Date') or datetime.max,e['Event_Key'])):
        er=list(req_event.get(ev['Event_Key'],{}).values()); memo={}; er.sort(key=lambda r:(depth(r,memo),norm(r['Training_Title'])))
        for rec in er:
            if rec['Disposition']!='PENDING_SCHEDULING': continue
            prereq=norm(rec.get('Prerequisite')); prereq_bound=None
            if prereq:
                pr=req_event[rec['Event_Key']].get(prereq)
                if pr:
                    rec['Prerequisite_Status']=pr['Disposition']
                    if pr['Disposition'] in ('PREVIOUSLY_COMPLETED','EQUIVALENT_COMPLETION'): rec['Prerequisite_Status']='SATISFIED_BY_COMPLETION'
                    elif pr['Disposition']=='PROPOSED_SCHEDULE' and dt(pr.get('Selected_Start')): prereq_bound=dt(pr['Selected_Start']); rec['Prerequisite_Scheduled_Start']=pr['Selected_Start']
                    elif pr['Disposition']=='ALREADY_ENROLLED' and dt(pr.get('Existing_Session_Start')): prereq_bound=dt(pr['Existing_Session_Start']); rec['Prerequisite_Scheduled_Start']=pr['Existing_Session_Start']
                    else:
                        rec['Disposition']='REVIEW_REQUIRED'; rec['Explanation']=f"Prerequisite '{rec.get('Prerequisite')}' is not in a schedulable state/date ({pr['Disposition']})."; continue
                else:
                    person_hist=hist_by_wid.get(rec['WID'],[]) or hist_by_eid.get(rec['Employee_ID'],[])
                    ph=[h for h in person_hist if norm(h.get('Record Learning Content'))==prereq and norm(h.get('Record Completion Status'))=='completed']
                    if ph: rec['Prerequisite_Status']='SATISFIED_BY_PRIOR_COMPLETION'
                    else:
                        pe=[o for o in orient_by_eid.get(rec['Employee_ID'],[]) if norm(o.get('Enrolled Course Offering'))==prereq and norm(o.get('Registration Status')) in ACTIVE_REGISTRATION_STATUSES and dt(o.get('Start Date'))]
                        if pe:
                            prereq_bound=min(dt(o.get('Start Date')) for o in pe); rec['Prerequisite_Status']='SATISFIED_BY_EXISTING_ENROLLMENT'; rec['Prerequisite_Scheduled_Start']=prereq_bound
                        else:
                            rec['Disposition']='REVIEW_REQUIRED'; rec['Prerequisite_Status']='PREREQUISITE_NOT_FOUND'; rec['Explanation']=f"Prerequisite '{rec.get('Prerequisite')}' has no prior completion or dated active enrollment."; continue

            all_s=sessions_by_title.get(norm(rec['Training_Title']),[])
            if not all_s:
                rec['Disposition']='REVIEW_REQUIRED'; rec['Explanation']='No session title matches this training requirement.'; continue
            candidates=[]
            for s in all_s:
                st=dt(s.get('Start Date')); k=skey(s)
                if norm(s.get('Availability Status'))!='open' or not st or remaining.get(k,0)<=0: continue
                if rec['Anchor_Date'] and st < rec['Anchor_Date']: continue
                if prereq_bound and st <= prereq_bound: continue
                if not timing_ok(rec,st): continue
                candidates.append(s)
            if not candidates:
                possible=[]
                for s in all_s:
                    st=dt(s.get('Start Date'))
                    if norm(s.get('Availability Status'))=='open' and st and (not rec['Anchor_Date'] or st>=rec['Anchor_Date']) and (not prereq_bound or st>prereq_bound) and timing_ok(rec,st): possible.append(s)
                exhausted=[s for s in possible if opening.get(skey(s),0)>0 and remaining.get(skey(s),0)<=0]
                rec['Disposition']='REVIEW_REQUIRED'; rec['Explanation']='Eligible session capacity was consumed by earlier assignments in this run.' if exhausted else 'No open session with remaining seats meets date/policy/prerequisite criteria.'; continue
            physical=[s for s in candidates if clean(s.get('Locations'))]; virtual=[s for s in candidates if not clean(s.get('Locations'))]
            mapped=[]; unmapped=[]
            for s in physical:
                d=dist(rec['Physical_Location'],clean(s.get('Locations')))
                (mapped if d is not None else unmapped).append((d,s) if d is not None else s)
            within=[(d,s) for d,s in mapped if d<=MAX_DISTANCE_MILES]
            chosen=None; chosen_dist=''; explanation=''
            if within:
                nearest=min(d for d,_ in within)
                nearest_loc=norm(min((s for d,s in within if abs(d-nearest)<0.01),key=lambda s:dt(s.get('Start Date'))).get('Locations'))
                eligible=[(d,s) for d,s in within if norm(s.get('Locations'))==nearest_loc]
                chosen_dist,chosen=min(eligible,key=lambda x:dt(x[1].get('Start Date'))); chosen_dist=round(chosen_dist,1)
                explanation=f'Selected earliest eligible physical session at nearest location ({chosen_dist:.1f} miles).'
            elif virtual:
                chosen=min(virtual,key=lambda s:dt(s.get('Start Date')))
                if mapped: explanation=f'No eligible physical session is within {MAX_DISTANCE_MILES:.0f} miles; selected virtual session.'
                else: explanation='No eligible mapped physical session; selected virtual session.'
            elif mapped:
                nearest=min(d for d,_ in mapped); rec['Disposition']='REVIEW_REQUIRED'; rec['Explanation']=f'Nearest eligible physical session is {nearest:.1f} miles away, over the {MAX_DISTANCE_MILES:.0f}-mile limit, and no virtual session exists.'; continue
            else:
                rec['Disposition']='REVIEW_REQUIRED'; rec['Explanation']='Eligible physical sessions exist but their locations cannot be mapped, and no virtual session exists.'; continue
            k=skey(chosen); before=remaining.get(k,0)
            if before<=0:
                rec['Disposition']='REVIEW_REQUIRED'; rec['Explanation']='Selected session lost its last remaining seat during processing.'; continue
            remaining[k]=before-1
            rec.update({'Disposition':'PROPOSED_SCHEDULE','Selected_Session_WID':id_norm(chosen.get('WID')),'Selected_Reference_ID':clean(chosen.get('Reference ID')),'Selected_Start':chosen.get('Start Date'),'Selected_End':chosen.get('End Date'),'Selected_Location':clean(chosen.get('Locations')) or 'Virtual','Distance_Miles':chosen_dist,'Original_Available_Seats':opening.get(k,0),'Remaining_Seats_After_Reservation':remaining[k],'Workday_Ready':'Yes','Explanation':explanation})
            if prereq_bound: rec['Explanation'] += f" Prerequisite scheduled earlier ({prereq_bound.strftime('%m/%d/%Y')})."
            reservations.append({'Session_Key':k,'Training_Title':rec['Training_Title'],'Start_Date':chosen.get('Start Date'),'Location':clean(chosen.get('Locations')) or 'Virtual','Employee_ID':rec['Employee_ID'],'Event_Key':rec['Event_Key'],'Seats_Before':before,'Seats_After':remaining[k]})

    for r in reqs:
        if r['Disposition']=='PENDING_SCHEDULING': r['Disposition']='REVIEW_REQUIRED'; r['Explanation']='Scheduling engine could not resolve this requirement.'
    return {
        'events':pd.DataFrame(events),
        'requirements':pd.DataFrame(reqs),
        'sessions':pd.DataFrame(sessions),
        'seat_reservations':pd.DataFrame(reservations),
        'summary':dict(Counter(r['Disposition'] for r in reqs)),
        'source_counts':{'new_hires':len(new_hires),'job_changes':len(job_changes),'history':len(history),'sessions':len(sessions)},
        'warning':'New Hire source has no explicit Job Code field; Candidate Position ID is used for matching.'
    }

def apply_manual_override(requirements_df, row_index, selected_session, reason):
    df=requirements_df.copy()
    if not reason or not str(reason).strip(): raise ValueError('Override reason is required.')
    for key,val in selected_session.items():
        if key in df.columns: df.at[row_index,key]=val
    df.at[row_index,'Disposition']='PROPOSED_SCHEDULE'
    df.at[row_index,'Workday_Ready']='Yes'
    df.at[row_index,'Override']='Yes'
    df.at[row_index,'Override_Reason']=str(reason).strip()
    df.at[row_index,'Explanation']='Manual override: '+str(reason).strip()
    return df
