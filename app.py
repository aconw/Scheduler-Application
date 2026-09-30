from pathlib import Path
from io import BytesIO
from zipfile import ZipFile,ZIP_DEFLATED
import pandas as pd
import openpyxl
import streamlit as st
from scheduler_engine import run_scheduler, read_sheet, norm, clean

ROOT=Path(__file__).parent; CONFIG=ROOT/'config'/'Scheduler_Configuration.xlsx'; ASSETS=ROOT/'assets'
st.set_page_config(page_title='Class Scheduling Application',page_icon='📚',layout='wide')
st.title('Class Scheduling Application')
st.caption('v2.9.3 • Strict cleanup of v2.9.2 • Diagnostics + Workday duplicate safeguards')

def rb(x):
    if x is None:return None
    if isinstance(x,(bytes,bytearray)):return bytes(x)
    if hasattr(x,'getvalue'):return bytes(x.getvalue())
    x.seek(0);return x.read()

def config_frames(raw=None):
    sheets=pd.read_excel(BytesIO(raw) if raw else CONFIG,sheet_name=None,engine='openpyxl')
    req={'Training Rules','Locations','Equivalencies','Manual Routing'}
    miss=req-set(sheets)
    if miss: raise ValueError(f'Missing configuration sheets: {sorted(miss)}')
    rules=sheets['Training Rules'].copy(); loc=sheets['Locations'].copy(); eq=sheets['Equivalencies'].copy(); route=sheets['Manual Routing'].copy()
    if 'active' in rules: rules=rules[rules.active.fillna(1).astype(str).str.strip().str.lower().isin(['1','1.0','true','yes','y'])]
    if 'active' in eq: eq=eq[eq.active.fillna(1).astype(str).str.strip().str.lower().isin(['1','1.0','true','yes','y'])]
    if 'active' in route: route=route[route.active.fillna(1).astype(str).str.strip().str.lower().isin(['1','1.0','true','yes','y'])]
    for d in (rules,loc,eq,route):
        for c in d.columns:d[c]=d[c].fillna('')
    if 'relationship_direction' not in eq: eq['relationship_direction']='ONE_WAY'
    return rules,loc,eq,route

def validate(raw, label, required, optional=None):
    if not raw:return (False,'Not uploaded',0)
    try:
        rows=read_sheet(raw, required_headers=required, optional_headers=optional or [])
        import openpyxl
        wb=openpyxl.load_workbook(BytesIO(raw),read_only=False,data_only=True)
        sheets=[]
        for ws in wb.worksheets:
            sheets.append(f'{ws.title} ({ws.max_row:,}x{ws.max_column:,})')
        return True,f'{len(raw):,} bytes • {len(rows):,} data rows • Worksheets: {"; ".join(sheets)}',len(rows)
    except Exception as e:return False,f'{label} validation failed: {e}',0

def workday_file(req, contingent=False):
    f=ASSETS/('Enroll_In_Learning_Content_vContingent_Worker_ID.xlsx' if contingent else 'Enroll_In_Learning_Content_vEmployee.xlsx')
    wb=openpyxl.load_workbook(f);ws=wb['Enroll In Learning Content']
    r=req[(req.Workday_Ready=='Yes')].copy()
    r=r[r.Worker_Type.map(norm).eq('contingent worker')] if contingent else r[~r.Worker_Type.map(norm).eq('contingent worker')]
    # Defense in depth: never export the same person/course or the same person/session twice.
    # This is intentionally independent of the scheduler's upstream duplicate protections.
    r['_course_key']=r['Training_Title'].map(lambda v:norm(v))
    r['_session_key']=r['Selected_Session_WID'].map(lambda v:id_norm(v))
    r=r.sort_values(['Person_Key','Selected_Start','Event_Key'],na_position='last')
    r=r.drop_duplicates(['Person_Key','_course_key'],keep='first')
    r=r.drop_duplicates(['Person_Key','_session_key'],keep='first')
    for i,(_,x) in enumerate(r.iterrows(),6):
        ws.cell(i,1,f"{x.Event_Key}-{i-5}"); ws.cell(i,2,str(x.Selected_Session_WID)); ws.cell(i,3,str(x.Current_Employee_ID if 'Current_Employee_ID' in x.index else x.Employee_ID)); ws.cell(i,4,'Y')
    b=BytesIO();wb.save(b);return b.getvalue(),len(r)

def audit_file(res):
    req=res['requirements']; out=BytesIO()
    selected=req[req.Disposition=='PROPOSED_SCHEDULE'].copy()
    with pd.ExcelWriter(out,engine='openpyxl') as w:
        selected.to_excel(w,index=False,sheet_name='Selected Sessions')
        req.to_excel(w,index=False,sheet_name='Requirement Audit')
        res['existing_workday'].to_excel(w,index=False,sheet_name='Existing Workday Training')
        res['duplicate_events'].to_excel(w,index=False,sheet_name='Duplicate Source Events')
        res['seat_reservations'].to_excel(w,index=False,sheet_name='Seat Audit')
        pd.DataFrame([res['summary']]).to_excel(w,index=False,sheet_name='Run Summary')
    out.seek(0);return out.getvalue()

def package(res):
    ef,en=workday_file(res['requirements'],False); cf,cn=workday_file(res['requirements'],True); af=audit_file(res)
    z=BytesIO()
    with ZipFile(z,'w',ZIP_DEFLATED) as zz:
        zz.writestr('Workday/Enroll_In_Learning_Content_Employees.xlsx',ef)
        if cn:zz.writestr('Workday/Enroll_In_Learning_Content_Contingent_Workers.xlsx',cf)
        zz.writestr('Scheduling_Results_and_Audit.xlsx',af)
        zz.writestr('README.txt',f'v2.9.3 export. Employee rows: {en}. Contingent worker rows: {cn}.\n')
    return z.getvalue(),en,cn

def eml_bytes(to_email, subject, body):
    from email.message import EmailMessage
    m=EmailMessage(); m['To']=to_email or ''; m['Subject']=subject; m['X-Unsent']='1'; m.set_content(body); return m.as_bytes()

def email_drafts(res, routing):
    req=res['requirements']; route={norm(r.get('training_title')):clean(r.get('recipient_email')) for _,r in routing.iterrows() if clean(r.get('training_title'))}
    from zipfile import ZipFile,ZIP_DEFLATED
    from io import BytesIO
    z=BytesIO(); manager_count=0; manual_count=0
    with ZipFile(z,'w',ZIP_DEFLATED) as zz:
        # One draft per employee/staffing event with selected sessions.
        for ek in sorted(req.loc[req['Disposition']=='PROPOSED_SCHEDULE','Event_Key'].astype(str).unique()):
            rows=req[req['Event_Key'].astype(str)==ek]
            if rows.empty: continue
            r0=rows.iloc[0]
            lines=[f"Employee: {r0['Worker_Name']}",f"Employee email: {r0.get('Work_Email') or r0.get('Home_Email')}",f"Employee ID: {r0.get('Current_Employee_ID') or r0.get('Employee_ID')}",f"Position: {r0['Position_Title']}",f"Cost Center: {r0['Cost_Center_Title']} ({r0['Cost_Center_ID']})",f"Hire / Position Effective Date: {r0['Anchor_Date']}",f"Hiring Manager: {r0['Hiring_Manager_AD']}",'','Scheduled Training','------------------']
            for _,r in rows[rows['Disposition']=='PROPOSED_SCHEDULE'].sort_values('Selected_Start').iterrows():
                st=pd.to_datetime(r['Selected_Start']) if pd.notna(r['Selected_Start']) else None; en=pd.to_datetime(r['Selected_End']) if pd.notna(r['Selected_End']) else None
                when=''
                if st is not None:
                    when=st.strftime('%m/%d/%Y %I:%M %p')
                    if en is not None: when += ' - ' + en.strftime('%I:%M %p')
                lines.append(f"• {r['Training_Title']} — {when} — {r['Selected_Location']}")
            done=rows[rows['Disposition'].isin(['PREVIOUSLY_COMPLETED','EQUIVALENT_COMPLETION','ALREADY_ENROLLED'])]
            if not done.empty:
                lines += ['','Existing/Completed Training — No New Assignment','----------------------------------------']
                for _,r in done.iterrows():
                    label=r.get('Existing_Enrollment_Training') or r.get('Completion_Training') or r['Training_Title']
                    when=r.get('Existing_Session_Start') or r.get('Completion_Date') or ''
                    lines.append(f"• {label} — {when}")
            zz.writestr(f"Manager Emails/{ek}.eml",eml_bytes(clean(r0['Hiring_Manager_AD']),f"Training Schedule - {r0['Worker_Name']}",'\n'.join(lines))); manager_count+=1
        for idx,r in req[req['Disposition']=='MANUAL_SCHEDULING_REQUIRED'].iterrows():
            recipient=route.get(norm(r['Training_Title']),'')
            body=f"Please schedule the following employee for {r['Training_Title']}.\n\nEmployee: {r['Worker_Name']}\nEmployee email: {r.get('Work_Email') or r.get('Home_Email')}\nEmployee ID: {r.get('Current_Employee_ID') or r.get('Employee_ID')}\nHire / Position Effective Date: {r['Anchor_Date']}\nPosition: {r['Position_Title']}\nCost Center: {r['Cost_Center_Title']} ({r['Cost_Center_ID']})\nHiring Manager: {r['Hiring_Manager_AD']}\n"
            zz.writestr(f"Manual Scheduling/{idx}.eml",eml_bytes(recipient,f"Manual scheduling request - {r['Training_Title']} - {r['Worker_Name']}",body)); manual_count+=1
    return z.getvalue(),manager_count,manual_count

conf_upload=st.file_uploader('Optional updated Scheduler_Configuration.xlsx',type=['xlsx'])
conf_raw=rb(conf_upload)
rules,loc,eq,route=config_frames(conf_raw)
st.success(f'Configuration: {len(rules):,} active rules • {len(loc):,} locations • {len(eq):,} equivalencies')

st.subheader('Upload Reports')
a,b=st.columns(2)
with a:
    nh=rb(st.file_uploader('New Hire Orientation Report',type=['xlsx']))
    jc=rb(st.file_uploader('New / Additional Job Change Report',type=['xlsx']))
with b:
    hist=rb(st.file_uploader('Training History / Learning Transcript',type=['xlsx']))
    ses=rb(st.file_uploader('Learning Content / Available Sessions',type=['xlsx']))
    ori=rb(st.file_uploader('Orientation Schedule — Lesson-Level Detail (required)',type=['xlsx']))

checks=[]
if nh:checks.append(('New Hire',validate(nh,'New Hire',['employee id','wid','hire date','candidate cost center id','candidate position id','worker'])))
if jc:checks.append(('Job Change',validate(jc,'Job Change',['employee id','wid','effective date','job code - proposed','cost center id - requested','worker'])))
checks += [('Training History',validate(hist,'Training History',['employee id','record learning content','record completion status'],['wid','registration status','record completion date',"learner's registration date"])),('Available Sessions',validate(ses,'Available Sessions',['title','reference id','start date','end date','available seats','wid'],['availability status','locations'])),('Existing Orientation Schedule',validate(ori,'Existing Orientation Schedule',['employee id','enrolled course offering','registration status','start date','end date'],['locations','wid']))]
for label,(ok,msg,n) in checks:(st.success if ok else st.error)(f'{label}: {msg}')
ready=bool((nh or jc) and hist and ses and ori and all(ok for _,(ok,_,_) in checks))
if st.button('Run Scheduling',type='primary',disabled=not ready,use_container_width=True):
    with st.spinner('Scheduling with policy precedence, priority, prerequisites, seat reservation, and multi-day conflict checking...'):
        st.session_state.result=run_scheduler(nh,jc,hist,ses,rules,loc,eq,ori)

if st.session_state.get('result') is not None:
    res=st.session_state.result; req=res['requirements']; counts=req.Disposition.value_counts().to_dict()
    st.subheader('Results')
    metrics=[('Requirements',len(req)),('Selected',counts.get('PROPOSED_SCHEDULE',0)),('TARGET_RANGE',int((req.Selection_Phase=='TARGET_RANGE').sum())),('FIRST_AVAILABLE',int((req.Selection_Phase=='FIRST_AVAILABLE').sum())),('Review',counts.get('REVIEW_REQUIRED',0)),('No documentation',counts.get('NO_TRAINING_DOCUMENTATION_FOUND',0)),('Already enrolled',counts.get('ALREADY_ENROLLED',0)),('Completed',counts.get('PREVIOUSLY_COMPLETED',0)+counts.get('EQUIVALENT_COMPLETION',0))]
    cc=st.columns(8)
    for c,(lab,val) in zip(cc,metrics):c.metric(lab,val)
    show=['Worker_Name','Employee_ID','Training_Title','Training_Documentation_Status','Priority','Scheduling_Policy','Selection_Order','Selection_Phase','Selected_Start','Selected_End','Selected_Location','Selected_Day_Count','Selected_Multi_Day','Selected_Session_Days','Disposition','Conflict_Detail','Explanation']
    st.dataframe(req[[c for c in show if c in req.columns]],use_container_width=True,height=460)
    st.subheader('Existing Workday Training')
    ex=res['existing_workday']
    if ex.empty:st.info('No active/upcoming sessions found.')
    else:st.dataframe(ex,use_container_width=True,height=300)
    p,en,cn=package(res)
    st.download_button(f'Download Export Package ({en} employee / {cn} contingent rows)',p,'Class_Scheduling_v2_9_2_Export.zip',mime='application/zip',type='primary')
    ed,mc,man=email_drafts(res,route)
    st.download_button(f'Download Email Drafts ({mc} manager / {man} manual)',ed,'Email_Drafts.zip',mime='application/zip')
