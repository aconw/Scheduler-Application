from __future__ import annotations
from pathlib import Path
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from email.message import EmailMessage
import pandas as pd
import openpyxl
import streamlit as st
from scheduler_engine import run_scheduler, read_sheet, norm, clean

ROOT=Path(__file__).parent
CONFIG=ROOT/'config'/'Scheduler_Configuration.xlsx'
ASSETS=ROOT/'assets'

st.set_page_config(page_title='Class Scheduling Application',page_icon='📚',layout='wide')
st.title('Class Scheduling Application')
st.caption('v2.9.2 • Strict change-control cleanup • TARGET_RANGE first • Priority • Multi-day • Training documentation status')


def rb(x):
    if x is None:return None
    if isinstance(x,(bytes,bytearray)):return bytes(x)
    if hasattr(x,'getvalue'):return bytes(x.getvalue())
    x.seek(0);return x.read()


def active(df):
    if 'active' not in df.columns:return df.copy()
    v=df['active'].fillna(1).astype(str).str.strip().str.lower()
    return df[v.isin(['1','1.0','true','yes','y'])].copy()


def config_frames(raw=None):
    sheets=pd.read_excel(BytesIO(raw) if raw else CONFIG,sheet_name=None,engine='openpyxl')
    missing={'Training Rules','Locations','Equivalencies','Manual Routing'}-set(sheets)
    if missing: raise ValueError(f'Missing configuration sheets: {sorted(missing)}')
    rules=active(sheets['Training Rules']).fillna('')
    loc=sheets['Locations'].fillna('')
    eq=active(sheets['Equivalencies']).fillna('')
    route=active(sheets['Manual Routing']).fillna('')
    if 'relationship_direction' not in eq.columns:eq['relationship_direction']='ONE_WAY'
    return rules,loc,eq,route


def validate(raw,label,required,optional=None):
    if not raw:return False,'Not uploaded'
    try:
        rows=read_sheet(raw,required_headers=required,optional_headers=optional or [])
        return True,f'{len(raw):,} bytes • {len(rows):,} data rows'
    except Exception as e:return False,str(e)


def eml(to,subject,body):
    m=EmailMessage();m['To']=to or '';m['Subject']=subject;m['X-Unsent']='1';m.set_content(body);return m.as_bytes()


def workday_export(req,contingent=False):
    template=ASSETS/('Enroll_In_Learning_Content_vContingent_Worker_ID.xlsx' if contingent else 'Enroll_In_Learning_Content_vEmployee.xlsx')
    wb=openpyxl.load_workbook(template);ws=wb['Enroll In Learning Content']
    rows=req[(req['Workday_Ready']=='Yes')].copy()
    if contingent:rows=rows[rows['Worker_Type'].map(norm).eq('contingent worker')]
    else:rows=rows[~rows['Worker_Type'].map(norm).eq('contingent worker')]
    rows=rows.drop_duplicates(['Person_Key','Training_Title'])
    for i,(_,r) in enumerate(rows.iterrows(),6):
        ws.cell(i,1,f"{r['Event_Key']}-{i-5}");ws.cell(i,2,str(r['Selected_Session_WID']));ws.cell(i,3,str(r['Employee_ID']));ws.cell(i,4,'Y')
    out=BytesIO();wb.save(out);return out.getvalue(),len(rows)


def audit_export(res):
    req=res['requirements'].copy();selected=req[req['Disposition']=='PROPOSED_SCHEDULE'].copy()
    out=BytesIO()
    with pd.ExcelWriter(out,engine='openpyxl') as w:
        selected.to_excel(w,index=False,sheet_name='Selected Sessions')
        req.to_excel(w,index=False,sheet_name='Requirement Audit')
        res['existing_workday'].to_excel(w,index=False,sheet_name='Existing Workday Training')
        res['duplicate_events'].to_excel(w,index=False,sheet_name='Duplicate Source Events')
        if not res['seat_reservations'].empty:res['seat_reservations'].to_excel(w,index=False,sheet_name='Seat Audit')
        pd.DataFrame([res['summary']]).to_excel(w,index=False,sheet_name='Run Summary')
    out.seek(0);return out.getvalue()


def output_package(res):
    req=res['requirements'];ef,en=workday_export(req,False);cf,cn=workday_export(req,True);af=audit_export(res)
    b=BytesIO()
    with ZipFile(b,'w',ZIP_DEFLATED) as z:
        z.writestr('Workday/Enroll_In_Learning_Content_Employees.xlsx',ef)
        if cn:z.writestr('Workday/Enroll_In_Learning_Content_Contingent_Workers.xlsx',cf)
        z.writestr('Scheduling_Results_and_Audit.xlsx',af)
        z.writestr('README.txt',f'v2.9.2 export. Employee rows: {en}. Contingent Worker rows: {cn}.\n')
    return b.getvalue(),en,cn


def email_package(req,routing):
    route={norm(r.get('training_title')):clean(r.get('recipient_email')) for _,r in routing.iterrows() if clean(r.get('training_title'))}
    b=BytesIO();manager=0;manual=0
    with ZipFile(b,'w',ZIP_DEFLATED) as z:
        for ek in req.loc[req['Disposition']=='PROPOSED_SCHEDULE','Event_Key'].astype(str).unique():
            rows=req[req['Event_Key'].astype(str)==ek];r0=rows.iloc[0]
            lines=[f"Employee: {r0['Worker_Name']}",f"Employee ID: {r0['Employee_ID']}",f"Position: {r0['Position_Title']}",f"Hire / Effective Date: {r0['Anchor_Date']}",'','Scheduled Training','------------------']
            for _,r in rows[rows['Disposition']=='PROPOSED_SCHEDULE'].sort_values('Selected_Start').iterrows():
                lines.append(f"• {r['Training_Title']} — {r['Selected_Start']} to {r['Selected_End']} — {r['Selected_Location']}")
            done=rows[rows['Disposition'].isin(['PREVIOUSLY_COMPLETED','EQUIVALENT_COMPLETION'])]
            if not done.empty:
                lines+=['','Previously Completed — No New Assignment','----------------------------------------']
                for _,r in done.iterrows():lines.append(f"• {r['Training_Title']} — {r['Completion_Date']}")
            z.writestr(f"Manager Emails/{ek}.eml",eml(clean(r0['Hiring_Manager_AD']),f"Training Schedule - {r0['Worker_Name']}",'\n'.join(lines)));manager+=1
        for idx,r in req[req['Disposition']=='MANUAL_SCHEDULING_REQUIRED'].iterrows():
            recipient=route.get(norm(r['Training_Title']),'')
            body=f"Please schedule the following employee for {r['Training_Title']}.\n\nEmployee: {r['Worker_Name']}\nEmployee email: {r['Work_Email'] or r['Home_Email']}\nEmployee ID: {r['Employee_ID']}\nHire / Position Effective Date: {r['Anchor_Date']}\nPosition: {r['Position_Title']}\nCost Center: {r['Cost_Center_Title']} ({r['Cost_Center_ID']})\nHiring Manager: {r['Hiring_Manager_AD']}\n"
            z.writestr(f"Manual Scheduling/{idx}.eml",eml(recipient,f"Manual scheduling request - {r['Training_Title']} - {r['Worker_Name']}",body));manual+=1
    return b.getvalue(),manager,manual

conf_up=st.file_uploader('Optional updated Scheduler_Configuration.xlsx',type=['xlsx'])
rules,locations,eq,routing=config_frames(rb(conf_up))
st.success(f'Configuration loaded: {len(rules):,} active rules • {len(eq):,} equivalencies')

st.subheader('Upload Reports')
a,b=st.columns(2)
with a:
    nh=rb(st.file_uploader('New Hire Orientation Report',type=['xlsx']))
    jc=rb(st.file_uploader('New / Additional Job Change Report',type=['xlsx']))
with b:
    hist=rb(st.file_uploader('Training History / Learning Transcript',type=['xlsx']))
    ses=rb(st.file_uploader('Learning Content / Available Sessions',type=['xlsx']))
    ori=rb(st.file_uploader('Orientation Schedule — Lesson-Level Detail',type=['xlsx']))

checks=[]
if nh:checks.append(('New Hire',validate(nh,'New Hire',['employee id','wid','hire date','candidate cost center id','candidate position id','worker'],['worker type','new hire position title','candidate department name','supervisory organization id','worker physical location','hiring manager ad','email - primary work','email - primary home'])))
if jc:checks.append(('Job Change',validate(jc,'Job Change',['employee id','wid','effective date','job code - proposed','cost center id - requested','worker'],['job title - proposed','cost center - proposed','sup org id - requested','location - proposed','worker type','hiring manager ad','candidate email'])))
checks.append(('Training History',validate(hist,'Training History',['employee id','record learning content','record completion status'],['wid','record completion date','registration status',"learner's registration date"])))
checks.append(('Available Sessions',validate(ses,'Available Sessions',['title','reference id','start date','end date','available seats','wid'],['availability status','locations'])))
checks.append(('Existing Orientation Schedule',validate(ori,'Existing Orientation Schedule',['employee id','enrolled course offering','registration status','start date','end date'],['wid','locations'])))
for label,(ok,msg) in checks:(st.success if ok else st.error)(f'{label}: {msg}')
ready=bool((nh or jc) and hist and ses and ori and all(ok for _,(ok,_) in checks))
if st.button('Run Scheduling',type='primary',disabled=not ready,use_container_width=True):
    try:
        with st.spinner('Scheduling...'):st.session_state.result=run_scheduler(nh,jc,hist,ses,rules,locations,eq,ori)
    except Exception as e:st.exception(e)

if st.session_state.get('result') is not None:
    res=st.session_state.result;req=res['requirements'];counts=res['summary'];st.subheader('Results')
    m=[('Requirements',len(req)),('Selected',counts.get('PROPOSED_SCHEDULE',0)),('TARGET_RANGE',int((req.Selection_Phase=='TARGET_RANGE').sum())),('FIRST_AVAILABLE',int((req.Selection_Phase=='FIRST_AVAILABLE').sum())),('Review',counts.get('REVIEW_REQUIRED',0)),('No documentation',counts.get('NO_TRAINING_DOCUMENTATION_FOUND',0)),('Already enrolled',counts.get('ALREADY_ENROLLED',0)),('Completed',counts.get('PREVIOUSLY_COMPLETED',0)+counts.get('EQUIVALENT_COMPLETION',0))]
    cc=st.columns(8)
    for c,(lab,val) in zip(cc,m):c.metric(lab,val)
    show=['Worker_Name','Employee_ID','Training_Title','Training_Documentation_Status','Priority','Scheduling_Policy','Selection_Order','Selection_Phase','Selected_Start','Selected_End','Selected_Location','Selected_Day_Count','Selected_Multi_Day','Selected_Session_Days','Disposition','Conflict_Detail','Explanation']
    st.dataframe(req[[c for c in show if c in req.columns]],use_container_width=True,height=460)
    st.subheader('Existing Workday Training')
    if res['existing_workday'].empty:st.info('No active/upcoming existing Workday training sessions found.')
    else:st.dataframe(res['existing_workday'],use_container_width=True,height=300)
    pkg,en,cn=output_package(res);st.download_button(f'Download Export Package ({en} employee / {cn} contingent)',pkg,'Class_Scheduling_v2_9_2_Export.zip',mime='application/zip',type='primary')
    ep,mc,mm=email_package(req,routing);st.download_button(f'Download Email Drafts ({mc} manager / {mm} manual)',ep,'Email_Drafts_v2_9_2.zip',mime='application/zip')
