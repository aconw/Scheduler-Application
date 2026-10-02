from __future__ import annotations
from pathlib import Path
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from email.message import EmailMessage
from datetime import datetime
import pandas as pd
import openpyxl
import streamlit as st

from scheduler_engine import run_scheduler, read_sheet, norm, clean

ROOT=Path(__file__).parent; CONFIG_DEFAULT=ROOT/'config'/'Scheduler_Configuration.xlsx'; ASSETS=ROOT/'assets'

st.set_page_config(page_title='Class Scheduling Batch Tool',page_icon='📚',layout='wide')
st.title('Class Scheduling Batch Tool')
st.caption('Simplified stateless build v2.8.2 • Training documentation status • TARGET_RANGE first • FIRST_AVAILABLE priority ranking • Multi-day session aware')


def raw_bytes(upload):
    if upload is None:return None
    if isinstance(upload,(bytes,bytearray)):return bytes(upload)
    if hasattr(upload,'getvalue'):return bytes(upload.getvalue())
    upload.seek(0); return upload.read()

def config_frames(raw=None):
    src=BytesIO(raw) if raw else CONFIG_DEFAULT
    sheets=pd.read_excel(src,sheet_name=None,engine='openpyxl')
    required=['Training Rules','Locations','Equivalencies','Manual Routing']; missing=[x for x in required if x not in sheets]
    if missing: raise ValueError(f'Configuration workbook is missing sheets: {missing}')
    def active(df):
        if 'active' not in df.columns:return df.copy()
        v=df['active'].fillna(1).astype(str).str.strip().str.lower()
        return df[v.isin(['1','1.0','true','yes','y'])].copy()
    rules=active(sheets['Training Rules']); loc=sheets['Locations'].copy(); eq=active(sheets['Equivalencies']); route=active(sheets['Manual Routing'])
    rules=rules.where(pd.notna(rules),''); loc=loc.where(pd.notna(loc),''); eq=eq.where(pd.notna(eq),''); route=route.where(pd.notna(route),'')
    if 'priority' not in rules.columns: rules['priority']=''
    if 'relationship_direction' not in eq.columns: eq['relationship_direction']='ONE_WAY'
    return rules,loc,eq,route

def validate(label,raw,required,optional=None):
    if raw is None:return (False,'Not uploaded',0)
    try:
        rows=read_sheet(raw,required,optional or [])
        wb=openpyxl.load_workbook(BytesIO(raw),read_only=False,data_only=True)
        sheets=[f'{ws.title} ({ws.max_row:,}x{ws.max_column:,})' for ws in wb.worksheets]
        return True,f'{len(raw):,} bytes • {"; ".join(sheets)} • {len(rows):,} data rows',len(rows)
    except Exception as e:return False,str(e),0

def workday_export(req,contingent=False):
    template=ASSETS/('Enroll_In_Learning_Content_vContingent_Worker_ID.xlsx' if contingent else 'Enroll_In_Learning_Content_vEmployee.xlsx')
    wb=openpyxl.load_workbook(template); ws=wb['Enroll In Learning Content']
    rows=req[(req['Workday_Ready']=='Yes') & req['Selected_Session_WID'].astype(str).ne('')].copy()
    rows=rows[rows['Worker_Type'].map(norm).eq('contingent worker')] if contingent else rows[~rows['Worker_Type'].map(norm).eq('contingent worker')]
    # Defense in depth: person + selected session cannot be exported twice.
    if not rows.empty:
        rows=rows.sort_values(['Person_Key','Selected_Start','Event_Key'],na_position='last').drop_duplicates(['Person_Key','Selected_Session_WID'],keep='first')
    for i,(_,r) in enumerate(rows.iterrows(),6):
        ws.cell(i,1,f"{r['Event_Key']}-{i-5}"); ws.cell(i,2,str(r['Selected_Session_WID'])); ws.cell(i,3,str(clean(r.get('Employee_ID')))); ws.cell(i,4,'Y'); ws.cell(i,5,None)
    bio=BytesIO(); wb.save(bio); return bio.getvalue(),len(rows)

def audit_workbook(result):
    req=result['requirements'].copy(); selected=req[req['Disposition']=='PROPOSED_SCHEDULE'].copy(); review=req[req['Disposition']=='REVIEW_REQUIRED'].copy(); seat=result['seat_reservations'].copy(); dup=result.get('duplicate_events',pd.DataFrame()).copy(); existing=result.get('existing_workday',pd.DataFrame()).copy(); doc=result.get('training_documentation_status',pd.DataFrame()).copy()
    cols=['Event_Key','Worker_Name','Employee_ID','WID','Event_Type','Anchor_Date','Job_Code','Position_Title','Cost_Center_ID','Cost_Center_Title','Training_Title','Priority','Scheduling_Policy','Selected_Session_WID','Selected_Reference_ID','Selected_Start','Selected_End','Selected_Location','Selected_Day_Count','Distance_Miles','Prerequisite','Prerequisite_Status','Disposition','Explanation']
    selected=selected[[c for c in cols if c in selected.columns]]
    summary=pd.DataFrame([{'raw_staffing_events':result['source_counts']['staffing_events_raw'],'canonical_staffing_events':result['source_counts']['staffing_events_canonical'],'duplicate_source_events':result['source_counts']['duplicate_source_events'],'training_documentation_found':result['source_counts'].get('training_documentation_found',0),'no_training_documentation_found':result['source_counts'].get('no_training_documentation_found',0),'requirements':len(req),'selected_sessions':int((req['Disposition']=='PROPOSED_SCHEDULE').sum()),'review_required':int((req['Disposition']=='REVIEW_REQUIRED').sum()),'manual_scheduling_required':int((req['Disposition']=='MANUAL_SCHEDULING_REQUIRED').sum()),'same_run_satisfied':int((req['Disposition']=='SATISFIED_BY_SAME_RUN_ASSIGNMENT').sum()),'already_enrolled':int((req['Disposition']=='ALREADY_ENROLLED').sum()),'completed_or_equivalent':int(req['Disposition'].isin(['PREVIOUSLY_COMPLETED','EQUIVALENT_COMPLETION']).sum())}])
    bio=BytesIO()
    with pd.ExcelWriter(bio,engine='openpyxl') as xw:
        selected.to_excel(xw,index=False,sheet_name='Selected Sessions'); req.to_excel(xw,index=False,sheet_name='Requirement Audit'); review.to_excel(xw,index=False,sheet_name='Review Queue'); seat.to_excel(xw,index=False,sheet_name='Seat Audit'); dup.to_excel(xw,index=False,sheet_name='Duplicate Source Events'); existing.to_excel(xw,index=False,sheet_name='Existing Workday Training'); doc.to_excel(xw,index=False,sheet_name='Training Documentation Status'); summary.to_excel(xw,index=False,sheet_name='Run Summary')
    bio.seek(0); wb=openpyxl.load_workbook(bio)
    from openpyxl.styles import Font,PatternFill,Alignment
    fill=PatternFill('solid',fgColor='1F4E78'); font=Font(color='FFFFFF',bold=True)
    for ws in wb.worksheets:
        ws.freeze_panes='A2'; ws.sheet_view.showGridLines=False; ws.auto_filter.ref=ws.dimensions
        for c in ws[1]: c.fill=fill; c.font=font; c.alignment=Alignment(wrap_text=True)
        for col in ws.columns:
            letter=col[0].column_letter; vals=list(col)[:300]; m=max([len(str(c.value)) if c.value is not None else 0 for c in vals]+[8]); ws.column_dimensions[letter].width=min(m+2,48)
    out=BytesIO(); wb.save(out); return out.getvalue()

def results_zip(result):
    req=result['requirements']
    emp,emp_n=workday_export(req,False); cw,cw_n=workday_export(req,True)
    audit=audit_workbook(result)
    z=BytesIO()
    with ZipFile(z,'w',ZIP_DEFLATED) as zipf:
        zipf.writestr('Scheduling_Results_and_Audit.xlsx',audit)
        zipf.writestr('Workday/Enroll_In_Learning_Content_Employees.xlsx',emp)
        if cw_n: zipf.writestr('Workday/Enroll_In_Learning_Content_Contingent_Workers.xlsx',cw)
        zipf.writestr('README.txt',f"v2.8.2 scheduling export generated {datetime.now():%Y-%m-%d %H:%M}. Employee rows: {emp_n}; contingent worker rows: {cw_n}.\n")
    return z.getvalue(),emp_n,cw_n

def email_drafts(req,routing):
    route={norm(r.get('training_title')):clean(r.get('recipient_email')) for _,r in routing.iterrows() if clean(r.get('training_title'))}
    z=BytesIO(); manager=0; manual=0
    with ZipFile(z,'w',ZIP_DEFLATED) as zipf:
        for ek in sorted(req.loc[req['Disposition']=='PROPOSED_SCHEDULE','Event_Key'].astype(str).unique()):
            rows=req[req['Event_Key'].astype(str)==ek]; r0=rows.iloc[0]; lines=[f"Employee: {r0['Worker_Name']}",f"Employee ID: {r0['Employee_ID']}",f"Position: {r0['Position_Title']}",'','Scheduled Training','------------------']
            for _,r in rows[rows['Disposition']=='PROPOSED_SCHEDULE'].sort_values('Selected_Start').iterrows():
                stx=pd.to_datetime(r['Selected_Start']); enx=pd.to_datetime(r['Selected_End']); daytext=f" ({int(r['Selected_Day_Count'])} days)" if pd.notna(r.get('Selected_Day_Count')) and int(r['Selected_Day_Count'])>1 else ''
                lines.append(f"• {r['Training_Title']} — {stx:%m/%d/%Y %I:%M %p}-{enx:%I:%M %p} — {r['Selected_Location']}{daytext}")
            body='\n'.join(lines); m=EmailMessage(); m['To']=clean(r0['Hiring_Manager_AD']); m['Subject']=f"Training Schedule - {r0['Worker_Name']}"; m['X-Unsent']='1'; m.set_content(body); zipf.writestr(f'Manager Emails/{ek}.eml',m.as_bytes()); manager+=1
        for idx,r in req[req['Disposition']=='MANUAL_SCHEDULING_REQUIRED'].iterrows():
            m=EmailMessage(); m['To']=route.get(norm(r['Training_Title']),''); m['Subject']=f"Manual scheduling request - {r['Training_Title']} - {r['Worker_Name']}"; m['X-Unsent']='1'; m.set_content(f"Please schedule the following employee for {r['Training_Title']}.\n\nEmployee: {r['Worker_Name']}\nEmployee email: {r.get('Work_Email') or r.get('Home_Email')}\nEmployee ID: {r['Employee_ID']}\nHire / Position Effective Date: {r['Anchor_Date']}\nPosition: {r['Position_Title']}\nCost Center: {r['Cost_Center_Title']} ({r['Cost_Center_ID']})\nHiring Manager: {r['Hiring_Manager_AD']}\n"); zipf.writestr(f'Manual Scheduling/{idx}.eml',m.as_bytes()); manual+=1
    return z.getvalue(),manager,manual

if 'result' not in st.session_state: st.session_state.result=None
if 'config_bytes' not in st.session_state: st.session_state.config_bytes=None

with st.expander('Configuration',expanded=False):
    conf=st.file_uploader('Optional Scheduler_Configuration.xlsx',type=['xlsx'],key='conf')
    if conf is not None: st.session_state.config_bytes=raw_bytes(conf)
    st.download_button('Download Current Configuration',CONFIG_DEFAULT.read_bytes(),'Scheduler_Configuration.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    try:
        cr,cl,ce,cm=config_frames(st.session_state.config_bytes); st.success(f'Configuration ready: {len(cr):,} active rules • {len(cl):,} locations • {len(ce):,} equivalencies • {len(cm):,} manual routes')
    except Exception as e: st.error(str(e))

st.subheader('1. Upload Workday Reports')
a,b=st.columns(2)
with a:
    nh=st.file_uploader('New Hire Orientation Report',type=['xlsx'],key='nh'); jc=st.file_uploader('New / Additional Job Change Report',type=['xlsx'],key='jc')
with b:
    hist=st.file_uploader('Training History / Learning Transcript',type=['xlsx'],key='hist'); sess=st.file_uploader('Learning Content / Available Sessions',type=['xlsx'],key='sess'); orient=st.file_uploader('Existing Orientation Schedule — required for conflict checking',type=['xlsx'],key='orient')
raw={k:raw_bytes(v) for k,v in {'nh':nh,'jc':jc,'hist':hist,'sess':sess,'orient':orient}.items()}

st.subheader('2. Preflight Validation')
checks=[]
if raw['nh'] is not None: checks.append(('New Hire',*validate('New Hire',raw['nh'],['Employee ID','WID','Hire Date','Candidate Cost Center ID'],['Candidate Position ID'])))
if raw['jc'] is not None: checks.append(('Job Change',*validate('Job Change',raw['jc'],['Employee ID','WID','Effective Date','Job Code - Proposed'])))
checks.append(('Training History',*validate('Training History',raw['hist'],['Employee ID','Record Learning Content','Record Completion Status'],['Learning Participant','WID','Registration Status',"Learner's Registration Date",'Record Completion Date'])))
checks.append(('Available Sessions',*validate('Available Sessions',raw['sess'],['Learning Content Type','Title','Reference ID','Start Date','End Date','Available Seats','WID'],['Availability Status','Locations'])))
checks.append(('Existing Orientation Schedule',*validate('Existing Orientation Schedule',raw['orient'],['Employee ID','Enrolled Course Offering','Registration Status','Start Date','End Date'],['WID','Locations'])))
for lab,ok,detail,rows in checks:
    (st.success if ok else st.error)(f'{lab}: {detail}')
required_labels={'Training History','Available Sessions','Existing Orientation Schedule','New Hire','Job Change'}
ready=raw['hist'] is not None and raw['sess'] is not None and raw['orient'] is not None and (raw['nh'] is not None or raw['jc'] is not None) and all(ok for lab,ok,_,_ in checks if lab in required_labels)

if st.button('Run Scheduling',type='primary',disabled=not ready,use_container_width=True):
    try:
        rules,loc,eq,route=config_frames(st.session_state.config_bytes)
        with st.spinner('Scheduling with policy precedence, priority ranking, seat reservation, and multi-day conflict checking...'):
            st.session_state.result=run_scheduler(raw['nh'],raw['jc'],raw['hist'],raw['sess'],rules,loc,eq,raw['orient'])
        st.success('Scheduling complete.')
    except Exception as e: st.exception(e)

if st.session_state.result is not None:
    result=st.session_state.result; req=result['requirements']; counts=req['Disposition'].value_counts().to_dict(); dup=len(result.get('duplicate_events',pd.DataFrame())); doc=result.get('training_documentation_status',pd.DataFrame())
    st.subheader('3. Results')
    no_doc=int((doc['Training_Documentation_Status']=='NO_TRAINING_DOCUMENTATION_FOUND').sum()) if not doc.empty else 0
    cols=st.columns(8); metrics=[('Requirements',len(req)),('Selected',counts.get('PROPOSED_SCHEDULE',0)),('Same-run satisfied',counts.get('SATISFIED_BY_SAME_RUN_ASSIGNMENT',0)),('Duplicate source events',dup),('No training documentation',no_doc),('Review',counts.get('REVIEW_REQUIRED',0)),('Manual',counts.get('MANUAL_SCHEDULING_REQUIRED',0)),('Already enrolled',counts.get('ALREADY_ENROLLED',0))]
    for c,(lab,v) in zip(cols,metrics): c.metric(lab,v)
    st.dataframe(req[['Worker_Name','Employee_ID','Event_Type','Training_Title','Priority','Scheduling_Policy','Selected_Session_WID','Selected_Start','Selected_End','Selected_Day_Count','Selected_Location','Distance_Miles','Disposition','Explanation']][[c for c in ['Worker_Name','Employee_ID','Event_Type','Training_Title','Priority','Scheduling_Policy','Selected_Session_WID','Selected_Start','Selected_End','Selected_Day_Count','Selected_Location','Distance_Miles','Disposition','Explanation'] if c in req.columns]],use_container_width=True,height=430)
    st.markdown('### Training Documentation Status')
    st.caption("Every canonical staffing event is shown. NO_TRAINING_DOCUMENTATION_FOUND means no active Training Rules matched that event's Job Code + Cost Center; the application does not interpret that as no training required.")
    if not doc.empty:
        st.dataframe(doc,use_container_width=True,height=280)
    else:
        st.info('No staffing events were available for training-documentation validation.')
    existing=result.get('existing_workday',pd.DataFrame())
    st.markdown('### Existing Workday Training')
    if not existing.empty: st.dataframe(existing,use_container_width=True,height=280)
    else: st.info('No active/upcoming existing Workday sessions found.')
    pkg,emp,cw=results_zip(result); st.subheader('4. Export'); st.write(f'Workday rows: {emp:,} employees / {cw:,} contingent workers. The audit workbook includes Selected Sessions, the complete Requirement Audit, Review Queue, Seat Audit, Duplicate Source Events, Existing Workday Training, Training Documentation Status, and Run Summary.')
    st.download_button('Download Scheduling Export Package',pkg,'Class_Scheduling_Export_Package_v2_8_2.zip',mime='application/zip',type='primary')
    st.subheader('5. Email Drafts (Optional)')
    _,_,_,route=config_frames(st.session_state.config_bytes); em,mc,mm=email_drafts(req,route); st.download_button(f'Download Email Drafts ({mc} manager / {mm} manual)',em,'Email_Drafts.zip',mime='application/zip')
