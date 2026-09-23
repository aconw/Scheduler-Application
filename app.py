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

ROOT=Path(__file__).parent
CONFIG_DEFAULT=ROOT/'config'/'Scheduler_Configuration.xlsx'
ASSETS=ROOT/'assets'

st.set_page_config(page_title='Class Scheduling Batch Tool', page_icon='📚', layout='wide')
st.title('Class Scheduling Batch Tool')
st.caption('Simplified stateless build v2.0 • Upload Workday reports → validate → schedule → download Workday files and audit package')


from batch_utils import raw_bytes, config_frames, validate_report, results_zip, make_approved_emails

# Session only holds the current in-memory run; nothing is persisted on the server.
if 'result' not in st.session_state: st.session_state.result=None
if 'config_bytes' not in st.session_state: st.session_state.config_bytes=None

with st.expander('Configuration',expanded=False):
    st.write('The bundled configuration contains the current training rules and location master. Upload an edited configuration workbook here only when you want to test/use changed rules.')
    conf=st.file_uploader('Optional Scheduler_Configuration.xlsx',type=['xlsx'],key='config_upload')
    if conf is not None: st.session_state.config_bytes=raw_bytes(conf)
    st.download_button('Download Current Configuration Workbook',CONFIG_DEFAULT.read_bytes(),'Scheduler_Configuration.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    try:
        cr,cl,ce,cm=config_frames(st.session_state.config_bytes); st.success(f'Configuration ready: {len(cr):,} active rules • {len(cl):,} locations • {len(ce):,} equivalencies • {len(cm):,} manual routes')
    except Exception as e: st.error(str(e))

st.subheader('1. Upload Workday Reports')
c1,c2=st.columns(2)
with c1:
    nh=st.file_uploader('New Hire Orientation Report (optional if Job Change uploaded)',type=['xlsx'],key='nh')
    jc=st.file_uploader('New / Additional Job Change Report (optional if New Hire uploaded)',type=['xlsx'],key='jc')
with c2:
    hist=st.file_uploader('Training History / Learning Transcript (required)',type=['xlsx'],key='hist')
    sess=st.file_uploader('Learning Content / Available Sessions (required)',type=['xlsx'],key='sess')
    orient=st.file_uploader('Existing Orientation Schedule (optional)',type=['xlsx'],key='orient')

raws={k:raw_bytes(v) for k,v in {'nh':nh,'jc':jc,'hist':hist,'sess':sess,'orient':orient}.items()}

st.subheader('2. Preflight Validation')
checks=[]
if raws['nh'] is not None: checks.append(validate_report('New Hire',raws['nh'],['Employee ID','WID','Hire Date','Candidate Cost Center ID']))
if raws['jc'] is not None: checks.append(validate_report('Job Change',raws['jc'],['Employee ID','WID','Effective Date','Job Code - Proposed']))
checks.append(validate_report('Training History',raws['hist'],['Employee ID','Record Learning Content','Record Completion Status'],['Learning Participant','WID','Registration Status',"Learner's Registration Date",'Record Completion Date']))
checks.append(validate_report('Available Sessions',raws['sess'],['Learning Content Type','Title','Reference ID','Start Date','Available Seats','WID']))
if raws['orient'] is not None: checks.append(validate_report('Existing Orientation Schedule',raws['orient'],['Employee ID','Enrolled Course Offering','Registration Status','Start Date']))
for ck in checks:
    (st.success if ck['ok'] else st.error)(f"{ck['label']}: {ck['detail']}")

required_ok=(raws['hist'] is not None and raws['sess'] is not None and (raws['nh'] is not None or raws['jc'] is not None) and all(c['ok'] for c in checks if c['label'] in ['Training History','Available Sessions','New Hire','Job Change']))

if st.button('Run Scheduling',type='primary',use_container_width=True,disabled=not required_ok):
    try:
        rules,locations,eq,routing=config_frames(st.session_state.config_bytes)
        with st.spinner('Scheduling...'):
            st.session_state.result=run_scheduler(raws['nh'],raws['jc'],raws['hist'],raws['sess'],rules,locations,eq,raws['orient'])
        st.success('Scheduling complete.')
    except Exception as e: st.exception(e)

if st.session_state.result is not None:
    result=st.session_state.result; req=result['requirements']; counts=req['Disposition'].value_counts().to_dict() if not req.empty and 'Disposition' in req else {}
    st.subheader('3. Results')
    cols=st.columns(6); metrics=[('Requirements',len(req)),('Proposed',counts.get('PROPOSED_SCHEDULE',0)),('Review',counts.get('REVIEW_REQUIRED',0)),('Manual',counts.get('MANUAL_SCHEDULING_REQUIRED',0)),('Completed',counts.get('PREVIOUSLY_COMPLETED',0)+counts.get('EQUIVALENT_COMPLETION',0)),('Already enrolled',counts.get('ALREADY_ENROLLED',0))]
    for col,(lab,val) in zip(cols,metrics): col.metric(lab,val)
    display=['Worker_Name','Employee_ID','Event_Type','Training_Title','Disposition','Selected_Start','Selected_Location','Distance_Miles','Explanation']
    st.dataframe(req[[c for c in display if c in req.columns]],use_container_width=True,height=420)
    pkg,emp_n,cw_n=results_zip(result)
    st.download_button(f'Download Results Package ({emp_n} employee / {cw_n} contingent Workday rows)',pkg,'Class_Scheduling_Results_Package.zip',mime='application/zip',type='primary')

    st.subheader('4. Generate Email Drafts After Review')
    st.write('Open `Scheduling_Audit_and_Approval.xlsx` from the results package. On the **Employee Approval** sheet, enter `APPROVED` or `DENIED` in the Approval column. If denied, enter a Denial Reason. Save the workbook, then upload it below.')
    reviewed=st.file_uploader('Reviewed Scheduling_Audit_and_Approval.xlsx',type=['xlsx'],key='reviewed')
    if reviewed is not None:
        try:
            rules,locations,eq,routing=config_frames(st.session_state.config_bytes)
            emails,nmgr,nmanual=make_approved_emails(raw_bytes(reviewed),req,routing)
            st.download_button(f'Download Email Drafts ({nmgr} manager / {nmanual} manual)',emails,'Email_Drafts.zip',mime='application/zip')
        except Exception as e: st.error(str(e))
