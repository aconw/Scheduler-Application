from __future__ import annotations
from pathlib import Path
import streamlit as st

from scheduler_engine import run_scheduler
from batch_utils import raw_bytes, config_frames, validate_report, results_zip, make_email_drafts

ROOT=Path(__file__).parent
CONFIG_DEFAULT=ROOT/'config'/'Scheduler_Configuration.xlsx'

st.set_page_config(page_title='Class Scheduling Batch Tool', page_icon='📚', layout='wide')
st.title('Class Scheduling Batch Tool')
st.caption('Simplified stateless build v2.2 • No approval step • Non-overlapping scheduling • Workday files + requirement audit')

if 'result' not in st.session_state: st.session_state.result=None
if 'config_bytes' not in st.session_state: st.session_state.config_bytes=None

with st.expander('Configuration',expanded=False):
    st.write('The bundled configuration contains the current training rules and location master. Upload an edited configuration workbook only when you want to use changed rules, equivalencies, locations, or manual-message routing.')
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
    orient=st.file_uploader('Existing Orientation Schedule (required — used to prevent conflicts with existing classes)',type=['xlsx'],key='orient')

raws={k:raw_bytes(v) for k,v in {'nh':nh,'jc':jc,'hist':hist,'sess':sess,'orient':orient}.items()}

st.subheader('2. Preflight Validation')
checks=[]
if raws['nh'] is not None: checks.append(validate_report('New Hire',raws['nh'],['Employee ID','WID','Hire Date','Candidate Cost Center ID']))
if raws['jc'] is not None: checks.append(validate_report('Job Change',raws['jc'],['Employee ID','WID','Effective Date','Job Code - Proposed']))
checks.append(validate_report('Training History',raws['hist'],['Employee ID','Record Learning Content','Record Completion Status'],['Learning Participant','WID','Registration Status',"Learner's Registration Date",'Record Completion Date']))
checks.append(validate_report('Available Sessions',raws['sess'],['Learning Content Type','Title','Reference ID','Start Date','End Date','Available Seats','WID']))
checks.append(validate_report('Existing Orientation Schedule',raws['orient'],['Employee ID','Enrolled Course Offering','Registration Status','Start Date','End Date']))
for ck in checks:
    (st.success if ck['ok'] else st.error)(f"{ck['label']}: {ck['detail']}")

required_ok=(raws['hist'] is not None and raws['sess'] is not None and raws['orient'] is not None and (raws['nh'] is not None or raws['jc'] is not None) and all(c['ok'] for c in checks if c['label'] in ['Training History','Available Sessions','Existing Orientation Schedule','New Hire','Job Change']))

if st.button('Run Scheduling',type='primary',use_container_width=True,disabled=not required_ok):
    try:
        rules,locations,eq,routing=config_frames(st.session_state.config_bytes)
        with st.spinner('Scheduling and checking employee time conflicts...'):
            st.session_state.result=run_scheduler(raws['nh'],raws['jc'],raws['hist'],raws['sess'],rules,locations,eq,raws['orient'])
        st.success('Scheduling complete. No selected employee sessions overlap one another. If an Existing Orientation Schedule was uploaded, those classes were also treated as blocked time.')
    except Exception as e: st.exception(e)

if st.session_state.result is not None:
    result=st.session_state.result; req=result['requirements']; counts=req['Disposition'].value_counts().to_dict() if not req.empty and 'Disposition' in req else {}
    st.subheader('3. Results')
    cols=st.columns(6); metrics=[('Requirements',len(req)),('Selected sessions',counts.get('PROPOSED_SCHEDULE',0)),('Review',counts.get('REVIEW_REQUIRED',0)),('Manual',counts.get('MANUAL_SCHEDULING_REQUIRED',0)),('Completed',counts.get('PREVIOUSLY_COMPLETED',0)+counts.get('EQUIVALENT_COMPLETION',0)),('Already enrolled',counts.get('ALREADY_ENROLLED',0))]
    for col,(lab,val) in zip(cols,metrics): col.metric(lab,val)
    display=['Worker_Name','Employee_ID','Event_Type','Training_Title','Disposition','Selected_Start','Selected_End','Selected_Location','Distance_Miles','Conflict_Detail','Explanation']
    st.dataframe(req[[c for c in display if c in req.columns]],use_container_width=True,height=440)

    pkg,emp_n,cw_n=results_zip(result)
    st.subheader('4. Export')
    st.write('The package contains exact Workday-format enrollment files plus `Scheduling_Results_and_Audit.xlsx`, which includes **Selected Sessions**, **Requirement Audit**, **Review Queue**, **Seat Audit**, and **Run Summary**. There is no approval/denial step.')
    st.download_button(f'Download Scheduling Export Package ({emp_n} employee / {cw_n} contingent Workday rows)',pkg,'Class_Scheduling_Export_Package.zip',mime='application/zip',type='primary')

    st.subheader('5. Email Drafts (Optional)')
    st.write('Email drafts are generated directly from the current scheduling results. They are drafts only; nothing is sent automatically.')
    try:
        _,_,_,routing=config_frames(st.session_state.config_bytes)
        emails,nmgr,nmanual=make_email_drafts(req,routing)
        st.download_button(f'Download Email Drafts ({nmgr} manager / {nmanual} manual)',emails,'Email_Drafts.zip',mime='application/zip')
    except Exception as e:
        st.error(str(e))
