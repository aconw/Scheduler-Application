from __future__ import annotations
from pathlib import Path
from io import BytesIO
from email.message import EmailMessage
from datetime import datetime
import sqlite3
import pandas as pd
import streamlit as st
import openpyxl

from database import connect, load_rules, load_equivalencies, load_locations, load_manual_routing, audit, DB_PATH
from scheduler_engine import run_scheduler, dt, norm, clean, MAX_DISTANCE_MILES

ROOT=Path(__file__).parent
ASSETS=ROOT/'assets'
conn=connect()

st.set_page_config(page_title='Class Scheduling Application', page_icon='📚', layout='wide')
st.title('Class Scheduling Application')
st.caption('Rules-driven training scheduling, Workday export, verification, approval, and messaging')

if 'user_name' not in st.session_state: st.session_state.user_name='Scheduler'
if 'run_result' not in st.session_state: st.session_state.run_result=None
if 'run_files' not in st.session_state: st.session_state.run_files={}

with st.sidebar:
    st.session_state.user_name=st.text_input('Your name / initials',st.session_state.user_name)
    page=st.radio('Navigation',[
        'Dashboard','New Scheduling Run','Scheduling Results','Exception Review','Workday Export',
        'Verification & Approval','Message Center','Training Configurator','Equivalencies','Locations','Manual Routing','Audit & Backup'
    ])
    st.caption('Identity note: WID is the durable person key; Employee ID is treated as changeable.')

def df_download(df, filename, label):
    bio=BytesIO()
    with pd.ExcelWriter(bio,engine='openpyxl') as xw:
        df.to_excel(xw,index=False,sheet_name='Data')
    st.download_button(label,bio.getvalue(),filename,mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def make_results_workbook(result):
    bio=BytesIO()
    with pd.ExcelWriter(bio,engine='openpyxl') as xw:
        result['events'].to_excel(xw,index=False,sheet_name='Staffing Events')
        result['requirements'].to_excel(xw,index=False,sheet_name='Requirement Results')
        result['seat_reservations'].to_excel(xw,index=False,sheet_name='Seat Audit')
    return bio.getvalue()

def workday_export(requirements, contingent=False):
    template=ASSETS/('Enroll_In_Learning_Content_vContingent_Worker_ID.xlsx' if contingent else 'Enroll_In_Learning_Content_vEmployee.xlsx')
    wb=openpyxl.load_workbook(template)
    ws=wb['Enroll In Learning Content']
    rows=requirements[(requirements['Workday_Ready']=='Yes') & requirements['Selected_Session_WID'].astype(str).ne('')].copy()
    if contingent:
        rows=rows[rows['Worker_Type'].map(norm).eq('contingent worker')]
    else:
        rows=rows[~rows['Worker_Type'].map(norm).eq('contingent worker')]
    for i,(_,r) in enumerate(rows.iterrows(),start=6):
        ws.cell(i,1,f"{r['Event_Key']}-{i-5}")
        ws.cell(i,2,str(r['Selected_Session_WID']))
        ws.cell(i,3,str(r['Employee_ID']))
        ws.cell(i,4,'Y')
        ws.cell(i,5,None)
    bio=BytesIO(); wb.save(bio); return bio.getvalue(),len(rows)

def make_eml(to_email, subject, body):
    msg=EmailMessage(); msg['To']=to_email or ''; msg['Subject']=subject; msg['X-Unsent']='1'; msg.set_content(body)
    return msg.as_bytes()

def manager_email_for_event(reqs,event_key):
    rows=reqs[reqs['Event_Key']==event_key]
    if rows.empty: return None
    r0=rows.iloc[0]; mgr=clean(r0.get('Hiring_Manager_AD'))
    lines=[f"Employee: {r0.get('Worker_Name','')}",f"Employee ID: {r0.get('Employee_ID','')}",f"Position: {r0.get('Position_Title','')}",f"Hire / Position Effective Date: {pd.to_datetime(r0.get('Anchor_Date')).strftime('%m/%d/%Y') if pd.notna(r0.get('Anchor_Date')) else ''}",'']
    sched=rows[rows['Disposition']=='PROPOSED_SCHEDULE']
    if not sched.empty:
        lines+=['Scheduled Training','------------------']
        for _,r in sched.iterrows():
            d=pd.to_datetime(r.get('Selected_Start')) if pd.notna(r.get('Selected_Start')) else None
            lines.append(f"• {r['Training_Title']} — {d.strftime('%m/%d/%Y') if d is not None else ''} — {r.get('Selected_Location','')}")
        lines.append('')
    existing=rows[rows['Disposition']=='ALREADY_ENROLLED']
    if not existing.empty:
        lines+=['Already Enrolled','----------------']
        for _,r in existing.iterrows(): lines.append(f"• {r['Training_Title']}")
        lines.append('')
    done=rows[rows['Disposition'].isin(['PREVIOUSLY_COMPLETED','EQUIVALENT_COMPLETION'])]
    if not done.empty:
        lines+=['Previously Completed — No New Assignment','----------------------------------------']
        for _,r in done.iterrows():
            d=pd.to_datetime(r.get('Completion_Date')) if pd.notna(r.get('Completion_Date')) and r.get('Completion_Date')!='' else None
            extra=f" (satisfied by {r.get('Completion_Training')})" if r['Disposition']=='EQUIVALENT_COMPLETION' else ''
            lines.append(f"• {r['Training_Title']}{extra} — completed {d.strftime('%m/%d/%Y') if d is not None else 'previously'}")
    subject=f"Training Schedule - {r0.get('Worker_Name','Employee')}"
    return mgr,subject,'\n'.join(lines)

def export_config_workbook():
    bio=BytesIO()
    with pd.ExcelWriter(bio,engine='openpyxl') as xw:
        load_rules(conn,False).to_excel(xw,index=False,sheet_name='Training Rules')
        load_equivalencies(conn,False).to_excel(xw,index=False,sheet_name='Equivalencies')
        load_locations(conn,False).to_excel(xw,index=False,sheet_name='Locations')
        load_manual_routing(conn,False).to_excel(xw,index=False,sheet_name='Manual Routing')
    return bio.getvalue()

if page=='Dashboard':
    st.subheader('Dashboard')
    rr=st.session_state.run_result
    if rr is None:
        st.info('No scheduling run is loaded in this session. Start with New Scheduling Run.')
    else:
        req=rr['requirements']; counts=req['Disposition'].value_counts().to_dict()
        cols=st.columns(6)
        metrics=[('Staffing events',len(rr['events'])),('Proposed',counts.get('PROPOSED_SCHEDULE',0)),('Needs review',counts.get('REVIEW_REQUIRED',0)),('Manual',counts.get('MANUAL_SCHEDULING_REQUIRED',0)),('Completed',counts.get('PREVIOUSLY_COMPLETED',0)+counts.get('EQUIVALENT_COMPLETION',0)),('Already enrolled',counts.get('ALREADY_ENROLLED',0))]
        for c,(lab,val) in zip(cols,metrics): c.metric(lab,val)
        st.warning(rr.get('warning',''))
        st.dataframe(req[['Worker_Name','Employee_ID','Event_Type','Training_Title','Disposition','Explanation']].head(200),use_container_width=True)

elif page=='New Scheduling Run':
    st.subheader('New Scheduling Run')
    st.write('Upload the operational reports for this run. Training documentation and locations are already stored in the application configuration.')
    c1,c2=st.columns(2)
    with c1:
        nh=st.file_uploader('New Hire Orientation Report (.xlsx)',type=['xlsx'],key='nh')
        jc=st.file_uploader('New / Additional Job Change Orientation Report (.xlsx)',type=['xlsx'],key='jc')
    with c2:
        hist=st.file_uploader('Training History / Learning Transcript (.xlsx)',type=['xlsx'],key='hist')
        sess=st.file_uploader('Learning Content / Available Sessions (.xlsx)',type=['xlsx'],key='sess')
        orient=st.file_uploader('Existing Orientation Schedule Report (optional)',type=['xlsx'],key='orient')
    st.caption('At least one staffing report (New Hire or Job Change) is required. Training History and Available Sessions are required.')
    if st.button('Run Scheduling',type='primary',use_container_width=True):
        if not (nh or jc) or not hist or not sess:
            st.error('Upload at least one staffing report, plus Training History and Available Sessions.')
        else:
            try:
                with st.spinner('Running scheduling logic...'):
                    result=run_scheduler(nh,jc,hist,sess,load_rules(conn),load_locations(conn,True),load_equivalencies(conn),orient)
                    st.session_state.run_result=result
                    audit(conn,'SCHEDULING_RUN','run',datetime.now().isoformat(timespec='seconds'),str(result['source_counts']),st.session_state.user_name)
                st.success('Scheduling run completed. Open Scheduling Results or Exception Review.')
                st.write(result['source_counts'])
            except Exception as e:
                st.exception(e)

elif page=='Scheduling Results':
    st.subheader('Scheduling Results')
    rr=st.session_state.run_result
    if rr is None: st.info('Run scheduling first.')
    else:
        req=rr['requirements']
        counts=req['Disposition'].value_counts().rename_axis('Disposition').reset_index(name='Count')
        st.dataframe(counts,use_container_width=True,hide_index=True)
        filters=st.multiselect('Disposition filter',sorted(req['Disposition'].dropna().unique()))
        show=req[req['Disposition'].isin(filters)] if filters else req
        cols=['Worker_Name','Employee_ID','WID','Event_Type','Anchor_Date','Job_Code','Cost_Center_ID','Training_Title','Scheduling_Policy','Disposition','Selected_Start','Selected_Location','Distance_Miles','Explanation']
        st.dataframe(show[[c for c in cols if c in show.columns]],use_container_width=True,height=550)
        st.download_button('Download Full Scheduling Results',make_results_workbook(rr),'Class_Scheduling_Results.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

elif page=='Exception Review':
    st.subheader('Exception Review & Manual Override')
    rr=st.session_state.run_result
    if rr is None: st.info('Run scheduling first.')
    else:
        req=rr['requirements']; exc=req[req['Disposition']=='REVIEW_REQUIRED']
        st.metric('Review items',len(exc))
        st.dataframe(exc[['Worker_Name','Employee_ID','Event_Key','Training_Title','Scheduling_Policy','Prerequisite','Explanation']],use_container_width=True,height=350)
        if not exc.empty:
            labels={idx:f"{row['Worker_Name']} | {row['Training_Title']} | {row['Explanation'][:70]}" for idx,row in exc.iterrows()}
            idx=st.selectbox('Select item to override',list(labels.keys()),format_func=lambda x:labels[x])
            row=req.loc[idx]; sessions=rr['sessions']
            cand=sessions[sessions['Title'].map(norm)==norm(row['Training_Title'])].copy() if not sessions.empty else pd.DataFrame()
            if not cand.empty:
                cand['label']=cand.apply(lambda r:f"{r.get('Start Date')} | {r.get('Locations') or 'Virtual'} | seats {r.get('Available Seats')} | {r.get('WID')}",axis=1)
                choice=st.selectbox('Session',list(cand.index),format_func=lambda x:cand.loc[x,'label'])
                reason=st.text_area('Override reason (required)')
                if st.button('Apply Manual Override'):
                    if not reason.strip(): st.error('Override reason is required.')
                    else:
                        s=cand.loc[choice]
                        req.at[idx,'Disposition']='PROPOSED_SCHEDULE'; req.at[idx,'Workday_Ready']='Yes'; req.at[idx,'Override']='Yes'; req.at[idx,'Override_Reason']=reason.strip(); req.at[idx,'Explanation']='Manual override: '+reason.strip(); req.at[idx,'Selected_Session_WID']=str(s.get('WID')); req.at[idx,'Selected_Reference_ID']=str(s.get('Reference ID')); req.at[idx,'Selected_Start']=s.get('Start Date'); req.at[idx,'Selected_End']=s.get('End Date'); req.at[idx,'Selected_Location']=s.get('Locations') or 'Virtual'
                        st.session_state.run_result['requirements']=req
                        audit(conn,'MANUAL_OVERRIDE','requirement',str(idx),f"{row['Worker_Name']} | {row['Training_Title']} | {reason}",st.session_state.user_name)
                        st.success('Override applied and audited.')
            else: st.warning('No sessions with an exact matching title are present in the uploaded session report.')

elif page=='Workday Export':
    st.subheader('Workday Export')
    rr=st.session_state.run_result
    if rr is None: st.info('Run scheduling first.')
    else:
        req=rr['requirements']
        emp_bytes,emp_n=workday_export(req,False); cw_bytes,cw_n=workday_export(req,True)
        st.write(f'Employee enrollment rows: **{emp_n}**')
        st.download_button('Download Employee Workday Enrollment File',emp_bytes,'Enroll_In_Learning_Content_Employees.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        st.write(f'Contingent worker enrollment rows: **{cw_n}**')
        if cw_n:
            st.download_button('Download Contingent Worker Workday Enrollment File',cw_bytes,'Enroll_In_Learning_Content_Contingent_Workers.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        st.caption('Only PROPOSED_SCHEDULE records marked Workday-ready are exported. Completed, already-enrolled, manual, and review-required items are excluded.')

elif page=='Verification & Approval':
    st.subheader('Workday Verification & Employee-Level Approval')
    rr=st.session_state.run_result
    if rr is None: st.info('Run scheduling first.')
    else:
        verification=st.file_uploader('Upload Workday Orientation Schedule Report after enrollment',type=['xlsx'],key='verify')
        req=rr['requirements']
        if verification:
            from scheduler_engine import read_sheet
            rows=read_sheet(verification,required_headers=['Employee ID','Enrolled Course Offering','Registration Status','Start Date'])
            vdf=pd.DataFrame(rows)
            proposed=req[req['Disposition']=='PROPOSED_SCHEDULE'].copy()
            def match_status(r):
                m=vdf[(vdf['Employee ID'].astype(str).str.strip()==str(r['Employee_ID']).strip()) & (vdf['Enrolled Course Offering'].map(norm)==norm(r['Training_Title']))]
                if m.empty: return 'MISSING'
                expected=dt(r.get('Selected_Start'))
                dates=[dt(x) for x in m['Start Date'] if dt(x)]
                if expected and dates and any(d.date()==expected.date() for d in dates): return 'MATCH'
                return 'FOUND_DIFFERENT_DATE'
            proposed['Verification']=proposed.apply(match_status,axis=1)
            st.session_state.verification_df=proposed
            st.dataframe(proposed[['Worker_Name','Employee_ID','Event_Key','Training_Title','Selected_Start','Selected_Location','Verification']],use_container_width=True,height=350)
        events=req[['Event_Key','Worker_Name','Employee_ID','Event_Type']].drop_duplicates()
        if not events.empty:
            labels={r.Event_Key:f"{r.Worker_Name} | {r.Employee_ID} | {r.Event_Type}" for r in events.itertuples()}
            ek=st.selectbox('Employee / staffing event',list(labels),format_func=lambda x:labels[x])
            event_rows=req[req['Event_Key']==ek]
            st.dataframe(event_rows[['Training_Title','Disposition','Selected_Start','Selected_Location','Completion_Date','Explanation']],use_container_width=True)
            existing=conn.execute('SELECT * FROM approvals WHERE event_key=?',(ek,)).fetchone()
            if existing: st.info(f"Current decision: {existing['decision']} by {existing['decided_by']} at {existing['decided_at']}"+(f" — {existing['reason']}" if existing['reason'] else ''))
            c1,c2=st.columns(2)
            with c1:
                if st.button('Approve Schedule',type='primary',use_container_width=True):
                    now=datetime.now().isoformat(timespec='seconds')
                    conn.execute('INSERT OR REPLACE INTO approvals(event_key,decision,reason,decided_by,decided_at) VALUES(?,?,?,?,?)',(ek,'APPROVED','',st.session_state.user_name,now)); conn.commit(); audit(conn,'SCHEDULE_APPROVED','event',ek,'',st.session_state.user_name); st.success('Approved.')
            with c2:
                reason=st.text_input('Denial reason',key='denyreason')
                if st.button('Deny Schedule',use_container_width=True):
                    if not reason.strip(): st.error('A denial reason is required.')
                    else:
                        now=datetime.now().isoformat(timespec='seconds')
                        conn.execute('INSERT OR REPLACE INTO approvals(event_key,decision,reason,decided_by,decided_at) VALUES(?,?,?,?,?)',(ek,'DENIED',reason.strip(),st.session_state.user_name,now)); conn.commit(); audit(conn,'SCHEDULE_DENIED','event',ek,reason.strip(),st.session_state.user_name); st.warning('Denied; return this employee to rework.')

elif page=='Message Center':
    st.subheader('Message Center')
    rr=st.session_state.run_result
    if rr is None: st.info('Run scheduling first.')
    else:
        req=rr['requirements']
        st.markdown('### Hiring Manager Drafts')
        approvals=pd.read_sql_query("SELECT * FROM approvals WHERE decision='APPROVED'",conn)
        if approvals.empty: st.info('No approved employee schedules yet.')
        else:
            for ek in approvals['event_key']:
                item=manager_email_for_event(req,ek)
                if not item: continue
                to_email,subject,body=item
                with st.expander(f"{subject} → {to_email or 'recipient missing'}"):
                    st.text(body)
                    st.download_button('Download Outlook-ready .eml draft',make_eml(to_email,subject,body),f"{ek}_manager_email.eml",mime='message/rfc822',key='mgr'+ek)
        st.markdown('### FLAG_MANUAL Drafts')
        manual=req[req['Disposition']=='MANUAL_SCHEDULING_REQUIRED']
        routing=load_manual_routing(conn)
        route={norm(r.training_title):r.recipient_email for r in routing.itertuples()}
        if manual.empty: st.info('No manual scheduling requirements in this run.')
        else:
            for idx,r in manual.iterrows():
                recipient=route.get(norm(r['Training_Title']),'')
                subject=f"Manual scheduling request - {r['Training_Title']} - {r['Worker_Name']}"
                body=f"Please schedule the following employee for {r['Training_Title']}.\n\nEmployee: {r['Worker_Name']}\nEmployee email: {r.get('Work_Email') or r.get('Home_Email')}\nEmployee ID: {r['Employee_ID']}\nHire / Position Effective Date: {pd.to_datetime(r['Anchor_Date']).strftime('%m/%d/%Y') if pd.notna(r['Anchor_Date']) else ''}\nPosition: {r['Position_Title']}\nCost Center: {r['Cost_Center_Title']} ({r['Cost_Center_ID']})\nHiring Manager: {r['Hiring_Manager_AD']}\n"
                with st.expander(f"{r['Worker_Name']} | {r['Training_Title']} → {recipient or 'routing not configured'}"):
                    st.text(body)
                    st.download_button('Download .eml draft',make_eml(recipient,subject,body),f"manual_{idx}.eml",mime='message/rfc822',key='man'+str(idx))

elif page=='Training Configurator':
    st.subheader('Training Documentation Configurator')
    rules=load_rules(conn,False)
    query=st.text_input('Search rules (training, job code, cost center, supervisory org)')
    view=rules.copy()
    if query:
        q=query.lower(); mask=view.astype(str).apply(lambda c:c.str.lower().str.contains(q,na=False)).any(axis=1); view=view[mask]
    st.dataframe(view[['id','job_code','cost_center','supervisory_org','training_title','prerequisite','scheduling_policy','timing_modifier','timing_min','timing_max','active','version']].head(1000),use_container_width=True,height=400)
    tabs=st.tabs(['Add Rule','Edit Rule','Export Documentation'])
    with tabs[0]:
        with st.form('add_rule'):
            job=st.text_input('Job Code'); cc=st.text_input('Cost Center'); so=st.text_input('Supervisory Organization (optional)'); title=st.text_input('Training Title'); pre=st.text_input('Prerequisite (optional)'); topic=st.text_input('Topic'); policy=st.selectbox('Scheduling Policy',['TARGET_RANGE','FIRST_AVAILABLE','FLAG_MANUAL']); mod=st.selectbox('Timing modifier',['','between','after','before']); mn=st.number_input('Timing min days',value=0.0); mx=st.number_input('Timing max days',value=0.0); submitted=st.form_submit_button('Add Rule')
            if submitted and title.strip():
                now=datetime.now().isoformat(timespec='seconds'); conn.execute('''INSERT INTO training_rules(job_code,cost_center,supervisory_org,training_title,prerequisite,topic,scheduling_policy,timing_modifier,timing_min,timing_max,active,created_by,created_at,modified_by,modified_at,version) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1)''',(job,cc,so,title,pre,topic,policy,mod,mn,mx,1,st.session_state.user_name,now,st.session_state.user_name,now)); conn.commit(); audit(conn,'RULE_CREATED','training_rule',title,'',st.session_state.user_name); st.success('Rule added.')
    with tabs[1]:
        rid=st.number_input('Rule ID',min_value=1,step=1)
        row=conn.execute('SELECT * FROM training_rules WHERE id=?',(int(rid),)).fetchone()
        if row:
            with st.form('edit_rule'):
                title=st.text_input('Training Title',row['training_title']); policy=st.selectbox('Scheduling Policy',['TARGET_RANGE','FIRST_AVAILABLE','FLAG_MANUAL'],index=max(0,['TARGET_RANGE','FIRST_AVAILABLE','FLAG_MANUAL'].index(row['scheduling_policy']) if row['scheduling_policy'] in ['TARGET_RANGE','FIRST_AVAILABLE','FLAG_MANUAL'] else 0)); pre=st.text_input('Prerequisite',row['prerequisite'] or ''); active=st.checkbox('Active',bool(row['active'])); note=st.text_input('Change note (optional)'); save=st.form_submit_button('Save Changes')
                if save:
                    now=datetime.now().isoformat(timespec='seconds'); conn.execute('UPDATE training_rules SET training_title=?,scheduling_policy=?,prerequisite=?,active=?,modified_by=?,modified_at=?,version=version+1 WHERE id=?',(title,policy,pre,int(active),st.session_state.user_name,now,int(rid))); conn.commit(); audit(conn,'RULE_MODIFIED','training_rule',str(rid),note,st.session_state.user_name); st.success('Rule updated and versioned.')
    with tabs[2]:
        st.download_button('Export Training Documentation',export_config_workbook(),'Training_Documentation_Export.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

elif page=='Equivalencies':
    st.subheader('Training Equivalencies')
    eq=load_equivalencies(conn,False); st.dataframe(eq,use_container_width=True)
    with st.form('addeq'):
        req=st.text_input('Required / current training title'); equiv=st.text_input('Equivalent historical training title'); save=st.form_submit_button('Add Equivalency')
        if save and req.strip() and equiv.strip():
            conn.execute('INSERT INTO equivalencies(required_training_title,equivalent_training_title,active,created_by,created_at) VALUES(?,?,?,?,?)',(req.strip(),equiv.strip(),1,st.session_state.user_name,datetime.now().isoformat(timespec='seconds'))); conn.commit(); audit(conn,'EQUIVALENCY_CREATED','equivalency',req,equiv,st.session_state.user_name); st.success('Equivalency added.')

elif page=='Locations':
    st.subheader('Location Master')
    loc=load_locations(conn,False); st.dataframe(loc[['location','city','region','postal_code','latitude','longitude','coordinate_precision','inactive']],use_container_width=True,height=550)
    st.caption(f'Physical sessions beyond {MAX_DISTANCE_MILES:.0f} miles are excluded. Physical sessions are preferred over virtual when an eligible physical option exists.')

elif page=='Manual Routing':
    st.subheader('FLAG_MANUAL Routing')
    rt=load_manual_routing(conn,False); st.dataframe(rt,use_container_width=True)
    with st.form('route'):
        title=st.text_input('Training title'); email=st.text_input('Recipient email'); save=st.form_submit_button('Save Routing')
        if save and title.strip():
            conn.execute('INSERT INTO manual_routing(training_title,recipient_email,active,modified_by,modified_at) VALUES(?,?,?,?,?) ON CONFLICT(training_title) DO UPDATE SET recipient_email=excluded.recipient_email,active=1,modified_by=excluded.modified_by,modified_at=excluded.modified_at',(title.strip(),email.strip(),1,st.session_state.user_name,datetime.now().isoformat(timespec='seconds'))); conn.commit(); audit(conn,'MANUAL_ROUTING_UPDATED','manual_routing',title,email,st.session_state.user_name); st.success('Routing saved.')

elif page=='Audit & Backup':
    st.subheader('Audit & Configuration Backup')
    logs=pd.read_sql_query('SELECT * FROM audit_log ORDER BY id DESC LIMIT 1000',conn); st.dataframe(logs,use_container_width=True,height=400)
    st.download_button('Download Configuration Backup',export_config_workbook(),'Class_Scheduler_Configuration_Backup.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.download_button('Download SQLite Configuration Database',Path(DB_PATH).read_bytes(),'scheduler_config.db',mime='application/octet-stream')
    st.markdown('### Restore configuration backup')
    backup=st.file_uploader('Configuration backup workbook',type=['xlsx'],key='config_restore')
    confirm=st.checkbox('I understand this will replace the current rules/equivalencies/locations/manual routing configuration.')
    if backup is not None and confirm and st.button('Restore Configuration Backup'):
        try:
            sheets=pd.read_excel(backup,sheet_name=None)
            required={'Training Rules','Equivalencies','Locations','Manual Routing'}
            if not required.issubset(set(sheets)):
                st.error('Backup workbook is missing one or more required sheets.')
            else:
                conn.execute('BEGIN')
                for table in ('training_rules','equivalencies','locations','manual_routing'):
                    conn.execute(f'DELETE FROM {table}')
                sheets['Training Rules'].to_sql('training_rules',conn,if_exists='append',index=False)
                sheets['Equivalencies'].to_sql('equivalencies',conn,if_exists='append',index=False)
                sheets['Locations'].to_sql('locations',conn,if_exists='append',index=False)
                sheets['Manual Routing'].to_sql('manual_routing',conn,if_exists='append',index=False)
                conn.commit(); audit(conn,'CONFIGURATION_RESTORED','configuration','backup','Configuration workbook restored',st.session_state.user_name)
                st.success('Configuration restored. Refresh the page before the next scheduling run.')
        except Exception as e:
            conn.rollback(); st.exception(e)
    st.warning('Free Streamlit Community Cloud does not guarantee persistence of files written by the running app. Download a configuration backup after rule/configuration changes. For long-term production use, connect the same app to a persistent database service.')
