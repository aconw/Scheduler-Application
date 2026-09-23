# START HERE — Simplified Stateless Build v2.0

This is a **complete replacement** for the earlier database-backed Streamlit build.

## What changed
- No SQLite database.
- No persistent application state or configuration writes on Streamlit Cloud.
- Every uploaded workbook is converted to a stable byte copy before parsing.
- Every required report is validated on screen before **Run Scheduling** is enabled.
- Training rules, locations, equivalencies, and manual-routing settings live in one Excel configuration workbook.
- Approval happens in the generated Excel audit workbook rather than being stored inside Streamlit.

## Replace the old GitHub app
For the cleanest deployment, replace the old repository contents with the contents of this folder.

At the repository root you should have:

```
app.py
batch_utils.py
scheduler_engine.py
requirements.txt
README.md
START_HERE.md
assets/
config/
```

You can remove the old `database.py`, `seed_config.py`, `data/`, and `seed_data/` files. They are not used by this build.

Commit the replacement to `main`. Streamlit should redeploy automatically. Reboot the app once after deployment.

## Verify the right build is live
Directly below the title you should see:

**Simplified stateless build v2.0**

If you do not see that text, the replacement deployment has not taken effect yet.

## First run
Upload the Workday reports. Before the run button becomes active, the screen should show green messages similar to:

- Training History: ... 4,135 data rows
- Available Sessions: ... 7,456 data rows

Only then select **Run Scheduling**.

## Results
Download `Class_Scheduling_Results_Package.zip`. It contains:

- `Scheduling_Audit_and_Approval.xlsx`
- Employee Workday enrollment file
- Contingent-worker Workday file when applicable

## Approval and email drafts
In `Scheduling_Audit_and_Approval.xlsx`, use the **Employee Approval** sheet. Enter `APPROVED` or `DENIED` in the Approval column and a Denial Reason when denied. Save it and upload it back to the app. The app will generate `.eml` manager-email drafts for approved staffing events plus FLAG_MANUAL scheduling drafts.

## Configuration
The master configuration is `config/Scheduler_Configuration.xlsx`. Keep your controlled copy in SharePoint or OneDrive. It contains:

- Training Rules
- Locations
- Equivalencies
- Manual Routing

For temporary testing, upload an edited configuration workbook in the app. For a permanent configuration change, review the workbook and replace the copy under `config/` in GitHub.
