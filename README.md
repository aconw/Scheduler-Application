# Class Scheduling Batch Tool — Simplified v2.1

This version fixes Workday Excel exports whose worksheet dimension metadata incorrectly reports `A1:A1`. The application deliberately opens Workday `.xlsx` files in normal OpenPyXL mode rather than read-only mode so all rows are detected.

# Class Scheduling Batch Tool — Simplified Build

This version is intentionally stateless. It does not use SQLite and does not rely on persistent Streamlit storage.

## Deploy
1. Create/update your GitHub repository.
2. Upload all files and folders from this application folder so `app.py` is at the repository root.
3. In Streamlit Community Cloud, deploy `main/app.py`.
4. Reboot the app after replacing the old build.

## Every scheduling run
1. Upload New Hire and/or Job Change report.
2. Upload Training History.
3. Upload Available Sessions.
4. Optionally upload Existing Orientation Schedule.
5. Confirm every required report shows a green preflight validation message.
6. Click **Run Scheduling**.
7. Download **Class_Scheduling_Results_Package.zip**.

The package contains the Workday files plus `Scheduling_Audit_and_Approval.xlsx`.

## Approval / emails
1. Open `Scheduling_Audit_and_Approval.xlsx`.
2. On **Employee Approval**, enter `APPROVED` or `DENIED` for each staffing event. Enter a denial reason when denied.
3. Save the workbook.
4. Upload it back to the app under **Generate Email Drafts After Review**.
5. Download the `.eml` email-draft ZIP.

## Configuration
`config/Scheduler_Configuration.xlsx` is the source of truth for:
- Training Rules
- Locations
- Equivalencies
- FLAG_MANUAL routing

Keep the master copy in SharePoint or OneDrive. The app includes a download button for the configuration workbook. You can edit it in Excel and upload the edited workbook in the Configuration section for a run. For permanent changes, replace `config/Scheduler_Configuration.xlsx` in GitHub with the reviewed version.

## Important mapping
For New Hire records, **Candidate Position ID = Job Code**.

## Business rules included
- WID is preferred as durable identity; Employee ID is used as the business identifier/fallback.
- Historical completion permanently satisfies a requirement.
- Configured equivalents satisfy the current course.
- Existing registration suppresses duplicate scheduling.
- Job Code + Cost Center and matching Sup Org rules are additive.
- TARGET_RANGE stays inside the range or goes to review.
- FIRST_AVAILABLE prioritizes nearest physical location, then earliest eligible session there.
- Physical wins over virtual when eligible.
- 130-mile maximum for physical training.
- Seats are consumed during a run.
- Prerequisites must be scheduled earlier.
- Employee and contingent-worker Workday files are separate.
