# Class Scheduling Application

A runnable Streamlit application for matching employees/staffing events to required training, selecting eligible sessions, suppressing prior completion/existing enrollment, enforcing prerequisites and seat capacity, exporting Workday enrollment files, approving schedules, and generating email drafts.

## What is already included

- 18,567 training rules seeded from the supplied Training Documentation workbook.
- 332 standardized locations with latitude/longitude.
- Empty Training Equivalency configuration ready for maintenance in the UI.
- Empty FLAG_MANUAL routing configuration ready for maintenance in the UI.
- Workday employee and contingent-worker upload templates.
- Audit logging for rule changes, overrides, approvals, denials, and scheduling runs.

## Fastest setup: Streamlit Community Cloud (no local install)

1. Create a new private GitHub repository.
2. Upload the contents of this folder to the repository root. Keep the `assets/` and `data/` folders.
3. Sign in at https://share.streamlit.io/ using your Streamlit account.
4. Select **Create app** / **New app**, choose the GitHub repository, branch `main`, and set the main file path to `app.py`.
5. Deploy.
6. Open **New Scheduling Run** in the app and upload the operational reports for that run.

> Important: Streamlit Community Cloud's local filesystem is ephemeral. Configuration edits made through the deployed app may be lost after a restart/redeploy. Use **Audit & Backup → Download Configuration Backup** after changes. For long-term production, point `database.py` at a persistent SQL database or hosted storage.

## Run locally later (optional)

Install Python 3.11+ and run:

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Streamlit will open the application in your browser.

## Operational reports uploaded each run

- New Hire Orientation Report and/or New/Additional Job Change Orientation Report
- Training History / Learning Transcript
- Learning Content / Available Sessions report
- Orientation Schedule Report is optional during initial scheduling and is uploaded later for Workday verification

Training documentation and location configuration are already stored in the application database and do not need to be uploaded each run.

## Scheduling rules implemented

- WID is the durable person identity; Employee ID is treated as a current business identifier.
- New hires use hire date; job changes use position effective date.
- Job Code + Cost Center rules and matching Supervisory Organization-specific rules are additive.
- Any historical exact or configured equivalent completion suppresses reassignment permanently.
- Existing active enrollment suppresses duplicate enrollment.
- Physical sessions are preferred over virtual.
- Nearest eligible physical location takes precedence, with a 130-mile maximum.
- TARGET_RANGE never auto-schedules outside the configured range; it flags the requirement for review.
- FIRST_AVAILABLE chooses the earliest eligible session at the nearest physical training location.
- Seats are reserved within a run so capacity cannot be double-assigned.
- Prerequisites must be scheduled earlier; a prerequisite may itself have prerequisites.
- FLAG_MANUAL generates one draft request per employee/class.
- Manual override is allowed and requires an audit reason.
- Employee and contingent-worker Workday exports are separate.
- Schedule approval is employee/staffing-event level; denial requires a reason.
- Approved schedules can generate one hiring-manager `.eml` draft per employee.

## Known assumption to validate

The supplied New Hire report does not expose a column literally named `Job Code`. The current engine uses `Candidate Position ID` for new-hire rule matching, because that was the best available field in the supplied export. If a true Job Code can be added to that Workday report, update the `Job_Code` mapping in `scheduler_engine.py`.

## Email drafts

The app downloads messages as `.eml` files with the `X-Unsent: 1` header. Outlook commonly opens these as editable unsent messages. No messages are automatically sent in this version.
