# Class Scheduling Batch Tool — Simplified v2.4

This version is a stateless Streamlit batch scheduler. It does not use SQLite and it does not include an approve/deny workflow.

## Run workflow

Upload these reports for each run:

- New Hire Orientation Report and/or New / Additional Job Change Report
- Training History / Learning Transcript
- Learning Content / Available Sessions
- Existing Orientation Schedule **(required for overlap prevention)**

Then click **Run Scheduling** and download the export package.

## Export package

`Class_Scheduling_Export_Package.zip` contains:

- `Workday/Enroll_In_Learning_Content_Employees.xlsx` — exact employee Workday upload format
- `Workday/Enroll_In_Learning_Content_Contingent_Workers.xlsx` — generated only when applicable
- `Scheduling_Results_and_Audit.xlsx`
  - `Selected Sessions`
  - `Requirement Audit`
  - `Review Queue`
  - `Seat Audit`
  - `Run Summary`

There is no approval sheet and no review workbook that must be uploaded back into the app.


## Person-level duplicate prevention (v2.4)

Every staffing event is still evaluated independently, but the same WID/person is enrolled in a given training course only once per run. Additional staffing events requiring that same course remain visible in the Requirement Audit as `SATISFIED_BY_SAME_RUN_ASSIGNMENT` and reference the enrollment that satisfies them. The Workday export also independently deduplicates by person + course and person + session.

## Non-overlap rule

The scheduler will not select a session if it overlaps another class for the same employee.

It checks against:

1. Classes already scheduled in the uploaded Existing Orientation Schedule report.
2. Classes newly selected earlier in the current scheduling run, including selections generated for another staffing event belonging to the same employee.

Intervals use standard half-open timing: a class ending at exactly 10:00 AM may be followed by a class starting at exactly 10:00 AM.

If otherwise-eligible sessions all conflict with the employee's schedule, the requirement becomes `REVIEW_REQUIRED` and the Requirement Audit explains the conflict.

## Other scheduling rules retained

- WID is the durable identity where available; Employee ID is the business identifier/fallback.
- Candidate Position ID is Job Code for new hires.
- General and supervisory-organization-specific training requirements are additive.
- Any historical completion permanently satisfies the requirement.
- Configured equivalent historical courses may satisfy a current requirement.
- Existing active enrollment suppresses duplicate registration.
- Physical sessions are preferred to virtual sessions.
- Nearest eligible physical site is prioritized, with a 130-mile limit.
- TARGET_RANGE does not automatically schedule outside the required range.
- FIRST_AVAILABLE chooses the earliest eligible session at the nearest eligible location.
- Seats are consumed within the run to prevent double assignment.
- Prerequisites must be scheduled earlier and can be chained.
- FLAG_MANUAL remains a manual-scheduling disposition.

## Configuration

`config/Scheduler_Configuration.xlsx` is the configuration source and contains:

- Training Rules
- Locations
- Equivalencies
- Manual Routing

You can download/edit/re-upload this workbook from the app. No server-side database is required.

## Email drafts

Email drafts can be downloaded directly after scheduling. They are `.eml` drafts only; the app does not send messages automatically.


## v2.4 startup simplification

`batch_utils.py` has been removed. All upload validation, export-package creation, and email-draft helpers now live directly in `app.py`, preventing mixed-version import errors.
