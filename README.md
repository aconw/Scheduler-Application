# Class Scheduling Batch Tool — Simplified v2.8

This is a stateless Streamlit batch scheduler. It does not use SQLite and does not include an approve/deny workflow.

## v2.8 changes

### Scheduling selection order
Within each staffing event, hard prerequisites are processed first. After dependency order is satisfied:

1. `TARGET_RANGE`
2. `FIRST_AVAILABLE`
3. `FLAG_MANUAL` remains a manual disposition

For `FIRST_AVAILABLE`, the scheduler considers the potential selection date/time. When different FIRST_AVAILABLE requirements have the same potential selection time, the lower numeric `priority` in Training Rules is processed first. Blank priority is lowest.

### Multi-day sessions
All Available Sessions rows sharing the same WID are one class offering. The scheduler reserves capacity once for the WID and stores every daily interval for conflict detection. The selected result displays the overall start, overall end, and day count.

## Configuration
`config/Scheduler_Configuration.xlsx` contains:
- Training Rules
- Locations
- Equivalencies
- Manual Routing

The packaged configuration is the current user-supplied workbook, including the Training Rules `priority` column.

## Existing features retained
- Candidate Position ID = Job Code for new hires.
- WID is the preferred durable person identifier.
- Duplicate source staffing events are collapsed before requirement generation.
- Historical completion permanently satisfies a requirement.
- Directional ONE_WAY/TWO_WAY equivalencies apply to completion and active-enrollment suppression.
- Existing active Workday training is displayed and blocks overlapping selections.
- Physical sessions are preferred to virtual; nearest eligible physical location is selected within 130 miles.
- Seats are reserved during the run.
- Prerequisites are scheduled earlier and may be chained.
- Employee sessions may not overlap.
- Employee and contingent worker Workday files remain separate.
- Requirement Audit and review/audit sheets remain in the exported package.
- Email drafts remain optional and are not sent automatically.

## Run workflow
Upload the New Hire and/or Job Change report, Training History, Available Sessions, and Existing Orientation Schedule. Confirm preflight validation, run scheduling, and download the results package.
