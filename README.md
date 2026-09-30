# Class Scheduling Application — Clean v2.9.2

v2.9.2 is a stateless Streamlit batch scheduler. It is based on the last known-good v2.8 behavior with only the requested scheduling changes applied.

## Run workflow
Upload:
- New Hire Orientation Report and/or New / Additional Job Change Report
- Training History / Learning Transcript
- Learning Content / Available Sessions
- Existing Orientation Schedule — Lesson-Level Detail

Click **Run Scheduling** and download the export package.

## Scheduling order
1. Resolve requirements and prerequisites.
2. Schedule ready `TARGET_RANGE` requirements before ready `FIRST_AVAILABLE` requirements.
3. For competing `FIRST_AVAILABLE` requirements with the same best feasible start date/time, lower numeric Priority is selected first. Blank Priority is lowest.
4. Existing training and newly selected sessions always block overlapping time.

Prerequisite dependencies can require a prerequisite to be scheduled before a TARGET_RANGE dependent course.

## Multi-day sessions
Available Sessions is lesson-level. Rows sharing the same session WID are grouped into one offering. One seat is consumed, and every daily Start/End interval participates in conflict checking.

## Documentation status
Each staffing event is explicitly classified as:
- `TRAINING_DOCUMENTATION_FOUND`
- `NO_TRAINING_DOCUMENTATION_FOUND`

A staffing event with no documentation does not generate automatic Workday enrollments.

## Preserved functionality
- WID-first durable person identity and current Employee ID handling
- existing Workday training visibility
- equivalencies for completion and active enrollment
- duplicate source-event collapse
- same-person/course and same-person/session protections
- seat reservation
- location matching and 130-mile limit
- physical-over-virtual preference
- prerequisites and chained prerequisites
- FLAG_MANUAL and configurable routing
- manager and manual email drafts (drafts only)
- separate employee and contingent-worker Workday files
- full Requirement Audit and scheduling audit package
