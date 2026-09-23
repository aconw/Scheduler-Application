# Release Notes — Simplified v2.5

## Primary fix: duplicate staffing events are collapsed before scheduling

v2.4 protected the Workday export from duplicate person/course and person/session rows, but identical staffing rows could still independently generate requirement chains before that protection ran.

v2.5 moves duplicate handling to the start of the scheduling pipeline.

A source staffing event is treated as a duplicate when all of the following match:

- durable person identity (WID when available; otherwise Employee ID)
- event type (New Hire or Job Change)
- hire/effective date
- job code
- position title
- cost center
- supervisory organization
- physical location
- worker type
- traveler designation

Only the first row becomes the canonical staffing event. The duplicate row does not generate training requirements, reserve seats, or create Workday enrollments.

Duplicate rows are preserved in the exported audit workbook on **Duplicate Source Events**, including the duplicate `Event_Key` and the `Canonical_Event_Key` it was collapsed into.

## Defense-in-depth protections retained

The application still independently enforces:

- one selected training course per person/WID per run;
- one selected session per person/WID per run;
- no overlapping selected sessions;
- no overlap with the uploaded Existing Orientation Schedule;
- Workday export deduplication by person + course and person + session.

## Audit change

`Scheduling_Results_and_Audit.xlsx` now includes:

- Selected Sessions
- Requirement Audit
- Review Queue
- Seat Audit
- **Duplicate Source Events**
- Run Summary

The Selected Sessions sheet now also shows `Disposition` explicitly to make prerequisite status and row disposition easier to distinguish.

## Regression validation

Against the original validation population:

- Raw staffing rows: 412
- Canonical staffing events after duplicate collapse: 381
- Duplicate source events identified: 31
- Selected sessions: 644
- Duplicate person/course selections: 0
- Duplicate person/session selections: 0

The reduction in requirement count is intentional because duplicate source events no longer generate duplicate requirement chains.
