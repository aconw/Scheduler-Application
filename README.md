# Class Scheduling Batch Tool — Simplified v2.6

This is a stateless Streamlit batch scheduler. It does not use SQLite and does not include an approve/deny workflow.

## Important v2.6 behavior

v2.6 makes the **Equivalencies** worksheet authoritative for both:

- historical completion suppression; and
- current / active enrollment suppression.

Each equivalency row can be `ONE_WAY` or `TWO_WAY`.

### Direction semantics

The Equivalencies worksheet uses these columns:

| required_training_title | equivalent_training_title | relationship_direction | active |
|---|---|---|---|

- `ONE_WAY`: `equivalent_training_title` satisfies `required_training_title`, but not the reverse.
- `TWO_WAY`: either title satisfies the other.
- Blank/missing direction in an older configuration workbook defaults to `ONE_WAY`.

Example already included in the bundled configuration:

`Cardiac Monitoring Blended Learning Session 1 - Sinus, Atrial & Junctional Rhythms` ↔ `Cardiac Monitoring` = `TWO_WAY`

This means an employee already enrolled in either class will not be assigned the other, and historical completion of either class can satisfy the other.

Equivalencies are direct mappings only; the application does not infer additional transitive relationships through chains of equivalencies.

## Duplicate staffing-event protection retained

Before training requirements are generated, the scheduler collapses duplicate source staffing events for the same person when event type, hire/effective date, job/position context, cost center, supervisory organization, physical location, worker type, and traveler designation are identical.

The canonical staffing event is scheduled once. Duplicate source rows remain visible on the **Duplicate Source Events** audit sheet but cannot reserve seats or create Workday enrollments.

## Run workflow

Upload for each run:

- New Hire Orientation Report and/or New / Additional Job Change Report
- Training History / Learning Transcript
- Learning Content / Available Sessions
- Existing Orientation Schedule (required for overlap prevention and existing-enrollment suppression)

Then click **Run Scheduling** and download the export package.

## Export package

`Class_Scheduling_Export_Package.zip` contains:

- `Workday/Enroll_In_Learning_Content_Employees.xlsx`
- `Workday/Enroll_In_Learning_Content_Contingent_Workers.xlsx` when applicable
- `Scheduling_Results_and_Audit.xlsx`
  - Selected Sessions
  - Requirement Audit
  - Review Queue
  - Seat Audit
  - Duplicate Source Events
  - Run Summary

The Requirement Audit includes fields showing the observed completion/enrollment title, equivalency used, direction, and whether a two-way relationship was matched in the forward or reverse direction.

## Scheduling protections

- Candidate Position ID is Job Code for new hires.
- WID is the preferred durable person identity.
- Duplicate source staffing events are collapsed before requirement generation.
- General and supervisory-organization-specific requirements are additive.
- Historical completion permanently satisfies a requirement.
- Configured equivalents apply to both historical completions and active enrollments.
- Existing active enrollment suppresses duplicate registration.
- A person/course is selected only once in a run.
- A person/session is selected only once in a run.
- Employee sessions may not overlap.
- Existing Workday orientation classes are treated as blocked time.
- Physical sessions are preferred to virtual sessions.
- Nearest eligible physical site is prioritized within the 130-mile limit.
- TARGET_RANGE does not schedule outside the configured window.
- FIRST_AVAILABLE chooses the earliest eligible session at the nearest eligible location.
- Seats are consumed inside the run to prevent double assignment.
- Prerequisites must occur earlier and may be chained.
- Configured equivalencies are also considered when checking an external prerequisite completion/enrollment.
- FLAG_MANUAL remains a manual-scheduling disposition.

## Configuration

`config/Scheduler_Configuration.xlsx` contains:

- Training Rules
- Locations
- Equivalencies
- Manual Routing

The configuration workbook may be downloaded, edited, and uploaded back into the app. No server-side database is required.
