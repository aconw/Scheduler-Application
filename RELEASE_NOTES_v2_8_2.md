# Release Notes — Simplified v2.8.2

## Strict change-control enhancement

v2.8.2 adds one functional enhancement to the exact v2.8.1 GitHub baseline supplied by the user: explicit training-documentation completeness status for every canonical staffing event.

### Status definitions
- `TRAINING_DOCUMENTATION_FOUND`: at least one active Training Rule exists for the staffing event's Job Code + Cost Center.
- `NO_TRAINING_DOCUMENTATION_FOUND`: no active Training Rule exists for that Job Code + Cost Center.

A missing-documentation event is not interpreted as “no training required.” No Training Rule assignments are generated for that event, but the employee/event remains visible in the application and exported audit.

### Results and audit
The application now shows a **Training Documentation Status** table. The exported `Scheduling_Results_and_Audit.xlsx` contains a **Training Documentation Status** worksheet with employee/event identifiers, Job Code, Cost Center, status, matching rule count, and explanation. The Run Summary also counts documentation-found and no-documentation-found events.

Requirement Audit rows generated from documented events also carry `Training_Documentation_Status`.

### Unchanged
No intentional changes were made to TARGET_RANGE/FIRST_AVAILABLE ordering, FIRST_AVAILABLE priority ranking, multi-day WID handling, overlap logic, seat reservation, prerequisites, equivalencies, existing-enrollment handling, location selection, duplicate-source-event handling, Workday templates, Workday export logic, or email drafts.

### Validation
A synthetic regression with one documented and one undocumented staffing event verified that the documented event schedules normally, the undocumented event receives `NO_TRAINING_DOCUMENTATION_FOUND`, the undocumented event generates no Training Rule requirements, and both events remain represented in the documentation-status output.
