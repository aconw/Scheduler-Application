# v2.9.3 Release Notes — Strict Cleanup

## Change control basis
v2.9.3 starts from the v2.9.2 application package. Only the following cleanup changes are included.

### 1. Workday export duplicate safeguard restored
The export now independently de-duplicates by both:
- Person + Training Course
- Person + Selected Session WID

This is defense in depth. The scheduler's existing person/course and person/session protections remain unchanged.

### 2. Workbook preflight diagnostics restored
Preflight results now show workbook byte size, data-row count, and worksheet dimensions/names. If a workbook is not recognized, the validation message is prefixed with the report name.

### 3. Location logic intentionally unchanged
The v2.9.2 engine already selects the nearest eligible physical location before evaluating the session timing at that location. No location-selection code was changed in v2.9.3; this behavior was regression-tested instead.

## Explicitly unchanged
TARGET_RANGE precedence, FIRST_AVAILABLE priority, multi-day session handling, training documentation status, equivalencies, existing Workday training visibility, WID/Employee ID handling, duplicate source-event collapse, prerequisites, seat reservation, 130-mile rule, physical-over-virtual preference, email drafts, and Workday file structure were not intentionally modified.
