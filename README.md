# Class Scheduling Application v2.9.2

This release is a strict-change-control cleanup of the previously validated v2.8 behavior. The only scheduling changes are the agreed v2.9 changes: training-rule Priority, policy precedence (`TARGET_RANGE` before `FIRST_AVAILABLE`), multi-day session handling, and explicit training documentation status. The upload reader is called with explicit keyword arguments to prevent header lists from being interpreted as worksheet names.

Existing email drafts, Workday exports, equivalencies, existing-enrollment suppression, duplicate source-event handling, location rules, seat reservation, prerequisites, Existing Workday Training, and no-overlap behavior are retained.

## Scheduling order
`TARGET_RANGE` requirements are scheduled before `FIRST_AVAILABLE` requirements when ready. Prerequisites can force prerequisite-first ordering. For different `FIRST_AVAILABLE` requirements competing at the same date/time, lower numeric Priority is used first; blank Priority is lowest.

## Multi-day sessions
Available Sessions rows sharing a session WID are treated as one offering and consume one seat. Every daily Start/End interval is used for conflict checking. The audit records day count, multi-day flag, and daily intervals.

## Training documentation status
A staffing event with matching rules receives `TRAINING_DOCUMENTATION_FOUND`. A staffing event with no matching rules receives `NO_TRAINING_DOCUMENTATION_FOUND` and is not scheduled until documentation is added.
