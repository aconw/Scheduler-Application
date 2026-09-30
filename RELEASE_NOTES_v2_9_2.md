# v2.9.2 — Strict Change Control

## Intended changes only
- Training Rules `priority` is supported.
- `TARGET_RANGE` is scheduled before `FIRST_AVAILABLE` when both are ready.
- `FIRST_AVAILABLE` same-date/time competition is resolved by lower numeric Priority.
- Available Sessions sharing a WID are grouped into one multi-day offering; all daily intervals block conflicts.
- `TRAINING_DOCUMENTATION_FOUND` and `NO_TRAINING_DOCUMENTATION_FOUND` are recorded explicitly.
- Upload validation uses explicit reader keyword arguments.

## Retained behavior
- Existing Workday Training is displayed and blocks conflicts.
- Existing enrollment and completion suppression, including configured equivalencies, is retained.
- WID-based person identity and Employee ID usage are retained for Workday output.
- Duplicate source staffing-event collapse is retained.
- Seat reservation, prerequisites, location preference, 130-mile limit, physical-over-virtual, manual routing, email drafts, and separate Workday employee/contingent exports are retained.
- No approval/deny step is introduced.
