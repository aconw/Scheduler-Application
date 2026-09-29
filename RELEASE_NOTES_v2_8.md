# v2.8 Release Notes

- Added Training Rules `priority` support.
- Scheduling processes `TARGET_RANGE` requirements before `FIRST_AVAILABLE` requirements, except prerequisite dependencies can require a prerequisite to be scheduled first.
- For competing `FIRST_AVAILABLE` requirements with the same feasible start date/time, lower numeric Priority is selected first; blank priority is lowest priority.
- Available Sessions are grouped by session WID. A multi-day offering is treated as one enrollment/seat and every daily lesson interval is used for overlap prevention.
- Requirement Audit now includes Priority, selection order/phase, multi-day flag, day count, and daily session intervals.
- The current user-supplied `Scheduler_Configuration.xlsx` is bundled as the v2.8 configuration source.
