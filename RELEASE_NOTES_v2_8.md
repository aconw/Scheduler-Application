# Release Notes — Simplified v2.8

## Strict change control
The intended v2.8 scope is limited to the two requested scheduling-engine changes:

1. Selection order: TARGET_RANGE requirements are selected before FIRST_AVAILABLE requirements, with prerequisite dependency order remaining a hard constraint. Among FIRST_AVAILABLE requirements, the earliest potential selection time is considered before numeric class priority; when candidates tie on selection time, lower numeric priority is processed first. Blank/unusable priority is treated as lowest priority.
2. Multi-day sessions: Available Sessions rows sharing the same session WID are treated as one class offering. Seat capacity is reserved once for the whole WID. Every daily start/end interval is retained and used for overlap detection. The selected result reports the overall earliest start, latest end, and day count, while conflict checking uses each individual day interval.

No approval/deny workflow, training-rule audience logic, equivalency configuration model, existing-enrollment workflow, duplicate-source-event behavior, location-distance rule, seat reservation concept, prerequisite concept, Workday templates, or email-draft concept was intentionally changed.

## Configuration
The packaged `config/Scheduler_Configuration.xlsx` is the user's current uploaded workbook and includes the new `priority` column on Training Rules. The workbook was copied without changing its business-rule content.

## Multi-day behavior
If the Available Sessions report contains:

- WID = `ABC`
- 10/01 8:00-12:00
- 10/02 8:00-12:00
- 10/03 8:00-12:00

v2.8 treats all three rows as one `ABC` class with one seat allocation and three occupied daily intervals. A different training on 10/02 10:00-11:00 will conflict with `ABC`; a training on 10/04 will not.

## Regression testing
Using the current configuration workbook and the established test population:

- raw staffing events: 412
- canonical staffing events after duplicate-source collapse: 381
- duplicate source events: 31
- available-session day rows: 7,456
- grouped session WIDs: 4,849
- selected requirements: 786
- review required: 506
- manual scheduling required: 98
- historical/equivalent completion: 2
- multi-day selected classes: 20
- selected-session overlap violations: 0

Synthetic tests also verified:

- TARGET_RANGE is selected before FIRST_AVAILABLE when their candidate sessions conflict at the same time;
- a higher-priority FIRST_AVAILABLE requirement wins a same-time conflict when priority is 1 vs 2;
- a multi-day WID consumes one seat for the class and blocks every daily interval;
- a multi-day class does not block a separate class on a non-overlapping day.
