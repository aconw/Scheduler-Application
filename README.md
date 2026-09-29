# Class Scheduling Application v2.9

## Scheduling order

For each canonical employee staffing event, the scheduler uses a dependency-aware policy order:

1. `TARGET_RANGE` requirements are selected before `FIRST_AVAILABLE` requirements when both are ready to schedule.
2. A prerequisite must be resolved before its dependent class, so a prerequisite can override the general policy phase.
3. For `FIRST_AVAILABLE`, location remains the primary session-selection rule: nearest eligible physical location first, then earliest session at that location; virtual is used when no eligible physical session is available within 130 miles.
4. When different `FIRST_AVAILABLE` requirements compete for the same date/time, lower numeric `Priority` wins (`1` is higher than `2`). Blank Priority is lowest.
5. Classes with the same session WID are treated as one offering. Multi-day offerings consume one seat and every daily interval is checked for employee conflicts.
6. Existing Orientation Schedule sessions block overlapping new selections.

## Additional protections

- WID is the durable person identifier.
- Exact duplicate staffing events are collapsed before requirements are created.
- One person/course and one person/session may reach the Workday export only once.
- Historical completion and active enrollment can use configured equivalencies with `ONE_WAY` or `TWO_WAY` relationships.
- Missing training documentation is reported as `NO_TRAINING_DOCUMENTATION_FOUND` rather than silently treated as no training required.

## Outputs

The ZIP export contains Workday enrollment files plus `Scheduling_Results_and_Audit.xlsx` with Selected Sessions, Requirement Audit, Existing Workday Training, Duplicate Source Events, Seat Audit, and Run Summary.
