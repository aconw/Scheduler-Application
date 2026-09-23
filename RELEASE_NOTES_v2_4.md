# Simplified Scheduler v2.4

## Primary fix: person-level duplicate enrollment prevention

The scheduler still evaluates every staffing event independently for training requirements, but it now separates **requirement reasons** from **Workday enrollments**.

If the same WID/person needs the same training because of more than one staffing event:

- the first valid assignment consumes one seat and is marked `PROPOSED_SCHEDULE`;
- a later requirement for the same person/course is marked `SATISFIED_BY_SAME_RUN_ASSIGNMENT` when that session also satisfies the later event's timing/prerequisite rule;
- the later audit row references `Satisfied_By_Event_Key` and `Satisfied_By_Session_WID`;
- no second seat is consumed;
- no second Workday enrollment row is created.

If the existing same-run assignment does **not** satisfy the later staffing event's timing/prerequisite rule, the later requirement becomes `REVIEW_REQUIRED`; the scheduler still does not create a duplicate enrollment.

## Export defense

The Workday export performs a second independent deduplication before writing rows:

- one row per person + training course;
- one row per person + selected session.

WID is the preferred person key. If Employee ID changes for a WID, the most recent staffing-event Employee ID is used for the Workday learner identifier.

## Existing protections retained

- no overlapping selected sessions for the same person;
- existing Workday schedule is blocked time;
- in-run seat reservation;
- prerequisite ordering;
- nearest eligible physical site within 130 miles;
- physical before virtual;
- completion/equivalency and existing-enrollment suppression;
- separate employee and contingent-worker Workday files;
- full Requirement Audit retained.

## Regression checks

Using the original validation population:

- 644 unique selected enrollments;
- 55 additional staffing-event requirements satisfied by a same-run assignment;
- 0 duplicate person + course selected rows;
- 0 duplicate person + session selected rows;
- 0 overlaps among selected sessions.

Using the newer `test history 2026-09-22 13_46 EDT.xlsx` transcript:

- 600 unique selected enrollments;
- 47 additional requirements satisfied by a same-run assignment;
- 0 duplicate person + course selected rows;
- 0 duplicate person + session selected rows.
