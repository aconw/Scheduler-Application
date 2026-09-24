# Release Notes — Simplified v2.7

## 1. Equivalency-aware active enrollment
The configured ONE_WAY/TWO_WAY equivalency relationships are used consistently for both historical completion suppression and active enrollment suppression.

## 2. Existing Workday Training visibility
The Orientation Schedule — Lesson-Level Detail report is now retained as a session-level view in the results. All active enrolled sessions found for the uploaded population are shown, whether or not they are required for the current role.

The source report is lesson-level, so duplicate rows are collapsed by person + course offering + start + end for display.

Every valid existing session also blocks overlapping newly selected sessions. Back-to-back sessions are allowed when one ends exactly as another begins.

## 3. Audit workbook
`Scheduling_Results_and_Audit.xlsx` now includes an `Existing Workday Training` sheet in addition to Selected Sessions, Requirement Audit, Review Queue, Seat Audit, Duplicate Source Events, and Run Summary.

## 4. No approval step
The workflow remains stateless and batch-based. There is no approve/deny stage.
