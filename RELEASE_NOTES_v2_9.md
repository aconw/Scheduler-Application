# v2.9 Release Notes

## Fixed scheduling precedence
The scheduler now explicitly chooses from ready `TARGET_RANGE` requirements before any ready `FIRST_AVAILABLE` requirement. This fixes the Adeira Wright scenario where a `FIRST_AVAILABLE` CRMH Nursing Services - Onboarding session was selected before a `TARGET_RANGE` Nursing Ancillary Academy - Day 1 Inpatient session.

The scheduler records `Selection_Order` and `Selection_Phase` so the decision sequence is visible in the Requirement Audit.

## Priority ranking
Training Rules `priority` is read from the configuration workbook. Lower numeric values are higher priority. Blank priority is lowest. Priority resolves a same-date/time competition between different `FIRST_AVAILABLE` requirements after location/session eligibility is established.

## Multi-day session handling
Available Sessions rows sharing the same session WID are grouped into one offering. The offering consumes one seat, and each daily Start Date/End Date interval participates in overlap prevention. The audit records day count, multi-day status, and all daily intervals.

## Training documentation status
Staffing events with no matching training documentation are explicitly recorded as `NO_TRAINING_DOCUMENTATION_FOUND`; they are not silently interpreted as having no training requirements.

## Existing Workday training
All active existing sessions from the Orientation Schedule report remain visible in the results, whether or not they are required for the employee's role, and their time blocks prevent overlapping new selections.

## Existing protections retained
- WID-first person identity
- duplicate source-event collapse
- person/course and person/session duplicate protection
- equivalency-aware completion and active-enrollment suppression
- seat reservation
- prerequisites
- physical-location preference and 130-mile maximum
- virtual fallback
- separate employee and contingent-worker Workday exports
- no approval/denial workflow
