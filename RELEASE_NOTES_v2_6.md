# Release Notes — Simplified v2.6

## Primary change: directional equivalencies apply to both completion and enrollment suppression

Prior versions used configured equivalencies when checking historical completions, but existing/current enrollment suppression still required an exact training-title match.

v2.6 uses the same Equivalencies worksheet for both checks.

If a requirement is satisfied by a configured active enrollment, the requirement is marked `ALREADY_ENROLLED`, no new seat is reserved, and no Workday enrollment row is created.

The Requirement Audit records:

- `Existing_Enrollment_Training`
- `Equivalency_Used`
- `Equivalency_Direction`
- `Equivalency_Match_Direction`
- a human-readable explanation of why the requirement was suppressed

## ONE_WAY vs TWO_WAY

An Equivalencies row now has `relationship_direction`.

### ONE_WAY

`equivalent_training_title` satisfies `required_training_title` only.

Example:

- Required: `Course B`
- Equivalent: `Legacy Course A`
- Direction: `ONE_WAY`

Completion/enrollment in `Legacy Course A` satisfies `Course B`, but completion/enrollment in `Course B` does not automatically satisfy a separate requirement for `Legacy Course A`.

### TWO_WAY

Either title satisfies the other.

The bundled v2.6 configuration contains:

- Required: `Cardiac Monitoring Blended Learning Session 1 - Sinus, Atrial & Junctional Rhythms`
- Equivalent: `Cardiac Monitoring`
- Direction: `TWO_WAY`

Therefore:

- `Cardiac Monitoring` satisfies the Session 1 requirement; and
- Session 1 satisfies a `Cardiac Monitoring` requirement.

Blank/missing direction values from older configuration files default to `ONE_WAY` for safe backward compatibility.

## Prerequisite behavior

When the scheduler must satisfy a prerequisite from prior history or an existing dated enrollment, configured equivalents are also considered.

## Protections retained from v2.5

- duplicate source staffing events are collapsed before requirement generation;
- one selected course per person/WID per run;
- one selected session per person/WID per run;
- no overlapping selected sessions;
- no overlap with the uploaded Existing Orientation Schedule;
- Workday export deduplication by person + course and person + session.

## Validation performed

Synthetic regression tests verified:

- ONE_WAY forward historical completion suppression;
- ONE_WAY reverse does not suppress;
- ONE_WAY forward active-enrollment suppression;
- TWO_WAY reverse historical completion suppression;
- TWO_WAY reverse active-enrollment suppression;
- the configured Cardiac Monitoring relationship suppresses a Session 1 assignment when `Cardiac Monitoring` is already actively enrolled.

The original full scheduling population also ran successfully after the change.
