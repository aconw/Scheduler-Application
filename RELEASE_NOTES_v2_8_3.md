# Release Notes — Simplified v2.8.3

## Requested change only
Read `Last Separation Date` from the Job Change report and apply it to prior-completion exemptions.

### Rule
For `JOB_CHANGE`: completed or equivalent-completed training satisfies a requirement only when `Completion Date > Last Separation Date`. Blank `Last Separation Date` means no prior separation and preserves existing behavior. Equality does not qualify.

### Audit
Requirement Audit retains `Last_Separation_Date`, `Completion_Date`, `Completion_Training`, and a dated explanation.

### Unchanged
Session policy ordering, priority, multi-day handling, overlap protection, existing enrollment suppression, equivalencies, duplicate source-event collapse, location logic, seat allocation, Workday export, and email drafts.
