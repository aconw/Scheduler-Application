# Class Scheduling Batch Tool — Simplified v2.8.3

Stateless Streamlit batch scheduler. No approval/deny workflow.

## v2.8.3 — Job Change completion exemption after separation
The Job Change report's `Last Separation Date` is now read. For `JOB_CHANGE` events, a previous/equivalent completion can exempt the employee only when the completion date is strictly later than the last separation date. Blank separation date means no prior separation, so the prior completion behavior is unchanged. A completion on or before the separation date does not exempt the job change and the requirement continues through normal scheduling.

The Requirement Audit retains `Last_Separation_Date`, `Completion_Date`, and the explanation.

All other scheduler behavior is retained from v2.8.2.
