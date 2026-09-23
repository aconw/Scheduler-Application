
## Important when replacing v2.2

Delete the old `batch_utils.py` file from GitHub. v2.3 no longer uses it. The repository root should contain `app.py` and `scheduler_engine.py`, but **not** `batch_utils.py`.

# Replace the Existing Streamlit App with v2.3

This is a complete replacement build, not a patch.

## GitHub

Replace the files in the current repository with the contents of this folder. Your repository root should include:

```
app.py
scheduler_engine.py
requirements.txt
README.md
START_HERE.md
assets/
config/
```

Remove older database-related files/folders if they are still present; v2.3 does not use them.

Commit the changes to `main`. Streamlit should redeploy automatically. Reboot once if needed.

## Confirm the correct version

The subtitle beneath the app title must say:

**Simplified stateless build v2.3 • No approval step • Non-overlapping scheduling • Workday files + requirement audit**

## Each scheduling run

Upload:

1. New Hire report and/or Job Change report
2. Training History
3. Available Sessions
4. Existing Orientation Schedule (required for conflict checking)

All preflight checks must be green before **Run Scheduling** becomes available.

## Download

Use **Download Scheduling Export Package**.

The ZIP contains Workday-ready enrollment files plus `Scheduling_Results_and_Audit.xlsx` with the selected-session list and complete Requirement Audit. No approve/deny step is required.


## v2.3 startup simplification

`batch_utils.py` has been removed. All upload validation, export-package creation, and email-draft helpers now live directly in `app.py`, preventing mixed-version import errors.
