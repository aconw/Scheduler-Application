# Start Here — Simplified v2.8.2

## Replace the current GitHub build
Replace the current application repository contents with the contents of this folder.

Repository root:

```
app.py
scheduler_engine.py
requirements.txt
README.md
START_HERE.md
RELEASE_NOTES_v2_8.md
assets/
config/
```

The packaged configuration workbook is the current workbook supplied for this v2.8.2 change and includes the Training Rules `priority` column.

After committing to `main`, let Streamlit redeploy and confirm the subtitle says:

**Simplified stateless build v2.8.2 • TARGET_RANGE first • FIRST_AVAILABLE priority ranking • Multi-day session aware • Non-overlapping scheduling**

## v2.8.2 selection behavior

- Hard prerequisites are processed before dependent classes.
- TARGET_RANGE requirements are processed before FIRST_AVAILABLE requirements within a staffing event.
- FIRST_AVAILABLE requirements are considered in potential selection-date order; when their potential selection time ties, lower numeric `priority` wins. Blank priority is lowest.
- Physical-location selection rules remain unchanged.

## v2.8.2 multi-day behavior

Each WID in the Available Sessions report represents one class offering. Multiple rows with the same WID are treated as days of that one offering.

- One seat is reserved for the whole WID.
- Every day's start/end interval blocks the employee's calendar.
- Selected results include the number of days.
- Back-to-back sessions are allowed when one interval ends exactly when another begins.

## Run inputs

Upload:

1. New Hire and/or Job Change report
2. Training History
3. Learning Content / Available Sessions
4. Existing Orientation Schedule

All required preflight checks must pass before Run Scheduling is enabled.


## v2.8.2 check
After deployment, confirm the app subtitle begins `Simplified stateless build v2.8.2`. Results now include a Training Documentation Status section, and the exported audit workbook includes a Training Documentation Status sheet.
