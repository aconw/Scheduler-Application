# Start Here — v2.5

## Replace the current Streamlit build

1. Unzip `Class_Scheduling_Application_Simplified_v2_5.zip`.
2. Replace the application files in your GitHub repository with the contents of the `simple_scheduler_v2_5` folder.
3. Commit the changes to the `main` branch.
4. Let Streamlit redeploy, then reboot the app once if needed.
5. Confirm the subtitle says:

**Simplified stateless build v2.5 • Duplicate source-event collapse • Person-level deduplication • Non-overlapping scheduling**

## Expected repository root

```
app.py
scheduler_engine.py
requirements.txt
README.md
START_HERE.md
RELEASE_NOTES_v2_5.md
assets/
config/
```

## What changed

v2.5 fixes duplicate source staffing rows. If Workday provides the same New Hire or Job Change event more than once for the same person, the app schedules only the first/canonical event. The duplicate row is preserved in the exported `Duplicate Source Events` audit sheet.

This is earlier and stronger than the v2.4 person/course export protection because a duplicated staffing row can no longer build its own prerequisite chain in the first place.
