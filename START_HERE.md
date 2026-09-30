# Start Here — v2.9.3 Strict Cleanup

v2.9.3 is a controlled cleanup of v2.9.2. Replace the application files as a complete set.

## Repository root
```
app.py
scheduler_engine.py
requirements.txt
README.md
START_HERE.md
RELEASE_NOTES_v2_9_3.md
assets/
config/
```

## Only v2.9.3 cleanup changes
- Restored Workday export duplicate protection for Person + Session WID in addition to Person + Course.
- Restored richer workbook preflight diagnostics.
- No change was made to the v2.9.2 nearest-location selection algorithm; it was verified by regression testing.

Confirm the app subtitle says **v2.9.3** after Streamlit redeploys.
