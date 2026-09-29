# v2.9.1 deployment

Replace the current GitHub application with the contents of this folder.

Repository root:

```text
app.py
scheduler_engine.py
requirements.txt
README.md
START_HERE.md
config/Scheduler_Configuration.xlsx
assets/...
```

Commit the changes to `main`, let Streamlit redeploy, and confirm the app subtitle starts with `v2.9.1`.

Use the current `Scheduler_Configuration.xlsx` included here. It contains the Training Rules Priority column and your current Equivalencies configuration.

Each run requires:

- New Hire and/or New / Additional Job Change report
- Training History / Learning Transcript
- Learning Content / Available Sessions
- Orientation Schedule — Lesson-Level Detail

The results package contains the Workday files and the Requirement Audit. There is no approval/denial step.


## v2.9.1
This release is a compatibility fix for the report preflight validator. The header arguments are passed explicitly so Streamlit cannot confuse them with worksheet names.
