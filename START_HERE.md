# Start Here — v2.8

## Replace the current Streamlit build

1. Unzip `Class_Scheduling_Application_Simplified_v2_6.zip`.
2. Replace the application files in your GitHub repository with the contents of the `simple_scheduler_v2_6` folder.
3. Commit the changes to the `main` branch.
4. Let Streamlit redeploy, then reboot the app once if needed.
5. Confirm the subtitle says:

**Simplified stateless build v2.8 • Directional equivalencies for completions + active enrollments • Duplicate source-event collapse • Non-overlapping scheduling**

## Expected repository root

```
app.py
scheduler_engine.py
requirements.txt
README.md
START_HERE.md
RELEASE_NOTES_v2_6.md
assets/
config/
```

## Equivalencies in v2.8

Open `config/Scheduler_Configuration.xlsx` and use the **Equivalencies** worksheet.

Columns:

- `required_training_title`
- `equivalent_training_title`
- `relationship_direction`
- `active`

`ONE_WAY` means the equivalent title satisfies the required title only.

`TWO_WAY` means either title satisfies the other.

The Cardiac Monitoring ↔ Cardiac Monitoring Blended Learning Session 1 relationship requested for this release is already included as `TWO_WAY`.

These mappings are checked against both completed training and active/current enrollments before the application selects a new session.


### v2.8
The results/audit now includes an Existing Workday Training view showing all active enrolled sessions, whether or not they satisfy a current-role requirement. Those sessions are also treated as blocked time for scheduling.


## v2.8 scheduling order

TARGET_RANGE requirements are scheduled before FIRST_AVAILABLE requirements, except when prerequisite dependencies require a prerequisite to be scheduled first. For FIRST_AVAILABLE requirements that compete for the same start date/time, lower numeric Priority wins (1 is higher than 2; blank is lowest). Available Sessions rows sharing the same WID are treated as one offering; all days of a multi-day offering participate in overlap checks and one seat is consumed.
