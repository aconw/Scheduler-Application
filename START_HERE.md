# Start Here — Deploy Today Without Installing Streamlit

You do **not** need Streamlit installed on your computer to use this build.

## 1. Create or use a GitHub account
Streamlit Community Cloud deploys from a GitHub repository. A private repository is recommended because this application contains internal training configuration.

## 2. Create a new private GitHub repository
A suggested name is `class-scheduling-application`.

## 3. Upload the application files
Unzip `Class_Scheduling_Application.zip` on your computer. In GitHub, upload **the contents of the `class_scheduler_app` folder** so that `app.py` is at the repository root.

The repository root should look like:

```
app.py
database.py
scheduler_engine.py
requirements.txt
README.md
START_HERE.md
assets/
data/
```

## 4. Connect Streamlit Community Cloud to GitHub
Sign in to Streamlit Community Cloud and connect your GitHub account/repository if you have not already done so.

## 5. Deploy
In Streamlit Community Cloud:

1. Select **Create app**.
2. Choose the GitHub repository you just created.
3. Choose branch `main`.
4. Set the entrypoint/main file to `app.py`.
5. Deploy.

Streamlit Cloud reads `requirements.txt` and installs Streamlit, pandas, and openpyxl automatically.

## 6. Run your first scheduling batch
Open **New Scheduling Run** and upload:

- New Hire Orientation Report and/or New/Additional Job Change Orientation Report
- Training History / Learning Transcript
- Learning Content / Available Sessions
- Existing Orientation Schedule Report is optional during the initial scheduling run

Then select **Run Scheduling**.

## 7. Recommended workflow

1. Review **Scheduling Results**.
2. Resolve true exceptions in **Exception Review**. Any manual override requires a reason and is audited.
3. Download Workday files in **Workday Export**.
4. Complete/upload enrollment in Workday.
5. Upload the resulting schedule in **Verification & Approval**.
6. Approve or deny each employee/staffing-event schedule.
7. Download hiring-manager and manual-scheduling email drafts from **Message Center**.

## 8. Configure these after launch

### Equivalencies
Open **Equivalencies** and add legacy/replacement course titles that satisfy current training requirements.

### FLAG_MANUAL recipients
Open **Manual Routing** and enter the recipient email for each manual course.

### Training rules
Open **Training Configurator**. The current training documentation is already seeded into the app.

## Important: configuration backup on free Streamlit Cloud
The free Community Cloud runtime does not guarantee that files written by a running app will persist forever across restarts/redeployments.

After changing training rules, equivalencies, or manual routing, open **Audit & Backup** and download **Configuration Backup**. The app also supports restoring this workbook.

For a future production deployment, the same application should be connected to a persistent hosted database.

## Current new-hire mapping assumption
The supplied New Hire report does not contain a field literally named `Job Code`. This build uses `Candidate Position ID` for new-hire rule matching. If Workday can add true Job Code to the report, that mapping should be changed before long-term production use.
