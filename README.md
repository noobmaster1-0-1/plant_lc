# Plant Maintenance Survey (MVP)

A mobile-first Android data-entry app for surveying plant → floor → machine → components.
Data is stored locally in SQLite. Excel is generated only when you tap **Export all data to Excel**.

## Current MVP features
- Add/select plants and floors.
- Record machines, including optional machine number, type, manufacturer, model, serial number and installation year.
- Record multiple components per machine.
- Capture required, installed and spare-stock quantities.
- Capture expected lifetime in days/months/years.
- Capture approximate installation month and year (or Unknown).
- Estimate days in service, remaining life and replacement date using the 15th of the selected month as an explicitly approximate reference.
- Export one workbook containing Summary, All Components, Procurement Summary and Machine Register.

## Important calculation notes
- Approximate installation date is represented internally by the 15th of the selected month. It is an estimate, not an exact installation date.
- If month/year is Unknown, service age and remaining life are left blank.
- Lifetime months and years use average conversion factors (30.4375 and 365.25 days).
- Procurement Summary is an indicative aggregation. Stock should be checked for shared/central stores before using it for an LC decision.
- The app does not submit LC requests or make purchasing decisions.

## Run on a computer
Python 3.10+ recommended.

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## Build Android APK
The included GitHub Actions workflow builds a debug APK using Buildozer on Ubuntu.
1. Create a GitHub repository and upload this project.
2. Push to the `main` branch (or run the workflow manually).
3. Open the repository's **Actions** tab.
4. Download the `plant-survey-apk` artifact from the completed workflow.

For first installation, Android may ask you to allow installation from that source.
This is a development/debug APK, not a Play Store release.

## Before factory deployment
Test the APK on your phone with sample data, verify backup/export, and confirm the workbook calculations with the maintenance/procurement team.
