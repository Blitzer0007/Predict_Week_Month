📌 Predict_Week_Month

A Node.js-based numerical prediction system that processes weekly and monthly 3-digit historical data from Excel, generates frequency tables, and outputs predictions for upcoming dates (including full 1-year predictions).

This repository contains two major scripts:

separate_predict_year.js
→ Generates a 1-year prediction for weekly and monthly data separately.

predict_with_freqs_exceljs_sep_year.js
→ Generates full probability tables + predictions for custom dates.

Both scripts automatically create their own output folder and save results inside it.

📦 Prerequisites

Before running anything, install:

1. Node.js

Version 16+ recommended
Download: https://nodejs.org/

Check install:

node -v
npm -v

2. Install project dependencies

Run inside the project root:

npm install


This installs only the safe packages:

exceljs

dayjs

⚠️ xlsx is not used (security advisories).
All Excel operations use exceljs instead.

📂 Folder Structure
Predict_Week_Month/
│
├── Prediction_Scripts/
│   ├── separate_predict_year.js
│   ├── predict_with_freqs_exceljs_sep_year.js
│
├── Test_Data/
│   ├── weekly.xlsx
│   ├── monthly.xlsx
│
├── README.md
└── package.json

🚀 Running the Scripts
✅ 1. Full 1-Year Prediction (Weekly & Monthly Separate)

Use:

node Prediction_Scripts/separate_predict_year.js Test_Data/weekly.xlsx Test_Data/monthly.xlsx

📌 Output location:

A folder is created automatically:

/separate_predict_year/


Inside it you will find:

Weekly output

weekly_year_predictions.xlsx

separate_predict_year_weekly_predictions.csv

(Use these for weekly prediction analysis)

Monthly output

monthly_year_predictions.xlsx

separate_predict_year_monthly_predictions.csv

(Use these for monthly prediction analysis)

✅ 2. Prediction for Custom Dates + Full Frequency Tables

Use:

node Prediction_Scripts/predict_with_freqs_exceljs_sep_year.js \
     Test_Data/weekly.xlsx \
     Test_Data/monthly.xlsx \
     output.xlsx \
     2025-11-16,2025-11-17,2025-11-18

📌 Output location:

A folder is created:

/predict_with_freqs_exceljs_sep_year/


You will get:

output.xlsx → contains:

positional_overall

triplets_overall

positional_weekday

triplets_weekday

positional_month

triplets_month

predictions (for given dates)

predict_with_freqs_exceljs_sep_year_weekly_predictions.csv

predict_with_freqs_exceljs_sep_year_monthly_predictions.csv

📈 Which File Should You Use for Weekly Predictions?
✔ Recommended File:
separate_predict_year/weekly_year_predictions.xlsx

For a given date:

Check the predictions sheet inside the workbook.

Key columns:
Column	Meaning
Date	Predicted date
Weekday	0=Sun … 6=Sat
ObservationsUsed	Count of historical draws used
TopCandidates	Ranked 3-digit predictions

You can also refer to:

triplets_weekday (exact historic weekday triplets)

positional_weekday (digit probability by position)

If you prefer CSV:
separate_predict_year/separate_predict_year_weekly_predictions.csv

📈 Which File Should You Use for Monthly Predictions?
Use:
separate_predict_year/monthly_year_predictions.xlsx


This uses:

month-day historical priority

month-only historical fallback

overall fallback

Sheets include:

Sheet	Purpose
predictions	Best candidates for each date
positional_month	digit probability by month
triplets_month	month-based frequency
triplets_month_day	exact date frequency
📝 Development Notes
✔ node_modules should NOT be committed

Ensure .gitignore includes:

/node_modules

✔ All output folders are auto-created

Each script self-creates a folder named after the script.

✔ xxx values are automatically ignored

The scripts do not treat xxx as numeric values.