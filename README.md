# Predict_Week_Month

Scripts to analyze weekly and monthly 3-digit draws, build frequency tables, run backtests, and produce predictions (including 1-year forecasts).  
Each Node.js script reads Excel input(s) and writes Excel/CSV outputs into a folder named after the JS file being executed.

# Recommended (reproducible)
npm ci

The scripts rely on:

exceljs — read/write .xlsx files

dayjs — date handling

These are included in package.json and will be installed by npm ci / npm install.

Repository layout (important files)

Prediction_Scripts/ — main Node.js scripts (examples below):

backtest.js (backtesting script)

predict_with_freqs_exceljs.js

predict_with_freqs_exceljs_sep_year.js

separate_predict_year.js

... (other helper / experimental scripts)

Test_Data/ — example Excel input files (weekly & monthly).

package.json, package-lock.json

.gitignore — should include /node_modules

Input Excel formats expected

Scripts try to parse common structures automatically. Typical acceptable formats:

Flattened table with columns like Date and Number (one draw per row).

Weekly grid: header row contains weekday names (SUN,MON,...), grid cells contain 3-digit draws.

Monthly grid: a DATE column (day-of-month) and month columns (JAN,FEB,...) containing 3-digit draws.

Scripts ignore blank cells and non-3-digit placeholders (e.g., XXX).

How outputs are organized

When you run a script, an output folder is created automatically whose name equals the JS filename (script basename). Example:

node Prediction_Scripts/separate_predict_year.js Test_Data/weekly.xlsx Test_Data/monthly.xlsx

# Creates:
./separate_predict_year/
  weekly_year_predictions.xlsx
  monthly_year_predictions.xlsx


Workbooks commonly include sheets:

positional_overall — digit counts/probabilities per position

positional_weekday / positional_month

triplets_weekday / triplets_month / triplets_month_day

predictions — predicted candidates per requested date / next 365 days

Backtest scripts output CSV files (per-case results) and are also placed under a script-named folder.

Common commands / examples

From repository root, replace paths as needed:

1) Run the separate-year predictions (weekly & monthly separate)
node Prediction_Scripts/separate_predict_year.js Test_Data/weekly.xlsx Test_Data/monthly.xlsx
# Outputs in ./separate_predict_year/

2) Predict for specific dates and produce a workbook with freq sheets
node Prediction_Scripts/predict_with_freqs_exceljs.js \
  Test_Data/weekly.xlsx Test_Data/monthly.xlsx \
  output.xlsx "2025-11-16,2025-11-17"
# If script is named predict_with_freqs_exceljs.js this will create ./predict_with_freqs_exceljs/output.xlsx

3) Produce one-year (365-day) predictions separately for weekly/monthly
node Prediction_Scripts/predict_with_freqs_exceljs_sep_year.js \
  Test_Data/weekly.xlsx Test_Data/monthly.xlsx
# Outputs:
# ./predict_with_freqs_exceljs_sep_year/weekly_year_predictions.xlsx
# ./predict_with_freqs_exceljs_sep_year/monthly_year_predictions.xlsx

4) Backtest (example)
node Prediction_Scripts/backtest_lottery.js Test_Data/weekly.xlsx Test_Data/monthly.xlsx
# Outputs CSV(s) in ./backtest_lottery/

Configuration & tuning

Many scripts expose constants at the top for:

smoothing (Laplace / pseudo-counts),

mixture weights between triplet-frequency and positional-product,

how many candidates to list (top-N),

prediction date-range (next 365 days is default for year predictions).

Edit those values in the script if you want alternate behavior.

Recommended workflow

Prepare/clean your Excel inputs and put them in Test_Data/ (or point to correct paths).

npm ci

Run the script you need (see commands above).

Open generated workbook(s) in the output folder (script basename).

Git / CI best-practices

Do not commit node_modules/. Add /node_modules to .gitignore.

Commit package.json and package-lock.json (for deterministic installs).

Example .gitignore entries:

/node_modules
/.env
.DS_Store




