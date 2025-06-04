Google Ads E-commerce Dashboard Script V1.0 💻📊🛍️
Author: Hassan El-Sisi

👋 Welcome! What's This All About?
Hey there, Google Ads aficionado! Ever wished for a powerful tool to automatically generate a comprehensive e-commerce dashboard right within your Google Sheets? One that slices and dices your Google Ads data, highlighting key trends, product performance, and potential areas for optimization? You've found it!

This script is your dedicated assistant for creating an in-depth e-commerce performance dashboard. It connects to your Google Ads account, fetches a wealth of data using GAQL, and then neatly organizes it into a multi-tab Google Sheet, complete with KPIs, data tables, and even charts on the overview tab.

Who is this for?

E-commerce Store Owners/Managers directly managing their Google Ads.

Digital Marketing Agencies looking to provide detailed and automated e-commerce reports to clients.

PPC Specialists focusing on e-commerce campaigns who want a quick and structured overview of performance.

Anyone who wants a clearer, data-backed understanding of their Google Ads e-commerce results.

Why is it awesome?

✅ Saves You Immense Time: Automates the tedious process of data extraction and report building.

🎯 E-commerce Focused Insights: Provides specific metrics crucial for e-commerce success like ROAS, AOV, and product-level performance.

⚙️ Highly Configurable: Tailor the reporting period and enable/disable specific analytical modules to fit your needs.

📊 Data-Driven Overview: Get a clear view of your campaign, product, and overall account performance.

📄 Organized & Comprehensive Reporting: All data is presented in a structured, multi-tab Google Sheet.

✨ Features at a Glance ✨
This script generates a detailed dashboard with various analytical sections:

📄 User Guide Tab: An in-sheet guide explaining the dashboard tabs and basic configuration.

⚙️ Settings & Config Tab: Displays key script settings and account information used for the report generation.

📊 Overview Dashboard:

High-level account performance KPIs (Cost, Impressions, Clicks, Conversions, Conv. Value, ROAS, CPA, CVR, AOV).

Top 5 Campaigns by Cost.

Metrics by Device (Cost, Clicks, Conversions, etc.).

Performance by Campaign Type.

Charts: Daily Performance Trends (Cost, Conversions, ROAS), Conversions by Device (Pie Chart), Top Products by various metrics (Bar Charts), Daily Conversion Value vs. Spend.

📈 Campaign Performance Tab: Detailed metrics for all active and enabled campaigns.

🛍️ Product Performance Analysis Tab: Aggregated data for individual products, including e-commerce specific KPIs and rule-based recommendations.

🖼️ Asset Performance Tab (Optional): Performance data for individual assets (e.g., headlines, descriptions, images). Default in script: Disabled.

📉 Impression Share & Rank Tab (Optional): Campaign-level Search Impression Share, Top IS, Absolute Top IS, and IS lost due to budget or rank. Default in script: Enabled.

💰 Budget Pacing Tab (Optional): Monitors campaign daily budgets against Month-To-Date (MTD) spend and flags pacing issues. Default in script: Enabled.

⏳ Lag-Adjusted Forecast Tab (Optional): Calculates historical conversion lag factors and applies them to recent performance to provide lag-adjusted conversion data and a short-term forecast. Includes a hidden zz - Conv. Lag Factors tab. Default in script: Enabled.

💡 Recommendations Tab: Automated, rule-based suggestions based on performance observed in other analytical tabs.

📚 Shared Library Export (Optional): Exports data from shared libraries, currently focused on Negative Keyword Lists (name, ID, member count). Default in script: Enabled.

📝 Raw Data Tabs (Hidden): Multiple hidden sheets prefixed with Raw Data -  store the direct output from Google Ads queries, serving as the foundation for the visible dashboard tabs.

❗ ERRORS Tab: Logs any errors encountered during script execution.

⚙️ How It Works (The "Magic" Explained Simply)
Connects to Google Ads: Securely accesses your Google Ads account data using Google Ads Query Language (GAQL). It only reads data.

Pulls the Data: Executes a series of pre-defined GAQL queries to gather performance metrics for campaigns, products, assets, daily trends, etc., based on the date range and features you've configured in the script.

Processes & Aggregates: The script then processes this raw data, calculates derived metrics (like ROAS, CPA, CTR, AOV), aggregates data where needed (e.g., for product summaries or daily totals), and prepares it for the dashboard.

Builds Your Dashboard: Finally, it creates (or clears and re-populates) a series of tabs in your specified Google Sheet. It writes the processed data, applies formatting for readability (including currency and percentages based on your account's settings), and generates charts on the Overview tab.

🚀 Getting Started: Your Step-by-Step Guide
Ready to get your automated e-commerce dashboard? Follow these steps:

Prerequisites:
✅ Access to a Google Ads account.

📄 A Google Sheet (create a new one, or use an existing one – the script will add/overwrite many tabs).

Step 1: Copy the Script Code
Grab all the code from the .js file containing this script. Select everything (Ctrl+A or Cmd+A) and copy it (Ctrl+C or Cmd+C).

Step 2: Open Google Ads Scripts
Log in to your Google Ads account.

Navigate to Tools & Settings (wrench icon 🔧) > BULK ACTIONS > Scripts.

Step 3: Paste and Name Your Script
Click the blue + button to create a new script.

Delete any placeholder code in the editor.

Paste the entire script code you copied.

Name your script (e.g., "E-commerce Dashboard Generator V1.0").

Step 4: Configure the Script (Crucial! 🛠️)
Scroll to the --- Section 1: SCRIPT_CONFIGURATION --- at the top of the script.

SPREADSHEET_ID:

Find the line: const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; (In the original script it's YOUR_SPREADSHEET_ID_HERE, change this).

Replace 'YOUR_SPREADSHEET_ID_HERE' with the actual ID of your Google Sheet. The ID is the long string of characters in the sheet's URL between /d/ and /edit.
Example: If URL is https://docs.google.com/spreadsheets/d/ABC123XYZ789/edit, the ID is ABC123XYZ789.

TARGET_CUSTOMER_ID:

const TARGET_CUSTOMER_ID = '';

If running from an MCC to target a specific client account, enter the client's Customer ID here (e.g., '123-456-7890'). Leave empty ('') if running directly in the client account.

REPORT_START_DATE and REPORT_END_DATE:

const REPORT_START_DATE = '2024-01-01';

const REPORT_END_DATE = '2025-12-31';

Set your desired reporting period in YYYY-MM-DD format.

Feature Toggles (ENABLE_..._TAB constants):

Review constants like ENABLE_IMPRESSION_SHARE_TAB, ENABLE_BUDGET_PACING_TAB, etc.

Set them to true to include the feature and its corresponding tab(s), or false to exclude them.

DEBUG_MODE (Optional):

const DEBUG_MODE = false;

Set to true for more detailed console logs if troubleshooting.

Step 5: Authorize the Script
Click Save (💾 icon).

Click Run (▶️ icon).

A pop-up "Authorization required" will appear. Click Review permissions.

Choose your Google account.

If you see "Google hasn’t verified this app," click Advanced, then "Go to [Your Script Name] (unsafe)".

Review permissions and click Allow.

Step 6: Run the Script! 🏃💨
Click Run (▶️ icon) again if needed.

The script may take several minutes. Check the Logs section below the code editor for progress.

Step 7: Check Your Google Sheet! 🎉
Once the logs indicate completion (e.g., "Script ... completed"), open your Google Sheet.

You'll find your new e-commerce dashboard with multiple tabs!

📊 Understanding the Report Tabs
The script generates several tabs in your Google Sheet. Here's a general idea:

⚪ 01 - User Guide: Instructions and overview.

⚪ 02 - Settings & Config: Summary of script settings and account info.

📊 03 - Overview Dashboard: High-level KPIs and charts. Start your analysis here!

📈 04 - Campaign Performance: Detailed campaign data.

🛍️ 05 - Product Performance Analysis: Product-specific metrics and recommendations.

🖼️ 06 - Asset Performance: (If enabled) Asset-level data.

📉 07 - Impression Share: (If enabled) IS metrics.

💰 08 - Budget Pacing: (If enabled) Campaign budget pacing status.

⏳ 09 - Lag-Adj. Forecast: (If enabled) Conversion data adjusted for lag.

💡 10 - Recommendations: Automated suggestions.

🕵️ zz - Conv. Lag Factors (Hidden): (If enabled) Data used for lag adjustments.

❗ ERRORS: Logs any script errors. Check this if something seems wrong.

👻 Raw Data - ... (Hidden Sheets): The foundational data pulled from Google Ads.

⚠️ Important Notes & Disclaimer
Execution Time: For large accounts or long date ranges, the script might approach Google Ads Scripts' 30-minute execution limit. If it times out, try a shorter date range.

Read-Only for Ads Account: The script DOES NOT MAKE ANY CHANGES to your Google Ads campaigns, ad groups, ads, or settings. It only reads data.

Data Accuracy: Data is pulled from the Google Ads API. Minor discrepancies with the UI can occasionally occur due to reporting lags or attribution differences.

Review Recommendations Critically: The automated recommendations are based on the script's logic and the thresholds defined. Always apply your expertise and judgment before making strategic decisions or changes to your account.

Permissions: Ensure you've granted the necessary permissions for the script to access Google Ads and Google Sheets.

🤔 Troubleshooting Common Issues
"Script timed out": Reduce the date range in SCRIPT_CONFIGURATION.

"#ERROR!" or "#VALUE!" in cells:

Check the "ERRORS" tab in the spreadsheet.

Verify SPREADSHEET_ID is correct.

Ensure all date formats and other settings in SCRIPT_CONFIGURATION are valid.

"No data" on a tab:

There might genuinely be no data for that report in the selected period.

Double-check your date range and feature toggle settings.

Happy Dashboarding! May your e-commerce insights lead to great success!