# Google Ads E-commerce Dashboard Script V1.0 🚀📊🛍️

**Author:** [Hassan El-Sisi](https://www.linkedin.com/in/hassan-elsisi/)

## 👋 Welcome! What's This All About?

Hey there, Google Ads aficionado! Ever wished for a powerful tool to automatically generate a comprehensive e-commerce dashboard right within your Google Sheets? One that slices and dices your Google Ads data, highlighting key trends, product performance, and potential areas for optimization? You've found it!

This script is your dedicated assistant for creating an in-depth e-commerce performance dashboard. It connects to your Google Ads account, fetches a wealth of data using GAQL, and then neatly organizes it into a multi-tab Google Sheet, complete with KPIs, data tables, and even charts on the overview tab.

**Who is this for?**

* 🛍️ **E-commerce Store Owners/Managers** directly managing their Google Ads.
* 🏢 **Digital Marketing Agencies** looking to provide detailed and automated e-commerce reports to clients.
* 🧑‍💻 **PPC Specialists** focusing on e-commerce campaigns who want a quick and structured overview of performance.
* 🤔 Anyone who wants a clearer, data-backed understanding of their Google Ads e-commerce results.

**Why is it awesome?**

* ✅ **Saves You Immense Time:** Automates the tedious process of data extraction and report building.
* 🎯 **E-commerce Focused Insights:** Provides specific metrics crucial for e-commerce success like ROAS, AOV, and product-level performance.
* ⚙️ **Highly Configurable:** Tailor the reporting period and enable/disable specific analytical modules to fit your needs.
* 📊 **Data-Driven Overview:** Get a clear view of your campaign, product, and overall account performance.
* 📄 **Organized & Comprehensive Reporting:** All data is presented in a structured, multi-tab Google Sheet.

## ✨ Features at a Glance ✨

This script generates a detailed dashboard with various analytical sections:

* 📄 **User Guide Tab**: (This file, adapted for in-sheet viewing!) Your guide to the dashboard.
* ⚙️ **Settings & Config Tab**: Displays key script settings (like date range) and account information used for the report generation. (Informational)
* ❗ **ERRORS Tab**: If anything goes wrong during the script run, details will be logged here. (Hopefully, it stays empty!)
* 📊 **Overview Dashboard**: A high-level snapshot of key e-commerce KPIs, trends, and charts. Your starting point! (Often themed with primary/summary colors)
    * High-level account performance KPIs (Cost, Impressions, Clicks, Conversions, Conv. Value, ROAS, CPA, CVR, AOV).
    * Top 5 Campaigns by Cost.
    * Metrics by Device (Cost, Clicks, Conversions, etc.).
    * Performance by Campaign Type.
    * **Charts**: Daily Performance Trends (Cost, Conversions, ROAS), Conversions by Device (Pie Chart), Top Products by various metrics (Bar Charts), Daily Conversion Value vs. Spend.
* 📈 **Campaign Performance Tab**: Detailed metrics for all active and enabled campaigns. (Often themed with performance/growth colors)
* 🛍️ **Product Performance Analysis Tab**: Aggregated data for individual products, including e-commerce specific KPIs and rule-based recommendations. (Often themed with e-commerce/product colors)
* 🖼️ **Asset Performance Tab** (Optional): Performance data for individual assets (e.g., headlines, descriptions, images). *Default in script: Disabled*.
* 📉 **Impression Share & Rank Tab** (Optional): Campaign-level Search Impression Share, Top IS, Absolute Top IS, and IS lost due to budget or rank. *Default in script: Enabled*. (Often themed with competitive metric colors)
* 💰 **Budget Pacing Tab** (Optional): Monitors campaign daily budgets against Month-To-Date (MTD) spend and flags pacing issues. *Default in script: Enabled*. (Often themed with financial/budget colors)
* ⏳ **Lag-Adjusted Forecast Tab** (Optional): Calculates historical conversion lag factors and applies them to recent performance to provide lag-adjusted conversion data and a short-term forecast. Includes a hidden `zz - Conv. Lag Factors` tab. *Default in script: Enabled*.
* 💡 **Recommendations Tab**: Automated, rule-based suggestions based on performance observed in other analytical tabs. (Often themed with insights/action colors)
* 📚 **Shared Library Export** (Optional): Exports data from shared libraries, currently focused on Negative Keyword Lists (name, ID, member count). *Default in script: Enabled*.
* 📝 **Raw Data Tabs (Hidden)**: Multiple hidden sheets prefixed with `Raw Data - ` store the direct output from Google Ads queries, serving as the foundation for the visible dashboard tabs.

## ⚙️ How It Works (The "Magic" Explained Simply)

Think of this script as a very smart, very fast data analyst for your e-commerce ads. Here's the basic idea:

1.  🔗 **Connects to Google Ads:** It uses Google's own tools (specifically, Google Ads Query Language or GAQL) to securely access your account data (don't worry, it only *reads* data, it doesn't change anything in your account!).
2.  📥 **Pulls the Data:** It runs a series of pre-defined queries to gather performance metrics for campaigns, products, assets, daily trends, etc., based on the date range and features you've configured in the script.
3.  🔢 **Processes & Aggregates:** The script then processes this raw data, calculates derived metrics (like ROAS, CPA, CTR, AOV), aggregates data where needed (e.g., for product summaries or daily totals), and prepares it for the dashboard.
4.  📝 **Builds Your Dashboard:** Finally, it creates (or clears and re-populates) a series of tabs in your specified Google Sheet. It writes the processed data, applies formatting for readability (including currency and percentages based on your account's settings), and generates charts on the Overview tab.

## 🚀 Getting Started: Your Step-by-Step Guide

Ready to get your automated e-commerce dashboard? Follow these steps carefully.

### Prerequisites:

* ✅ Access to a Google Ads account.
* 📄 A Google Sheet (create a new one, or use an existing one – the script will add/overwrite many tabs).

### Step 1: Copy the Script Code

* 📋 Grab all the code from the `.js` file containing this script. Select everything (Ctrl+A or Cmd+A) and copy it (Ctrl+C or Cmd+C).

### Step 2: Open Google Ads Scripts

* ➡️ Log in to your Google Ads account.
* ⚙️ In the top menu, click on **Tools & Settings** (it looks like a wrench 🔧).
* 🖱️ Under "BULK ACTIONS," click on **Scripts**.

### Step 3: Paste and Name Your Script

* ➕ Click the big blue **+** button to create a new script.
* 🗑️ You'll see a script editor. Delete any existing code in there (usually a simple `function main() {}`).
* ✍️ Paste the entire code you copied in Step 1 into the script editor (Ctrl+V or Cmd+V).
* 🏷️ At the top, where it says "Untitled script," give your script a memorable name, like "My E-commerce Dashboard V1" or "GA Ecom Report".

### Step 4: Configure the Script (Super Important! 🛠️)

This is where you tell the script where to put the report and what settings to use.
*Scroll to the very top of the script code you just pasted.* You're looking for **`--- Section 1: SCRIPT_CONFIGURATION ---`**.

1.  **`SPREADSHEET_ID`**: This tells the script which Google Sheet to use.
    * 📄 Open the Google Sheet you want the report to be generated in (or create a new one).
    * 🔗 Look at the URL (the web address) in your browser. It will look something like this:
        `https://docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit#gid=0`
    * 🎯 You need to copy the long string of letters and numbers that's between `/d/` and `/edit`. That's your Spreadsheet ID!
    * ✏️ In the script, find this line:
        ```javascript
        const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
        ```
    * 🔄 Replace `'YOUR_SPREADSHEET_ID_HERE'` with *your* actual Spreadsheet ID, keeping the single quotes. (The original script had `13lu2NEXvZxfAZ-2P40Am5uF1scagRD3b0A86kAA-7N4` as an example, make sure to use your own).

2.  **`TARGET_CUSTOMER_ID`**: This is for MCC users.
    * Find this line:
        ```javascript
        const TARGET_CUSTOMER_ID = '';
        ```
    * 🏢 If running from an MCC to target a specific client account, enter the client's Customer ID here (e.g., `'123-456-7890'`). Leave empty (`''`) if running directly in the client account.

3.  **`REPORT_START_DATE` and `REPORT_END_DATE`**: This sets the date range for your dashboard.
    * Find these lines:
        ```javascript
        const REPORT_START_DATE = '2024-01-01';
        const REPORT_END_DATE = '2025-12-31';
        ```
    * 🗓️ Change the dates (inside the single quotes) to the start and end dates you want for your report. Use the `YYYY-MM-DD` format.

4.  **Feature Toggles (`ENABLE_..._TAB` constants)**: Customize which reports are generated.
    * 🔎 Still in `--- Section 1: SCRIPT_CONFIGURATION ---`, you'll see variables like `ENABLE_IMPRESSION_SHARE_TAB`, `ENABLE_BUDGET_PACING_TAB`, etc.
    * Read the comments next to each one to understand what it does.
    * Toggle these between `true` (to include the feature and its tab) or `false` (to exclude it).

5.  **`DEBUG_MODE`** (Optional): For more detailed logging.
    * Find:
        ```javascript
        const DEBUG_MODE = false;
        ```
    * 🐞 Set to `true` if you're troubleshooting and want more detailed messages in the Google Ads script logs.

### Step 5: Authorize the Script (Give it Permission)

Scripts need your permission to access your Ads data and write to your Google Sheet.

* 💾 Click the **Save** icon at the top of the script editor.
* ▶️ Now, click the **Run** button.
* 🔔 A pop-up titled "Authorization required" will appear. Click **Review permissions**.
* 👤 Choose the Google account that has access to the Google Ads account you want to audit.
* 🛡️ You might see a screen saying "Google hasn’t verified this app." This is normal for custom scripts.
    * Click on **Advanced** (it might be a small link).
    * Then click on **"Go to [Your Script Name] (unsafe)"**.
* 👍 Review the permissions the script needs (it will ask for access to your Google Ads data and your Google Spreadsheets). Click **Allow**.

### Step 6: Run the Script! 🏃💨

* ▶️ After authorizing, you might be taken back to the script editor. Click the **Run** button again.
* ⏳ **Be Patient!** This script is doing a lot of work. It might take several minutes to run, especially for larger accounts or long date ranges (Google Ads Scripts have a 30-minute time limit).
* 📋 You can see the script's progress in the **Logs** section below the code editor. It will print messages like "Processing raw data for: Campaigns", "Overview dashboard populated", etc.

### Step 7: Check Your Google Sheet! 🎉

* ✅ Once the logs say something like "Script ... completed ... Duration: ...", open the Google Sheet you specified in `SPREADSHEET_ID`.
* ✨ You should see a whole bunch of new tabs filled with your e-commerce dashboard data, ready for review!

## 📊 Understanding the Report Tabs

The script creates many tabs. Here's a quick guide to what you'll typically find (colors are illustrative of common report themes):

* ⚪ **`01 - User Guide`**: Instructions and overview. (Default/White)
* ⚪ **`02 - Settings & Config`**: Summary of script settings and account info. (Default/White)
* 📊 **`03 - Overview Dashboard`**: High-level KPIs and charts. **Start your analysis here!** (Often themed with primary/summary colors like Pastel Blue)
* 📈 **`04 - Campaign Performance`**: Detailed campaign data. (Often themed with performance/growth colors like Light Teal)
* 🛍️ **`05 - Product Performance Analysis`**: Product-specific metrics and recommendations. (Often themed with e-commerce/product colors like Light Purple)
* 🖼️ **`06 - Asset Performance`**: (If enabled) Asset-level data.
* 📉 **`07 - Impression Share`**: (If enabled) IS metrics.
* 💰 **`08 - Budget Pacing`**: (If enabled) Campaign budget pacing status.
* ⏳ **`09 - Lag-Adj. Forecast`**: (If enabled) Conversion data adjusted for lag.
* 💡 **`10 - Recommendations`**: Automated suggestions. (Often themed with insights/action colors like Light Green)
* 🕵️ **`zz - Conv. Lag Factors (Hidden)`**: (If enabled) Data used for lag adjustments.
* ❗ **`ERRORS`**: Logs any script errors. Check this if something seems wrong! (Often themed with alert/error colors like Red)
* 👻 **`Raw Data - ...` (Hidden Sheets)**: The foundational data pulled from Google Ads. (Pale Yellow or similar for data/utility sheets)

## ⚠️ Important Notes & Disclaimer

* ⏱️ **Execution Time:** For large accounts or long date ranges, the script might approach Google Ads Scripts' 30-minute execution limit. If it times out, try a shorter date range.
* 🔒 **Read-Only for Ads Account:** This script **DOES NOT MAKE ANY CHANGES** to your Google Ads campaigns, ad groups, ads, or settings. It only reads data.
* 📊 **Data Accuracy:** Data is pulled from the Google Ads API. Minor discrepancies with the UI can occasionally occur due to reporting lags or attribution differences.
* 🧠 **Review Recommendations Critically:** The automated recommendations are based on the script's logic and the thresholds defined. **Always use your own expertise and judgment** before making strategic decisions or changes to your account. This script is a tool to help you, not a replacement for a skilled account manager.
* 🔑 **Permissions:** Ensure you've granted the necessary permissions during the authorization step for the script to access Google Ads and Google Sheets.

## 🤔 Troubleshooting Common Issues

* **"Script timed out"**: Reduce the date range in `SCRIPT_CONFIGURATION`.
* **"\#ERROR!" or "\#VALUE!" in cells**:
    * Check the "ERRORS" tab in the spreadsheet for any messages.
    * Verify `SPREADSHEET_ID` is correct.
    * Ensure all date formats and other settings in `SCRIPT_CONFIGURATION` are valid.
* **"No data" on a tab**:
    * There might genuinely be no data for that report in the selected period (e.g., no Performance Max campaigns ran if the PMax-related tabs are empty).
    * Double-check your date range and feature toggle settings in `SCRIPT_CONFIGURATION`.

Happy Dashboarding! May your e-commerce insights lead to great success! 🌟
