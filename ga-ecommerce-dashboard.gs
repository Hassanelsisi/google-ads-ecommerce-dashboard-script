/****************************************************************************************
 * Google Ads E-commerce Dashboard Script
 * Version: 1.0 (Layout, Chart, Totals, Currency Fixes & Enhancements - 29-May-2025)
 * ─────────────────────────────────────────────────────────────────────────────────────
 * ▸ Generates a comprehensive e-commerce dashboard in Google Sheets using Google Ads Query Language (GAQL).
 * ▸ Features include: campaign performance, asset analysis, impression share & rank metrics,
 * budget pacing, conversion lag adjustments, product-level analytics, and shared library exports.
 * ▸ Reporting date range and feature processing are configurable below.
 * ▸ This script requires authorization for Google Ads and Google Sheets services.
 *
 * Script Structure (Renamed and Reorganized Sections):
 * 1. SCRIPT_CONFIGURATION:           All user-configurable settings, constants, and feature toggles.
 * 2. HELPER_UTILITIES:               Core helper functions for data types, math, date operations, sheet styling.
 * 3. LOGGING_FRAMEWORK:              Centralized functions for logging script activity and errors.
 * 4. SCRIPT_INITIALIZATION_AND_CHECKS: Initial script setup, runtime guards, and API/Sheet access checks.
 * 5. GOOGLE_ADS_DATA_EXTRACTION:     Functions dedicated to retrieving data from Google Ads via GAQL.
 * 6. RAW_DATA_SHEET_PROCESSING:      Management of "Raw Data" sheets (creation, clearing, writing data, formatting).
 * 7. ANALYTICAL_DASHBOARD_MODULES:   Functions for "Impression Share", "Budget Pacing", "Conversion Lag" tabs.
 * 8. CORE_DASHBOARD_CONSTRUCTION:    Functions to build main user-facing dashboard tabs and charts, including User Guide.
 * 9. MAIN_ORCHESTRATION_FLOW:        The main orchestrator function that controls the script's execution flow.
 ****************************************************************************************/

'use strict'; // Enforces stricter parsing and error handling in JavaScript.

// --- Section 1: SCRIPT_CONFIGURATION ---
// All user-configurable settings, constants, and feature toggles.

const SCRIPT_VERSION = '1.0'; // FIX: Updated version

/** ── Spreadsheet & Account Settings ─────────────────────────────────────── **/
const SPREADSHEET_ID = 'YOUR_SHEET_ID_HERE'; // Replace with your Google Sheet ID
const TARGET_CUSTOMER_ID = ''; // Leave empty if running in the client account, or set MCC client ID (e.g., '123-456-7890')

/** ── Reporting Date Range (Format:YYYY-MM-DD) ─────────────────────────── **/
const REPORT_START_DATE = '2024-01-01';
const REPORT_END_DATE = '2025-12-31';

/** ── Feature Toggles (for enabling/disabling entire features/tabs) ─────────── **/
const ENABLE_IMPRESSION_SHARE_TAB = true;
const ENABLE_BUDGET_PACING_TAB = true;
const ENABLE_CONV_LAG_ADJUST_TABS = true;
const ENABLE_LIBRARY_EXPORT = true;
const ENABLE_ASSET_PERFORMANCE_TAB = false;
const DEBUG_MODE = false; // If true, logs more detailed information.

/** ── Data Processing Toggles (controls if raw data for these features is fetched and processed) ── **/
const PROCESS_IMPRESSION_SHARE_DATA = ENABLE_IMPRESSION_SHARE_TAB;
const PROCESS_ASSET_PERFORMANCE_DATA = ENABLE_ASSET_PERFORMANCE_TAB;
const PROCESS_LIBRARY_EXPORT_DATA = ENABLE_LIBRARY_EXPORT;
const PROCESS_CONV_LAG_DATA = ENABLE_CONV_LAG_ADJUST_TABS;


/** ── Logging & Error Handling Settings ───────────────────────────────────── **/
const LOG_VERBOSE = true; // Enables detailed [VERBOSE] logs.
const ERROR_SHEET_NAME = 'ERRORS'; // Name of the sheet where script errors will be recorded. (Preserved during cleanup)

/** ── Sheet Naming Conventions ────────────────────────────────────────────── **/
const TEMP_SHEET_PREFIX = '_TMP_SCRIPT_RUNNING_';
const TAB_PREFIX_RAW = 'Raw Data - ';

// Renumbered Dashboard Tabs
const USER_GUIDE_TAB_NAME = '01 - User Guide';
const SETTINGS_CONFIG_TAB_NAME = '02 - Settings & Config';
const OVERVIEW_TAB_NAME = '03 - Overview Dashboard';
const CAMPAIGN_PERF_TAB_NAME = '04 - Campaign Performance';
const PRODUCT_ANALYSIS_TAB_NAME = '05 - Product Performance Analysis';
const ASSET_PERFORMANCE_TAB_NAME = '06 - Asset Performance';
const SHARE_RANK_TAB_NAME = '07 - Impression Share';
const BUDGET_PACING_TAB_NAME = '08 - Budget Pacing';
const LAG_FORECAST_TAB_NAME = '09 - Lag-Adj. Forecast';
const RECOMMENDATIONS_TAB_NAME = '10 - Recommendations';

// Hidden/Analytical Data Tab
const LAG_FACTORS_TAB_NAME = 'zz - Conv. Lag Factors (Hidden)';


/** ── Formatting & Runtime Limits ─────────────────────────────────────────── **/
const CHART_COLORS_ARRAY = ['#4285F4', '#DB4437', '#F4B400', '#0F9D58',
  '#AB47BC', '#00ACC1', '#FF7043', '#9E9D24', '#5E35B1', '#039BE5'
];
const ADOPTED_CHART_COLORS = {
  CTR: '#9900FF',
  IMPRESSIONS: '#FF00FF',
  CONV_RATE: '#00CCFF',
  CONV_VALUE_ANALYSIS: '#0000CC',
  PRODUCT_CONVERSIONS: '#9900FF',
  PRODUCT_CONV_VALUE: '#0000CC',
  DEVICE_CONVERSIONS_PIE: ['#4285F4', '#DB4437', '#F4B400', '#0F9D58', '#AB47BC'],
  DAILY_SPEND: '#DB4437',
  DAILY_CONV_VALUE: '#4285F4'
};

const GAQL_PAGE_SIZE = 10000;
const GAQL_FETCH_LIMIT = 25000;
const MIN_TIME_BUFFER_SEC = 60;
const SHEET_FORMAT_CHUNK_SIZE = 2000;

/** ── Thresholds & Business Logic (for Recommendations & Analysis) ────────── **/
const REC_TARGET_ROAS_THRESHOLD = 2.0;
const REC_CPA_FACTOR_THRESHOLD = 1.5;
const REC_DEFAULT_ACCOUNT_CPA = 50.0;
const REC_LOW_CTR_THRESHOLD = 0.01;
const REC_MIN_IMPR_FOR_CTR = 100;
const REC_MIN_CLICKS_NO_CONV = 20;

const BUDGET_PACE_LOW_THRESHOLD = 0.90;
const BUDGET_PACE_HIGH_THRESHOLD = 1.10;
const CONV_LAG_LOOKBACK_DAYS = 60;
const CONV_LAG_STABLE_PERIOD_OFFSET = 30;
const CONV_LAG_FORECAST_DAYS = 14;

/** ── UI Styles (JSON objects for easy re-use in styling sheets) ────────── **/
const STYLE_UNIFIED_HEADER = {
  fontFamily: 'Arial',
  fontSize: 10,
  fontWeight: 'bold',
  fontColor: '#FFFFFF',
  backgroundColor: '#4A86E8',
  horizontalAlignment: 'center',
  verticalAlignment: 'middle',
  wrapStrategy: SpreadsheetApp.WrapStrategy.WRAP,
  textRotation: 0
};
const STYLE_SECTION_HEADER = {
    fontFamily: 'Arial',
    fontSize: 14,
    fontWeight: 'bold',
    fontColor: '#212121',
    backgroundColor: '#EFEFEF',
    horizontalAlignment: 'left',
    verticalAlignment: 'middle',
    wrapStrategy: SpreadsheetApp.WrapStrategy.WRAP,
    textRotation: 0,
    padding: 5
};
const STYLE_TOTALS_ROW = {
  fontFamily: 'Arial',
  fontSize: 10,
  fontWeight: 'bold',
  fontColor: '#FFFFFF',
  backgroundColor: '#4A86E8', // Match UNIFIED_HEADER
  verticalAlignment: 'middle'
};
const STYLE_KPI_METRIC = {
  fontFamily: 'Arial',
  fontSize: 10,
  fontWeight: 'bold',
  fontColor: '#212121'
};
const STYLE_KPI_VALUE = {
  fontFamily: 'Arial',
  fontSize: 10,
  fontColor: '#212121'
};
const STYLE_CHART_TITLE = {
  color: '#000000',
  fontName: 'Arial',
  fontSize: 11,
  bold: false
};
const STYLE_CHART_AXIS_TITLE = {
  color: '#000000',
  fontName: 'Arial',
  fontSize: 10,
  bold: false
};
const STYLE_CHART_AXIS_LABEL = {
  color: '#000000',
  fontName: 'Arial',
  fontSize: 9,
  bold: false
};
const STYLE_CHART_LEGEND = {
  color: '#424242',
  fontName: 'Arial',
  fontSize: 10
};
const STYLE_TABLE_ROW_EVEN_BG = '#F9F9F9';
const STYLE_TABLE_ROW_ODD_BG = '#FFFFFF';

// Conditional Formatting Colors
const LIGHT_GREEN_BACKGROUND = '#D9EAD3';
const LIGHT_YELLOW_BACKGROUND = '#FFF2CC';
const LIGHT_RED_BACKGROUND = '#F4CCCC';

// --- End SCRIPT_CONFIGURATION ---

// --- Section 2: HELPER_UTILITIES ---

function safeMicrosToCurrency_(micros) {
  const num = Number(micros);
  return !isNaN(num) && micros !== null && typeof micros !== 'undefined' ? num / 1000000 : 0;
}

function safeParseFloat_(value) {
  if (value === null || typeof value === 'undefined' || String(value).trim() === '') return 0;
  const num = parseFloat(String(value).replace(/,/g, ''));
  return isNaN(num) ? 0 : num;
}

function safeParseInt_(value) {
  if (value === null || typeof value === 'undefined' || String(value).trim() === '') return 0;
  const num = parseInt(String(value).replace(/,/g, ''), 10);
  return isNaN(num) ? 0 : num;
}

function div(numerator, denominator) {
  const den = safeParseFloat_(denominator);
  return den !== 0 ? safeParseFloat_(numerator) / den : 0;
}

function pct(numerator, denominator, returnAsDecimal) {
    const den = safeParseFloat_(denominator);
    const result = den !== 0 ? (safeParseFloat_(numerator) / den) : 0;
    return returnAsDecimal ? result : result * 100;
}


function getCurrencyFormatString_(currencyCode) {
  const upperCode = String(currencyCode || '').toUpperCase();
  const robustFormatBase = `"${upperCode} "#,##0.00;"-${upperCode} "#,##0.00;"${upperCode} "0.00`;
  const formats = {
    'USD': '"USD "#,##0.00;"-USD "#,##0.00;"USD "0.00',
    'EUR': '"EUR "#,##0.00;"-EUR "#,##0.00;"EUR "0.00',
    'GBP': '"GBP "#,##0.00;"-GBP "#,##0.00;"GBP "0.00',
    'JPY': '"JPY "¥#,##0;"-JPY "¥#,##0;"JPY "0', // JPY typically has no decimals
    'CAD': '"CAD "#,##0.00;"-CAD "#,##0.00;"CAD "0.00',
    'AUD': '"AUD "#,##0.00;"-AUD "#,##0.00;"AUD "0.00',
    'INR': '"INR "₹#,##,##0.00;"-INR "₹#,##,##0.00;"INR "0.00',
    'EGP': '"EGP "#,##0.00;"-EGP "#,##0.00;"EGP "0.00',
    'SAR': '"SAR "#,##0.00;"-SAR "#,##0.00;"SAR "0.00',
    'AED': '"AED "#,##0.00;"-AED "#,##0.00;"AED "0.00'
  };
  return formats[upperCode] || robustFormatBase; // Fallback to robust format with the code
}

function applyChunkedFormatting_(sheet, formatsArray, startDataRow) {
  if (!sheet || !formatsArray || formatsArray.length === 0 || (formatsArray[0] && formatsArray[0].length === 0)) {
    logVerbose_(`Skipping chunked formatting for ${sheet ? sheet.getName() : 'unknown sheet'}: no formats or data provided.`);
    return;
  }
  const numDataRows = formatsArray.length;
  const numColumns = formatsArray[0].length;

  logVerbose_(`Applying chunked formatting to ${sheet.getName()}. Total data rows to format: ${numDataRows}`);
  for (let r = 0; r < numDataRows; r += SHEET_FORMAT_CHUNK_SIZE) {
    if (!checkRemainingTimeGraceful_()) {
      logError_('applyChunkedFormatting_', new Error('Low execution time remaining, aborting further formatting.'), `Sheet: ${sheet.getName()}, Row chunk starting at: ${r}`);
      return;
    }
    const chunkEndRowOffset = Math.min(r + SHEET_FORMAT_CHUNK_SIZE, numDataRows);
    const chunkNumRows = chunkEndRowOffset - r;
    const currentChunkFormats = formatsArray.slice(r, chunkEndRowOffset);

    try {
      sheet.getRange(startDataRow + r, 1, chunkNumRows, numColumns).setNumberFormats(currentChunkFormats);
      logVerbose_(`Formatted chunk for ${sheet.getName()}: Spreadsheet rows ${startDataRow + r} to ${startDataRow + r + chunkNumRows - 1}`);
    } catch (e) {
      logError_('applyChunkedFormatting_', e, `Error formatting chunk on sheet: ${sheet.getName()}, Spreadsheet rows: ${startDataRow + r} to ${startDataRow + r + chunkNumRows -1}`);
    }
  }
  logInfo_(`Finished applying chunked formatting for ${sheet.getName()}.`);
}


function applyChartOptionsRecursively_(builder, optionsObject, basePath) {
  basePath = basePath || '';
  for (const key in optionsObject) {
    if (optionsObject.hasOwnProperty(key)) {
      const fullPath = basePath ? `${basePath}.${key}` : key;
      const value = optionsObject[key];
      if (typeof value === 'object' && value !== null && !Array.isArray(value) && !(value instanceof Date)) {
        applyChartOptionsRecursively_(builder, value, fullPath);
      } else {
        try {
          builder.setOption(fullPath, value);
        } catch (e) {
          logError_('applyChartOptionsRecursively_', e, `Failed to set chart option: ${fullPath} = ${value}`);
        }
      }
    }
  }
}

function getTodayDateString_() {
  return Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
}

function addDaysToDate_(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function getDaysBetweenDates_(startDate, endDate) {
  const msPerDay = 24 * 60 * 60 * 1000;
  const start = (startDate instanceof Date) ? startDate : new Date(startDate);
  const end = (endDate instanceof Date) ? endDate : new Date(endDate);
  return Math.round((end.getTime() - start.getTime()) / msPerDay);
}

function autoResizeSheetColumns_(sheet, numColumns) {
  if (!sheet) return;
  const colsToResize = numColumns || sheet.getLastColumn();
  if (colsToResize > 0 && sheet.getLastRow() > 0) {
    for (let c = 1; c <= colsToResize; c++) {
      try {
        sheet.autoResizeColumn(c);
        const currentWidth = sheet.getColumnWidth(c);
        sheet.setColumnWidth(c, currentWidth + 15);
      } catch (e) {
        logVerbose_(`Could not auto-resize/pad column ${c} for sheet ${sheet.getName()}: ${e.message}`);
      }
    }
  }
}

function ensureSheet_(ss, sheetName, isHidden) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    try {
      sheet.clearContents().clearFormats().clearNotes();
      const charts = sheet.getCharts();
      for (let i = 0; i < charts.length; i++) {
        sheet.removeChart(charts[i]);
      }
      sheet.clearConditionalFormatRules();
      logVerbose_(`Cleared existing sheet: "${sheetName}"`);
    } catch (e) {
      logError_('ensureSheet_ (Clearing)', e, `Failed to clear sheet: ${sheetName}. Attempting to delete and recreate.`);
      try {
        ss.deleteSheet(sheet);
        sheet = ss.insertSheet(sheetName);
        logInfo_(`Recreated sheet after clearing error: "${sheetName}"`);
      } catch (e2) {
        logError_('ensureSheet_ (Recreating)', e2, `Failed to recreate sheet: ${sheetName}. This sheet may be unusable.`);
        return null;
      }
    }
  } else {
    try {
      sheet = ss.insertSheet(sheetName);
      logInfo_(`Created new sheet: "${sheetName}"`);
    } catch (e) {
      logError_('ensureSheet_ (Inserting)', e, `Failed to insert new sheet: ${sheetName}.`);
      return null;
    }
  }

  if (!sheet) return null;


  if (isHidden) {
    try {
      if (!sheet.isSheetHidden()) sheet.hideSheet();
    } catch (e) {
      logError_('ensureSheet_', e, `Failed to hide sheet: ${sheetName}`);
    }
  } else {
    try {
      if (sheet.isSheetHidden()) sheet.showSheet();
    } catch (e) {
      logError_('ensureSheet_', e, `Failed to show sheet: ${sheetName}`);
    }
  }
  return sheet;
}

function styleHeader_(range, styleObject) {
  const style = styleObject || STYLE_UNIFIED_HEADER;
  range.setFontFamily(style.fontFamily)
    .setFontSize(style.fontSize)
    .setFontWeight(style.fontWeight)
    .setFontColor(style.fontColor)
    .setBackground(style.backgroundColor)
    .setHorizontalAlignment(style.horizontalAlignment)
    .setVerticalAlignment(style.verticalAlignment);

  if (style.wrapStrategy !== undefined) {
    range.setWrapStrategy(style.wrapStrategy);
  }
  if (style.textRotation !== undefined) {
    range.setTextRotation(style.textRotation);
  }
  range.setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function createSectionHeader_(sheet, row, text, numColumnsToMerge) {
    if (!sheet) return;
    const mergeCols = numColumnsToMerge || sheet.getMaxColumns();
    const range = sheet.getRange(row, 1, 1, mergeCols);
    range.mergeAcross();
    range.setValue(text);
    styleHeader_(range, STYLE_SECTION_HEADER);
    sheet.setRowHeight(row, 35);
    return row + 1;
}


function applyGeneralSheetFormatting_(sheet) {
  if (!sheet) return;
  try {
    sheet.setHiddenGridlines(true);
    sheet.setRowHeight(1, 30);
    if (sheet.getMaxRows() > 1) {
        const numDataRows = sheet.getMaxRows() - 1;
        if (numDataRows > 0) {
            sheet.setRowHeights(2, numDataRows, 25);
        }
    }
  } catch (e) {
    logError_('applyGeneralSheetFormatting_', e, `Failed to apply general formatting for sheet: ${sheet.getName()}`);
  }
}

function applyTableStyles(sheet, startRow, numRows, numCols) {
  if (!sheet || numRows <= 0 || numCols <= 0 || startRow <= 0) return;
  try {
    const maxSheetRows = sheet.getMaxRows();
    if (startRow + numRows - 1 > maxSheetRows) {
      logWarn_('applyTableStyles', `Calculated end row ${startRow + numRows - 1} exceeds max rows ${maxSheetRows} for sheet ${sheet.getName()}. Adjusting numRows.`);
      numRows = maxSheetRows - startRow + 1;
      if (numRows <= 0) return;
    }

    const tableRange = sheet.getRange(startRow, 1, numRows, numCols);

    tableRange.setBorder(true, true, true, true, null, null, "#424242", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    if (numRows > 1) {
        sheet.getRange(startRow + 1, 1, numRows - 1, numCols).setBorder(null, null, null, null, true, true, "#D0D0D0", SpreadsheetApp.BorderStyle.SOLID_THIN);
    } else {
         tableRange.setBorder(true, true, true, true, true, true, "#424242", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }


    if (numRows > 1) {
        const isLastRowTotal = sheet.getRange(startRow + numRows - 1, 1).getDisplayValue().toUpperCase() === 'TOTAL';
        const dataBodyRowCount = isLastRowTotal ? numRows - 2 : numRows - 1;

        if (dataBodyRowCount > 0) {
            const dataBodyRange = sheet.getRange(startRow + 1, 1, dataBodyRowCount, numCols);
            const backgrounds = [];
            for (let i = 0; i < dataBodyRowCount; i++) {
                backgrounds.push(new Array(numCols).fill(
                    i % 2 === 0 ? STYLE_TABLE_ROW_ODD_BG : STYLE_TABLE_ROW_EVEN_BG
                ));
            }
            if (backgrounds.length > 0) {
                dataBodyRange.setBackgrounds(backgrounds);
            }
        }
    }
    styleHeader_(sheet.getRange(startRow, 1, 1, numCols));
    logVerbose_(`Applied enhanced table styles to range ${tableRange.getA1Notation()} on sheet ${sheet.getName()}`);
  } catch (e) {
    logError_('applyTableStyles', e, `Failed to apply table styles to sheet: ${sheet.getName()}`);
  }
}

function centerAlignSheetData(sheet) {
  if (!sheet || sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) return;
  try {
    const dataRange = sheet.getDataRange();
    dataRange.setVerticalAlignment("middle");
    for (let i = 1; i <= dataRange.getNumColumns(); i++) {
        const columnRange = dataRange.offset(0, i - 1, dataRange.getNumRows(), 1);
        const firstValue = columnRange.offset(1,0,1,1).getValue();
        if (typeof firstValue === 'number' && !isDateColumnByName_(sheet.getRange(1,i).getValue())) {
             columnRange.setHorizontalAlignment("right");
        } else {
             columnRange.setHorizontalAlignment("left");
        }
    }
    if (sheet.getFrozenRows() > 0) {
        sheet.getRange(1, 1, sheet.getFrozenRows(), sheet.getLastColumn()).setHorizontalAlignment("center");
    } else if (sheet.getLastRow() > 0) {
        sheet.getRange(1, 1, 1, sheet.getLastColumn()).setHorizontalAlignment("center");
    }


  } catch(e) {
    logError_('centerAlignSheetData', e, `Failed to center align data for sheet: ${sheet.getName()}`);
  }
}

function isDateColumnByName_(headerName) {
    if (!headerName || typeof headerName !== 'string') return false;
    return headerName.toLowerCase().includes('date');
}
// --- End HELPER_UTILITIES ---

// --- Section 3: LOGGING_FRAMEWORK ---
const LOG_LEVELS_ENUM = {
  ERROR: 0,
  INFO: 1,
  VERBOSE: 2
};
let CURRENT_LOG_LEVEL = LOG_VERBOSE ? LOG_LEVELS_ENUM.VERBOSE : LOG_LEVELS_ENUM.INFO;

function setLogLevel_(levelStr) {
  const upperLevel = String(levelStr || '').toUpperCase();
  if (LOG_LEVELS_ENUM.hasOwnProperty(upperLevel)) {
    CURRENT_LOG_LEVEL = LOG_LEVELS_ENUM[upperLevel];
    Logger.log(`[INFO] Log level set to ${upperLevel}`);
  } else {
    Logger.log(`[WARNING] Invalid log level specified: ${levelStr}. Current log level remains unchanged.`);
  }
}

function logError_(functionName, errorObject, context) {
  const timestamp = Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd HH:mm:ss z');
  const errorMessage = errorObject.message || String(errorObject);
  const stackTrace = DEBUG_MODE && errorObject.stack ? errorObject.stack : (errorObject.stack ? 'Stack trace available in DEBUG_MODE' : 'N/A');
  const queryContext = context || 'N/A';

  Logger.log(`[ERROR] [${functionName}] ${errorMessage}`);
  if (DEBUG_MODE && errorObject.stack) {
    Logger.log(`Stack Trace: ${errorObject.stack}`);
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let errorSheet = spreadsheet.getSheetByName(ERROR_SHEET_NAME);
    if (!errorSheet) {
      errorSheet = spreadsheet.insertSheet(ERROR_SHEET_NAME, 0);
      const headers = ['Timestamp', 'Function', 'Error Message', 'Context / GAQL', 'Stack Trace (Verbose)'];
      const headerRange = errorSheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      styleHeader_(headerRange);
      errorSheet.setFrozenRows(1);
    }
    errorSheet.insertRowAfter(1);
    errorSheet.getRange(2, 1, 1, 5).setValues([[timestamp, functionName, errorMessage, queryContext, stackTrace]]);
    autoResizeSheetColumns_(errorSheet, 5);
  } catch (e) {
    Logger.log(`[CRITICAL] Failed to write error to ERRORS sheet: ${e.message}. Original error in ${functionName}: ${errorMessage}`);
  }
}

function logInfo_(message) {
  if (CURRENT_LOG_LEVEL >= LOG_LEVELS_ENUM.INFO) {
    Logger.log(`[INFO] ${message}`);
  }
}

function logWarn_(functionName, message) {
  if (CURRENT_LOG_LEVEL >= LOG_LEVELS_ENUM.INFO) {
    Logger.log(`[WARNING] [${functionName}] ${message}`);
  }
}

function logVerbose_(message) {
  if (CURRENT_LOG_LEVEL >= LOG_LEVELS_ENUM.VERBOSE) {
    Logger.log(`[VERBOSE] ${message}`);
  }
}
// --- End LOGGING_FRAMEWORK ---

// --- Section 4: SCRIPT_INITIALIZATION_AND_CHECKS ---
function checkRemainingTimeGraceful_() {
  return AdsApp.getExecutionInfo().getRemainingTime() >= MIN_TIME_BUFFER_SEC;
}

function checkRuntimeBudgetOrThrow_() {
  if (!checkRemainingTimeGraceful_()) {
    const remaining = AdsApp.getExecutionInfo().getRemainingTime();
    logError_('checkRuntimeBudgetOrThrow_', new Error(`Low execution time: ${remaining}s`), 'Aborting script.');
    throw new Error(`Less than ${MIN_TIME_BUFFER_SEC} seconds of script execution time remaining – aborting.`);
  }
  logVerbose_(`Runtime budget check passed. Remaining: ${AdsApp.getExecutionInfo().getRemainingTime()}s`);
}

function assertSheetsAccessEnabled_() {
  try {
    SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.openById(SPREADSHEET_ID);
    SpreadsheetApp.flush();
    logVerbose_('Google Sheets API access confirmed.');
  } catch (e) {
    logError_('assertSheetsAccessEnabled_', e, 'Sheets API access error.');
    throw new Error('Sheets API is disabled or permission not granted. Please run the script once in the editor, grant the necessary OAuth permissions when prompted, and then try running it again. Error details: ' + e.message);
  }
}

function assertAdsApiVersionCompatible_() {
  logVerbose_('Ads API version check: Using default version provided by Google Ads Scripts environment.');
}

function initializeScriptEnvironment_() {
  checkRuntimeBudgetOrThrow_();
  assertSheetsAccessEnabled_();
  assertAdsApiVersionCompatible_();
  logInfo_('Initialization checks passed. Script environment is ready.');
}
// --- End SCRIPT_INITIALIZATION_AND_CHECKS ---

// --- Section 5: GOOGLE_ADS_DATA_EXTRACTION ---
function getDateClause_() {
  return ` BETWEEN '${REPORT_START_DATE}' AND '${REPORT_END_DATE}' `;
}

function executeGAQLQuery_(query, queryLabel) {
  if (!checkRemainingTimeGraceful_()) {
    logError_('executeGAQLQuery_', new Error('Not enough time remaining to execute GAQL query.'), `Query Label: ${queryLabel}`);
    return null;
  }
  try {
    logVerbose_(`Executing GAQL [${queryLabel}]: ${query.replace(/\s+/g, ' ')}`);
    const report = AdsApp.search(query);
    const rows = [];
    while (report.hasNext()) {
      rows.push(report.next());
    }
    logInfo_(`GAQL [${queryLabel}] successfully returned ${rows.length} rows.`);
    return rows;
  } catch (e) {
    logError_(`executeGAQLQuery_ (${queryLabel})`, e, query);
    return null;
  }
}

function fetchAccountConfigData_() {
  const query = `
    SELECT customer.currency_code, customer.time_zone, customer.descriptive_name
    FROM customer LIMIT 1`;
  return executeGAQLQuery_(query, 'AccountConfig');
}

function fetchCampaignData_() {
  const query = `
    SELECT
      campaign.id, campaign.name, campaign.status, 
      campaign.advertising_channel_type, campaign.advertising_channel_sub_type,
      campaign.bidding_strategy_type, campaign_budget.amount_micros,
      metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.ctr,
      metrics.average_cpc, metrics.conversions, metrics.conversions_value,
      metrics.search_impression_share, 
      metrics.search_budget_lost_impression_share, 
      metrics.search_rank_lost_impression_share,
      metrics.search_top_impression_share,
      metrics.search_absolute_top_impression_share
    FROM campaign
    WHERE segments.date ${getDateClause_()}
      AND campaign.status != 'REMOVED'
    ORDER BY campaign.name
    LIMIT ${GAQL_FETCH_LIMIT}`;
  return executeGAQLQuery_(query, 'CampaignPerformance');
}

function fetchDailyData_() {
  const query = `
    SELECT
      segments.date, segments.device, campaign.id, campaign.name, 
      metrics.cost_micros, metrics.impressions, metrics.clicks,
      metrics.conversions, metrics.conversions_value
    FROM campaign 
    WHERE segments.date ${getDateClause_()}
    ORDER BY segments.date ASC, campaign.id
    LIMIT ${GAQL_FETCH_LIMIT * 3}`;
  return executeGAQLQuery_(query, 'DailyPerformance');
}

function fetchAdGroupData_() {
  const query = `
    SELECT
      campaign.id, campaign.name, 
      ad_group.id, ad_group.name, ad_group.status,
      metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.ctr,
      metrics.average_cpc, metrics.conversions, metrics.conversions_value
    FROM ad_group
    WHERE segments.date ${getDateClause_()}
      AND campaign.status != 'REMOVED'
      AND ad_group.status != 'REMOVED'
    ORDER BY campaign.name, ad_group.name
    LIMIT ${GAQL_FETCH_LIMIT}`;
  return executeGAQLQuery_(query, 'AdGroupPerformance');
}

function fetchAdData_() {
  const query = `
    SELECT
      campaign.name, ad_group.name,
      ad_group_ad.ad.id, ad_group_ad.ad.type, ad_group_ad.status,
      ad_group_ad.ad.final_urls,
      metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.ctr,
      metrics.conversions, metrics.conversions_value
    FROM ad_group_ad
    WHERE segments.date ${getDateClause_()}
      AND campaign.status != 'REMOVED'
      AND ad_group.status != 'REMOVED'
      AND ad_group_ad.status != 'REMOVED'
    ORDER BY campaign.name, ad_group.name, ad_group_ad.ad.id
    LIMIT ${GAQL_FETCH_LIMIT}`;
  return executeGAQLQuery_(query, 'AdPerformance');
}

function fetchKeywordData_() {
  const query = `
    SELECT
      campaign.name, ad_group.name,
      ad_group_criterion.criterion_id, ad_group_criterion.keyword.text, 
      ad_group_criterion.keyword.match_type, ad_group_criterion.status,
      ad_group_criterion.quality_info.quality_score,
      metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.ctr,
      metrics.average_cpc, metrics.conversions, metrics.conversions_value
    FROM keyword_view
    WHERE segments.date ${getDateClause_()}
      AND campaign.status != 'REMOVED'
      AND ad_group.status != 'REMOVED'
      AND ad_group_criterion.status != 'REMOVED'
    ORDER BY campaign.name, ad_group.name, metrics.cost_micros DESC
    LIMIT ${GAQL_FETCH_LIMIT}`;
  return executeGAQLQuery_(query, 'KeywordPerformance');
}

function fetchProductData_() {
  const query = `
    SELECT
      campaign.id, campaign.name, ad_group.id, ad_group.name,
      segments.product_item_id, segments.product_title, segments.product_brand,
      segments.product_type_l1, segments.product_type_l2, 
      segments.product_custom_attribute0,
      metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.ctr,
      metrics.average_cpc, metrics.conversions, metrics.conversions_value
    FROM shopping_performance_view
    WHERE segments.date ${getDateClause_()}
      AND metrics.impressions > 0 
    ORDER BY metrics.cost_micros DESC
    LIMIT ${GAQL_FETCH_LIMIT}`;
  return executeGAQLQuery_(query, 'ProductPerformance');
}

function fetchAssetPerformanceData_() {
  const query = `
    SELECT
        campaign.name, campaign.id,
        ad_group.name, ad_group.id,
        asset.id, asset.type, 
        asset_field_type_view.field_type,
        metrics.impressions,
        metrics.clicks,
        metrics.cost_micros
    FROM asset_field_type_view
    WHERE segments.date ${getDateClause_()}
      AND campaign.status != 'REMOVED'
      AND ad_group.status != 'REMOVED' 
    ORDER BY campaign.name, ad_group.name, asset.type, asset_field_type_view.field_type, metrics.impressions DESC
    LIMIT ${GAQL_FETCH_LIMIT}`;
  return executeGAQLQuery_(query, 'AssetPerformance');
}

function fetchPMaxAssetGroupData_() {
  const query = `
    SELECT
      campaign.id, campaign.name,
      asset_group.id, asset_group.name, asset_group.status,
      metrics.cost_micros, metrics.impressions, metrics.clicks,
      metrics.conversions, metrics.conversions_value
    FROM asset_group
    WHERE segments.date ${getDateClause_()}
      AND campaign.advertising_channel_type = 'PERFORMANCE_MAX'
      AND campaign.status != 'REMOVED'
      AND asset_group.status != 'REMOVED'
    ORDER BY campaign.name, asset_group.name
    LIMIT ${GAQL_FETCH_LIMIT}`;
  return executeGAQLQuery_(query, 'PMaxAssetGroupPerformance');
}

function fetchPMaxListingGroupData_() {
  const query = `
    SELECT
      campaign.id, campaign.name, asset_group.id, asset_group.name,
      asset_group_listing_group_filter.id,
      asset_group_listing_group_filter.type,
      asset_group_listing_group_filter.case_value.product_brand.value,
      asset_group_listing_group_filter.case_value.product_item_id.value,
      asset_group_listing_group_filter.case_value.product_type.value, 
      asset_group_listing_group_filter.case_value.product_type.level
    FROM asset_group_listing_group_filter
    WHERE campaign.advertising_channel_type = 'PERFORMANCE_MAX' 
      AND campaign.status != 'REMOVED' AND asset_group.status != 'REMOVED'
      AND asset_group_listing_group_filter.type != 'SUBDIVISION' 
    ORDER BY campaign.name, asset_group.name, asset_group_listing_group_filter.id
    LIMIT ${GAQL_FETCH_LIMIT}`;
  return executeGAQLQuery_(query, 'PMaxListingGroupStructure');
}

function fetchYouTubeData_() {
  const query = `
    SELECT
      campaign.name, campaign.status,
      metrics.cost_micros, metrics.impressions, metrics.video_views,
      metrics.video_view_rate, metrics.average_cpv,
      metrics.video_quartile_p25_rate, metrics.video_quartile_p50_rate,
      metrics.video_quartile_p75_rate, metrics.video_quartile_p100_rate,
      metrics.clicks, metrics.conversions, metrics.conversions_value
    FROM campaign
    WHERE segments.date ${getDateClause_()}
      AND campaign.advertising_channel_type = 'VIDEO'
      AND campaign.status != 'REMOVED'
    ORDER BY campaign.name, metrics.cost_micros DESC
    LIMIT ${GAQL_FETCH_LIMIT}`;
  return executeGAQLQuery_(query, 'YouTubePerformance');
}

function fetchAudienceData_() {
  const query = `
    SELECT
      campaign.name, ad_group.name, 
      ad_group_criterion.display_name, ad_group_criterion.type,
      metrics.cost_micros, metrics.impressions, metrics.clicks,
      metrics.conversions, metrics.conversions_value
    FROM ad_group_audience_view
    WHERE segments.date ${getDateClause_()}
      AND campaign.status != 'REMOVED'
      AND ad_group.status != 'REMOVED'
      AND ad_group_criterion.status != 'REMOVED'
    ORDER BY campaign.name, ad_group.name, metrics.cost_micros DESC
    LIMIT ${GAQL_FETCH_LIMIT}`;
  return executeGAQLQuery_(query, 'AudiencePerformance');
}

function fetchImpressionShareRankData_() {
  const query = `
    SELECT
      segments.date,
      campaign.id,
      campaign.name,
      metrics.search_impression_share,
      metrics.search_budget_lost_impression_share, 
      metrics.search_rank_lost_impression_share,
      metrics.search_top_impression_share,
      metrics.search_absolute_top_impression_share
    FROM campaign
    WHERE segments.date ${getDateClause_()}
      AND campaign.advertising_channel_type IN ('SEARCH', 'SHOPPING') 
      AND campaign.status != 'REMOVED'
    ORDER BY segments.date, campaign.name
    LIMIT ${GAQL_FETCH_LIMIT}`;
  return executeGAQLQuery_(query, 'ImpressionShareRank');
}

function fetchAssetLibraryData_() {
  const negKeywordLists = executeGAQLQuery_(`
    SELECT shared_set.id, shared_set.name, shared_set.status, shared_set.type, shared_set.member_count
    FROM shared_set
    WHERE shared_set.type = 'NEGATIVE_KEYWORDS'
    ORDER BY shared_set.name`, 'SharedNegativeKeywordLists');

  return {
    negativeKeywordLists: negKeywordLists || []
  };
}
// --- End GOOGLE_ADS_DATA_EXTRACTION ---

// --- Section 6: RAW_DATA_SHEET_PROCESSING ---
const HDR_CAMPAIGNS = ['Campaign ID', 'Campaign Name', 'Status', 'Budget',
  'Channel', 'Sub-Channel', 'Bidding', 'Cost', 'Impr.',
  'Clicks', 'CTR', 'Avg. CPC', 'Conv.', 'Conv. Value',
  'ROAS', 'CPA', 'AOV',
  'Search IS', 'IS Lost (Budget)', 'IS Lost (Rank)', 'Top IS', 'Abs Top IS'
];
const HDR_DAILY = ['Date', 'Device', 'Campaign ID', 'Campaign Name', 'Cost', 'Impr.', 'Clicks', 'Conv.', 'Conv. Value', 'AOV'];
const HDR_ADGROUPS = ['Campaign ID', 'Campaign Name', 'AdGroup ID', 'AdGroup Name',
  'Status', 'Cost', 'Impr.', 'Clicks', 'CTR', 'Avg. CPC',
  'Conv.', 'Conv. Value', 'ROAS', 'CPA', 'AOV'
];
const HDR_ADS = ['Campaign Name', 'AdGroup Name', 'Ad ID', 'Ad Type', 'Status',
  'Final URLs', 'Cost', 'Impr.', 'Clicks', 'CTR',
  'Conv.', 'Conv. Value', 'ROAS', 'CPA', 'AOV'
];
const HDR_KEYWORDS = ['Campaign Name', 'AdGroup Name', 'KW ID', 'Keyword', 'Match Type',
  'Status', 'QS', 'Cost', 'Impr.', 'Clicks', 'CTR',
  'Avg. CPC', 'Conv.', 'Conv. Value', 'ROAS', 'CPA', 'AOV'
];
const HDR_PRODUCTS = ['Campaign ID', 'Campaign Name', 'AdGroup ID', 'AdGroup Name',
  'Product ID', 'Product Title', 'Brand', 'Product Type L1', 'Product Type L2',
  'Custom Label 0', 'Cost', 'Impr.', 'Clicks', 'CTR', 'Avg. CPC',
  'Conv.', 'Conv. Value', 'ROAS', 'CPA', 'AOV'
];
const HDR_ASSETS = ['Campaign Name', 'Campaign ID', 'Ad Group Name', 'Ad Group ID', 'Asset ID', 'Asset Type', 'Field Type', 'Impressions', 'Clicks', 'Cost'];

const HDR_PMAX_AG = ['Campaign ID', 'Campaign Name', 'AssetGroup ID', 'AssetGroup Name',
  'Status', 'Cost', 'Impr.', 'Clicks', 'Conv.', 'Conv. Value',
  'ROAS', 'CPA', 'AOV'
];
const HDR_PMAX_LG = ['Campaign ID', 'Campaign Name', 'AssetGroup ID', 'AssetGroup Name',
  'ListingGroup ID', 'ListingGroup Type', 'Dimension Value', 'Dimension Level'
];
const HDR_YT = ['Campaign Name', 'Status', 'Cost', 'Impr.', 'Video Views', 'View Rate',
  'Avg. CPV', 'Q25 %', 'Q50 %', 'Q75 %', 'Q100 %',
  'Clicks', 'Conv.', 'Conv. Value', 'AOV'
];
const HDR_AUDIENCE = ['Campaign Name', 'AdGroup Name', 'Audience Display Name', 'Criterion Type', 'Cost',
  'Impr.', 'Clicks', 'Conv.', 'Conv. Value', 'ROAS', 'CPA', 'AOV'
];
const HDR_IS_RANK = ['Date', 'Campaign ID', 'Campaign Name', 'Search IS', 'IS Lost (Budget)', 'IS Lost (Rank)', 'Top IS', 'Abs Top IS'];
const HDR_SHARED_NEG_LISTS = ['List ID', 'List Name', 'Status', 'Type', 'Member Count'];
// FIX: Reordered columns for HDR_LAG_FORECAST
const HDR_LAG_FORECAST = ['Applied Lag Bucket', 'Date', 'Reported Conversions', 'Reported Conv. Value', 'Days Ago', 'Calculated Uplift Factor', 'Adjusted Conversions', 'Adjusted Conv. Value'];


const RAW_SHEET_DEFINITIONS_ = [{
  name: 'Campaigns',
  headers: HDR_CAMPAIGNS,
  fetchFn: fetchCampaignData_,
  writerFn: writeCampaignRawData_,
  enabled: () => true
}, {
  name: 'Daily Performance',
  headers: HDR_DAILY,
  fetchFn: fetchDailyData_,
  writerFn: writeDailyRawData_,
  enabled: () => true
}, {
  name: 'Impression Share & Rank',
  headers: HDR_IS_RANK,
  fetchFn: fetchImpressionShareRankData_,
  writerFn: writeImpressionShareRankRawData_,
  enabled: () => PROCESS_IMPRESSION_SHARE_DATA
}, {
  name: 'Ad Groups',
  headers: HDR_ADGROUPS,
  fetchFn: fetchAdGroupData_,
  writerFn: writeAdGroupRawData_,
  enabled: () => true
}, {
  name: 'Ads',
  headers: HDR_ADS,
  fetchFn: fetchAdData_,
  writerFn: writeAdRawData_,
  enabled: () => true
}, {
  name: 'Keywords',
  headers: HDR_KEYWORDS,
  fetchFn: fetchKeywordData_,
  writerFn: writeKeywordRawData_,
  enabled: () => true
}, {
  name: 'Products',
  headers: HDR_PRODUCTS,
  fetchFn: fetchProductData_,
  writerFn: writeProductRawData_,
  enabled: () => true
}, {
  name: 'Asset Performance',
  headers: HDR_ASSETS,
  fetchFn: fetchAssetPerformanceData_,
  writerFn: writeAssetPerformanceRawData_,
  enabled: () => PROCESS_ASSET_PERFORMANCE_DATA
}, {
  name: 'PMax Asset Groups',
  headers: HDR_PMAX_AG,
  fetchFn: fetchPMaxAssetGroupData_,
  writerFn: writePMaxAssetGroupRawData_,
  enabled: () => true
}, {
  name: 'PMax Listing Groups',
  headers: HDR_PMAX_LG,
  fetchFn: fetchPMaxListingGroupData_,
  writerFn: writePMaxListingGroupRawData_,
  enabled: () => true
}, {
  name: 'YouTube Campaigns',
  headers: HDR_YT,
  fetchFn: fetchYouTubeData_,
  writerFn: writeYouTubeRawData_,
  enabled: () => true
}, {
  name: 'Audience Performance',
  headers: HDR_AUDIENCE,
  fetchFn: fetchAudienceData_,
  writerFn: writeAudienceRawData_,
  enabled: () => true
}, {
  name: 'Asset Library - Neg KW Lists',
  headers: HDR_SHARED_NEG_LISTS,
  fetchFn: () => fetchAssetLibraryData_().negativeKeywordLists,
  writerFn: writeAssetLibraryNegKWLists_,
  enabled: () => PROCESS_LIBRARY_EXPORT_DATA
}, ];

function initializeRawDataSheets_(ss) {
  logInfo_('Initializing raw data sheets structure...');
  SpreadsheetApp.flush();

  RAW_SHEET_DEFINITIONS_.forEach(def => {
    const sheetFullName = TAB_PREFIX_RAW + def.name;
    if (typeof def.enabled === 'function' && !def.enabled()) {
      logInfo_(`Skipping raw sheet initialization for "${def.name}" as its processing is disabled.`);
      const existingSheet = ss.getSheetByName(sheetFullName);
      if (existingSheet) {
        try {
          ss.deleteSheet(existingSheet);
          logVerbose_(`Deleted existing raw sheet for disabled feature: "${sheetFullName}"`);
        } catch (e) {
          logError_('initializeRawDataSheets_', e, `Failed to delete raw sheet for disabled feature: ${sheetFullName}`);
        }
      }
      return;
    }

    const sheet = ensureSheet_(ss, sheetFullName, true);
    if (sheet && def.headers && def.headers.length > 0) {
      try {
        const headerRange = sheet.getRange(1, 1, 1, def.headers.length);
        headerRange.setValues([def.headers]);
        styleHeader_(headerRange);
        sheet.setFrozenRows(1);
        applyGeneralSheetFormatting_(sheet);
      } catch (e) {
        logError_('initializeRawDataSheets_', e, `Failed to set headers for sheet: ${sheetFullName}`);
      }
    } else if (!sheet) {
      logError_('initializeRawDataSheets_', new Error('Sheet object is null after ensureSheet_ call.'), `Sheet: ${sheetFullName}`);
    }
  });
  SpreadsheetApp.flush();
  logInfo_('Raw data sheets structure initialized and sheets hidden (where applicable).');
}

function writeDataToRawSheet_(ss, simpleSheetName, headers, dataRows, currencyCode) {
  const fullSheetName = TAB_PREFIX_RAW + simpleSheetName;
  const sheet = ss.getSheetByName(fullSheetName);
  if (!sheet) {
    logError_('writeDataToRawSheet_', new Error(`Sheet ${fullSheetName} not found during write operation.`), `Sheet: ${simpleSheetName}`);
    return;
  }

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent().clearFormat();
  }

  if (dataRows && dataRows.length > 0) {
    if (dataRows[0].length !== headers.length) {
      logError_('writeDataToRawSheet_', new Error(`Header/data column mismatch for ${fullSheetName}.`), `Headers: ${headers.length}, DataCols: ${dataRows[0].length}`);
      return;
    }
    const dataRange = sheet.getRange(2, 1, dataRows.length, headers.length);
    dataRange.setValues(dataRows);
    logInfo_(`Successfully wrote ${dataRows.length} rows to ${fullSheetName}.`);

    const backgrounds = [];
    for (let i = 0; i < dataRows.length; i++) {
        backgrounds.push(new Array(headers.length).fill(
            i % 2 === 0 ? STYLE_TABLE_ROW_ODD_BG : STYLE_TABLE_ROW_EVEN_BG
        ));
    }
    if (backgrounds.length > 0) {
        dataRange.setBackgrounds(backgrounds);
    }


    const columnFormatStrings = headers.map(header => getColumnFormatForHeader_(header, currencyCode));
    const formatsArrayForSheet = [];
    for (let r = 0; r < dataRows.length; r++) {
      formatsArrayForSheet.push(columnFormatStrings);
    }
    applyChunkedFormatting_(sheet, formatsArrayForSheet, 2);
  } else {
    logInfo_(`No data rows to write to ${fullSheetName}.`);
  }
  autoResizeSheetColumns_(sheet, headers.length);
  styleHeader_(sheet.getRange(1, 1, 1, headers.length));
  sheet.setFrozenRows(1);
}

function getColumnFormatForHeader_(header, currencyCode) {
  const lowerHeader = String(header || '').toLowerCase();

  if (['search is', 'is lost (budget)', 'is lost (rank)', 'top is', 'abs top is', 'abs. top is'].indexOf(lowerHeader) !== -1) {
    return '0.00%';
  }
  if (lowerHeader.includes('cost') || lowerHeader.includes('cpc') || lowerHeader.includes('cpm') ||
    lowerHeader.includes('conv. value') || lowerHeader.includes('budget') ||
    lowerHeader.includes('value') || lowerHeader.includes('aov') || lowerHeader.includes('cpv') ||
    lowerHeader.includes('cpa')) {
    return getCurrencyFormatString_(currencyCode);
  } else if (lowerHeader.includes('ctr') || lowerHeader.includes('rate') || lowerHeader.includes('cvr') ||
    lowerHeader.includes('%') || lowerHeader.includes('share')) {
    return '0.00%';
  } else if (lowerHeader.includes('clicks') || lowerHeader.includes('impr.') ||
    lowerHeader.includes('conv.') || lowerHeader.includes('views') ||
    (lowerHeader.includes('id') && !lowerHeader.includes('video id') && !lowerHeader.includes('product id')) ||
    lowerHeader.includes('qs') || lowerHeader.includes('quality score') ||
    lowerHeader.includes('member count') || lowerHeader.includes('days ago')) {
    return '#,##0';
  } else if (lowerHeader.includes('roas') || lowerHeader.includes('uplift factor')) {
    return '#,##0.00';
  } else if (lowerHeader.includes('date')) {
      return 'yyyy-mm-dd';
  }
  return '@';
}

function writeCampaignRawData_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'Campaigns');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) {
    logError_('writeCampaignRawData_', new Error('No data from fetch function.'), def.name);
    return;
  }
  const data = gaqlRows.map(row => {
    const camp = row.campaign;
    const mets = row.metrics;
    const campBudget = row.campaignBudget;
    let budgetDisplay = 'N/A';
    if (campBudget && typeof campBudget.amountMicros !== 'undefined' && campBudget.amountMicros !== null) {
      budgetDisplay = safeMicrosToCurrency_(campBudget.amountMicros);
    }
    return [
      camp.id, camp.name, camp.status, budgetDisplay,
      camp.advertisingChannelType, camp.advertisingChannelSubType, camp.biddingStrategyType,
      safeMicrosToCurrency_(mets.costMicros), safeParseInt_(mets.impressions), safeParseInt_(mets.clicks),
      safeParseFloat_(mets.ctr),
      safeMicrosToCurrency_(mets.averageCpc),
      safeParseFloat_(mets.conversions),
      safeParseFloat_(mets.conversionsValue),
      div(safeParseFloat_(mets.conversionsValue), safeMicrosToCurrency_(mets.costMicros)),
      div(safeMicrosToCurrency_(mets.costMicros), safeParseFloat_(mets.conversions)),
      div(safeParseFloat_(mets.conversionsValue), safeParseFloat_(mets.conversions)),
      safeParseFloat_(mets.searchImpressionShare),
      safeParseFloat_(mets.searchBudgetLostImpressionShare),
      safeParseFloat_(mets.searchRankLostImpressionShare),
      safeParseFloat_(mets.searchTopImpressionShare),
      safeParseFloat_(mets.searchAbsoluteTopImpressionShare)
    ];
  });
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writeDailyRawData_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'Daily Performance');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) { logError_('writeDailyRawData_', new Error('No data from fetch function.'), def.name); return; }
  const data = gaqlRows.map(row => {
    const segs = row.segments; const mets = row.metrics; const camp = row.campaign;
    return [
      segs.date, segs.device, camp.id, camp.name,
      safeMicrosToCurrency_(mets.costMicros), safeParseInt_(mets.impressions), safeParseInt_(mets.clicks),
      safeParseFloat_(mets.conversions),
      safeParseFloat_(mets.conversionsValue),
      div(safeParseFloat_(mets.conversionsValue), safeParseFloat_(mets.conversions))
    ];
  });
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writeImpressionShareRankRawData_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'Impression Share & Rank');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) { logError_('writeImpressionShareRankRawData_', new Error('No data from fetch function.'), def.name); return; }
  const data = gaqlRows.map(row => [
    row.segments.date, row.campaign.id, row.campaign.name,
    safeParseFloat_(row.metrics.searchImpressionShare),
    safeParseFloat_(row.metrics.searchBudgetLostImpressionShare),
    safeParseFloat_(row.metrics.searchRankLostImpressionShare),
    safeParseFloat_(row.metrics.searchTopImpressionShare),
    safeParseFloat_(row.metrics.searchAbsoluteTopImpressionShare)
  ]);
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writeAdGroupRawData_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'Ad Groups');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) { logError_('writeAdGroupRawData_', new Error('No data from fetch function.'), def.name); return; }
  const data = gaqlRows.map(row => {
    const mets = row.metrics;
    return [
      row.campaign.id, row.campaign.name, row.adGroup.id, row.adGroup.name, row.adGroup.status,
      safeMicrosToCurrency_(mets.costMicros), safeParseInt_(mets.impressions), safeParseInt_(mets.clicks),
      safeParseFloat_(mets.ctr), safeMicrosToCurrency_(mets.averageCpc),
      safeParseFloat_(mets.conversions),
      safeParseFloat_(mets.conversionsValue),
      div(safeParseFloat_(mets.conversionsValue), safeMicrosToCurrency_(mets.costMicros)),
      div(safeMicrosToCurrency_(mets.costMicros), safeParseFloat_(mets.conversions)),
      div(safeParseFloat_(mets.conversionsValue), safeParseFloat_(mets.conversions))
    ];
  });
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writeAdRawData_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'Ads');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) { logError_('writeAdRawData_', new Error('No data from fetch function.'), def.name); return; }
  const data = gaqlRows.map(row => {
    const adAd = row.adGroupAd; const ad = adAd.ad; const mets = row.metrics;
    return [
      row.campaign.name, row.adGroup.name, ad.id, ad.type, adAd.status,
      (ad.finalUrls || []).join(', '),
      safeMicrosToCurrency_(mets.costMicros), safeParseInt_(mets.impressions), safeParseInt_(mets.clicks),
      safeParseFloat_(mets.ctr),
      safeParseFloat_(mets.conversions),
      safeParseFloat_(mets.conversionsValue),
      div(safeParseFloat_(mets.conversionsValue), safeMicrosToCurrency_(mets.costMicros)),
      div(safeMicrosToCurrency_(mets.costMicros), safeParseFloat_(mets.conversions)),
      div(safeParseFloat_(mets.conversionsValue), safeParseFloat_(mets.conversions))
    ];
  });
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writeKeywordRawData_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'Keywords');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) { logError_('writeKeywordRawData_', new Error('No data from fetch function.'), def.name); return; }
  const data = gaqlRows.map(row => {
    const crit = row.adGroupCriterion; const mets = row.metrics;
    return [
      row.campaign.name, row.adGroup.name, crit.criterionId, crit.keyword.text, crit.keyword.matchType,
      crit.status, crit.qualityInfo ? crit.qualityInfo.qualityScore : null,
      safeMicrosToCurrency_(mets.costMicros), safeParseInt_(mets.impressions), safeParseInt_(mets.clicks),
      safeParseFloat_(mets.ctr), safeMicrosToCurrency_(mets.averageCpc),
      safeParseFloat_(mets.conversions),
      safeParseFloat_(mets.conversionsValue),
      div(safeParseFloat_(mets.conversionsValue), safeMicrosToCurrency_(mets.costMicros)),
      div(safeMicrosToCurrency_(mets.costMicros), safeParseFloat_(mets.conversions)),
      div(safeParseFloat_(mets.conversionsValue), safeParseFloat_(mets.conversions))
    ];
  });
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writeProductRawData_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'Products');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) { logError_('writeProductRawData_', new Error('No data from fetch function.'), def.name); return; }
  const data = gaqlRows.map(row => {
    const segs = row.segments; const mets = row.metrics;
    return [
      row.campaign.id, row.campaign.name, row.adGroup.id, row.adGroup.name,
      segs.productItemId, segs.productTitle, segs.productBrand,
      segs.productTypeL1, segs.productTypeL2, segs.productCustomAttribute0,
      safeMicrosToCurrency_(mets.costMicros), safeParseInt_(mets.impressions), safeParseInt_(mets.clicks),
      safeParseFloat_(mets.ctr), safeMicrosToCurrency_(mets.averageCpc),
      safeParseFloat_(mets.conversions),
      safeParseFloat_(mets.conversionsValue),
      div(safeParseFloat_(mets.conversionsValue), safeMicrosToCurrency_(mets.costMicros)),
      div(safeMicrosToCurrency_(mets.costMicros), safeParseFloat_(mets.conversions)),
      div(safeParseFloat_(mets.conversionsValue), safeParseFloat_(mets.conversions))
    ];
  });
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writeAssetPerformanceRawData_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'Asset Performance');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) { logError_('writeAssetPerformanceRawData_', new Error('No data from fetch for Asset Performance.'), def.name); return; }
  const data = gaqlRows.map(row => {
    const camp = row.campaign; const ag = row.adGroup; const asset = row.asset;
    const fieldTypeView = row.assetFieldTypeView; const mets = row.metrics;
    return [
      camp.name, camp.id,
      ag.name, ag.id,
      asset.id,
      asset.type,
      fieldTypeView.fieldType,
      safeParseInt_(mets.impressions),
      safeParseInt_(mets.clicks),
      safeMicrosToCurrency_(mets.costMicros)
    ];
  });
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writePMaxAssetGroupRawData_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'PMax Asset Groups');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) { logError_('writePMaxAssetGroupRawData_', new Error('No data from fetch function.'), def.name); return; }
  const data = gaqlRows.map(row => {
    const mets = row.metrics;
    return [
      row.campaign.id, row.campaign.name, row.assetGroup.id, row.assetGroup.name, row.assetGroup.status,
      safeMicrosToCurrency_(mets.costMicros), safeParseInt_(mets.impressions), safeParseInt_(mets.clicks),
      safeParseFloat_(mets.conversions),
      safeParseFloat_(mets.conversionsValue),
      div(safeParseFloat_(mets.conversionsValue), safeMicrosToCurrency_(mets.costMicros)),
      div(safeMicrosToCurrency_(mets.costMicros), safeParseFloat_(mets.conversions)),
      div(safeParseFloat_(mets.conversionsValue), safeParseFloat_(mets.conversions))
    ];
  });
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writePMaxListingGroupRawData_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'PMax Listing Groups');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) { logError_('writePMaxListingGroupRawData_', new Error('No data from fetch function.'), def.name); return; }
  const data = gaqlRows.map(row => {
    const lgf = row.assetGroupListingGroupFilter; let dimVal = "N/A", dimLvl = "N/A";
    if (lgf.caseValue) {
      if (lgf.caseValue.productBrand && lgf.caseValue.productBrand.value) { dimVal = lgf.caseValue.productBrand.value; dimLvl = 'Brand'; }
      else if (lgf.caseValue.productItemId && lgf.caseValue.productItemId.value) { dimVal = lgf.caseValue.productItemId.value; dimLvl = 'Item ID'; }
      else if (lgf.caseValue.productType && lgf.caseValue.productType.value) { dimVal = lgf.caseValue.productType.value; dimLvl = `Product Type L${lgf.caseValue.productType.level || '?'}`; }
    }
    return [
      row.campaign.id, row.campaign.name, row.assetGroup.id, row.assetGroup.name,
      lgf.id, lgf.type, dimVal, dimLvl
    ];
  });
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writeYouTubeRawData_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'YouTube Campaigns');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) { logError_('writeYouTubeRawData_', new Error('No data from fetch function.'), def.name); return; }
  const data = gaqlRows.map(row => {
    const mets = row.metrics;
    return [
      row.campaign.name, row.campaign.status,
      safeMicrosToCurrency_(mets.costMicros), safeParseInt_(mets.impressions), safeParseInt_(mets.videoViews),
      safeParseFloat_(mets.videoViewRate), safeMicrosToCurrency_(mets.averageCpv),
      safeParseFloat_(mets.videoQuartileP25Rate), safeParseFloat_(mets.videoQuartileP50Rate),
      safeParseFloat_(mets.videoQuartileP75Rate), safeParseFloat_(mets.videoQuartileP100Rate),
      safeParseInt_(mets.clicks), safeParseFloat_(mets.conversions),
      safeParseFloat_(mets.conversionsValue),
      div(safeParseFloat_(mets.conversionsValue), safeParseFloat_(mets.conversions))
    ];
  });
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writeAudienceRawData_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'Audience Performance');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) { logError_('writeAudienceRawData_', new Error('No data from fetch function.'), def.name); return; }
  const data = gaqlRows.map(row => {
    const crit = row.adGroupCriterion; const mets = row.metrics;
    return [
      row.campaign.name, row.adGroup.name, crit.displayName || 'N/A', crit.type,
      safeMicrosToCurrency_(mets.costMicros), safeParseInt_(mets.impressions), safeParseInt_(mets.clicks),
      safeParseFloat_(mets.conversions),
      safeParseFloat_(mets.conversionsValue),
      div(safeParseFloat_(mets.conversionsValue), safeMicrosToCurrency_(mets.costMicros)),
      div(safeMicrosToCurrency_(mets.costMicros), safeParseFloat_(mets.conversions)),
      div(safeParseFloat_(mets.conversionsValue), safeParseFloat_(mets.conversions))
    ];
  });
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writeAssetLibraryNegKWLists_(ss, currencyCode) {
  const def = RAW_SHEET_DEFINITIONS_.find(d => d.name === 'Asset Library - Neg KW Lists');
  const gaqlRows = def.fetchFn();
  if (!gaqlRows) { logError_('writeAssetLibraryNegKWLists_', new Error('No data from fetch for Neg KW Lists.'), def.name); return; }
  const data = gaqlRows.map(row => [
    row.sharedSet.id, row.sharedSet.name, row.sharedSet.status, row.sharedSet.type,
    safeParseInt_(row.sharedSet.memberCount)
  ]);
  writeDataToRawSheet_(ss, def.name, def.headers, data, currencyCode);
}

function writeAccountConfigSheet_(ss, customerData, currencyCode) {
  const sheet = ensureSheet_(ss, SETTINGS_CONFIG_TAB_NAME, false);
  if (!sheet) { logError_('writeAccountConfigSheet_', new Error('Failed to ensure Settings & Config sheet.')); return; }
  applyGeneralSheetFormatting_(sheet);
  let currentRow = 1;
  currentRow = createSectionHeader_(sheet, currentRow, 'Account & Script Configuration', 2);

  const data = [
    ['Google Ads Account Name', customerData.descriptiveName],
    ['Google Ads Customer ID', AdsApp.currentAccount().getCustomerId()],
    ['Currency Code', currencyCode],
    ['Time Zone', customerData.timeZone],
    ['Report Start Date', REPORT_START_DATE],
    ['Report End Date', REPORT_END_DATE],
    ['Script Last Run', Utilities.formatDate(new Date(), customerData.timeZone, 'yyyy-MM-dd HH:mm:ss z')],
    ['Script Version', SCRIPT_VERSION]
  ];
  const dataRange = sheet.getRange(currentRow, 1, data.length, data[0].length);
  dataRange.setValues(data);
  dataRange.setFontFamily('Arial').setFontSize(10);
  sheet.getRange(currentRow, 1, data.length, 1).setFontWeight('bold');
  sheet.getRange(currentRow, 2, data.length, 1).setHorizontalAlignment('left');

  autoResizeSheetColumns_(sheet, data[0].length);
  logInfo_(`Settings & Config sheet populated.`);
}
// --- End RAW_DATA_SHEET_PROCESSING ---

// --- Section 7: ANALYTICAL_DASHBOARD_MODULES ---

function populateShareRankTab_(ss, currencyCode) {
  if (!ENABLE_IMPRESSION_SHARE_TAB) {
    logInfo_('ShareRank tab generation is disabled via Config.');
    return;
  }
  const sheetName = SHARE_RANK_TAB_NAME;
  const analysisSheet = ensureSheet_(ss, sheetName, false);
  if (!analysisSheet) { logError_('populateShareRankTab_', new Error(`Failed to ensure sheet: ${sheetName}`)); return; }
  applyGeneralSheetFormatting_(analysisSheet);
  let currentRow = 1;
  currentRow = createSectionHeader_(analysisSheet, currentRow, 'Impression Share & Rank Analysis', 11);


  const rawISSheetName = TAB_PREFIX_RAW + 'Impression Share & Rank';
  const rawISSheet = ss.getSheetByName(rawISSheetName);
  const rawCampaignSheetName = TAB_PREFIX_RAW + 'Campaigns';
  const rawCampaignSheet = ss.getSheetByName(rawCampaignSheetName);


  if (!rawISSheet || rawISSheet.getLastRow() < 2) {
    logError_('populateShareRankTab_', new Error(`Raw data sheet '${rawISSheetName}' is empty or missing.`));
    analysisSheet.getRange(currentRow, 1).setValue(`Data source '${rawISSheetName}' is empty or missing.`);
    return;
  }
  if (!rawCampaignSheet || rawCampaignSheet.getLastRow() < 2) {
    logError_('populateShareRankTab_', new Error(`Raw data sheet '${rawCampaignSheetName}' is empty or missing for budget/cost info.`));
    analysisSheet.getRange(currentRow, 1).setValue(`Data source '${rawCampaignSheetName}' is empty or missing.`);
    return;
  }

  const isHeaders = rawISSheet.getRange(1, 1, 1, rawISSheet.getLastColumn()).getValues()[0];
  const isDataRows = rawISSheet.getRange(2, 1, rawISSheet.getLastRow() - 1, isHeaders.length).getValues();

  const campaignHeaders = rawCampaignSheet.getRange(1, 1, 1, rawCampaignSheet.getLastColumn()).getValues()[0];
  const campaignDataRows = rawCampaignSheet.getRange(2, 1, rawCampaignSheet.getLastRow() - 1, campaignHeaders.length).getValues();

  const campaignMetricsMap = {};
  const campIdIdx_rh = campaignHeaders.indexOf('Campaign ID');
  const campBudgetIdx_rh = campaignHeaders.indexOf('Budget');
  const campCostIdx_rh = campaignHeaders.indexOf('Cost');

  campaignDataRows.forEach(row => {
    const budgetVal = row[campBudgetIdx_rh];
    let dailyBudget = 0;
    if (typeof budgetVal === 'number') {
      dailyBudget = budgetVal;
    } else if (typeof budgetVal === 'string' && !isNaN(parseFloat(String(budgetVal).replace(/[^0-9.-]+/g,"")))) {
      dailyBudget = parseFloat(String(budgetVal).replace(/[^0-9.-]+/g,""));
    } else {
      logVerbose_(`ShareRank: Budget for campaign ${row[campIdIdx_rh]} is non-numeric: "${budgetVal}". Pacing will be N/A.`);
    }
    campaignMetricsMap[row[campIdIdx_rh]] = {
      cost: safeParseFloat_(row[campCostIdx_rh]),
      budget: dailyBudget
    };
  });

  const aggregatedIS = {};
  const dateIdx_is = isHeaders.indexOf('Date');
  const campIdIdx_is = isHeaders.indexOf('Campaign ID');
  const campNameIdx_is = isHeaders.indexOf('Campaign Name');
  const searchISIdx_is = isHeaders.indexOf('Search IS');
  const topISIdx_is = isHeaders.indexOf('Top IS');
  const absTopISIdx_is = isHeaders.indexOf('Abs Top IS');
  const rankLostISIdx_is = isHeaders.indexOf('IS Lost (Rank)');
  const budgetLostISIdx_is = isHeaders.indexOf('IS Lost (Budget)');


  isDataRows.forEach(row => {
    const campaignId = row[campIdIdx_is];
    const currentDateValue = row[dateIdx_is];
    const currentDateTime = (currentDateValue instanceof Date) ? currentDateValue.getTime() : new Date(currentDateValue).getTime();

    let aggregatedDate = 0;
    if (aggregatedIS[campaignId] && aggregatedIS[campaignId].date) {
      const aggDateValue = aggregatedIS[campaignId].date;
      aggregatedDate = (aggDateValue instanceof Date) ? aggDateValue.getTime() : new Date(aggDateValue).getTime();
    }

    if (!aggregatedIS[campaignId] || currentDateTime > aggregatedDate) {
      aggregatedIS[campaignId] = {
        date: currentDateValue,
        name: row[campNameIdx_is],
        searchIS: safeParseFloat_(row[searchISIdx_is]),
        topIS: safeParseFloat_(row[topISIdx_is]),
        absTopIS: safeParseFloat_(row[absTopISIdx_is]),
        rankLostIS: safeParseFloat_(row[rankLostISIdx_is]),
        budgetLostIS: safeParseFloat_(row[budgetLostISIdx_is])
      };
    }
  });

  const outputHeaders = ['Campaign ID', 'Campaign Name', 'Search IS', 'Top IS', 'Abs. Top IS', 'IS Lost (Rank)', 'IS Lost (Budget)', 'Cost (Period)', 'Daily Budget', 'Pacing Status', 'Recommendation'];
  const outputData = [];
  const daysInPeriod = getDaysBetweenDates_(new Date(REPORT_START_DATE), new Date(REPORT_END_DATE)) + 1;

  let totalCost = 0;
  const totals = {
      cost: 0,
      searchIS_weighted_numerator: 0,
      topIS_weighted_numerator: 0,
      absTopIS_weighted_numerator: 0,
      rankLostIS_weighted_numerator: 0,
      budgetLostIS_weighted_numerator: 0,
      totalImpressions_for_IS_calc: 0
  };


  for (const campId in aggregatedIS) {
    const isEntry = aggregatedIS[campId];
    const campMetrics = campaignMetricsMap[campId] || {
      cost: 0,
      budget: 0
    };

    let pacingStatusText = 'N/A (Budget Info Missing)';
    let recommendation = [];

    if (isEntry.budgetLostIS > 0.20) recommendation.push('High IS Lost (Budget) - Consider Budget Increase.');
    if (isEntry.rankLostIS > 0.20) recommendation.push('High IS Lost (Rank) - Review Bids/QS.');

    if (campMetrics.budget > 0 && daysInPeriod > 0) {
      const avgDailySpend = campMetrics.cost / daysInPeriod;
      const paceRatio = avgDailySpend / campMetrics.budget;
      if (paceRatio > BUDGET_PACE_HIGH_THRESHOLD) {
        pacingStatusText = `Over Paced (${(paceRatio * 100).toFixed(0)}%)`;
        recommendation.push('Over Pacing.');
      } else if (paceRatio < BUDGET_PACE_LOW_THRESHOLD && paceRatio > 0) {
        pacingStatusText = `Under Paced (${(paceRatio * 100).toFixed(0)}%)`;
        recommendation.push('Under Pacing.');
      } else if (paceRatio > 0) {
        pacingStatusText = `On Track (${(paceRatio * 100).toFixed(0)}%)`;
      } else {
        pacingStatusText = 'On Track (No Spend or Budget)';
      }
    } else if (campMetrics.cost > 0) {
      pacingStatusText = 'Spending w/o Numeric Budget';
      recommendation.push('Spending without defined numeric daily budget for pacing.');
    }

    if (recommendation.length === 0 && isEntry.searchIS > 0.70) recommendation.push('IS Healthy.');
    else if (recommendation.length === 0) recommendation.push('Monitor IS.');

    outputData.push([
      campId, isEntry.name,
      isEntry.searchIS, isEntry.topIS, isEntry.absTopIS,
      isEntry.rankLostIS, isEntry.budgetLostIS,
      campMetrics.cost, campMetrics.budget > 0 ? campMetrics.budget : 'N/A',
      pacingStatusText,
      recommendation.join(' ')
    ]);
    totals.cost += campMetrics.cost;
    if(campMetrics.cost > 0) {
        totals.searchIS_weighted_numerator += isEntry.searchIS * campMetrics.cost;
        totals.topIS_weighted_numerator += isEntry.topIS * campMetrics.cost;
        totals.absTopIS_weighted_numerator += isEntry.absTopIS * campMetrics.cost;
        totals.rankLostIS_weighted_numerator += isEntry.rankLostIS * campMetrics.cost;
        totals.budgetLostIS_weighted_numerator += isEntry.budgetLostIS * campMetrics.cost;
        totals.totalImpressions_for_IS_calc += campMetrics.cost;
    }
  }

  outputData.sort((a, b) => b[outputHeaders.indexOf('Search IS')] - a[outputHeaders.indexOf('Search IS')]);

  if (outputData.length > 0) {
      const totalRow = ['TOTAL', '',
          totals.totalImpressions_for_IS_calc > 0 ? totals.searchIS_weighted_numerator / totals.totalImpressions_for_IS_calc : 0,
          totals.totalImpressions_for_IS_calc > 0 ? totals.topIS_weighted_numerator / totals.totalImpressions_for_IS_calc : 0,
          totals.totalImpressions_for_IS_calc > 0 ? totals.absTopIS_weighted_numerator / totals.totalImpressions_for_IS_calc : 0,
          totals.totalImpressions_for_IS_calc > 0 ? totals.rankLostIS_weighted_numerator / totals.totalImpressions_for_IS_calc : 0,
          totals.totalImpressions_for_IS_calc > 0 ? totals.budgetLostIS_weighted_numerator / totals.totalImpressions_for_IS_calc : 0,
          totals.cost,
          '', '', ''
      ];
      outputData.push(totalRow);
  }


  const analysisSheetObj = ss.getSheetByName(sheetName);
  if (analysisSheetObj) {
    analysisSheetObj.getRange(currentRow, 1, 1, outputHeaders.length).setValues([outputHeaders]);
    if (outputData && outputData.length > 0) {
      const dataStartRow = currentRow + 1;
      analysisSheetObj.getRange(dataStartRow, 1, outputData.length, outputHeaders.length).setValues(outputData);
      const columnFormatStrings = outputHeaders.map(header => getColumnFormatForHeader_(header, currencyCode));
      const formatsArray = [];
      for (let r = 0; r < outputData.length; r++) formatsArray.push(columnFormatStrings);
      applyChunkedFormatting_(analysisSheetObj, formatsArray, dataStartRow);
      applyTableStyles(analysisSheetObj, currentRow, outputData.length + 1, outputHeaders.length);

      if (outputData.length > 0) {
          const totalRowRange = analysisSheetObj.getRange(dataStartRow + outputData.length -1, 1, 1, outputHeaders.length);
          totalRowRange.setBackground(STYLE_TOTALS_ROW.backgroundColor)
                       .setFontWeight(STYLE_TOTALS_ROW.fontWeight)
                       .setFontColor(STYLE_TOTALS_ROW.fontColor)
                       .setBorder(true, null, true, null, null, null, "#212121", SpreadsheetApp.BorderStyle.SOLID_MEDIUM); // Distinct border for totals
      }

    } else {
        analysisSheetObj.getRange(currentRow + 1, 1).setValue("No data to display for Impression Share.");
    }
    autoResizeSheetColumns_(analysisSheetObj, outputHeaders.length);
  }
  logInfo_(`ShareRank tab '${sheetName}' populated with ${outputData.length} campaigns.`);
}


function populateBudgetPacingTab_(ss, currencyCode) {
  if (!ENABLE_BUDGET_PACING_TAB) {
    logInfo_('Budget Pacing tab generation is disabled via Config.');
    return;
  }
  const sheetName = BUDGET_PACING_TAB_NAME;
  const analysisSheet = ensureSheet_(ss, sheetName, false);
  if (!analysisSheet) { logError_('populateBudgetPacingTab_', new Error(`Failed to ensure sheet: ${sheetName}`)); return; }
  applyGeneralSheetFormatting_(analysisSheet);
  let currentRow = 1;
  currentRow = createSectionHeader_(analysisSheet, currentRow, 'Campaign Budget Pacing (Month-to-Date)', 10);


  const rawCampaignSheetName = TAB_PREFIX_RAW + 'Campaigns';
  const rawCampaignSheet = ss.getSheetByName(rawCampaignSheetName);

  const rawDailySheetName = TAB_PREFIX_RAW + 'Daily Performance';
  const rawDailySheet = ss.getSheetByName(rawDailySheetName);

  if (!rawCampaignSheet || rawCampaignSheet.getLastRow() < 2) {
    logError_('populateBudgetPacingTab_', new Error(`Raw data sheet '${rawCampaignSheetName}' is empty or missing.`));
    analysisSheet.getRange(currentRow,1).setValue(`Data source '${rawCampaignSheetName}' is empty or missing.`);
    return;
  }
  if (!rawDailySheet || rawDailySheet.getLastRow() < 2) {
    logError_('populateBudgetPacingTab_', new Error(`Raw data sheet '${rawDailySheetName}' is empty or missing for MTD cost.`));
    analysisSheet.getRange(currentRow,1).setValue(`Data source '${rawDailySheetName}' is empty or missing for MTD cost.`);
    return;
  }

  const dailyHeaders = rawDailySheet.getRange(1, 1, 1, rawDailySheet.getLastColumn()).getValues()[0];
  const dailyDataRows = rawDailySheet.getRange(2, 1, rawDailySheet.getLastRow() - 1, dailyHeaders.length).getValues();

  const dateIdx_daily = dailyHeaders.indexOf('Date');
  const costIdx_daily = dailyHeaders.indexOf('Cost');
  const campIdIdx_daily = dailyHeaders.indexOf('Campaign ID');

  if ([dateIdx_daily, costIdx_daily, campIdIdx_daily].includes(-1)) {
    logError_('populateBudgetPacingTab_', new Error('Required columns missing in Raw Daily Performance.'));
    analysisSheet.getRange(currentRow,1).setValue('Error: Columns missing in Raw Daily Performance for MTD cost.');
    return;
  }

  const today = new Date(getTodayDateString_());
  const currentMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const campaignMtdCosts = {};

  // FIX: Define daysElapsedInMonth, etc. outside the loop
  const daysInCurrentMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const daysElapsedInMonth = today.getDate();
  const daysLeftInMonth = daysInCurrentMonth - daysElapsedInMonth;


  dailyDataRows.forEach(row => {
    const rowDate = new Date(row[dateIdx_daily]);
    const campaignId = row[campIdIdx_daily];
    if (rowDate >= currentMonthStart && rowDate <= today) {
      if (!campaignMtdCosts[campaignId]) {
        campaignMtdCosts[campaignId] = 0;
      }
      campaignMtdCosts[campaignId] += safeParseFloat_(row[costIdx_daily]);
    }
  });
  logInfo_(`Calculated MTD costs for ${Object.keys(campaignMtdCosts).length} campaigns for Budget Pacing.`);

  const campaignHeaders = rawCampaignSheet.getRange(1, 1, 1, rawCampaignSheet.getLastColumn()).getValues()[0];
  const campaignDataRows = rawCampaignSheet.getRange(2, 1, rawCampaignSheet.getLastRow() - 1, campaignHeaders.length).getValues();

  const campIdIdx_rh = campaignHeaders.indexOf('Campaign ID');
  const campNameIdx_rh = campaignHeaders.indexOf('Campaign Name');
  const campStatusIdx_rh = campaignHeaders.indexOf('Status');
  const campBudgetIdx_rh = campaignHeaders.indexOf('Budget');

  const outputHeaders = ['Campaign ID', 'Campaign Name', 'Campaign Status', 'Daily Budget', 'MTD Cost', 'Pacing % (MTD)', 'Status', 'Days Left in Month', 'Est. Month End Spend', 'Recommendation'];
  const outputData = [];
  let totalDailyBudget = 0;
  let totalMtdCost = 0;
  let totalEstMonthEndSpend = 0;


  campaignDataRows.forEach(row => {
    const campaignId = row[campIdIdx_rh];
    const campaignName = row[campNameIdx_rh];
    const campaignStatus = row[campStatusIdx_rh];
    const budgetVal = row[campBudgetIdx_rh];
    let dailyBudget = 0;

    if (typeof budgetVal === 'number') {
      dailyBudget = budgetVal;
    } else if (typeof budgetVal === 'string' && !isNaN(parseFloat(String(budgetVal).replace(/[^0-9.-]+/g,"")))) {
      dailyBudget = parseFloat(String(budgetVal).replace(/[^0-9.-]+/g,""));
    } else {
      logVerbose_(`Budget for campaign ${campaignName} (${campaignId}) is non-numeric: "${budgetVal}". Pacing may be N/A.`);
    }

    const mtdCost = campaignMtdCosts[campaignId] || 0;

    let pacingPercent = 0;
    let status = 'N/A (Budget Unknown/Zero)';
    let recommendation = 'Ensure numeric daily budget is available for accurate pacing.';
    let estimatedMonthEndSpend = mtdCost;

    if (campaignStatus === 'PAUSED') {
      status = 'N/A (Paused)';
      pacingPercent = 0;
      recommendation = 'Campaign is paused. No pacing calculation applicable.';
    } else if (dailyBudget > 0) {
      const expectedMtdSpend = dailyBudget * daysElapsedInMonth;
      pacingPercent = expectedMtdSpend > 0 ? (mtdCost / expectedMtdSpend) : (mtdCost > 0 ? Infinity : 0);
      status = 'On Track';
      recommendation = 'Maintain current pacing.';
      if (pacingPercent > BUDGET_PACE_HIGH_THRESHOLD) {
        status = `Over Paced (${(pacingPercent * 100).toFixed(0)}%)`;
        recommendation = 'Review: Significantly Over Pacing. Consider budget/bid adjustments.';
      } else if (pacingPercent < BUDGET_PACE_LOW_THRESHOLD) {
        status = `Under Paced (${(pacingPercent * 100).toFixed(0)}%)`;
        recommendation = 'Review: Significantly Under Pacing. Consider increasing activity.';
      }
      estimatedMonthEndSpend = dailyBudget * daysInCurrentMonth;
    } else if (mtdCost > 0) {
      status = 'Spending, Budget N/A or Zero';
      recommendation = 'Campaign is spending but daily budget is not numerically defined or is zero.';
      estimatedMonthEndSpend = (daysElapsedInMonth > 0 ? (mtdCost / daysElapsedInMonth) * daysInCurrentMonth : mtdCost);
    }

    if (campaignStatus !== 'PAUSED') {
        totalDailyBudget += dailyBudget;
        totalMtdCost += mtdCost;
        totalEstMonthEndSpend += estimatedMonthEndSpend;
    }

    outputData.push([
      campaignId, campaignName, campaignStatus,
      dailyBudget > 0 ? dailyBudget : 'N/A',
      mtdCost,
      (campaignStatus !== 'PAUSED' && dailyBudget > 0) ? pacingPercent : 'N/A',
      status, daysLeftInMonth, estimatedMonthEndSpend, recommendation
    ]);
  });

  if (outputData.length > 0) {
      const overallPacing = totalDailyBudget > 0 && daysElapsedInMonth > 0 ? (totalMtdCost / (totalDailyBudget * daysElapsedInMonth)) : 0;
      let overallStatus = 'N/A';
      if (totalDailyBudget > 0) {
          if (overallPacing > BUDGET_PACE_HIGH_THRESHOLD) overallStatus = `Over Paced (${(overallPacing * 100).toFixed(0)}%)`;
          else if (overallPacing < BUDGET_PACE_LOW_THRESHOLD) overallStatus = `Under Paced (${(overallPacing * 100).toFixed(0)}%)`;
          else overallStatus = `On Track (${(overallPacing * 100).toFixed(0)}%)`;
      }

      const totalRow = ['TOTAL', '', '', totalDailyBudget, totalMtdCost, overallPacing, overallStatus, daysLeftInMonth, totalEstMonthEndSpend, ''];
      outputData.push(totalRow);
  }


  const analysisSheetObj = ss.getSheetByName(sheetName);
  if (analysisSheetObj) {
    analysisSheetObj.getRange(currentRow, 1, 1, outputHeaders.length).setValues([outputHeaders]);
    if (outputData && outputData.length > 0) {
      const dataStartRow = currentRow + 1;
      analysisSheetObj.getRange(dataStartRow, 1, outputData.length, outputHeaders.length).setValues(outputData);
      const columnFormatStrings = outputHeaders.map(header => getColumnFormatForHeader_(header, currencyCode));
      const formatsArray = [];
      for (let r = 0; r < outputData.length; r++) formatsArray.push(columnFormatStrings);
      applyChunkedFormatting_(analysisSheetObj, formatsArray, dataStartRow);
      applyTableStyles(analysisSheetObj, currentRow, outputData.length + 1, outputHeaders.length);

      if (outputData.length > 0) {
          const totalRowRange = analysisSheetObj.getRange(dataStartRow + outputData.length -1, 1, 1, outputHeaders.length);
          totalRowRange.setBackground(STYLE_TOTALS_ROW.backgroundColor)
                       .setFontWeight(STYLE_TOTALS_ROW.fontWeight)
                       .setFontColor(STYLE_TOTALS_ROW.fontColor);
          totalRowRange.setBorder(true,null,true,null,null,null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }

    } else {
        analysisSheetObj.getRange(currentRow + 1, 1).setValue("No data to display for Budget Pacing.");
    }
    autoResizeSheetColumns_(analysisSheetObj, outputHeaders.length);
  }
  logInfo_(`Budget Pacing tab '${sheetName}' populated with ${outputData.length} campaigns.`);
}


function populateConversionLagTabs_(ss, currencyCode) {
  if (!ENABLE_CONV_LAG_ADJUST_TABS) {
    logInfo_('Conversion Lag Adjustment tabs generation is disabled via Config.');
    return;
  }
  const factorsSheetName = LAG_FACTORS_TAB_NAME;
  const forecastSheetName = LAG_FORECAST_TAB_NAME;

  const factorsSheet = ensureSheet_(ss, factorsSheetName, true);
  const forecastSheet = ensureSheet_(ss, forecastSheetName, false);

  if (!factorsSheet || !forecastSheet) {
    logError_('populateConversionLagTabs_', new Error('Failed to ensure one or more Conversion Lag sheets.'));
    return;
  }
  applyGeneralSheetFormatting_(factorsSheet);
  applyGeneralSheetFormatting_(forecastSheet);
  let forecastCurrentRow = 1;
  forecastCurrentRow = createSectionHeader_(forecastSheet, forecastCurrentRow, 'Lag-Adjusted Conversion Forecast', HDR_LAG_FORECAST.length);
  let factorsCurrentRow = 1;
  factorsCurrentRow = createSectionHeader_(factorsSheet, factorsCurrentRow, 'Conversion Lag Factors (Historical)', 5);


  const timeZone = AdsApp.currentAccount().getTimeZone();
  const lagLearnStartDate = Utilities.formatDate(addDaysToDate_(new Date(getTodayDateString_()), -CONV_LAG_LOOKBACK_DAYS - CONV_LAG_STABLE_PERIOD_OFFSET), timeZone, 'yyyy-MM-dd');
  const lagLearnEndDate = Utilities.formatDate(addDaysToDate_(new Date(getTodayDateString_()), -CONV_LAG_STABLE_PERIOD_OFFSET - 1), timeZone, 'yyyy-MM-dd');

  const lagDataQuery = `
    SELECT
      segments.conversion_lag_bucket,
      metrics.conversions_by_conversion_date,
      metrics.conversions_value_by_conversion_date
    FROM customer
    WHERE segments.date BETWEEN '${lagLearnStartDate}' AND '${lagLearnEndDate}'`;

  const lagRows = executeGAQLQuery_(lagDataQuery, 'ConversionLagLearning');
  if (!lagRows || lagRows.length === 0) {
    logError_('populateConversionLagTabs_', new Error('No data returned for lag learning.'));
    factorsSheet.getRange(factorsCurrentRow, 1).setValue('No data available to calculate conversion lag factors.');
    forecastSheet.getRange(forecastCurrentRow, 1).setValue('Lag factors not available for forecast.');
    return;
  }

  const lagBuckets = {};
  lagRows.forEach(row => {
    const bucket = row.segments.conversionLagBucket;
    if (!lagBuckets[bucket]) {
      lagBuckets[bucket] = {
        conversions: 0,
        value: 0
      };
    }
    lagBuckets[bucket].conversions += safeParseFloat_(row.metrics.conversionsByConversionDate);
    lagBuckets[bucket].value += safeParseFloat_(row.metrics.conversionsValueByConversionDate);
  });

  const totalLagConversions = Object.values(lagBuckets).reduce((sum, b) => sum + b.conversions, 0);

  const lagFactorsTable = [];
  let cumulativeConversionShare = 0;

  const BUCKET_ORDER = [
    "LESS_THAN_ONE_DAY", "ONE_TO_TWO_DAYS", "TWO_TO_THREE_DAYS", "THREE_TO_FOUR_DAYS",
    "FOUR_TO_FIVE_DAYS", "FIVE_TO_SIX_DAYS", "SIX_TO_SEVEN_DAYS", "SEVEN_TO_EIGHT_DAYS",
    "EIGHT_TO_NINE_DAYS", "NINE_TO_TEN_DAYS", "TEN_TO_ELEVEN_DAYS", "ELEVEN_TO_TWELVE_DAYS",
    "TWELVE_TO_THIRTEEN_DAYS", "THIRTEEN_TO_FOURTEEN_DAYS", "FOURTEEN_TO_TWENTY_ONE_DAYS",
    "TWENTY_ONE_TO_THIRTY_DAYS", "THIRTY_TO_FORTY_FIVE_DAYS", "FORTY_FIVE_TO_SIXTY_DAYS",
  ];

  BUCKET_ORDER.forEach(bucketKey => {
    if (lagBuckets[bucketKey]) {
      const convShare = totalLagConversions > 0 ? lagBuckets[bucketKey].conversions / totalLagConversions : 0;
      cumulativeConversionShare += convShare;
      const upliftFactor = cumulativeConversionShare > 0 ? 1 / cumulativeConversionShare : 1;
      lagFactorsTable.push([
        bucketKey,
        lagBuckets[bucketKey].conversions,
        convShare,
        cumulativeConversionShare,
        upliftFactor
      ]);
    }
  });

  const factorsHeaders = ['Lag Bucket', 'Historical Conversions (in bucket)', 'Conversion Share (of total lagged)', 'Cumulative Conv. Share', 'Uplift Factor'];
  const factorsSheetObj = ss.getSheetByName(factorsSheetName);
  if (factorsSheetObj) {
    factorsSheetObj.getRange(factorsCurrentRow, 1, 1, factorsHeaders.length).setValues([factorsHeaders]);
    if (lagFactorsTable && lagFactorsTable.length > 0) {
      factorsSheetObj.getRange(factorsCurrentRow + 1, 1, lagFactorsTable.length, factorsHeaders.length).setValues(lagFactorsTable);
      const columnFormatStrings = factorsHeaders.map(header => getColumnFormatForHeader_(header, currencyCode));
      const formatsArray = [];
      for (let r = 0; r < lagFactorsTable.length; r++) formatsArray.push(columnFormatStrings);
      applyChunkedFormatting_(factorsSheetObj, formatsArray, factorsCurrentRow + 1);
      applyTableStyles(factorsSheetObj, factorsCurrentRow, lagFactorsTable.length + 1, factorsHeaders.length);
    } else {
        factorsSheetObj.getRange(factorsCurrentRow + 1, 1).setValue("No lag factor data to display.");
    }
    autoResizeSheetColumns_(factorsSheetObj, factorsHeaders.length);
  }

  const forecastStartDate = Utilities.formatDate(addDaysToDate_(new Date(getTodayDateString_()), -CONV_LAG_FORECAST_DAYS + 1), timeZone, 'yyyy-MM-dd');
  const forecastEndDate = getTodayDateString_();

  const recentPerfQuery = `
    SELECT segments.date, metrics.conversions, metrics.conversions_value
    FROM customer 
    WHERE segments.date BETWEEN '${forecastStartDate}' AND '${forecastEndDate}'
    ORDER BY segments.date ASC`;

  const recentPerfRows = executeGAQLQuery_(recentPerfQuery, 'RecentPerformanceForLagForecast');
  const forecastTable = [];
  // FIX: Use the reordered HDR_LAG_FORECAST
  const forecastHeaders = ['Applied Lag Bucket', 'Date', 'Reported Conversions', 'Reported Conv. Value', 'Days Ago', 'Calculated Uplift Factor', 'Adjusted Conversions', 'Adjusted Conv. Value'];


  const todayForDiff = new Date(getTodayDateString_());

  if (recentPerfRows) {
    recentPerfRows.forEach(row => {
      const reportDate = new Date(row.segments.date);
      const reportDateInAccountTZ = new Date(Utilities.formatDate(reportDate, timeZone, "yyyy-MM-dd'T00:00:00'"));
      const daysAgo = getDaysBetweenDates_(reportDateInAccountTZ, todayForDiff);

      let upliftFactor = 1;
      let cumulativeShareForDay = 0;
      let appliedLagBucket = "N/A (Beyond Max Lag)";

      for (const factorEntry of lagFactorsTable) {
        const bucketKey = factorEntry[0];
        const bucketMaxDays = getBucketMaxDays_(bucketKey);
        cumulativeShareForDay = factorEntry[3];
        if (daysAgo <= bucketMaxDays) {
          appliedLagBucket = bucketKey;
          break;
        }
      }
      upliftFactor = cumulativeShareForDay > 0 ? 1 / cumulativeShareForDay : 1;
      upliftFactor = Math.min(upliftFactor, 5.0);

      const reportedConversions = safeParseFloat_(row.metrics.conversions);
      const reportedConvValue = safeParseFloat_(row.metrics.conversionsValue);

      // FIX: Reorder data to match new header order
      forecastTable.push([
        appliedLagBucket,
        row.segments.date,
        reportedConversions,
        reportedConvValue,
        daysAgo,
        upliftFactor,
        reportedConversions * upliftFactor,
        reportedConvValue * upliftFactor
      ]);
    });
  }

  const forecastSheetObj = ss.getSheetByName(forecastSheetName);
  if (forecastSheetObj) {
    forecastSheetObj.getRange(forecastCurrentRow, 1, 1, forecastHeaders.length).setValues([forecastHeaders]);
    const lagBucketHeaderCell = forecastSheetObj.getRange(forecastCurrentRow, 1); // Applied Lag Bucket is now first
    lagBucketHeaderCell.setNote('Describes the typical time lag between an ad interaction and a conversion. E.g., "LESS_THAN_ONE_DAY" means conversions typically happen within 24 hours of the interaction; "ONE_TO_TWO_DAYS" means 24-48 hours, etc.');


    if (forecastTable && forecastTable.length > 0) {
      forecastSheetObj.getRange(forecastCurrentRow + 1, 1, forecastTable.length, forecastHeaders.length).setValues(forecastTable);
      const columnFormatStrings = forecastHeaders.map(header => getColumnFormatForHeader_(header, currencyCode));
      const formatsArray = [];
      for (let r = 0; r < forecastTable.length; r++) formatsArray.push(columnFormatStrings);
      applyChunkedFormatting_(forecastSheetObj, formatsArray, forecastCurrentRow + 1);
      applyTableStyles(forecastSheetObj, forecastCurrentRow, forecastTable.length + 1, forecastHeaders.length);
    } else {
        forecastSheetObj.getRange(forecastCurrentRow + 1, 1).setValue("No forecast data to display.");
    }
    autoResizeSheetColumns_(forecastSheetObj, forecastHeaders.length);
  }
  logInfo_(`Conversion Lag tabs populated. Factors: ${lagFactorsTable.length} buckets. Forecast: ${forecastTable.length} days.`);
}

function getBucketMaxDays_(bucketString) {
  if (!bucketString) return 999;
  if (bucketString === "LESS_THAN_ONE_DAY") return 0;

  const dayMap = {
    "ONE": 1, "TWO": 2, "THREE": 3, "FOUR": 4, "FIVE": 5, "SIX": 6, "SEVEN": 7,
    "EIGHT": 8, "NINE": 9, "TEN": 10, "ELEVEN": 11, "TWELVE": 12, "THIRTEEN": 13,
    "FOURTEEN": 14, "TWENTY_ONE": 21, "THIRTY": 30, "FORTY_FIVE": 45,
    "SIXTY": 60, "NINETY": 90
  };

  const parts = bucketString.split('_TO_');
  if (parts.length === 2) {
    const endPart = parts[1].replace('_DAYS', '');
    return dayMap[endPart.toUpperCase()] !== undefined ? dayMap[endPart.toUpperCase()] : 999;
  }
  if (bucketString.startsWith("MORE_THAN_")) return 999;

  return 999;
}
// --- End ANALYTICAL_DASHBOARD_MODULES ---

// --- Section 8: CORE_DASHBOARD_CONSTRUCTION ---
function populateUserGuideTab_(ss) {
  const sheetName = USER_GUIDE_TAB_NAME;
  const sheet = ensureSheet_(ss, sheetName, false);
  if (!sheet) { logError_('populateUserGuideTab_', new Error(`Failed to ensure sheet: ${sheetName}`)); return; }

  applyGeneralSheetFormatting_(sheet);
  let currentRow = 1;
  currentRow = createSectionHeader_(sheet, currentRow, 'Dashboard User Guide & Notes', 2);

  sheet.getRange(currentRow, 1).setValue('Welcome to your Google Ads E-commerce Dashboard! This document provides an overview of key performance indicators for your account.').setWrap(true);
  currentRow += 2;

  sheet.getRange(currentRow, 1).setValue('Key Dashboard Tabs:').setFontWeight('bold');
  currentRow++;

  const tabDescriptions = [
    [USER_GUIDE_TAB_NAME, 'This guide.'],
    [SETTINGS_CONFIG_TAB_NAME, 'Displays current script settings and account information.'],
    [OVERVIEW_TAB_NAME, 'High-level account performance KPIs, daily trends, and device performance charts.'],
    [CAMPAIGN_PERF_TAB_NAME, 'Detailed metrics for all active campaigns. Uses QUERY from raw data.'],
    [PRODUCT_ANALYSIS_TAB_NAME, 'Aggregated product data with e-commerce specific KPIs and rule-based recommendations.'],
  ];
  if (ENABLE_ASSET_PERFORMANCE_TAB) {
    tabDescriptions.push([ASSET_PERFORMANCE_TAB_NAME, 'Performance data for individual assets.']);
  }
  if (ENABLE_IMPRESSION_SHARE_TAB) {
    tabDescriptions.push([SHARE_RANK_TAB_NAME, 'Campaign-level impression share, top impression share, and share lost due to budget or rank.']);
  }
  if (ENABLE_BUDGET_PACING_TAB) {
    tabDescriptions.push([BUDGET_PACING_TAB_NAME, 'Monitors campaign budgets against Month-To-Date (MTD) spend and flags pacing issues.']);
  }
  if (ENABLE_CONV_LAG_ADJUST_TABS) {
    tabDescriptions.push([LAG_FORECAST_TAB_NAME, `Estimated and forecasted conversion numbers adjusted for typical lag. Forecast window: ${CONV_LAG_FORECAST_DAYS} days.`]);
    tabDescriptions.push([LAG_FACTORS_TAB_NAME, 'Historical conversion lag factors used for adjustments. (Usually hidden)']);
  }
  tabDescriptions.push([RECOMMENDATIONS_TAB_NAME, 'Automated, rule-based suggestions based on performance observed in other tabs.']);

  const outputTabLinks = [];
  tabDescriptions.forEach(descPair => {
      const tabNameToLink = descPair[0];
      const description = descPair[1];
      const targetSheet = ss.getSheetByName(tabNameToLink);
      if (targetSheet) {
          const sheetGid = targetSheet.getSheetId();
          outputTabLinks.push([`=HYPERLINK("#gid=${sheetGid}","${tabNameToLink}")`, description]);
      } else {
          outputTabLinks.push([tabNameToLink + " (Sheet not found)", description]);
      }
  });

  if (outputTabLinks.length > 0) {
    sheet.getRange(currentRow, 1, outputTabLinks.length, 2).setValues(outputTabLinks).setWrap(true);
    sheet.getRange(currentRow, 1, outputTabLinks.length, 1).setFontWeight('bold');
    currentRow += outputTabLinks.length + 2;
  }


  sheet.getRange(currentRow, 1).setValue('Configuration:').setFontWeight('bold');
  currentRow++;
  sheet.getRange(currentRow, 1).setValue(`This script operates on data between ${REPORT_START_DATE} and ${REPORT_END_DATE}. These dates are set in the SCRIPT_CONFIGURATION section of the script code.`).setWrap(true);
  currentRow++;
  sheet.getRange(currentRow, 1).setValue('Feature toggles (e.g., enabling specific analytical tabs) are also managed within the SCRIPT_CONFIGURATION section.').setWrap(true);
  currentRow += 2;

  sheet.getRange(currentRow, 1).setValue('Error Log:').setFontWeight('bold');
  currentRow++;
  const errorSheetGid = ss.getSheetByName(ERROR_SHEET_NAME) ? ss.getSheetByName(ERROR_SHEET_NAME).getSheetId() : null;
  const errorText = `Any script errors encountered during execution are logged in the "${ERROR_SHEET_NAME}" tab in this spreadsheet, as well as in the Google Ads script logs.`;
  const errorSheetNameIndex = errorText.indexOf(ERROR_SHEET_NAME);

  const richText = SpreadsheetApp.newRichTextValue().setText(errorText);
  if (errorSheetGid && errorSheetNameIndex !== -1) {
      richText.setLinkUrl(errorSheetNameIndex, errorSheetNameIndex + ERROR_SHEET_NAME.length, `#gid=${errorSheetGid}`);
  }
  sheet.getRange(currentRow, 1).setRichTextValue(richText.build()).setWrap(true);
  currentRow += 2;

  sheet.getRange(currentRow, 1).setValue('Data Source:').setFontWeight('bold');
  currentRow++;
  sheet.getRange(currentRow, 1).setValue('All data is pulled directly from your Google Ads account using the Google Ads Query Language (GAQL). Raw data is stored in hidden sheets prefixed with "Raw Data - ".').setWrap(true);

  autoResizeSheetColumns_(sheet, 1);
  sheet.setColumnWidth(2, 450);
  logInfo_(`User Guide tab '${sheetName}' populated.`);
}


function populateOverviewDashboard_(ss, currencyCode) {
  const sheetName = OVERVIEW_TAB_NAME;
  const sheet = ensureSheet_(ss, sheetName, false);
  if (!sheet) { logError_('populateOverviewDashboard_', new Error(`Failed to ensure sheet: ${sheetName}`)); return; }

  applyGeneralSheetFormatting_(sheet);
  let currentRow = 1;
  sheet.getRange('A1').setValue('Overall Account Performance Overview (incl. AOV)').setFontSize(16).setFontWeight('bold').setFontFamily('Arial');
  sheet.getRange('A2').setValue(`Reporting Period: ${REPORT_START_DATE} to ${REPORT_END_DATE} (Last Updated: ${new Date().toLocaleString()})`)
    .setFontSize(9).setFontStyle('italic').setFontFamily('Arial');
  currentRow = 3;

  currentRow = createSectionHeader_(sheet, currentRow, 'Key Performance Indicators', 2);
  const kpiTableHeaderRow = currentRow;
  sheet.getRange(kpiTableHeaderRow, 1, 1, 2).setValues([['Metric', 'Value']]);
  currentRow++;

  const rawCampaignsSheetName = TAB_PREFIX_RAW + 'Campaigns';
  const rawCampaignSheet = ss.getSheetByName(rawCampaignsSheetName);
  if (!rawCampaignSheet || rawCampaignSheet.getLastRow() < 1) {
    logError_('populateOverviewDashboard_', new Error(`Raw Campaigns sheet '${rawCampaignsSheetName}' is empty or missing data.`));
    sheet.getRange(currentRow, 1).setValue('Campaign data not available for KPIs.');
    return;
  }
  if (rawCampaignSheet.getLastRow() < 2 && rawCampaignSheet.getRange(1,1).getValue() !== "") {
       logWarn_('populateOverviewDashboard_', `Raw Campaigns sheet '${rawCampaignsSheetName}' only has headers. KPIs will be zero or show Data Error.`);
  }

  const lastDataRow = Math.max(2, rawCampaignSheet.getLastRow());

  function getColLetter(headerName, actualHeaders) {
    const index = actualHeaders.indexOf(headerName);
    if (index === -1) {
      logError_('populateOverviewDashboard_', new Error(`Header '${headerName}' not found in ${rawCampaignsSheetName}.`), `Sheet: ${rawCampaignsSheetName}`);
      return "INVALID_COLUMN";
    }
    return String.fromCharCode(65 + index);
  }

  const campaignSheetHeaders = rawCampaignSheet.getRange(1, 1, 1, rawCampaignSheet.getLastColumn()).getValues()[0];

  const costCol = getColLetter('Cost', campaignSheetHeaders);
  const imprCol = getColLetter('Impr.', campaignSheetHeaders);
  const clicksCol = getColLetter('Clicks', campaignSheetHeaders);
  const convCol = getColLetter('Conv.', campaignSheetHeaders);
  const convValueCol = getColLetter('Conv. Value', campaignSheetHeaders);

  const kpiDefinitions = [
    { label: 'Total Cost', type: 'sum', col: costCol, formatType: 'currency' },
    { label: 'Total Impressions', type: 'sum', col: imprCol, formatType: 'integer' },
    { label: 'Total Clicks', type: 'sum', col: clicksCol, formatType: 'integer' },
    { label: 'Overall CTR', type: 'calc', dep: ['Total Clicks', 'Total Impressions'], formatType: 'percent' },
    { label: 'Total Conversions', type: 'sum', col: convCol, formatType: 'decimal' },
    { label: 'Total Conv. Value', type: 'sum', col: convValueCol, formatType: 'currency' },
    { label: 'Overall CPA', type: 'calc', dep: ['Total Cost', 'Total Conversions'], formatType: 'currency' },
    { label: 'Overall ROAS', type: 'calc', dep: ['Total Conv. Value', 'Total Cost'], formatType: 'decimal_2' },
    { label: 'Overall CVR', type: 'calc', dep: ['Total Conversions', 'Total Clicks'], formatType: 'percent' },
    { label: 'Overall AOV', type: 'calc', dep: ['Total Conv. Value', 'Total Conversions'], formatType: 'currency' }
  ];

  const kpiTableData = [];
  kpiDefinitions.forEach(kpi => kpiTableData.push([kpi.label, 'Loading...']));
  const kpiTableRange = sheet.getRange(currentRow, 1, kpiTableData.length, 2);
  kpiTableRange.setValues(kpiTableData);

  sheet.getRange(currentRow, 1, kpiTableData.length, 1).setFontWeight('bold').setHorizontalAlignment('left');
  sheet.getRange(currentRow, 2, kpiTableData.length, 1).setHorizontalAlignment('right');


  const kpiValueCellCol = 'B';
  kpiDefinitions.forEach((kpi, index) => {
    const targetCell = sheet.getRange(currentRow + index, 2);
    if (kpi.type === 'sum') {
      if (kpi.col !== "INVALID_COLUMN") {
        targetCell.setFormula(`=SUM('${rawCampaignsSheetName}'!${kpi.col}2:${kpi.col}${lastDataRow})`);
      } else {
        targetCell.setValue('Data Error');
      }
    } else if (kpi.type === 'calc') {
      let formula = "";
      let depsValid = true;
      const depCells = {};

      kpi.dep.forEach(depLabel => {
        const depIndex = kpiDefinitions.findIndex(d => d.label === depLabel);
        if (depIndex === -1 || (kpiDefinitions[depIndex].type === 'sum' && kpiDefinitions[depIndex].col === "INVALID_COLUMN")) {
          depsValid = false;
        }
        depCells[depLabel] = `${kpiValueCellCol}${currentRow + depIndex}`;
      });

      if (!depsValid) {
        targetCell.setValue('Data Error');
      } else {
        if (kpi.label === 'Overall CTR') formula = `=IFERROR(${depCells['Total Clicks']}/${depCells['Total Impressions']},0)`;
        else if (kpi.label === 'Overall CPA') formula = `=IFERROR(${depCells['Total Cost']}/${depCells['Total Conversions']},0)`;
        else if (kpi.label === 'Overall ROAS') formula = `=IFERROR(${depCells['Total Conv. Value']}/${depCells['Total Cost']},0)`;
        else if (kpi.label === 'Overall CVR') formula = `=IFERROR(${depCells['Total Conversions']}/${depCells['Total Clicks']},0)`;
        else if (kpi.label === 'Overall AOV') formula = `=IFERROR(${depCells['Total Conv. Value']}/${depCells['Total Conversions']},0)`;

        if (formula) targetCell.setFormula(formula);
        else targetCell.setValue('Formula Error');
      }
    }
    if (kpi.formatType === 'currency') targetCell.setNumberFormat(getCurrencyFormatString_(currencyCode));
    else if (kpi.formatType === 'integer') targetCell.setNumberFormat('#,##0');
    else if (kpi.formatType === 'percent') targetCell.setNumberFormat('0.00%');
    else if (kpi.formatType === 'decimal') targetCell.setNumberFormat('#,##0.00');
     else if (kpi.formatType === 'decimal_2') targetCell.setNumberFormat('0.00');
  });
  applyTableStyles(sheet, kpiTableHeaderRow , kpiTableData.length + 1, 2);
  currentRow += kpiTableData.length + 2;


  currentRow = createSectionHeader_(sheet, currentRow, "Top 5 Campaigns by Cost", 2);
  const topCampaignsTableHeaderRow = currentRow;
  const topCampaignsHeader = ['Campaign Name', 'Cost'];
  sheet.getRange(topCampaignsTableHeaderRow, 1, 1, 2).setValues([topCampaignsHeader]);
  const topCampaignsDataStartRow = currentRow + 1;

  const campNameColLetterForTop = getColLetter('Campaign Name', campaignSheetHeaders);
  const costColLetterForTop = getColLetter('Cost', campaignSheetHeaders);
  const lastColLetterForCampaigns = String.fromCharCode(64 + campaignSheetHeaders.length);

  if (costColLetterForTop !== "INVALID_COLUMN" && campNameColLetterForTop !== "INVALID_COLUMN") {
    const topCampaignsQuery = `=IFERROR(QUERY('${rawCampaignsSheetName}'!A2:${lastColLetterForCampaigns}${lastDataRow}, "SELECT ${campNameColLetterForTop}, ${costColLetterForTop} WHERE ${costColLetterForTop} > 0 ORDER BY ${costColLetterForTop} DESC LIMIT 5", 0), "Error loading Top Campaigns")`;
    sheet.getRange(topCampaignsDataStartRow, 1).setFormula(topCampaignsQuery);
    Utilities.sleep(500);
    applyTableStyles(sheet, topCampaignsTableHeaderRow, 6, 2);
    sheet.getRange(topCampaignsDataStartRow, 2, 5, 1).setNumberFormat(getCurrencyFormatString_(currencyCode));
  } else {
    sheet.getRange(topCampaignsDataStartRow, 1).setValue('Data Error for Top Campaigns');
  }
  currentRow += 7;

  const deviceMetricsHeader = ['Device', 'Cost', 'Clicks', 'Conversions', 'Conv. Value', 'CTR', 'CVR', 'CPA', 'ROAS', 'AOV'];
  currentRow = createSectionHeader_(sheet, currentRow, "Metrics by Device", deviceMetricsHeader.length);
  const deviceTableHeaderRow = currentRow;
  sheet.getRange(deviceTableHeaderRow, 1, 1, deviceMetricsHeader.length).setValues([deviceMetricsHeader]); // Explicitly set header
  const deviceDataStartRow = currentRow + 1;

  const rawDailySheetName = TAB_PREFIX_RAW + 'Daily Performance';
  const rawDailySheet = ss.getSheetByName(rawDailySheetName);
  let numDeviceDataRows = 0;
  if (rawDailySheet && rawDailySheet.getLastRow() > 1) {
    const dailyLastRow = rawDailySheet.getLastRow();
    const dailyHeaders = rawDailySheet.getRange(1,1,1,rawDailySheet.getLastColumn()).getValues()[0];

    const dailyDeviceColLetter = getColLetter('Device', dailyHeaders);
    const dailyCostColLetter = getColLetter('Cost', dailyHeaders);
    const dailyClicksColLetter = getColLetter('Clicks', dailyHeaders);
    const dailyImprColLetter = getColLetter('Impr.', dailyHeaders);
    const dailyConvColLetter = getColLetter('Conv.', dailyHeaders);
    const dailyConvValueColLetter = getColLetter('Conv. Value', dailyHeaders);
    const dailyLastColLetter = String.fromCharCode(64 + rawDailySheet.getLastColumn());

    if ([dailyDeviceColLetter, dailyCostColLetter, dailyClicksColLetter, dailyImprColLetter, dailyConvColLetter, dailyConvValueColLetter].some(c => c === "INVALID_COLUMN")) {
      logError_('populateOverviewDashboard_', new Error('One or more critical columns not found in Raw Data - Daily Performance for Device Breakdown.'));
      sheet.getRange(deviceDataStartRow, 1).setValue('Data Error: Missing columns in Daily Data for Device Breakdown.');
    } else {
      // FIX: Ensure QUERY header argument is 0 because we set headers manually
      const deviceMetricsQuery = `=IFERROR(QUERY('${rawDailySheetName}'!A2:${dailyLastColLetter}${dailyLastRow}, "SELECT ${dailyDeviceColLetter}, SUM(${dailyCostColLetter}), SUM(${dailyClicksColLetter}), SUM(${dailyConvColLetter}), SUM(${dailyConvValueColLetter}), SUM(${dailyClicksColLetter})/SUM(${dailyImprColLetter}), SUM(${dailyConvColLetter})/SUM(${dailyClicksColLetter}), SUM(${dailyCostColLetter})/SUM(${dailyConvColLetter}), SUM(${dailyConvValueColLetter})/SUM(${dailyCostColLetter}), SUM(${dailyConvValueColLetter})/SUM(${dailyConvColLetter}) WHERE ${dailyDeviceColLetter} IS NOT NULL GROUP BY ${dailyDeviceColLetter} LABEL SUM(${dailyCostColLetter}) '', SUM(${dailyClicksColLetter}) '', SUM(${dailyConvColLetter}) '', SUM(${dailyConvValueColLetter}) '', SUM(${dailyClicksColLetter})/SUM(${dailyImprColLetter}) '', SUM(${dailyConvColLetter})/SUM(${dailyClicksColLetter}) '', SUM(${dailyCostColLetter})/SUM(${dailyConvColLetter}) '', SUM(${dailyConvValueColLetter})/SUM(${dailyCostColLetter}) '', SUM(${dailyConvValueColLetter})/SUM(${dailyConvColLetter}) ''", 0), "Error loading Device Metrics")`;
      sheet.getRange(deviceDataStartRow, 1).setFormula(deviceMetricsQuery);

      Utilities.sleep(1000);
      const actualDeviceData = sheet.getRange(deviceDataStartRow, 1, sheet.getMaxRows() - deviceDataStartRow, deviceMetricsHeader.length).getDisplayValues();
      for (let i = 0; i < actualDeviceData.length; i++) {
          if (actualDeviceData[i][0] === "" || actualDeviceData[i][0] === null || actualDeviceData[i][0].toString().startsWith("#N/A") || actualDeviceData[i][0].toString().startsWith("Error loading")) {
              break;
          }
          numDeviceDataRows++;
      }
      logVerbose_(`Device Metrics Query returned ${numDeviceDataRows} data rows.`);


      if (numDeviceDataRows > 0) {
        applyTableStyles(sheet, deviceTableHeaderRow, numDeviceDataRows + 1, deviceMetricsHeader.length);

        sheet.getRange(deviceDataStartRow, 2, numDeviceDataRows, 1).setNumberFormat(getCurrencyFormatString_(currencyCode)); // Cost
        sheet.getRange(deviceDataStartRow, 3, numDeviceDataRows, 1).setNumberFormat('#,##0'); // Clicks
        sheet.getRange(deviceDataStartRow, 4, numDeviceDataRows, 1).setNumberFormat('#,##0.00'); // Conversions
        sheet.getRange(deviceDataStartRow, 5, numDeviceDataRows, 1).setNumberFormat(getCurrencyFormatString_(currencyCode)); // Conv Value
        sheet.getRange(deviceDataStartRow, 6, numDeviceDataRows, 1).setNumberFormat('0.00%'); // CTR
        sheet.getRange(deviceDataStartRow, 7, numDeviceDataRows, 1).setNumberFormat('0.00%'); // CVR
        sheet.getRange(deviceDataStartRow, 8, numDeviceDataRows, 1).setNumberFormat(getCurrencyFormatString_(currencyCode)); // CPA
        sheet.getRange(deviceDataStartRow, 9, numDeviceDataRows, 1).setNumberFormat('0.00'); // ROAS
        sheet.getRange(deviceDataStartRow, 10, numDeviceDataRows, 1).setNumberFormat(getCurrencyFormatString_(currencyCode)); // AOV
      } else {
         sheet.getRange(deviceDataStartRow, 1).setValue("No device data to display.");
         logWarn_('populateOverviewDashboard_', 'Device Metrics QUERY did not return valid data or returned an error.');
      }
    }
  } else {
       sheet.getRange(deviceDataStartRow, 1).setValue('Daily data for device breakdown not available.');
  }
  currentRow += Math.max(1, numDeviceDataRows) + 2; // Header + data + buffer


  const campTypeHeader = ['Campaign Type', 'Cost', 'Conversions', 'Conv. Value', 'ROAS', 'CPA', 'AOV'];
  currentRow = createSectionHeader_(sheet, currentRow, "Performance by Campaign Type", campTypeHeader.length);
  const campTypeTableHeaderRow = currentRow;
  sheet.getRange(campTypeTableHeaderRow, 1, 1, campTypeHeader.length).setValues([campTypeHeader]); // Explicitly set header
  const campTypeDataStartRow = currentRow + 1;
  let numCampTypeRows = 0;

  const campTypeColLetter = getColLetter('Channel', campaignSheetHeaders);
  const campCostColLetter = getColLetter('Cost', campaignSheetHeaders);
  const campConvColLetter = getColLetter('Conv.', campaignSheetHeaders);
  const campConvValueColLetter = getColLetter('Conv. Value', campaignSheetHeaders);

  if ([campTypeColLetter, campCostColLetter, campConvColLetter, campConvValueColLetter].some(c => c === "INVALID_COLUMN")) {
    logError_('populateOverviewDashboard_', new Error('One or more critical columns not found in Raw Data - Campaigns for Campaign Type Breakdown.'));
    sheet.getRange(campTypeDataStartRow, 1).setValue('Data Error: Missing columns in Campaign Data for Type Breakdown.');
  } else {
    // FIX: Ensure QUERY header argument is 0
    const campTypeQuery = `=IFERROR(QUERY('${rawCampaignsSheetName}'!A2:${lastColLetterForCampaigns}${lastDataRow}, "SELECT ${campTypeColLetter}, SUM(${campCostColLetter}), SUM(${campConvColLetter}), SUM(${campConvValueColLetter}), SUM(${campConvValueColLetter})/SUM(${campCostColLetter}), SUM(${campCostColLetter})/SUM(${campConvColLetter}), SUM(${campConvValueColLetter})/SUM(${campConvColLetter}) WHERE ${campTypeColLetter} IS NOT NULL GROUP BY ${campTypeColLetter} LABEL SUM(${campCostColLetter}) '', SUM(${campConvColLetter}) '', SUM(${campConvValueColLetter}) '', SUM(${campConvValueColLetter})/SUM(${campCostColLetter}) '', SUM(${campCostColLetter})/SUM(${campConvColLetter}) '', SUM(${campConvValueColLetter})/SUM(${campConvColLetter}) ''", 0), "Error loading Campaign Type data")`;
    sheet.getRange(campTypeDataStartRow, 1).setFormula(campTypeQuery);
    Utilities.sleep(500);

    const actualCampTypeData = sheet.getRange(campTypeDataStartRow, 1, sheet.getMaxRows() - campTypeDataStartRow, campTypeHeader.length).getDisplayValues();
    for (let i = 0; i < actualCampTypeData.length; i++) {
        if (actualCampTypeData[i][0] === "" || actualCampTypeData[i][0] === null || actualCampTypeData[i][0].toString().startsWith("#N/A") || actualCampTypeData[i][0].toString().startsWith("Error loading")) {
            break;
        }
        numCampTypeRows++;
    }
    logVerbose_(`Campaign Type Query returned ${numCampTypeRows} data rows.`);


    if (numCampTypeRows > 0 && !sheet.getRange(campTypeDataStartRow, 1).getDisplayValue().startsWith("Error loading")) {
      applyTableStyles(sheet, campTypeTableHeaderRow, numCampTypeRows + 1, campTypeHeader.length);

      sheet.getRange(campTypeDataStartRow, 2, numCampTypeRows, 1).setNumberFormat(getCurrencyFormatString_(currencyCode)); // Cost
      sheet.getRange(campTypeDataStartRow, 3, numCampTypeRows, 1).setNumberFormat('#,##0.00'); // Conversions
      sheet.getRange(campTypeDataStartRow, 4, numCampTypeRows, 1).setNumberFormat(getCurrencyFormatString_(currencyCode)); // Conv. Value
      sheet.getRange(campTypeDataStartRow, 5, numCampTypeRows, 1).setNumberFormat('0.00'); // ROAS
      sheet.getRange(campTypeDataStartRow, 6, numCampTypeRows, 1).setNumberFormat(getCurrencyFormatString_(currencyCode)); // CPA
      sheet.getRange(campTypeDataStartRow, 7, numCampTypeRows, 1).setNumberFormat(getCurrencyFormatString_(currencyCode)); // AOV
    } else {
        sheet.getRange(campTypeDataStartRow, 1).setValue("No campaign type data to display.");
        logWarn_('populateOverviewDashboard_', 'Campaign Type QUERY did not return valid data or returned an error.');
    }
  }
  currentRow += Math.max(1, numCampTypeRows) + 2; // Header + data + buffer

  autoResizeSheetColumns_(sheet);
  logInfo_(`Overview dashboard '${sheetName}' populated.`);
}

function populateCampaignPerformanceDashboardTab_(ss, currencyCode) {
  const sheetName = CAMPAIGN_PERF_TAB_NAME;
  const sheet = ensureSheet_(ss, sheetName, false);
  if (!sheet) { logError_('populateCampaignPerformanceDashboardTab_', new Error(`Failed to ensure sheet: ${sheetName}`)); return; }

  applyGeneralSheetFormatting_(sheet);
  let currentRow = 1;
  currentRow = createSectionHeader_(sheet, currentRow, 'Detailed Campaign Performance', HDR_CAMPAIGNS.length);


  const rawSheetName = TAB_PREFIX_RAW + 'Campaigns';
  const rawSheet = ss.getSheetByName(rawSheetName);
  if (!rawSheet || rawSheet.getLastRow() < 1) {
    logError_('populateCampaignPerformanceDashboardTab_', new Error(`Raw sheet ${rawSheetName} missing or empty.`));
    sheet.getRange(currentRow, 1).setValue(`Raw campaign data not found in '${rawSheetName}'.`);
    return;
  }
  const rawSheetHeaders = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getValues()[0];
  const numRawCols = rawSheetHeaders.length;
  const lastColLetter = String.fromCharCode(64 + numRawCols);
  const lastRowInRaw = rawSheet.getLastRow();

  const queryString = `=IFERROR(QUERY('${rawSheetName}'!A1:${lastColLetter}${lastRowInRaw > 0 ? lastRowInRaw : 1},"SELECT * WHERE Col1 IS NOT NULL",1),"Error loading campaign data. Check Raw Data - Campaigns.")`;
  sheet.getRange(currentRow, 1).setFormula(queryString);
  sheet.setFrozenRows(currentRow);
  SpreadsheetApp.flush();

  Utilities.sleep(1500);
  const dataRange = sheet.getDataRange();
  const numRowsInSheet = dataRange.getNumRows(); // Total rows in the sheet after QUERY
  const queryOutputStartRow = currentRow;

  if (numRowsInSheet > queryOutputStartRow) {
    const headersFromQuery = sheet.getRange(queryOutputStartRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const formats = headersFromQuery.map(header => getColumnFormatForHeader_(header, currencyCode));

    const dataBodyStartRow = queryOutputStartRow + 1;
    const numDataBodyRows = numRowsInSheet - queryOutputStartRow;

    if (numDataBodyRows > 0) {
        const dataBodyRange = sheet.getRange(dataBodyStartRow, 1, numDataBodyRows, headersFromQuery.length);
        const backgrounds = [];
        for (let i = 0; i < numDataBodyRows; i++) {
            backgrounds.push(new Array(headersFromQuery.length).fill(
                i % 2 === 0 ? STYLE_TABLE_ROW_ODD_BG : STYLE_TABLE_ROW_EVEN_BG
            ));
        }
        if (backgrounds.length > 0) {
            dataBodyRange.setBackgrounds(backgrounds);
        }

        for (let c = 0; c < headersFromQuery.length; c++) {
          if (formats[c] !== '@') {
            try {
              sheet.getRange(dataBodyStartRow, c + 1, numDataBodyRows, 1).setNumberFormat(formats[c]);
            } catch(e) {
              logWarn_('populateCampaignPerformanceDashboardTab_', `Could not apply format ${formats[c]} to column ${headersFromQuery[c]}: ${e.message}`);
            }
          }
        }
        // Add Totals Row
        const totalsRowData = ['TOTAL'];
        let totalCost = 0, totalImpr = 0, totalClicks = 0, totalConv = 0, totalConvValue = 0;
        const dataValues = dataBodyRange.getValues(); // Get values from the sheet after QUERY
        dataValues.forEach(row => {
            totalCost += safeParseFloat_(row[headersFromQuery.indexOf('Cost')]);
            totalImpr += safeParseInt_(row[headersFromQuery.indexOf('Impr.')]);
            totalClicks += safeParseInt_(row[headersFromQuery.indexOf('Clicks')]);
            totalConv += safeParseFloat_(row[headersFromQuery.indexOf('Conv.')]);
            totalConvValue += safeParseFloat_(row[headersFromQuery.indexOf('Conv. Value')]);
        });

        for (let i = 1; i < headersFromQuery.length; i++) { // Start from 1 to skip 'Campaign ID' or 'Campaign Name'
            const header = headersFromQuery[i];
            if (header === 'Cost') totalsRowData.push(totalCost);
            else if (header === 'Impr.') totalsRowData.push(totalImpr);
            else if (header === 'Clicks') totalsRowData.push(totalClicks);
            else if (header === 'CTR') totalsRowData.push(pct(totalClicks, totalImpr, true));
            else if (header === 'Avg. CPC') totalsRowData.push(div(totalCost, totalClicks));
            else if (header === 'Conv.') totalsRowData.push(totalConv);
            else if (header === 'Conv. Value') totalsRowData.push(totalConvValue);
            else if (header === 'ROAS') totalsRowData.push(div(totalConvValue, totalCost));
            else if (header === 'CPA') totalsRowData.push(div(totalCost, totalConv));
            else if (header === 'AOV') totalsRowData.push(div(totalConvValue, totalConv));
            else totalsRowData.push('');
        }
        const totalsRowNumber = dataBodyStartRow + numDataBodyRows;
        const totalsRange = sheet.getRange(totalsRowNumber, 1, 1, headersFromQuery.length);
        totalsRange.setValues([totalsRowData]);
        totalsRange.setBackground(STYLE_TOTALS_ROW.backgroundColor)
                   .setFontWeight(STYLE_TOTALS_ROW.fontWeight)
                   .setFontColor(STYLE_TOTALS_ROW.fontColor);
        totalsRange.setBorder(true, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

        for (let c = 0; c < headersFromQuery.length; c++) {
            if (formats[c] !== '@') {
                try {
                    sheet.getRange(totalsRowNumber, c + 1).setNumberFormat(formats[c]);
                } catch(e) { /* ignore */ }
            }
        }
    }
    styleHeader_(sheet.getRange(queryOutputStartRow, 1, 1, headersFromQuery.length));
  }
  autoResizeSheetColumns_(sheet);
  logInfo_(`Campaign Performance dashboard '${sheetName}' set up with QUERY and enhanced formatting attempted.`);
}

function populateProductPerformanceAnalysisDashboardTab_(ss, currencyCode) {
  const sheetName = PRODUCT_ANALYSIS_TAB_NAME;
  const analysisSheet = ensureSheet_(ss, sheetName, false);
  if (!analysisSheet) { logError_('populateProductPerformanceAnalysisDashboardTab_', new Error(`Failed to ensure sheet: ${sheetName}`)); return; }

  applyGeneralSheetFormatting_(analysisSheet);
  let currentRow = 1;
  currentRow = createSectionHeader_(analysisSheet, currentRow, 'Product Performance Analysis & Recommendations', HDR_PRODUCTS.length);


  const rawProductSheetName = TAB_PREFIX_RAW + 'Products';
  const rawProductSheet = ss.getSheetByName(rawProductSheetName);

  if (!rawProductSheet || rawProductSheet.getLastRow() < 2) {
    logError_('populateProductPerformanceAnalysisDashboardTab_', new Error(`Raw Products sheet '${rawProductSheetName}' is empty or missing.`));
    analysisSheet.getRange(currentRow, 1).setValue(`Data source '${rawProductSheetName}' is empty or missing.`);
    return;
  }

  const rawHeaders = rawProductSheet.getRange(1, 1, 1, rawProductSheet.getLastColumn()).getValues()[0];
  const rawData = rawProductSheet.getRange(2, 1, rawProductSheet.getLastRow() - 1, rawHeaders.length).getValues();

  const idIdx = rawHeaders.indexOf('Product ID');
  const titleIdx = rawHeaders.indexOf('Product Title');
  const costIdx = rawHeaders.indexOf('Cost');
  const imprIdx = rawHeaders.indexOf('Impr.');
  const clicksIdx = rawHeaders.indexOf('Clicks');
  const convIdx = rawHeaders.indexOf('Conv.');
  const convValIdx = rawHeaders.indexOf('Conv. Value');

  if ([idIdx, titleIdx, costIdx, imprIdx, clicksIdx, convIdx, convValIdx].includes(-1)) {
    logError_('populateProductPerformanceAnalysisDashboardTab_', new Error('Required columns missing in Raw Data - Products.'));
    analysisSheet.getRange(currentRow, 1).setValue('Error: Header mismatch in Raw Data - Products.');
    return;
  }

  const aggregatedProducts = {};
  let totalAccountCost = 0;
  let totalAccountConversions = 0;

  rawData.forEach(row => {
    const productId = row[idIdx];
    if (!productId) return;

    const cost = safeParseFloat_(row[costIdx]);
    const conversions = safeParseFloat_(row[convIdx]);
    totalAccountCost += cost;
    totalAccountConversions += conversions;

    if (!aggregatedProducts[productId]) {
      aggregatedProducts[productId] = {
        title: row[titleIdx] || 'N/A',
        cost: 0, impr: 0, clicks: 0, conv: 0, convVal: 0
      };
    }
    aggregatedProducts[productId].cost += cost;
    aggregatedProducts[productId].impr += safeParseInt_(row[imprIdx]);
    aggregatedProducts[productId].clicks += safeParseInt_(row[clicksIdx]);
    aggregatedProducts[productId].conv += conversions;
    aggregatedProducts[productId].convVal += safeParseFloat_(row[convValIdx]);
    if (aggregatedProducts[productId].title === 'N/A' && row[titleIdx]) {
      aggregatedProducts[productId].title = row[titleIdx];
    }
  });

  const accountAvgCPA = totalAccountConversions > 0 ? totalAccountCost / totalAccountConversions : REC_DEFAULT_ACCOUNT_CPA;

  const outputHeaders = ['Product ID', 'Product Title', 'Cost', 'Impr.', 'Clicks', 'CTR', 'Conv.', 'CVR', 'CPA', 'Avg. CPC', 'Conv. Value', 'ROAS', 'AOV', 'Recommendation'];
  const outputData = [];

  let grandTotalCost = 0, grandTotalImpr = 0, grandTotalClicks = 0, grandTotalConv = 0, grandTotalConvVal = 0;


  for (const productId in aggregatedProducts) {
    const p = aggregatedProducts[productId];
    const ctr = pct(p.clicks, p.impr, true);
    const cvr = pct(p.conv, p.clicks, true);
    const cpa = div(p.cost, p.conv);
    const avgCpc = div(p.cost, p.clicks);
    const roas = div(p.convVal, p.cost);
    const aov = div(p.convVal, p.conv);

    grandTotalCost += p.cost;
    grandTotalImpr += p.impr;
    grandTotalClicks += p.clicks;
    grandTotalConv += p.conv;
    grandTotalConvVal += p.convVal;

    let recText = [];
    if (p.cost > 0 && p.conv === 0 && p.clicks >= REC_MIN_CLICKS_NO_CONV) recText.push('High spend, no conv.');
    if (roas > 0 && roas < REC_TARGET_ROAS_THRESHOLD) recText.push(`ROAS below target (${REC_TARGET_ROAS_THRESHOLD.toFixed(2)}).`);
    if (cpa > 0 && cpa > (accountAvgCPA * REC_CPA_FACTOR_THRESHOLD)) recText.push(`CPA above target.`);
    if (p.impr >= REC_MIN_IMPR_FOR_CTR && ctr < REC_LOW_CTR_THRESHOLD) recText.push(`Low CTR (${(ctr * 100).toFixed(1)}%).`);
    if (recText.length === 0 && p.conv > 0) recText.push('Good performance.');
    else if (recText.length === 0) recText.push('Monitor data.');

    outputData.push([
      productId, p.title, p.cost, p.impr, p.clicks, ctr, p.conv, cvr, cpa, avgCpc, p.convVal, roas, aov, recText.join(' ')
    ]);
  }

  outputData.sort((a, b) => b[outputHeaders.indexOf('Conv.')] - a[outputHeaders.indexOf('Conv.')]);

  if (outputData.length > 0) {
      const totalCTR = pct(grandTotalClicks, grandTotalImpr, true);
      const totalCVR = pct(grandTotalConv, grandTotalClicks, true);
      const totalCPA = div(grandTotalCost, grandTotalConv);
      const totalAvgCPC = div(grandTotalCost, grandTotalClicks);
      const totalROAS = div(grandTotalConvVal, grandTotalCost);
      const totalAOV = div(grandTotalConvVal, grandTotalConv);
      const totalRow = ['TOTAL', '', grandTotalCost, grandTotalImpr, grandTotalClicks, totalCTR, grandTotalConv, totalCVR, totalCPA, totalAvgCPC, grandTotalConvVal, totalROAS, totalAOV, ''];
      outputData.push(totalRow);
  }


  const analysisSheetObj = ss.getSheetByName(sheetName);
  if (analysisSheetObj) {
    analysisSheetObj.getRange(currentRow, 1, 1, outputHeaders.length).setValues([outputHeaders]);
    if (outputData && outputData.length > 0) {
      const dataStartRow = currentRow + 1;
      analysisSheetObj.getRange(dataStartRow, 1, outputData.length, outputHeaders.length).setValues(outputData);
      const columnFormatStrings = outputHeaders.map(header => getColumnFormatForHeader_(header, currencyCode));
      const formatsArray = [];
      for (let r = 0; r < outputData.length; r++) formatsArray.push(columnFormatStrings);
      applyChunkedFormatting_(analysisSheetObj, formatsArray, dataStartRow);
      applyTableStyles(analysisSheetObj, currentRow, outputData.length + 1, outputHeaders.length);

      if (outputData.length > 0) {
          const totalRowRange = analysisSheetObj.getRange(dataStartRow + outputData.length -1, 1, 1, outputHeaders.length);
          totalRowRange.setBackground(STYLE_TOTALS_ROW.backgroundColor)
                       .setFontWeight(STYLE_TOTALS_ROW.fontWeight)
                       .setFontColor(STYLE_TOTALS_ROW.fontColor);
          totalRowRange.setBorder(true,null,true,null,null,null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    } else {
        analysisSheetObj.getRange(currentRow + 1, 1).setValue("No product data to display.");
    }
    autoResizeSheetColumns_(analysisSheetObj, outputHeaders.length);
  }
  logInfo_(`Product Performance Analysis tab '${sheetName}' populated.`);
}

function populateAssetPerformanceDashboardTab_(ss, currencyCode) {
  if (!ENABLE_ASSET_PERFORMANCE_TAB) {
    logInfo_('Asset Performance tab generation is disabled via Config.');
    return;
  }
  const sheetName = ASSET_PERFORMANCE_TAB_NAME;
  const analysisSheet = ensureSheet_(ss, sheetName, false);
   if (!analysisSheet) { logError_('populateAssetPerformanceDashboardTab_', new Error(`Failed to ensure sheet: ${sheetName}`)); return; }

  applyGeneralSheetFormatting_(analysisSheet);
  let currentRow = 1;
  currentRow = createSectionHeader_(analysisSheet, currentRow, 'Asset Performance', HDR_ASSETS.length);


  const rawAssetSheetName = TAB_PREFIX_RAW + 'Asset Performance';
  const rawAssetSheet = ss.getSheetByName(rawAssetSheetName);

  if (!rawAssetSheet || rawAssetSheet.getLastRow() < 2) {
    logError_('populateAssetPerformanceDashboardTab_', new Error(`Raw Asset Performance sheet '${rawAssetSheetName}' is empty or missing.`));
    analysisSheet.getRange(currentRow, 1).setValue(`Data source '${rawAssetSheetName}' is empty or missing.`);
    return;
  }

  const rawHeaders = rawAssetSheet.getRange(1, 1, 1, rawAssetSheet.getLastColumn()).getValues()[0];
  const rawData = rawAssetSheet.getRange(2, 1, rawAssetSheet.getLastRow() - 1, rawHeaders.length).getValues();

  const outputHeaders = ['Asset ID', 'Asset Type', 'Field Type', 'Campaign Name', 'Ad Group Name', 'Impressions', 'Clicks', 'Cost', 'CTR'];
  const outputData = [];
  let totalImpressions = 0, totalClicks = 0, totalCost = 0;

  const assetIdIdx = rawHeaders.indexOf('Asset ID');
  const assetTypeIdx = rawHeaders.indexOf('Asset Type');
  const fieldTypeIdx = rawHeaders.indexOf('Field Type');
  const campaignNameIdx = rawHeaders.indexOf('Campaign Name');
  const adGroupNameIdx = rawHeaders.indexOf('Ad Group Name');
  const impressionsIdx = rawHeaders.indexOf('Impressions');
  const clicksIdx = rawHeaders.indexOf('Clicks');
  const costIdx = rawHeaders.indexOf('Cost');

  if ([assetIdIdx, assetTypeIdx, fieldTypeIdx, campaignNameIdx, adGroupNameIdx, impressionsIdx, clicksIdx, costIdx].includes(-1)) {
    logError_('populateAssetPerformanceDashboardTab_', new Error('Required columns missing in Raw Data - Asset Performance. Check HDR_ASSETS and query.'), `Missing indices for: ${outputHeaders.filter((h,i) => [assetIdIdx, assetTypeIdx, fieldTypeIdx, campaignNameIdx, adGroupNameIdx, impressionsIdx, clicksIdx, costIdx][i] === -1 )}`);
    analysisSheet.getRange(currentRow, 1).setValue('Error: Header mismatch in Raw Data - Asset Performance.');
    return;
  }

  rawData.forEach(row => {
    const impressions = safeParseInt_(row[impressionsIdx]);
    const clicks = safeParseInt_(row[clicksIdx]);
    const cost = safeParseFloat_(row[costIdx]);
    const ctr = pct(clicks, impressions, true);

    totalImpressions += impressions;
    totalClicks += clicks;
    totalCost += cost;

    outputData.push([
      row[assetIdIdx],
      row[assetTypeIdx],
      row[fieldTypeIdx],
      row[campaignNameIdx],
      row[adGroupNameIdx],
      impressions,
      clicks,
      cost,
      ctr
    ]);
  });

  outputData.sort((a, b) => b[outputHeaders.indexOf('Impressions')] - a[outputHeaders.indexOf('Impressions')]);

  if (outputData.length > 0) {
      const totalCTR = pct(totalClicks, totalImpressions, true);
      const totalRow = ['', '', '', '', 'TOTAL', totalImpressions, totalClicks, totalCost, totalCTR];
      outputData.push(totalRow);
  }

  const analysisSheetObj = ss.getSheetByName(sheetName);
  if (analysisSheetObj) {
     analysisSheetObj.getRange(currentRow, 1, 1, outputHeaders.length).setValues([outputHeaders]);
     if (outputData && outputData.length > 0) {
      const dataStartRow = currentRow + 1;
      analysisSheetObj.getRange(dataStartRow, 1, outputData.length, outputHeaders.length).setValues(outputData);
      const columnFormatStrings = outputHeaders.map(header => getColumnFormatForHeader_(header, currencyCode));
      const formatsArray = [];
      for (let r = 0; r < outputData.length; r++) formatsArray.push(columnFormatStrings);
      applyChunkedFormatting_(analysisSheetObj, formatsArray, dataStartRow);
      applyTableStyles(analysisSheetObj, currentRow, outputData.length + 1, outputHeaders.length);

      if (outputData.length > 0) {
          const totalRowRange = analysisSheetObj.getRange(dataStartRow + outputData.length -1, 1, 1, outputHeaders.length);
          totalRowRange.setBackground(STYLE_TOTALS_ROW.backgroundColor)
                       .setFontWeight(STYLE_TOTALS_ROW.fontWeight)
                       .setFontColor(STYLE_TOTALS_ROW.fontColor);
          totalRowRange.setBorder(true,null,true,null,null,null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }

    } else {
        analysisSheetObj.getRange(currentRow + 1, 1).setValue("No asset data to display.");
    }
    autoResizeSheetColumns_(analysisSheetObj, outputHeaders.length);
  }
  logInfo_(`Asset Performance dashboard tab '${sheetName}' populated with ${outputData.length} assets.`);
}


function populateRecommendationsDashboardTab_(ss) {
  const sheetName = RECOMMENDATIONS_TAB_NAME;
  const sheet = ensureSheet_(ss, sheetName, false);
  if (!sheet) { logError_('populateRecommendationsDashboardTab_', new Error(`Failed to ensure sheet: ${sheetName}`)); return; }

  applyGeneralSheetFormatting_(sheet);
  let currentRow = 1;
  sheet.getRange('A1').setValue('Automated Recommendations & Insights').setFontSize(16).setFontWeight('bold').setFontFamily('Arial');
  sheet.getRange('A2').setValue('This tab provides rule-based recommendations. Review underlying data for context.').setWrap(true).setFontFamily('Arial');
  currentRow = 4;

  const recHeaders = ['Area', 'Observation / Rule', 'Status / Details', 'Relevant Tab(s)'];

  const recommendations = [];
  if (ENABLE_BUDGET_PACING_TAB) {
    const pacingSheet = ss.getSheetByName(BUDGET_PACING_TAB_NAME);
    if (pacingSheet && pacingSheet.getLastRow() > 2) {
      const pacingDataRange = pacingSheet.getRange(3, 1, pacingSheet.getLastRow() - 2, pacingSheet.getLastColumn());
      const pacingData = pacingDataRange.getValues();
      const pacingHeadersActual = pacingSheet.getRange(2,1,1,pacingSheet.getLastColumn()).getValues()[0];
      const pacingStatusCol = pacingHeadersActual.indexOf('Status');
      const pacingCampNameCol = pacingHeadersActual.indexOf('Campaign Name');
      pacingData.forEach(row => {
          if (pacingStatusCol !== -1 && pacingCampNameCol !== -1 && row[pacingStatusCol] && (String(row[pacingStatusCol]).toLowerCase().includes('over') || String(row[pacingStatusCol]).toLowerCase().includes('under'))) {
            recommendations.push(['Budget Pacing', `Campaign "${row[pacingCampNameCol]}" is ${row[pacingStatusCol]}.`, 'Review budget allocation or bids.', BUDGET_PACING_TAB_NAME]);
          }
      });
    } else {
      recommendations.push(['Budget Pacing', 'Pacing data not available or tab not populated.', 'N/A', BUDGET_PACING_TAB_NAME]);
    }
  }

  if (ENABLE_IMPRESSION_SHARE_TAB) {
    const shareRankSheet = ss.getSheetByName(SHARE_RANK_TAB_NAME);
    if (shareRankSheet && shareRankSheet.getLastRow() > 2) {
      const srDataRange = shareRankSheet.getRange(3, 1, shareRankSheet.getLastRow() - 2, shareRankSheet.getLastColumn());
      const srData = srDataRange.getValues();
      const srHeadersActual = shareRankSheet.getRange(2,1,1,shareRankSheet.getLastColumn()).getValues()[0];
      const srCampNameCol = srHeadersActual.indexOf('Campaign Name');
      const srRankLostCol = srHeadersActual.indexOf('IS Lost (Rank)');
      const srBudgetLostCol = srHeadersActual.indexOf('IS Lost (Budget)');
      srData.forEach(row => {
          if (srCampNameCol !== -1 && srRankLostCol !== -1 && safeParseFloat_(row[srRankLostCol]) > 0.20) {
            recommendations.push(['Impression Share', `Campaign "${row[srCampNameCol]}" has high IS Lost to Rank (${(safeParseFloat_(row[srRankLostCol]) * 100).toFixed(0)}%).`, 'Improve Ad Rank (Bids/QS).', SHARE_RANK_TAB_NAME]);
          }
          if (srCampNameCol !== -1 && srBudgetLostCol !== -1 && safeParseFloat_(row[srBudgetLostCol]) > 0.20) {
            recommendations.push(['Impression Share', `Campaign "${row[srCampNameCol]}" has high IS Lost to Budget (${(safeParseFloat_(row[srBudgetLostCol]) * 100).toFixed(0)}%).`, 'Consider budget increase if profitable.', SHARE_RANK_TAB_NAME]);
          }
      });
    } else {
      recommendations.push(['Impression Share', 'Share/Rank data not available or tab not populated.', 'N/A', SHARE_RANK_TAB_NAME]);
    }
  }

  if (recommendations.length > 0) {
    sheet.getRange(currentRow + 1, 1, recommendations.length, recommendations[0].length).setValues(recommendations).setWrap(true).setVerticalAlignment('top');
    applyTableStyles(sheet, currentRow, recommendations.length + 1, recHeaders.length);
    sheet.getRange(currentRow, 1, 1, recHeaders.length).setValues([recHeaders]);
    styleHeader_(sheet.getRange(currentRow, 1, 1, recHeaders.length));

    applyConditionalFormattingRecommendations_(sheet, recommendations.length, currentRow + 1);
  } else {
    sheet.getRange(currentRow + 1, 1).setValue('No specific recommendations triggered based on current rules or enabled features.');
  }
  autoResizeSheetColumns_(sheet, recHeaders.length);
  logInfo_(`Recommendations tab '${sheetName}' populated.`);
}

function applyConditionalFormattingRecommendations_(sheet, numRecommendationRows, dataStartRow) {
  if (!sheet || numRecommendationRows === 0) return;

  const range = sheet.getRange(dataStartRow, 1, numRecommendationRows, 4);

  const rules = [];

  const issueKeywords = ['issue', 'high', 'low', 'over', 'under', 'lost', 'n/a', 'error', 'not available', 'not populated', 'missing', 'below target', 'above target'];
  let formulaRedConditions = issueKeywords.map(kw => `REGEXMATCH(LOWER(C${dataStartRow}),"${kw.replace('(','\\(').replace(')','\\)')}")`);
  let formulaRed = `OR(${formulaRedConditions.join(',')})`;

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=${formulaRed}`)
    .setBackground(LIGHT_RED_BACKGROUND)
    .setRanges([range])
    .build());

  const positiveKeywords = ['good performance', 'is healthy', 'maintain current pacing', 'on track'];
  let formulaGreenConditions = positiveKeywords.map(kw => `REGEXMATCH(LOWER(C${dataStartRow}),"${kw}")`);
  let formulaGreen = `AND(OR(${formulaGreenConditions.join(',')}),NOT(OR(${formulaRedConditions.join(',')})))`;

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=${formulaGreen}`)
    .setBackground(LIGHT_GREEN_BACKGROUND)
    .setRanges([range])
    .build());

  sheet.setConditionalFormatRules(rules);
  logVerbose_('Applied conditional formatting to Recommendations tab.');
}


function createOverviewDashboardCharts_(ss, currencyCode) {
  const overviewSheetName = OVERVIEW_TAB_NAME;
  const rawDailySheetName = TAB_PREFIX_RAW + 'Daily Performance';
  const rawProductSheetName = TAB_PREFIX_RAW + 'Products';

  const overviewSheet = ss.getSheetByName(overviewSheetName);
  const rawDailySheet = ss.getSheetByName(rawDailySheetName);
  const rawProductSheet = ss.getSheetByName(rawProductSheetName);


  if (!overviewSheet) {
    logError_('createOverviewDashboardCharts_', new Error(`Sheet ${overviewSheetName} missing.`));
    return;
  }
  SpreadsheetApp.flush();
  let chartStartRow = overviewSheet.getLastRow() + 3;
  let chartStartCol = 1;
  const chartsPerRow = 2;
  let chartsInCurrentRow = 0;
  const chartWidth = 650; // Standardized width
  const chartHeight = 350; // Standardized height
  const chartColOffset = 8; // Approx columns for a chart + small gap

  overviewSheet.getCharts().forEach(chart => overviewSheet.removeChart(chart));

  let tempSheet = ss.getSheetByName("_ChartDataSource_Temp");
  if (tempSheet) {
      tempSheet.clear();
  } else {
      tempSheet = ss.insertSheet("_ChartDataSource_Temp").hideSheet();
  }

  // --- Chart 1: Daily Performance Trends (Cost, Conversions, ROAS) ---
  if (rawDailySheet && rawDailySheet.getLastRow() >=2) {
      const rawDailyHeaders = rawDailySheet.getRange(1, 1, 1, rawDailySheet.getLastColumn()).getValues()[0];
      const dateIdx = rawDailyHeaders.indexOf('Date');
      const costIdx = rawDailyHeaders.indexOf('Cost');
      const convIdx = rawDailyHeaders.indexOf('Conv.');
      const convValIdx = rawDailyHeaders.indexOf('Conv. Value');

      if ([dateIdx, costIdx, convIdx, convValIdx].every(idx => idx !== -1)) {
          const rawDailyData = rawDailySheet.getRange(2, 1, rawDailySheet.getLastRow() - 1, rawDailyHeaders.length).getValues();
          const trendChartDataAggregated = {};
          rawDailyData.forEach(row => {
            const dateStr = Utilities.formatDate(new Date(row[dateIdx]), AdsApp.currentAccount().getTimeZone(), "yyyy-MM-dd");
            if (!trendChartDataAggregated[dateStr]) {
              trendChartDataAggregated[dateStr] = {
                date: new Date(row[dateIdx]), cost: 0, conversions: 0, convValue: 0
              };
            }
            trendChartDataAggregated[dateStr].cost += safeParseFloat_(row[costIdx]);
            trendChartDataAggregated[dateStr].conversions += safeParseFloat_(row[convIdx]);
            trendChartDataAggregated[dateStr].convValue += safeParseFloat_(row[convValIdx]);
          });

          const trendChartSourceData = [['Date', 'Cost', 'Conversions', 'ROAS']];
          Object.keys(trendChartDataAggregated).sort().forEach(dateStr => {
            const dayData = trendChartDataAggregated[dateStr];
            trendChartSourceData.push([dayData.date, dayData.cost, dayData.conversions, div(dayData.convValue, dayData.cost)]);
          });

          if (trendChartSourceData.length > 1) {
            tempSheet.clearContents().clearFormats();
            const tempRange = tempSheet.getRange(1, 1, trendChartSourceData.length, 4);
            tempRange.setValues(trendChartSourceData);
            SpreadsheetApp.flush();
            tempRange.offset(1, 0, trendChartSourceData.length - 1, 1).setNumberFormat('MMM d, yy');
            tempRange.offset(1, 1, trendChartSourceData.length - 1, 1).setNumberFormat(getCurrencyFormatString_(currencyCode));
            tempRange.offset(1, 2, trendChartSourceData.length - 1, 1).setNumberFormat('#,##0.00');
            tempRange.offset(1, 3, trendChartSourceData.length - 1, 1).setNumberFormat('0.00');

            const trendChart = overviewSheet.newChart().setChartType(Charts.ChartType.COMBO)
              .addRange(tempRange)
              .setPosition(chartStartRow, chartStartCol, 5, 5)
              .setOption('title', 'Daily Performance Trends (Cost, Conversions, ROAS)')
              .setOption('titleTextStyle', STYLE_CHART_TITLE)
              .setOption('series', {
                0: { type: 'bars', color: ADOPTED_CHART_COLORS.DAILY_SPEND, targetAxisIndex: 0, labelInLegend: 'Cost' },
                1: { type: 'line', color: ADOPTED_CHART_COLORS.PRODUCT_CONVERSIONS, targetAxisIndex: 0, labelInLegend: 'Conversions' },
                2: { type: 'line', color: ADOPTED_CHART_COLORS.CONV_RATE, targetAxisIndex: 1, labelInLegend: 'ROAS' }
              })
              .setOption('vAxes', {
                0: { title: 'Cost / Conversions', titleTextStyle: STYLE_CHART_AXIS_TITLE, textStyle: STYLE_CHART_AXIS_LABEL, format: getCurrencyFormatString_(currencyCode) },
                1: { title: 'ROAS', titleTextStyle: STYLE_CHART_AXIS_TITLE, textStyle: STYLE_CHART_AXIS_LABEL, format: '0.00', viewWindow: { min: 0 } }
              })
              .setOption('hAxis', { title: 'Date', titleTextStyle: STYLE_CHART_AXIS_TITLE, textStyle: STYLE_CHART_AXIS_LABEL, format: 'MMM d' })
              .setOption('legend', { position: 'top', alignment: 'center', textStyle: STYLE_CHART_LEGEND })
              .setOption('width', chartWidth).setOption('height', chartHeight)
              .setOption('useFirstColumnAsDomain', true)
              .setOption('backgroundColor', '#FFFFFF')
              .build();
            overviewSheet.insertChart(trendChart);
            chartsInCurrentRow++;
            logInfo_('Daily Performance Trend Chart created.');
          } else {
            overviewSheet.getRange(chartStartRow, chartStartCol).setValue('Not enough data for Daily Performance Trend Chart.');
          }
      } else {
          overviewSheet.getRange(chartStartRow, chartStartCol).setValue('Daily Performance Trend Chart: Missing required columns in raw data.');
      }
  } else {
      overviewSheet.getRange(chartStartRow, chartStartCol).setValue('Daily Performance Trend Chart: Raw daily data sheet missing or empty.');
  }
  if (chartsInCurrentRow >= chartsPerRow) { chartStartRow += (Math.floor(chartHeight/20) + 2); chartStartCol = 1; chartsInCurrentRow = 0; } else { chartStartCol += chartColOffset; }


  // --- Chart 2: Pie Chart for Devices per Conversion (Overall) ---
  if (rawDailySheet && rawDailySheet.getLastRow() >=2) {
      const rawDailyHeaders = rawDailySheet.getRange(1, 1, 1, rawDailySheet.getLastColumn()).getValues()[0];
      const deviceIdx = rawDailyHeaders.indexOf('Device');
      const convIdx = rawDailyHeaders.indexOf('Conv.');

      if (deviceIdx !== -1 && convIdx !== -1) {
          const rawDailyData = rawDailySheet.getRange(2, 1, rawDailySheet.getLastRow() - 1, rawDailyHeaders.length).getValues();
          const deviceConvAggregated = {};
          rawDailyData.forEach(row => {
              const device = row[deviceIdx] || 'Unknown';
              if (!deviceConvAggregated[device]) deviceConvAggregated[device] = 0;
              deviceConvAggregated[device] += safeParseFloat_(row[convIdx]);
          });
          const deviceConvChartData = [['Device', 'Conversions']];
          for (const device in deviceConvAggregated) {
              if (deviceConvAggregated[device] > 0) {
                  deviceConvChartData.push([device, deviceConvAggregated[device]]);
              }
          }

          if (deviceConvChartData.length > 1) {
            tempSheet.clearContents().clearFormats();
            const tempRange = tempSheet.getRange(1, 1, deviceConvChartData.length, 2);
            tempRange.setValues(deviceConvChartData);
            SpreadsheetApp.flush();
            tempRange.offset(1, 1, deviceConvChartData.length - 1, 1).setNumberFormat('#,##0.00');


            const deviceConvPie = overviewSheet.newChart().setChartType(Charts.ChartType.PIE)
              .addRange(tempRange)
              .setPosition(chartStartRow, chartStartCol, 5, 5)
              .setOption('title', 'Conversions by Device (Overall)')
              .setOption('titleTextStyle', STYLE_CHART_TITLE)
              .setOption('legend', { position: 'right', textStyle: STYLE_CHART_LEGEND })
              .setOption('pieSliceText', 'percentage')
              .setOption('colors', ADOPTED_CHART_COLORS.DEVICE_CONVERSIONS_PIE)
              .setOption('width', chartWidth).setOption('height', chartHeight)
              .setOption('useFirstColumnAsDomain', true)
              .setOption('backgroundColor', '#FFFFFF')
              .build();
            overviewSheet.insertChart(deviceConvPie);
            chartsInCurrentRow++;
            logInfo_('Device Conversions Pie Chart created.');
          } else {
            overviewSheet.getRange(chartStartRow, chartStartCol).setValue('Not enough data for Device Conversions Pie Chart.');
          }
      } else {
          overviewSheet.getRange(chartStartRow, chartStartCol).setValue('Device Conversions Pie Chart: Missing required columns in raw data.');
      }
  } else {
      overviewSheet.getRange(chartStartRow, chartStartCol).setValue('Device Conversions Pie Chart: Raw daily data sheet missing or empty.');
  }
  if (chartsInCurrentRow >= chartsPerRow) { chartStartRow += (Math.floor(chartHeight/20) + 2); chartStartCol = 1; chartsInCurrentRow = 0; } else { chartStartCol += chartColOffset; }

  // --- Product Title Charts ---
  if (rawProductSheet && rawProductSheet.getLastRow() >=2) {
      const rawProductHeaders = rawProductSheet.getRange(1, 1, 1, rawProductSheet.getLastColumn()).getValues()[0];
      const productTitleIdx_prod = rawProductHeaders.indexOf('Product Title');
      const convIdx_prod = rawProductHeaders.indexOf('Conv.');
      const convValIdx_prod = rawProductHeaders.indexOf('Conv. Value');
      const imprIdx_prod = rawProductHeaders.indexOf('Impr.');
      const clicksIdx_prod = rawProductHeaders.indexOf('Clicks');

      if ([productTitleIdx_prod, convIdx_prod, convValIdx_prod, imprIdx_prod, clicksIdx_prod].every(idx => idx !== -1)) {
          const rawProductData = rawProductSheet.getRange(2, 1, rawProductSheet.getLastRow() - 1, rawProductHeaders.length).getValues();
          const productAggData = {};

          rawProductData.forEach(row => {
              const title = row[productTitleIdx_prod] || 'N/A';
              if (!productAggData[title]) {
                  productAggData[title] = { conv: 0, convVal: 0, impr: 0, clicks: 0 };
              }
              productAggData[title].conv += safeParseFloat_(row[convIdx_prod]);
              productAggData[title].convVal += safeParseFloat_(row[convValIdx_prod]);
              productAggData[title].impr += safeParseInt_(row[imprIdx_prod]);
              productAggData[title].clicks += safeParseInt_(row[clicksIdx_prod]);
          });

          function createProductBarChart(chartTitle, data, valueKey, color, format, isPercentage, yAxisTitle) {
              let chartData = [['Product Title', yAxisTitle || chartTitle.split(' by ')[0]]];
              let sortedData = Object.entries(data);

              if (isPercentage) {
                   sortedData = sortedData.filter(([,d]) => d.clicks > 0)
                                      .map(([title, d]) => [title, div(d.conv, d.clicks)]) // CVR calculation
                                      .sort(([,a],[,b]) => b - a).slice(0,10);
              } else {
                  sortedData = sortedData.map(([title, d]) => [title, d[valueKey]])
                                     .sort(([,a],[,b]) => b - a).slice(0,10);
              }
              sortedData.forEach(arr => chartData.push(arr));


              if (chartData.length > 1) {
                  tempSheet.clearContents().clearFormats();
                  const tempRange = tempSheet.getRange(1, 1, chartData.length, 2);
                  tempRange.setValues(chartData);
                  SpreadsheetApp.flush();
                  tempRange.offset(1, 1, chartData.length - 1, 1).setNumberFormat(format);

                  const chart = overviewSheet.newChart().setChartType(Charts.ChartType.BAR)
                    .addRange(tempRange)
                    .setPosition(chartStartRow, chartStartCol, 5, 5)
                    .setOption('title', chartTitle)
                    .setOption('titleTextStyle', STYLE_CHART_TITLE)
                    .setOption('colors', [color])
                    .setOption('hAxis', { title: yAxisTitle || chartTitle.split(' by ')[0] , titleTextStyle: STYLE_CHART_AXIS_TITLE, textStyle: STYLE_CHART_AXIS_LABEL, format: format })
                    .setOption('vAxis', { title: 'Product Title', titleTextStyle: STYLE_CHART_AXIS_TITLE, textStyle: STYLE_CHART_AXIS_LABEL })
                    .setOption('legend', { position: 'none' })
                    .setOption('width', chartWidth).setOption('height', chartHeight)
                    .setOption('useFirstColumnAsDomain', true)
                    .setOption('backgroundColor', '#FFFFFF')
                    .build();
                  overviewSheet.insertChart(chart);
                  chartsInCurrentRow++;
                  logInfo_(`${chartTitle} created.`);
              } else {
                  overviewSheet.getRange(chartStartRow, chartStartCol).setValue(`Not enough data for ${chartTitle}.`);
              }
              if (chartsInCurrentRow >= chartsPerRow) { chartStartRow += (Math.floor(chartHeight/20) + 2); chartStartCol = 1; chartsInCurrentRow = 0; } else { chartStartCol += chartColOffset; }
          }

          createProductBarChart('Top 10 Products by Conversions (Overall)', productAggData, 'conv', ADOPTED_CHART_COLORS.PRODUCT_CONVERSIONS, '#,##0.00', false, 'Conversions');
          createProductBarChart('Top 10 Products by Conversion Value (Overall)', productAggData, 'convVal', ADOPTED_CHART_COLORS.PRODUCT_CONV_VALUE, getCurrencyFormatString_(currencyCode), false, 'Conversion Value');
          createProductBarChart('Top 10 Products by Impressions (Overall)', productAggData, 'impr', ADOPTED_CHART_COLORS.IMPRESSIONS, '#,##0', false, 'Impressions');
          createProductBarChart('Top 10 Products by Conversion Rate (Overall)', productAggData, 'cvr', ADOPTED_CHART_COLORS.CONV_RATE, '0.00%', true, 'Conversion Rate');


      } else {
          overviewSheet.getRange(chartStartRow, 1).setValue('Product Charts: Missing required columns in raw product data.'); chartStartRow += 2;
      }
  } else {
      overviewSheet.getRange(chartStartRow, 1).setValue('Product Charts: Raw product data sheet missing or empty.'); chartStartRow += 2;
  }
   if (chartsInCurrentRow > 0 && chartsInCurrentRow < chartsPerRow) { // If the last row of charts wasn't full, still advance row
       chartStartRow += (Math.floor(chartHeight/20) + 2);
       chartStartCol = 1;
       chartsInCurrentRow = 0;
   }


  // --- Chart: Daily Conversion Value vs. Spend ---
  if (rawDailySheet && rawDailySheet.getLastRow() >=2) {
      const rawDailyHeaders = rawDailySheet.getRange(1, 1, 1, rawDailySheet.getLastColumn()).getValues()[0];
      const dateIdx = rawDailyHeaders.indexOf('Date');
      const costIdx = rawDailyHeaders.indexOf('Cost');
      const convValIdx = rawDailyHeaders.indexOf('Conv. Value');

      if ([dateIdx, costIdx, convValIdx].every(idx => idx !== -1)) {
          const rawDailyData = rawDailySheet.getRange(2, 1, rawDailySheet.getLastRow() - 1, rawDailyHeaders.length).getValues();
          const trendChartDataAggregated = {};
          rawDailyData.forEach(row => {
            const dateStr = Utilities.formatDate(new Date(row[dateIdx]), AdsApp.currentAccount().getTimeZone(), "yyyy-MM-dd");
            if (!trendChartDataAggregated[dateStr]) {
              trendChartDataAggregated[dateStr] = {
                date: new Date(row[dateIdx]), cost: 0, convValue: 0
              };
            }
            trendChartDataAggregated[dateStr].cost += safeParseFloat_(row[costIdx]);
            trendChartDataAggregated[dateStr].convValue += safeParseFloat_(row[convValIdx]);
          });

          const convValSpendTrendData = [['Date', 'Conversion Value', 'Spend']];
          Object.keys(trendChartDataAggregated).sort().forEach(dateStr => {
              const dayData = trendChartDataAggregated[dateStr];
              convValSpendTrendData.push([dayData.date, dayData.convValue, dayData.cost]);
          });

          if (convValSpendTrendData.length > 1) {
              tempSheet.clearContents().clearFormats();
              const tempRange = tempSheet.getRange(1, 1, convValSpendTrendData.length, 3);
              tempRange.setValues(convValSpendTrendData);
              SpreadsheetApp.flush();
              tempRange.offset(1, 0, convValSpendTrendData.length - 1, 1).setNumberFormat('yyyy-mm-dd');
              tempRange.offset(1, 1, convValSpendTrendData.length - 1, 1).setNumberFormat(getCurrencyFormatString_(currencyCode));
              tempRange.offset(1, 2, convValSpendTrendData.length - 1, 1).setNumberFormat(getCurrencyFormatString_(currencyCode));

              const convValSpendChart = overviewSheet.newChart().setChartType(Charts.ChartType.LINE)
                  .addRange(tempRange)
                  .setPosition(chartStartRow, 1, 5, 5) // Start this chart in column 1 of its row
                  .setOption('title', 'Daily Conversion Value vs. Spend')
                  .setOption('titleTextStyle', STYLE_CHART_TITLE)
                  .setOption('series', {
                      0: { color: ADOPTED_CHART_COLORS.DAILY_CONV_VALUE, targetAxisIndex: 0, labelInLegend: 'Conversion Value' },
                      1: { color: ADOPTED_CHART_COLORS.DAILY_SPEND, targetAxisIndex: 0, labelInLegend: 'Spend' }
                  })
                  .setOption('vAxes', { 0: { title: 'Amount (' + currencyCode + ')', titleTextStyle: STYLE_CHART_AXIS_TITLE, textStyle: STYLE_CHART_AXIS_LABEL, format: getCurrencyFormatString_(currencyCode) }})
                  .setOption('hAxis', { title: 'Date', titleTextStyle: STYLE_CHART_AXIS_TITLE, textStyle: STYLE_CHART_AXIS_LABEL, format: 'MMM d' })
                  .setOption('legend', { position: 'top', alignment: 'center', textStyle: STYLE_CHART_LEGEND })
                  .setOption('width', 800).setOption('height', 400) // Make this chart wider as it's the last one in its "row"
                  .setOption('useFirstColumnAsDomain', true)
                  .setOption('backgroundColor', '#FFFFFF')
                  .build();
              overviewSheet.insertChart(convValSpendChart);
              logInfo_('Daily Conv Value vs Spend Chart created.');
          } else {
              overviewSheet.getRange(chartStartRow, 1).setValue('Not enough data for Daily Conv Value vs Spend Chart.');
          }
      } else {
          overviewSheet.getRange(chartStartRow, 1).setValue('Daily Conv Value vs Spend Chart: Missing required columns in raw data.');
      }
  } else {
      overviewSheet.getRange(chartStartRow, 1).setValue('Daily Conv Value vs Spend Chart: Raw daily data sheet missing or empty.');
  }


  const finalTempSheet = ss.getSheetByName("_ChartDataSource_Temp");
  if (finalTempSheet && !finalTempSheet.isSheetHidden()) {
    // ss.deleteSheet(finalTempSheet); // Keep hidden
  }
}
// --- End CORE_DASHBOARD_CONSTRUCTION ---

// --- Section 9: MAIN_ORCHESTRATION_FLOW ---

function main() {
  const scriptStartTime = new Date();
  logInfo_('--------------------------------------------------------------------');
  logInfo_(`Script "Google Ads E-commerce Dashboard v${SCRIPT_VERSION}" started at ${scriptStartTime.toLocaleString()}.`);

  let accountCustomerData = null;
  let accountCurrency = 'USD';

  try {
    initializeScriptEnvironment_();

    logInfo_(`Reporting Window: ${REPORT_START_DATE} to ${REPORT_END_DATE}`);

    if (TARGET_CUSTOMER_ID && TARGET_CUSTOMER_ID.trim() !== '') {
      try {
        const targetAccount = AdsManagerApp.accounts().withIds([TARGET_CUSTOMER_ID.replace(/-/g, '')]).get().next();
        AdsManagerApp.select(targetAccount);
        logInfo_(`Successfully switched to MCC client account: ${TARGET_CUSTOMER_ID} (${AdsApp.currentAccount().getName()})`);
      } catch (e) {
        logError_('main (MCC Switch)', e, `Failed to switch to CID: ${TARGET_CUSTOMER_ID}`);
        throw e;
      }
    }
    logInfo_(`Operating on Ads Account: ${AdsApp.currentAccount().getName()} (${AdsApp.currentAccount().getCustomerId()})`);

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    SpreadsheetApp.setActiveSpreadsheet(spreadsheet);

    logInfo_('Starting sheet cleanup: Deleting existing report-related tabs...');
    const allSheetsInSpreadsheet = spreadsheet.getSheets();
    const tempPlaceholderSheetName = TEMP_SHEET_PREFIX + Date.now();
    const tempPlaceholderSheet = spreadsheet.insertSheet(tempPlaceholderSheetName, 0);

    for (let i = 0; i < allSheetsInSpreadsheet.length; i++) {
      const currentSheetName = allSheetsInSpreadsheet[i].getName();
      if (currentSheetName !== ERROR_SHEET_NAME &&
          currentSheetName !== tempPlaceholderSheetName) {
        try {
          spreadsheet.deleteSheet(allSheetsInSpreadsheet[i]);
          logVerbose_(`Deleted existing sheet: "${currentSheetName}"`);
        } catch (e) {
          logError_('main (Sheet Cleanup)', e, `Failed to delete sheet: ${currentSheetName}`);
        }
      }
    }
    SpreadsheetApp.flush();
    logInfo_('Sheet cleanup complete.');

    let errorSheetForCheck = spreadsheet.getSheetByName(ERROR_SHEET_NAME);
    if (!errorSheetForCheck) {
      errorSheetForCheck = spreadsheet.insertSheet(ERROR_SHEET_NAME, 0);
      const headers = ['Timestamp', 'Function', 'Error Message', 'Context / GAQL', 'Stack Trace (Verbose)'];
      const headerRange = errorSheetForCheck.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      styleHeader_(headerRange);
      errorSheetForCheck.setFrozenRows(1);
      logInfo_(`Ensured ${ERROR_SHEET_NAME} tab exists.`);
    }


    initializeRawDataSheets_(spreadsheet);

    const accountConfigResult = fetchAccountConfigData_();

    if (accountConfigResult && accountConfigResult.length > 0) {
      accountCustomerData = accountConfigResult[0].customer;
      accountCurrency = accountCustomerData.currencyCode;
      logInfo_(`Account Config fetched: Currency=${accountCurrency}, TimeZone=${accountCustomerData.timeZone}`);
    } else {
      logError_('main', new Error('Failed to fetch account configuration. Using defaults.'));
      try {
          accountCurrency = AdsApp.currentAccount().getCurrencyCode();
          accountCustomerData = {
              descriptiveName: AdsApp.currentAccount().getName(),
              timeZone: AdsApp.currentAccount().getTimeZone(),
              currencyCode: accountCurrency
          };
          logInfo_(`Used fallback account info: Currency=${accountCurrency}, TimeZone=${accountCustomerData.timeZone}`);
      } catch(e) {
          logError_('main', e, 'Failed to get fallback account info.');
      }
    }

    logInfo_('Starting raw data fetch and write operations for all defined reports...');
    RAW_SHEET_DEFINITIONS_.forEach(definition => {
      if (!checkRemainingTimeGraceful_()) {
        logError_('main (Raw Data Loop)', new Error('Low execution time remaining.'), `Skipping: ${definition.name}`);
        throw new Error('Exiting due to low time before raw data processing for: ' + definition.name);
      }

      if (typeof definition.enabled === 'function' && !definition.enabled()) {
        logInfo_(`Skipping data processing for: ${definition.name} (feature processing disabled).`);
        return;
      }

      logInfo_(`Processing raw data for: ${definition.name}`);
      try {
        definition.writerFn(spreadsheet, accountCurrency);
      } catch (e) {
        logError_('main (Raw Data Loop)', e, `Error processing raw data for: ${definition.name}. Script will attempt to continue.`);
      }
    });
    logInfo_('All raw data fetch and write operations attempted.');

    populateUserGuideTab_(spreadsheet);

    if (accountCustomerData) {
      writeAccountConfigSheet_(spreadsheet, accountCustomerData, accountCurrency);
    } else {
      const settingsSheet = ensureSheet_(spreadsheet, SETTINGS_CONFIG_TAB_NAME, false);
      if(settingsSheet) settingsSheet.getRange('A1').setValue('ERROR: Account configuration data was not available to fully populate this sheet.');
    }
    populateOverviewDashboard_(spreadsheet, accountCurrency);
    populateCampaignPerformanceDashboardTab_(spreadsheet, accountCurrency);
    populateProductPerformanceAnalysisDashboardTab_(spreadsheet, accountCurrency);

    if (ENABLE_ASSET_PERFORMANCE_TAB) {
      populateAssetPerformanceDashboardTab_(spreadsheet, accountCurrency);
    }
    if (ENABLE_IMPRESSION_SHARE_TAB) {
      populateShareRankTab_(spreadsheet, accountCurrency);
    }
    if (ENABLE_BUDGET_PACING_TAB) {
      populateBudgetPacingTab_(spreadsheet, accountCurrency);
    }
    if (ENABLE_CONV_LAG_ADJUST_TABS) {
      populateConversionLagTabs_(spreadsheet, accountCurrency);
    }
    populateRecommendationsDashboardTab_(spreadsheet);

    createOverviewDashboardCharts_(spreadsheet, accountCurrency);
    logInfo_('All dashboard tabs populated and charts created.');

    const dashboardTabsToFinalFormat = [
      USER_GUIDE_TAB_NAME, SETTINGS_CONFIG_TAB_NAME, OVERVIEW_TAB_NAME,
      CAMPAIGN_PERF_TAB_NAME, PRODUCT_ANALYSIS_TAB_NAME, RECOMMENDATIONS_TAB_NAME
    ];
    if(ENABLE_ASSET_PERFORMANCE_TAB && PROCESS_ASSET_PERFORMANCE_DATA) dashboardTabsToFinalFormat.push(ASSET_PERFORMANCE_TAB_NAME);
    if(ENABLE_IMPRESSION_SHARE_TAB) dashboardTabsToFinalFormat.push(SHARE_RANK_TAB_NAME);
    if(ENABLE_BUDGET_PACING_TAB) dashboardTabsToFinalFormat.push(BUDGET_PACING_TAB_NAME);
    if(ENABLE_CONV_LAG_ADJUST_TABS) dashboardTabsToFinalFormat.push(LAG_FORECAST_TAB_NAME);

    dashboardTabsToFinalFormat.forEach(function(name) {
      const sheetToProcess = spreadsheet.getSheetByName(name);
      if (sheetToProcess) {
        centerAlignSheetData(sheetToProcess);
        logVerbose_(`Applied final formatting touches for tab: ${name}`);
      }
    });


    try {
      spreadsheet.deleteSheet(tempPlaceholderSheet);
      logInfo_(`Deleted temporary placeholder sheet: ${tempPlaceholderSheetName}`);
    } catch (e) {
      logError_('main (Final Cleanup)', e, `Failed to delete final temporary placeholder sheet: ${tempPlaceholderSheetName}`);
    }

    const errorSheet = spreadsheet.getSheetByName(ERROR_SHEET_NAME);
    if (errorSheet && errorSheet.getIndex() > 1) {
        spreadsheet.setActiveSheet(errorSheet);
        spreadsheet.moveActiveSheet(1);
    }
    const userGuideSheet = spreadsheet.getSheetByName(USER_GUIDE_TAB_NAME);
    if (userGuideSheet) {
        const targetIndex = errorSheet ? 2 : 1;
        if (userGuideSheet.getIndex() > targetIndex) {
            spreadsheet.setActiveSheet(userGuideSheet);
            spreadsheet.moveActiveSheet(targetIndex);
        }
    }


    const scriptEndTime = new Date();
    const durationMs = scriptEndTime.getTime() - scriptStartTime.getTime();
    const durationStr = `${Math.floor(durationMs / 60000)}m ${Math.floor((durationMs % 60000) / 1000)}s`;

    if (errorSheet && errorSheet.getLastRow() <= 1) {
      errorSheet.appendRow([Utilities.formatDate(scriptEndTime, AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd HH:mm:ss z'),
        'SCRIPT_EXECUTION', 'SUCCESS: Script completed without logging errors to this sheet.', `Duration: ${durationStr}`, SCRIPT_VERSION
      ]);
    } else if (errorSheet) {
      errorSheet.insertRowAfter(1);
      errorSheet.getRange(2, 1, 1, 5).setValues([[Utilities.formatDate(scriptEndTime, AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd HH:mm:ss z'),
        'SCRIPT_EXECUTION', 'COMPLETED WITH ERRORS: Please review errors logged above.', `Duration: ${durationStr}`, SCRIPT_VERSION
      ]]);
    }
    if (errorSheet) autoResizeSheetColumns_(errorSheet);


    logInfo_(`Script "Google Ads E-commerce Dashboard v${SCRIPT_VERSION}" completed at ${scriptEndTime.toLocaleString()}. Duration: ${durationStr}`);
    logInfo_(`Dashboard URL: ${spreadsheet.getUrl()}`);
    logInfo_('--------------------------------------------------------------------');

  } catch (e) {
    logError_('main (CRITICAL SCRIPT FAILURE)', e, 'Script execution halted due to a critical error.');
    try {
      const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      let errorSheet = spreadsheet.getSheetByName(ERROR_SHEET_NAME);
      if (!errorSheet) {
        errorSheet = spreadsheet.insertSheet(ERROR_SHEET_NAME, 0);
        const headers = ['Timestamp', 'Function', 'Error Message', 'Context / GAQL', 'Stack Trace (Verbose)'];
        const headerRange = errorSheet.getRange(1, 1, 1, headers.length);
        headerRange.setValues([headers]);
        styleHeader_(headerRange);
        errorSheet.setFrozenRows(1);
      }
      errorSheet.insertRowAfter(1);
      errorSheet.getRange(2, 1, 1, 5).setValues([[
        Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd HH:mm:ss z'),
        'main_CRITICAL_FAILURE_HANDLER', e.message, 'Global script failure point.', e.stack || 'N/A'
      ]]);
       if (errorSheet) autoResizeSheetColumns_(errorSheet);
    } catch (sheetErr) {
      Logger.log(`[CRITICAL] Also failed to write CRITICAL SCRIPT FAILURE to sheet: ${sheetErr.message}`);
    }
  }
}
// --- End MAIN_ORCHESTRATION_FLOW ---
