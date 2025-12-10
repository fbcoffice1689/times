// --- CORE CONSTANTS ---
const REMINDER_LIST_KEY = 'customReminderList';
const REMINDER_ENABLED_KEY = 'isReminderEnabled';
const CLOCK_IN_KEY = 'isClockedIn'; // Stores the overall clock status (IN/OUT)
const ON_BREAK_KEY = 'isUserOnBreak'; // NEW: Stores the break status (true/false)
const POPUP_TITLE = "‼️ Important Sheet Reminders ‼️"; // Retain for completeness, even if not used.
const LOG_SHEET_NAME = 'logs'; // The sheet name for time logging

// Default reminder list structure
const DEFAULT_LIST_JSON = JSON.stringify([
  {text: "Update Monthly Sales Figures.", isCompleted: false, targetDate: null},
  {text: "Verify currency exchange rates.", isCompleted: false, targetDate: null},
  {text: "Send final report to management.", isCompleted: false, targetDate: null}
]);

// --- LIFECYCLE FUNCTIONS ---

function onOpen() {
  const scriptProperties = PropertiesService.getScriptProperties();

  // 1. Ensure default properties are set up
  initializeProperties(scriptProperties);
  
  // 2. Setup the "logs" sheet if it doesn't exist
  setupLogsSheet();

  SpreadsheetApp.getUi()
      .createMenu('⏰ Time Clock') // Name of the custom menu
      .addItem('Open Time Clock', 'showSidebar') // Menu item
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Hybrid Time & Tasks')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * HANDLES WEB APP LAUNCH: Required entry point for the deployed web app URL.
 */
function doGet() {
  // 1. Get the content of the HTML file
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Index'); 
  // 2. Set the necessary parameters for embedding and responsive design
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  htmlOutput.setSandboxMode(HtmlService.SandboxMode.IFRAME);

  return htmlOutput;
}  

// --- GAS UTILITY FUNCTIONS (Called by onOpen) ---

/**
 * Sets up the default values for the reminder list, enabled state, and clock status.
 */
function initializeProperties(scriptProperties) {
  const reminderListValue = scriptProperties.getProperty(REMINDER_LIST_KEY);
  
  // Ensure REMINDER_LIST_KEY is never null or an empty string
  if (reminderListValue === null || reminderListValue === "") {
    scriptProperties.setProperty(REMINDER_LIST_KEY, DEFAULT_LIST_JSON);
  }
  
  if (scriptProperties.getProperty(REMINDER_ENABLED_KEY) === null) {
    scriptProperties.setProperty(REMINDER_ENABLED_KEY, 'true');
  }
  if (scriptProperties.getProperty(CLOCK_IN_KEY) === null) {
    scriptProperties.setProperty(CLOCK_IN_KEY, 'false');
  }
  // NEW: Initialize break status
  if (scriptProperties.getProperty(ON_BREAK_KEY) === null) {
    scriptProperties.setProperty(ON_BREAK_KEY, 'false');
  }
}

/**
 * Ensures the 'logs' sheet exists and has the correct headers.
 */
function setupLogsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    const headers = ["Timestamp", "Action", "Status"]; 
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
}

// --- REPORTING UTILITY (Required by calculateReportFromSyncedLogs) ---

/**
 * Utility function to convert the report data duration from milliseconds.
 * @param {number} totalMilliseconds The duration in milliseconds.
 * @returns {string} The duration formatted as HH:mm:ss.
 */
function formatMsToDuration(totalMilliseconds) {
  const totalSeconds = Math.floor(totalMilliseconds / 1000);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

// --- FIREBASE SYNC FUNCTIONS (Called by index.html via google.script.run) ---

/**
 * Receives the final, aggregated report data from the client and writes it to a Sheet.
 * Called by handleMarkReported() in index.html.
 * @param {Object} data - Contains cutoffTimestamp, totalWorkHours, totalBreakHours
 */
function recordFinalReport(data) {
  const sheetName = "Aggregated Reports";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // Set headers on a newly created sheet
    const headers = ["Report Date", "User ID", "Work Hours", "Break Hours", "Cutoff Timestamp"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setFontWeight("bold").setBackground("#d9ead3");
  }
  
  // Format the report date
  const reportDate = new Date(data.cutoffTimestamp).toLocaleDateString();
  const currentUserId = Session.getTemporaryActiveUserKey(); 

  const rowData = [
    reportDate,
    currentUserId,
    data.totalWorkHours,
    data.totalBreakHours,
    data.cutoffTimestamp
  ];
  
  sheet.appendRow(rowData);
  
  Logger.log(`Aggregated Report recorded for user ${currentUserId} at ${reportDate}.`);
}

/**
 * Receives raw log entries as a JSON string from the client, 
 * safely parses it, and writes the data to the 'Time Logs Report (Raw Sync)' sheet.
 * Called by syncRawLogsToSheet() in index.html.
 * @param {string} rawDataJsonString A JSON string representing the array of log entries.
 */
function writeRawLogsFromJson(rawDataJsonString) {
  const sheetName = "Time Logs Report (Raw Sync)";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // 1. Parse the JSON string into a JavaScript array
  let rawData;
  try {
    rawData = JSON.parse(rawDataJsonString);
  } catch (e) {
    Logger.log("FATAL ERROR: Could not parse JSON data from client. " + e.toString());
    return;
  }

  if (rawData.length === 0) {
    sheet.clear();
    sheet.getRange('A1').setValue("No Log Entries Found in Firestore.");
    return;
  }
  
  // 2. Map the data, enforcing all values are pure strings or safely-created Dates.
  const processedData = rawData.map(row => {
    // Row structure is expected to be [Timestamp String, Action String, Status String]
    const timestampString = String(row[0] || 'CORRUPT_TIMESTAMP');
    const actionString = String(row[1] || 'ACTION_MISSING');        
    const statusString = String(row[2] || '');                      
    
    let timestampValue;
    try {
      // Attempt to create a Date object on the server.
      const dateObject = new Date(timestampString);
      
      // If conversion is valid, use the Date object; otherwise, use the string.
      timestampValue = isNaN(dateObject.getTime()) ? timestampString : dateObject;
      
    } catch (e) {
      // Final fallback to the original string if parsing fails
      timestampValue = timestampString;
    }

    // Return the final clean array.
    return [
      timestampValue, 
      actionString, 
      statusString
    ];
  }).filter(row => row !== null); 

  if (processedData.length === 0) {
      sheet.clear();
      sheet.getRange('A1').setValue("All log entries were corrupted or invalid.");
      return;
  }
  
  // Define header row
  const headers = ["Timestamp", "Action", "Status"];
  
  // Clear existing content and set headers
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#d9ead3");
  
  // 3. Write data to the sheet
  const startRow = 2;
  const numRows = processedData.length;
  const numCols = 3; 
  
  sheet.getRange(startRow, 1, numRows, numCols).setValues(processedData);
  
  Logger.log(`Successfully synced ${numRows} raw log entries from client via JSON.`);
}

// --- REPORTING FUNCTION FOR GOOGLE SHEETS ---

/**
 * Custom function that reads the raw data from the 'Time Logs Report (Raw Sync)' 
 * sheet and calculates the final work sessions and total time between two timestamps.
 * * @param {Date|null} startTime The start timestamp (e.g., from a spreadsheet cell).
 * @param {Date|null} endTime The end timestamp (e.g., from a spreadsheet cell).
 * @returns {Array<Array<any>>} A two-dimensional array of report data (5 columns).
 */
function calculateReportFromSyncedLogs(startTime, endTime) {
  const SYNC_SHEET_NAME = "Time Logs Report (Raw Sync)";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SYNC_SHEET_NAME);
  const tz = ss.getSpreadsheetTimeZone();
  
  // Fallback if the sheet hasn't been created yet
  if (!sheet) return [["Error: Sync sheet not found. Run the 'Sync Full Log' button first."]];
  
  // 1. Establish the filter boundaries
  // Use 0 (start of time) if startTime is null, and Now if endTime is null.
  let startFilterTimeMs = (startTime instanceof Date) ? startTime.getTime() : 0;
  let endFilterTimeMs = (endTime instanceof Date) ? endTime.getTime() : new Date().getTime();

  // Ensure start time is not greater than end time (swap if needed)
  if (startFilterTimeMs > endFilterTimeMs) {
      [startFilterTimeMs, endFilterTimeMs] = [endFilterTimeMs, startFilterTimeMs];
  }
  
  // Read all data from the synced sheet, skipping headers (Row 1)
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [["No data available in synced sheet."]];

  // 2. Filter the raw data based on the time window
  // Note: Action column is index 1
  const logsToProcess = data.slice(1).filter(row => {
      const timestamp = row[0];
      
      // Ensure the timestamp is a valid Date object for comparison
      if (!(timestamp instanceof Date) || isNaN(timestamp.getTime())) {
          return false;
      }
      
      const timeMs = timestamp.getTime();
      return timeMs >= startFilterTimeMs && timeMs <= endFilterTimeMs;
  });
  
  let sessions = []; // Format: [Date, Time IN, Time OUT, Duration]
  let sessionStartTime = null; 
  let sessionStartLog = null; 
  let grandTotalDurationMs = 0; 
  
  // 3. Process logs looking for sequential IN/OUT cycles within the filtered data
  for (let i = 0; i < logsToProcess.length; i++) {
    const timestamp = logsToProcess[i][0];
    const action = logsToProcess[i][1];
    
    // Ensure timestamp is a valid date object (filtered above, but for safety)
    if (!(timestamp instanceof Date)) {
      continue;
    }

    // Note: The Firebase client uses CLOCK_IN and CLOCK_OUT, and BREAK_START/BREAK_END.
    // This logic only tracks CLOCK_IN/CLOCK_OUT cycles, intentionally ignoring break logs,
    // which aligns with the original code's approach to clean up reporting logic.
    if (action === 'CLOCK_IN') {
      // Start of a new work segment/block
      sessionStartTime = timestamp;
      sessionStartLog = timestamp; 
    } else if (action === 'CLOCK_OUT') {
      if (sessionStartTime && sessionStartLog) {
        // Calculate work duration
        const totalWorkDuration = timestamp.getTime() - sessionStartTime.getTime();
        
        grandTotalDurationMs += totalWorkDuration;

        // Col 1: Date
        const dateString = Utilities.formatDate(sessionStartLog, tz, 'MMM dd, yyyy');
        // Col 2: Clock In Time
        const timeIn = Utilities.formatDate(sessionStartLog, tz, 'HH:mm:ss');
        // Col 3: Clock Out Time
        const timeOut = Utilities.formatDate(timestamp, tz, 'HH:mm:ss');
        // Col 4: Duration
        const durationString = formatMsToDuration(totalWorkDuration);
        
        sessions.push([dateString, timeIn, timeOut, durationString]);
      }
      // Reset for the next cycle
      sessionStartTime = null;
      sessionStartLog = null;
    }
  }
  
  // --- Final Output Formatting ---
  const reportHeaders = [["", "Date", "Time IN", "Time OUT", "Total Duration"]];
  
  if (sessions.length === 0) {
    return reportHeaders.concat([["", "No completed IN/OUT sessions found in synced data.", "", "", ""]]);
  }
  
  // 1. Map the 4-column session data back to the 5-column structure (adding the dummy empty column)
  const sheetSessions = sessions.map(session => ["", ...session]);

  // 2. Return headers and sessions
  return reportHeaders.concat(sheetSessions);
}