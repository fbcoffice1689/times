// --- CORE CONSTANTS ---
const REMINDER_LIST_KEY = 'customReminderList';
const REMINDER_ENABLED_KEY = 'isReminderEnabled';
const CLOCK_IN_KEY = 'isClockedIn'; // Stores the overall clock status (IN/OUT)
const ON_BREAK_KEY = 'isUserOnBreak'; // NEW: Stores the break status (true/false)
const POPUP_TITLE = "‼️ Important Sheet Reminders ‼️";
const LOG_SHEET_NAME = 'logs'; // The sheet name for time logging

// Default reminder list structure
const DEFAULT_LIST_JSON = JSON.stringify([
  {text: "Update Monthly Sales Figures.", isCompleted: false, targetDate: null},
  {text: "Verify currency exchange rates.", isCompleted: false, targetDate: null},
  {text: "Send final report to management.", isCompleted: false, targetDate: null}
]);

/**
 * HANDLES WEB APP LAUNCH: Required entry point for the deployed web app URL.
 * It serves the content of the 'Index.html' file (your new time clock app).
 * * @returns {HtmlOutput} The compiled HTML template of the Index.html file.
 */
function doGet() {
  // 1. Get the content of the HTML file
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Index'); // Use 'Index' (or your file name)

  // 2. Set the necessary parameters for embedding and responsive design
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  htmlOutput.setSandboxMode(HtmlService.SandboxMode.IFRAME);

  // Set initial dimensions (GAS often defaults to a small size if not set)
  htmlOutput.setWidth(1000); 
  htmlOutput.setHeight(800);

  return htmlOutput;
}

// --- Code.gs file ---

/**
 * Universal gateway to handle POST requests from external hosts.
 */
function doPost(e) {
  try {
    // Content-Type is set to text/plain by the external fetch request
    const request = JSON.parse(e.postData.contents); 
    const functionName = request.functionName;
    const parameters = request.parameters || [];

    // Ensure the function exists and execute it
    if (this[functionName] && typeof this[functionName] === 'function') {
      const result = this[functionName].apply(this, parameters);
      
      // Return success to the client
      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'success', 
        result: result || 'Function executed successfully.' 
      }))
      .setMimeType(ContentService.MimeType.JSON);
    } else {
      throw new Error(`Function ${functionName} not found or is not accessible.`);
    }
  } catch (error) {
    // Return error to the client
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'error', 
      message: error.message 
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles the client data sent from the external host.
 * This is the function called via the doPost gateway.
 */
function recordFinalReport(data) {
  const sheetName = "Aggregated Reports";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ... (Your implementation of writing data to the sheet) ...
  
  // Example implementation (ensure your Sheet exists and is active)
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = ["Report Date", "User ID", "Work Hours", "Break Hours", "Cutoff Timestamp"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#d9ead3");
  }
  
  const reportDate = new Date(data.cutoffTimestamp).toLocaleDateString();
  const currentUserId = Session.getTemporaryActiveUserKey() || "EXTERNAL_HOST"; 

  const rowData = [
    reportDate,
    currentUserId,
    data.totalWorkHours,
    data.totalBreakHours,
    data.cutoffTimestamp
  ];
  
  sheet.appendRow(rowData);
  
  return `Synced ${data.totalWorkHours} hours.`;
}


/**
 * Fetches the current clock status and the active reminder list for the client UI.
 * This combines two data points into a single server call for efficiency.
 * @returns {object} An object containing the clock status and the JSON string of reminders.
 */
function getAppStartupData() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // CLOCK_IN_KEY and REMINDER_LIST_KEY are defined at the top of your Code.gs
  const isClockedIn = scriptProperties.getProperty(CLOCK_IN_KEY) === 'true';
  
  // Fetch the reminder list JSON string. Use the default list if nothing is set yet.
  const remindersJson = scriptProperties.getProperty(REMINDER_LIST_KEY) || DEFAULT_LIST_JSON; 

  // Return the combined object expected by TimeClock.html
  return {
    isClockedIn: isClockedIn,
    remindersJson: remindersJson
  };
}

/**
 * The main onOpen function. This is a simple trigger that runs automatically 
 * whenever a user opens the Google Sheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();

  // 1. Ensure default properties are set up for both Reminders and Clock Status
  initializeProperties(scriptProperties);
  
  // 2. Setup the "logs" sheet if it doesn't exist
  setupLogsSheet();

  // 3. Display Reminder Pop-up if enabled and due
  showReminderPopup(ui, scriptProperties);

  // 4. Create Custom Menus (Time Clock and Reminders)
  createCustomMenus(ui, scriptProperties);
}


// --- REMINDER FUNCTIONS ---

/**
 * Handles the logic for displaying the reminder pop-up.
 */
function showReminderPopup(ui, scriptProperties) {
  const isEnabled = scriptProperties.getProperty(REMINDER_ENABLED_KEY) === 'true';
  
  if (!isEnabled) return;
  
  try {
    const listJson = scriptProperties.getProperty(REMINDER_LIST_KEY);
    const allReminders = JSON.parse(listJson);
    const today = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    
    // Filter the reminders: incomplete AND (no date OR date is today or in the past)
    const remindersToShow = allReminders.filter(r => 
      !r.isCompleted && 
      (r.targetDate === null || r.targetDate <= today)
    );

    // Format the filtered list for the alert box
    const formattedMessage = remindersToShow
      .map((r, i) => {
        let dateMarker = '';
        if (r.targetDate) {
          if (r.targetDate === today) {
            dateMarker = ' (TODAY)';
          } else if (r.targetDate < today) {
            dateMarker = ' (PAST DUE)';
          }
        }
        return `${i + 1}. ${r.text}${dateMarker}`;
      })
      .join('\n');
    
    if (remindersToShow.length > 0) {
      ui.alert(POPUP_TITLE, formattedMessage, ui.ButtonSet.OK);
    }
  } catch (e) {
    Logger.log("Error in showReminderPopup: " + e.toString());
  }
}


/**
 * Creates the custom menus for both Reminders and Time Clock.
 */
function createCustomMenus(ui, scriptProperties) {
  // Reminder Menu (Existing Logic)
  const isReminderEnabled = scriptProperties.getProperty(REMINDER_ENABLED_KEY) === 'true';
  const toggleReminderText = isReminderEnabled ? '🔴 Disable Pop-up on Open' : '🟢 Enable Pop-up on Open';
  
  const reminderMenu = ui.createMenu('📝 Reminder Tools')
    .addItem('✏️ Edit Reminder List (Sidebar)', 'openReminderEditor')
    .addSeparator()
    .addItem(toggleReminderText, 'toggleReminder');

  // Time Clock Menu (NEW)
  const isClockedIn = scriptProperties.getProperty(CLOCK_IN_KEY) === 'true';
  const isUserOnBreak = scriptProperties.getProperty(ON_BREAK_KEY) === 'true'; // Get break status

  const clockMenu = ui.createMenu('⏱️ Time Clock');

  if (isClockedIn) {
    // Options available when clocked IN
    // NEW: Added Preview option
    clockMenu.addItem('👁️ Preview Report Email (Test Mode)', 'previewReportEmail'); 
    clockMenu.addSeparator();
    clockMenu.addItem('⏰ Clock OUT (Select Time)', 'clockInOutDialog');
    clockMenu.addItem('✅ Clock OUT and Report NOW (Send Email)', 'clockOutNowAndReport'); 

    // Break Toggle: Show Start Break or Resume Work
    const breakMenuItemText = isUserOnBreak ? '▶️ Resume Work (Break IN)' : '⏸️ Start Break (Break OUT)';
    clockMenu.addItem(breakMenuItemText, 'toggleBreak'); // Renamed function in GS

  } else {
    // Options available when clocked OUT
    clockMenu.addItem('👁️ Preview Report Email (Test Mode)', 'previewReportEmail'); 
    clockMenu.addSeparator();
    clockMenu.addItem('✅ Clock IN', 'clockInOutDialog');
    clockMenu.addSeparator();
    // Manual reporting option
    clockMenu.addItem('✅ Mark Hours as Reported (Clear Report)', 'markHoursReported');
  }
  
  // Combine and add to UI
  clockMenu.addToUi();
  reminderMenu.addToUi();
}


/**
 * Sets up the default values for the reminder list, enabled state, and clock status.
 */
function initializeProperties(scriptProperties) {
  const reminderListValue = scriptProperties.getProperty(REMINDER_LIST_KEY);
  
  // --- FIX: Ensure REMINDER_LIST_KEY is never null or an empty string ---
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


// --- TIME CLOCK FUNCTIONS ---

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

/**
 * Utility function to get the 'logs' sheet.
 */
function getLogSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
}

/**
 * Utility function to get the current clock status.
 */
function getClockStatus() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const isClockedIn = scriptProperties.getProperty(CLOCK_IN_KEY) === 'true';
  return isClockedIn;
}

/**
 * Function linked to the menu that opens the custom Clock In/Out dialog.
 */
function clockInOutDialog() {
  const html = HtmlService.createHtmlOutputFromFile('TimeClock')
      .setTitle('⏱️ Clock In / Out')
      .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Server-side function called by the sidebar to record the time log.
 * @param {string} dateTimeString The edited date and time string from the dialog.
 */
function recordTimeLog(dateTimeString) {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheet = getLogSheet();
  const isClockedIn = scriptProperties.getProperty(CLOCK_IN_KEY) === 'true';
  const newStatus = !isClockedIn;
  const action = newStatus ? 'IN' : 'OUT';

  // If clocking OUT, make sure to clear the break status
  if (!newStatus) {
      scriptProperties.setProperty(ON_BREAK_KEY, 'false');
  }

  // Parse the string into a Date object (Apps Script will handle timezone conversion)
  const logTime = new Date(dateTimeString);
  
  // Log the entry
  sheet.appendRow([logTime, action, '']);
  
  // Update Clock Status
  scriptProperties.setProperty(CLOCK_IN_KEY, newStatus.toString());
  
  // Rebuild the menu to reflect the new state immediately
  createCustomMenus(ui, scriptProperties);
  
  // Use Utilities.formatDate with 24-hour HH format based on sheet's timezone for clarity
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const displayTime = Utilities.formatDate(logTime, tz, 'yyyy-MM-dd HH:mm:ss');
  
  // Updated alert to use unambiguous 24-hour time based on sheet's timezone
  ui.alert(`Successfully Clocked ${action}!`, `Time recorded: ${displayTime}`, ui.ButtonSet.OK);
}

/**
 * Finds the timestamp of the last REPORTED action.
 * @returns {Date} The timestamp of the last REPORTED action, or the time the logs sheet was created.
 */
function getLastReportedTimestamp() {
  const sheet = getLogSheet();
  const data = sheet.getDataRange().getValues();
  
  // Look backwards for the last REPORTED marker
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === 'REPORTED' && data[i][0] instanceof Date) {
      return data[i][0]; // Return the timestamp of the last report
    }
  }
  
  // If no REPORTED marker is found, assume the log started when the sheet was created (Row 2 data start).
  if (data.length > 1 && data[1][0] instanceof Date) {
      return data[1][0];
  }
  
  // Fallback to a very old date or current date if no valid logs exist
  return new Date(0); 
}

/**
 * Calculates the duration from milliseconds.
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


/**
 * Utility function to convert the report data into an HTML string for email.
 * It includes only the daily summary and the grand total.
 * @param {{sessions: Array<Array<any>>, totalMs: number}} reportData The structured report data.
 * @returns {string} The daily summary formatted as an HTML table with inline CSS.
 */
function generateHtmlReportContent(reportData) {
  const sessions = reportData.sessions;
  
  if (sessions.length === 0) {
    return "<p style='font-style: italic; color: #555;'>* No new clock-in/out sessions to report since the last reset. *</p>";
  }
  
  const grandTotalDurationString = formatMsToDuration(reportData.totalMs);
  
  // 1. Aggregate total duration per day
  const dailyTotals = {}; // key: YYYY-MM-DD, value: total duration in MS

  sessions.forEach(session => {
    // Session format: [Date, Time IN, Time OUT, Duration String]
    const date = session[0]; 
    const durationString = session[3]; 
    
    // Convert duration string back to milliseconds for calculation
    const parts = durationString.split(':');
    const ms = (parseInt(parts[0], 10) * 3600 + parseInt(parts[1], 10) * 60 + parseInt(parts[2], 10)) * 1000;
    
    dailyTotals[date] = (dailyTotals[date] || 0) + ms;
  });
  
  // 2. Build the HTML Table
  let html = `
    <style>
      .report-container { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #ddd; border-radius: 8px; overflow: hidden; }
      .report-header { background-color: #007bff; color: white; padding: 15px; text-align: center; font-size: 1.2em; border-bottom: 2px solid #0056b3; }
      .report-table { width: 100%; border-collapse: collapse; margin-top: 15px; }
      .report-table th, .report-table td { border: 1px solid #eee; padding: 10px; text-align: left; }
      .report-table th { background-color: #f8f9fa; color: #333; font-weight: bold; border-bottom: 3px solid #dee2e6; }
      .report-table tr:nth-child(even) { background-color: #f2f2f2; }
      .grand-total-row { background-color: #e9ecef; font-weight: bold; }
      .total-time-cell { text-align: right; font-family: monospace; font-size: 1.1em;}
    </style>
    <div class="report-container">
        <div class="report-header">
            Daily Work Summary
        </div>
        <table class="report-table">
            <thead>
                <tr>
                    <th>Date</th>
                    <th class="total-time-cell">Total Daily Time</th>
                </tr>
            </thead>
            <tbody>
  `;
  
  // Add daily totals
  const sortedDates = Object.keys(dailyTotals).sort();
  sortedDates.forEach(date => {
    const dailyDurationString = formatMsToDuration(dailyTotals[date]);
    html += `
        <tr>
            <td>${date}</td>
            <td class="total-time-cell">${dailyDurationString}</td>
        </tr>
    `;
  });
  
  // Add Grand Total row
  html += `
            <tr class="grand-total-row">
                <td>GRAND TOTAL</td>
                <td class="total-time-cell">${grandTotalDurationString}</td>
            </tr>
        </tbody>
      </table>
    </div>
  `;
  
  return html;
}

/**
 * Function linked to the menu that runs the report calculation and displays the email content
 * without logging the OUT or REPORTED actions and without sending an email.
 */
function previewReportEmail() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // 1. Get reporting period info (DOES NOT LOG)
  const report = getReportData();
  const totalWorkDurationString = formatMsToDuration(report.totalMs);
  const lastReportedTime = getLastReportedTimestamp();
  const now = new Date();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const startDate = Utilities.formatDate(lastReportedTime, tz, 'yyyy-MM-dd HH:mm:ss');
  const endDate = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
  const recipients = "roseneedsgrace@gmail.com, will@dartagency.com";
  const subject = `Time Report: ${startDate.split(' ')[0]} to ${endDate.split(' ')[0]}`;

  // 2. Generate Daily Summary Report HTML Content
  const htmlReportContent = generateHtmlReportContent(report);
  
  // 3. Construct Full HTML Email Body for Preview (Including wrapper styles)
  let emailHtmlBody = `
    <html>
      <body style="font-family: Arial, sans-serif; color: #333; margin: 0; padding: 20px; background-color: #f4f4f4;">
        <div style="max-width: 650px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">

          <h1 style="color: #dc3545; text-align: center; border-bottom: 3px dashed #dc3545; padding-bottom: 10px; font-size: 1.5em;">PREVIEW MODE - TEST REPORT</h1>
          
          <p style="font-size: 0.9em; color: #6c757d; text-align: center; margin-bottom: 20px;">
            * This email has NOT been sent, and your time has NOT been logged as OUT/REPORTED. *
          </p>
          
          <p>Hello Will,</p>
          
          <div style="background-color: #e9f7ef; padding: 15px; border-radius: 8px; margin-bottom: 20px; border: 1px solid #c3e6cb;">
            <h2 style="color: #007bff; border-bottom: 2px solid #007bff; padding-bottom: 5px; font-size: 1.2em;">Report Details</h2>
            <p><strong>Recipients:</strong> ${recipients}</p>
            <p><strong>Subject:</strong> ${subject}</p>
            <p><strong>Reporting Period:</strong> ${startDate.split(' ')[0]} to ${endDate.split(' ')[0]}</p>
            <p><strong>Total Work Time:</strong> <span style="font-size: 1.2em; color: #28a745; font-weight: bold;">${totalWorkDurationString}</span></p>
          </div>
          
          <div style="background-color: #f0f0f0; padding: 10px; border-left: 5px solid #999; margin-bottom: 20px; font-style: italic; color: #555;">
            --- User Message Placeholder ---<br>
            [Your optional message will be inserted here when you click 'Send.']
          </div>
          
          ${htmlReportContent}
          
          <p style="margin-top: 30px; text-align: right; color: #6c757d;">Thank you!</p>
        </div>
      </body>
    </html>
  `;
  
  // 4. Create and display the HTML dialog
  const htmlTemplate = HtmlService.createTemplateFromFile('EmailPreview');
  htmlTemplate.emailHtmlBody = emailHtmlBody; // Pass the full HTML body
  
  const html = htmlTemplate.evaluate()
      .setTitle('Time Report Email Preview')
      .setWidth(650)
      .setHeight(500);
      
  ui.showModalDialog(html, 'Time Report Email Preview (Test Mode)');
}


/**
 * Function linked to the menu that opens the custom Email Prompt dialog.
 */
function clockOutNowAndReport() {
  const ui = SpreadsheetApp.getUi();
  if (getClockStatus() !== true) {
    ui.alert("Clock Out Error", "You are already Clocked OUT.", ui.ButtonSet.OK);
    return;
  }
  
  const html = HtmlService.createHtmlOutputFromFile('EmailPrompt')
      .setTitle('Send Final Report')
      .setWidth(400);
  // Using showModalDialog to keep the user focused on the action
  ui.showModalDialog(html, 'Send Final Report');
}

/**
 * Server-side function to handle the report generation, logging, and email sending.
 * This is called by the EmailPrompt dialog.
 * @param {string} additionalMessage The extra text provided by the user.
 */
function sendReportEmail(additionalMessage) {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheet = getLogSheet();
  
  // Critical check: ensure the user is still clocked in before proceeding
  if (scriptProperties.getProperty(CLOCK_IN_KEY) !== 'true') {
    ui.alert("Error", "Action cancelled: You were already Clocked OUT.", ui.ButtonSet.OK);
    return;
  }
  
  // Ensure break status is cleared on report/clock out
  scriptProperties.setProperty(ON_BREAK_KEY, 'false');

  const now = new Date();
  
  // 1. Log the 'OUT' entry
  sheet.appendRow([now, 'OUT', '']);
  scriptProperties.setProperty(CLOCK_IN_KEY, 'false'); // Set status to OUT
  
  // Get reporting period info
  const report = getReportData();
  const totalWorkDurationString = formatMsToDuration(report.totalMs);
  const lastReportedTime = getLastReportedTimestamp();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const startDate = Utilities.formatDate(lastReportedTime, tz, 'yyyy-MM-dd HH:mm:ss');
  const endDate = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
  
  // 2. Log the 'REPORTED' entry with status metadata (using the calculated total work time)
  const statusMessage = `${startDate} | ${endDate} | Total Work: ${totalWorkDurationString}`;
  sheet.appendRow([now, 'REPORTED', statusMessage]);

  // 3. Generate Daily Summary Report HTML Content
  const htmlReportContent = generateHtmlReportContent(report);
  
  // 4. Construct Email Content
  const recipients = "roseneedsgrace@gmail.com, rosemarypyles@icloud.com";
  const subject = `Time Report: ${startDate.split(' ')[0]} to ${endDate.split(' ')[0]}`;
  
  // Basic Text Fallback (good practice for MailApp)
  let textBody = `Hello,

Reporting Period: ${startDate.split(' ')[0]} to ${endDate.split(' ')[0]}
Total Work Time for this period: ${totalWorkDurationString}

${additionalMessage ? `User Message: ${additionalMessage.trim()}\n\n` : ''}

(See HTML table for daily breakdown)

Thank you!
`;

  // Full HTML Body
  let htmlBody = `
    <html>
      <body style="font-family: Arial, sans-serif; color: #333;">
        <p>Hello,</p>
        
        <div style="background-color: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px; border: 1px solid #e0e0e0;">
          <h2 style="color: #007bff; border-bottom: 2px solid #007bff; padding-bottom: 5px; margin-top: 0; font-size: 1.2em;">Work Hours Report</h2>
          <p><strong>Reporting Period:</strong> ${startDate.split(' ')[0]} to ${endDate.split(' ')[0]}</p>
          <p><strong>Total Work Time:</strong> <span style="font-size: 1.2em; color: #28a745; font-weight: bold;">${totalWorkDurationString}</span></p>
        </div>
        
        ${additionalMessage ? `<div style="background-color: #fff3cd; padding: 10px; border-left: 5px solid #ffc107; margin-bottom: 20px; font-style: italic;">
          <strong>User Message:</strong><br>${additionalMessage.trim().replace(/\n/g, '<br>')}
        </div>` : ''}
        
        ${htmlReportContent}
        
        <p style="margin-top: 30px;">Thank you!</p>
      </body>
    </html>
  `;
  
  // 5. Send Email
  try {
    // Note: MailApp.sendEmail requires authorization the first time it is run.
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      body: textBody, // Use the plain text body as a fallback
      htmlBody: htmlBody // Use the structured HTML body
    });
    
    // 6. Update Menu and Alert User
    createCustomMenus(ui, scriptProperties);
    
    ui.alert(
      'Report Sent & Clocked Out', 
      `Successfully Clocked OUT, marked hours as Reported, and sent the email to ${recipients}.\n\nPeriod: ${startDate} - ${endDate}\nTotal Work Time: ${totalWorkDurationString}`,
      ui.ButtonSet.OK
    );
  } catch(e) {
    ui.alert("Email Error", "Failed to send email. You might need to grant the script permission to send emails the first time you run this function. Error: " + e.message, ui.ButtonSet.OK);
    Logger.log("Email sending failed: " + e.message);
  }
}

/**
 * Function linked to the menu that toggles the break status.
 * Renamed from recordBreak to toggleBreak.
 */
function toggleBreak() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheet = getLogSheet();
  
  const isClockedIn = scriptProperties.getProperty(CLOCK_IN_KEY) === 'true';
  if (!isClockedIn) {
    ui.alert("Cannot Record Break", "You must be Clocked IN to start or resume work.", ui.ButtonSet.OK);
    return;
  }
  
  const isUserOnBreak = scriptProperties.getProperty(ON_BREAK_KEY) === 'true';
  const newBreakStatus = !isUserOnBreak;
  
  const action = newBreakStatus ? 'BREAK-OUT' : 'BREAK-IN';
  const statusMessage = newBreakStatus ? 'Break Started' : 'Work Resumed';

  // 1. Log the entry
  const now = new Date();
  sheet.appendRow([now, action, '']);

  // 2. Update Break Status
  scriptProperties.setProperty(ON_BREAK_KEY, newBreakStatus.toString());

  // 3. Update Menu and Alert User
  createCustomMenus(ui, scriptProperties);

  ui.alert(statusMessage, `${action} logged at ${now.toLocaleTimeString()}.\n\nLog: ${action}`, ui.ButtonSet.OK);
}

/**
 * Function linked to the menu that records a REPORTED marker to reset the report window.
 */
function markHoursReported() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getLogSheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  const isClockedIn = getClockStatus();
  
  if (isClockedIn) {
    ui.alert("Cannot Report Hours", "Please Clock OUT before marking hours as reported.", ui.ButtonSet.OK);
    return;
  }
  
  // Clear break status just in case
  scriptProperties.setProperty(ON_BREAK_KEY, 'false');
  
  const now = new Date();
  
  // Get reporting period info 
  const report = getReportData();
  const totalWorkDurationString = formatMsToDuration(report.totalMs);
  const lastReportedTime = getLastReportedTimestamp();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const startDate = Utilities.formatDate(lastReportedTime, tz, 'yyyy-MM-dd HH:mm:ss');
  const endDate = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
  
  // Log the REPORTED entry with status metadata (using the calculated total work time)
  const statusMessage = `${startDate} | ${endDate} | Total Work: ${totalWorkDurationString}`;
  sheet.appendRow([now, 'REPORTED', statusMessage]);

  // Update the alert to reflect the new function name
  ui.alert('Hours Reported', `A 'REPORTED' marker was logged at ${now.toLocaleTimeString()}. The reporting period logged was from ${startDate} to ${endDate}.\nTotal Work Time: ${totalWorkDurationString}`, ui.ButtonSet.OK);
}

// --- REPORTING FUNCTION UTILITY ---

/**
 * Calculates and returns structured report data, including the grand total work duration.
 * @returns {{sessions: Array<Array<any>>, totalMs: number}} An object containing the session array (4 columns) and the total duration in milliseconds.
 */
function getReportData() {
  const sheet = getLogSheet();
  if (!sheet) return { sessions: [["Error: 'logs' sheet not found."]], totalMs: 0 };
  
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { sessions: [["No log data available."]], totalMs: 0 };
  
  // 1. Find the starting point (after the last REPORTED marker)
  let startIndex = 1; // Start after headers
  for (let i = data.length - 1; i >= 1; i--) { // Iterate backwards from the last entry
    if (data[i][1] === 'REPORTED') {
      startIndex = i + 1; // Start processing from the next row
      break;
    }
  }

  // Slice the data to only include relevant entries (excluding headers and reported logs)
  const logsToProcess = data.slice(startIndex);
  
  let sessions = []; // Format: [Date, Time IN, Time OUT, Duration]
  let sessionStartTime = null; 
  let sessionStartLog = null; 
  let totalWorkDuration = 0; // Accumulated duration in milliseconds for the current work block
  let grandTotalDurationMs = 0; // Grand total of all completed work blocks

  if (logsToProcess.length === 0) {
    return { sessions: [], totalMs: 0 };
  }
  
  // We process logs looking for sequential IN/BREAK-OUT/BREAK-IN/OUT cycles.
  // A "work block" is defined as starting with an IN and ending with an OUT.
  for (let i = 0; i < logsToProcess.length; i++) {
    const timestamp = logsToProcess[i][0];
    const action = logsToProcess[i][1];
    
    // Ensure timestamp is a valid date object
    if (!(timestamp instanceof Date)) {
      continue;
    }

    if (action === 'IN') {
      // Start of a new work segment/block
      sessionStartTime = timestamp;
      sessionStartLog = timestamp; 
      totalWorkDuration = 0; // Reset duration for a new block
    } else if (action === 'BREAK-OUT' || action === 'BREAK') { 
      if (sessionStartTime && sessionStartLog) {
        // Calculate work duration from the last active start (IN or BREAK-IN) to the BREAK-OUT
        const durationSegment = timestamp.getTime() - sessionStartTime.getTime();
        totalWorkDuration += durationSegment;
        sessionStartTime = null; // Mark that work is paused
      }
    } else if (action === 'BREAK-IN') { // New Break End/Work Resume
       if (sessionStartLog) {
         // Resume work, setting the new session start time
         sessionStartTime = timestamp;
       }
    } else if (action === 'OUT') {
      if (sessionStartTime && sessionStartLog) {
        // 1. Finalize the last work segment
        const durationSegment = timestamp.getTime() - sessionStartTime.getTime();
        totalWorkDuration += durationSegment;
        
        // 2. Add to grand total
        grandTotalDurationMs += totalWorkDuration;

        // 3. Report the completed work block
        
        // Col 2: Date
        const dateString = Utilities.formatDate(sessionStartLog, tz, 'yyyy-MM-dd');
        // Col 3: Clock In Time
        const timeIn = Utilities.formatDate(sessionStartLog, tz, 'HH:mm:ss');
        // Col 4: Clock Out Time
        const timeOut = Utilities.formatDate(timestamp, tz, 'HH:mm:ss');
        // Col 5: Duration
        const durationString = formatMsToDuration(totalWorkDuration);
        
        // Pushes FOUR separate strings into the array row 
        sessions.push([dateString, timeIn, timeOut, durationString]);
        
      }
      // 4. Reset ALL variables for the next complete IN/OUT cycle
      sessionStartTime = null;
      sessionStartLog = null;
      totalWorkDuration = 0; 
    }
  }
  
  return { sessions: sessions, totalMs: grandTotalDurationMs };
}


// --- REPORTING FUNCTION FOR GOOGLE SHEETS ---

/**
 * Utility function to convert the report data duration from milliseconds.
 * (This function already exists in Code.js but is included for context)
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


/**
 * Custom function that reads the raw data from the 'Time Logs Report (Raw Sync)' 
 * sheet and calculates the final work sessions and total time.
 * * @returns {Array<Array<any>>} A two-dimensional array of report data (5 columns).
 */
function calculateReportFromSyncedLogs(dummy) {
  const SYNC_SHEET_NAME = "Time Logs Report (Raw Sync)";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SYNC_SHEET_NAME);
  
  // Fallback if the sheet hasn't been created yet
  if (!sheet) return [["Error: Sync sheet not found. Run the 'Sync Full Log' button first."]];
  
  // Read all data from the synced sheet, skipping headers (Row 1)
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [["No data available in synced sheet."]];
  
  // The synced sheet structure is: [Timestamp (Date Obj), Action (String), Status (String)]
  const logsToProcess = data.slice(1); 
  
  let sessions = []; // Format: [Date, Time IN, Time OUT, Duration]
  let sessionStartTime = null; 
  let sessionStartLog = null; 
  let grandTotalDurationMs = 0; 
  const tz = ss.getSpreadsheetTimeZone();

  // Process logs looking for sequential IN/OUT cycles.
  for (let i = 0; i < logsToProcess.length; i++) {
    const timestamp = logsToProcess[i][0];
    const action = logsToProcess[i][1];
    
    // Ensure timestamp is a valid date object (Sheets service converts ISO strings to Date objects automatically)
    if (!(timestamp instanceof Date)) {
      continue;
    }

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
    // NOTE: This logic ignores break actions, as those would require tracking BREAK-IN/BREAK-OUT status, 
    // and the core goal is to display the simple IN/OUT pairings from the raw log sheet.
  }
  
  // --- Final Output Formatting ---
  const reportHeaders = [["", "Date (Session)", "Time IN", "Time OUT", "Total Duration"]];
  
  if (sessions.length === 0) {
    return reportHeaders.concat([["", "No completed IN/OUT sessions found in synced data.", "", "", ""]]);
  }
  
  // 1. Map the 4-column session data back to the 5-column structure (adding the dummy empty column)
  const sheetSessions = sessions.map(session => ["", ...session]);

  // 2. Add Grand Total row at the bottom
  //const totalDurationString = formatMsToDuration(grandTotalDurationMs);
  //const totalRow = ["", "GRAND TOTAL", "", "", totalDurationString];
  
  //return reportHeaders.concat(sheetSessions, [totalRow]);
  return reportHeaders.concat(sheetSessions);
}

// --- EXISTING REMINDER UTILITY FUNCTIONS (included for completeness) ---

function openReminderEditor() {
  const html = HtmlService.createHtmlOutputFromFile('ReminderEditor')
      .setTitle('📝 Edit Reminder List')
      .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getReminderList() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const listJson = scriptProperties.getProperty(REMINDER_LIST_KEY);
  return listJson;
}

function saveReminderList(remindersJsonString) {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();

  try {
    const reminders = JSON.parse(remindersJsonString);
    const cleanedReminders = reminders
        .filter(r => r.text && r.text.trim().length > 0)
        .map(r => ({
            text: r.text.trim(),
            isCompleted: r.isCompleted === true,
            targetDate: r.targetDate || null
        }));

    scriptProperties.setProperty(REMINDER_LIST_KEY, JSON.stringify(cleanedReminders));
    
    if (cleanedReminders.length === 0) {
      // Re-initialize to default if the list is empty after cleaning, for safety
      scriptProperties.setProperty(REMINDER_LIST_KEY, DEFAULT_LIST_JSON);
      ui.alert('List Cleared', 'No valid reminders were saved. The default list has been restored.', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Save Error', 'An error occurred while parsing and saving the reminders. Please ensure your input is valid.', ui.ButtonSet.OK);
    Logger.log("Save error: " + e.toString());
  }
}

function toggleReminder() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const isEnabled = scriptProperties.getProperty(REMINDER_ENABLED_KEY) === 'true';
  const newState = !isEnabled;
  scriptProperties.setProperty(REMINDER_ENABLED_KEY, newState.toString());
  
  const statusMessage = newState ? "Reminder pop-up is now **ENABLED**." : "Reminder pop-up is now **DISABLED**.";
  
  createCustomMenus(ui, scriptProperties);
  ui.alert('Status Updated', statusMessage + '\n\nThe new status will take effect the next time the sheet is opened.', ui.ButtonSet.OK);
}

// --- NEW FUNCTIONS FOR GOOGLE SHEET SYNC ---

/**
 * Receives the final, aggregated report data from the client and writes it to a Sheet.
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
  
  // Get the active user's ID (used for security/tracking in GAS)
  const currentUserId = Session.getTemporaryActiveUserKey(); 

  // Data row to append
  const rowData = [
    reportDate,
    currentUserId,
    data.totalWorkHours,
    data.totalBreakHours,
    data.cutoffTimestamp
  ];
  
  // Append the data as a new row
  sheet.appendRow(rowData);
  
  Logger.log(`Aggregated Report recorded for user ${currentUserId} at ${reportDate}.`);
}

/**
 * Receives raw log entries as a JSON string from the client, 
 * safely parses it, and writes the data to the 'Time Logs Report' sheet.
 * * This version uses strict string parsing to prevent "Illegal Value" TypeErrors.
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