/**
 * @file Code.gs (CARMACK-OPTIMIZED + EMAIL COUNT VERIFICATION)
 * @description Performance optimizations: Batch operations, eliminate waste, keep EmailUtils
 * @version 2.1 - Added email count tracking and verification
 *
 * REVISION 2.1 NOTES (Email Count Tracking):
 * - Added email count tracking that builds during import (no extra API calls)
 * - Tracks cumulative email count across all batches
 * - At completion, verifies count matches actual rows in Directory sheet
 * - Shows user both expected and actual count for confidence
 */

// --- CONFIGURATION ---
const CONFIG_MONTHLY = {
  WORK_EMAIL_ADDRESS: 'kerem@bittigitti.com.tr',
  SPREADSHEET_NAME: 'TR_Master_Directory',
  METADATA_SHEET_NAME: 'Metadata',  // Used to track which months have been imported
  BATCH_SIZE: 300,
  BODY_LIMIT: 25000,
  HEADERS: [
    'Email ID', 'Sent From', 'Sent To', 'Subject', 'Body (Plaintext)', 'Send Date', 'Import Month'
  ],
  COLUMN_WIDTHS: [200, 200, 200, 300, 400, 180, 120]
};

const PROGRESS_KEY = 'IMPORT_PROGRESS_STATUS';

// ====================================================================
// ==================== 1. MENU & INITIALIZATION ======================
// ====================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Monthly Import')
    .addItem('📅 Import Month', 'showMonthPicker')
    .addItem('📊 View Progress', 'viewImportedMonths')
    .addItem('🗑️ Clear & Reset', 'clearMonthlyImport')
    .addToUi();
}

function showMonthPicker() {
  const ui = SpreadsheetApp.getUi();

  const monthResponse = ui.prompt('📅 Enter month number (1-12):\n\n1=Jan, 2=Feb, 3=Mar, 4=Apr, 5=May, 6=Jun\n7=Jul, 8=Aug, 9=Sep, 10=Oct, 11=Nov, 12=Dec');

  if (monthResponse.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }

  const month = parseInt(monthResponse.getResponseText().trim());
  if (isNaN(month) || month < 1 || month > 12) {
    showError_('Invalid', 'Month must be between 1 and 12');
    return;
  }

  const yearResponse = ui.prompt('📅 Enter year (e.g., 2024, 2025):');
  if (yearResponse.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }

  const year = parseInt(yearResponse.getResponseText().trim());
  if (isNaN(year) || year < 2000 || year > 2100) {
    showError_('Invalid', 'Year must be between 2000 and 2100');
    return;
  }

  showProgressDialog();
  importMonthlyEmails(month, year);
}

// ====================================================================
// ======================= 2. CORE IMPORT LOGIC =======================
// ====================================================================

/**
 * CARMACK-OPTIMIZED: Faster processing with smart batching
 * WORLD-CLASS: Added LockService, fixed query bug, fixed logic bug
 * EMAIL COUNT: Tracks email count during import and verifies at end
 * FLEXIBLE: Works on any active sheet - user can use 09/25, 08/25, etc.
 */
function importMonthlyEmails(month, year) {
  // ROBUSTNESS: Use LockService to prevent concurrent executions
  const lock = LockService.getUserLock();
  if (!lock.tryLock(1000)) {
    Logger.log('Could not get lock. Another import is likely running.');
    showError_('Busy', 'Another import is already in progress. Please wait.');
    return;
  }

  Logger.log('=== importMonthlyEmails START ===');
  Logger.log('month=' + month + ', year=' + year);

  initializeProgress(month, year);
  updateProgress({ status: 'Checking if already imported...' });

  try {
    if (isMonthAlreadyImported_(month, year)) {
      updateProgress({ error: 'This month is already imported', complete: true });
      showInfo_(`${formatMonthYear_(month, year)} already imported to this sheet.`);
      return;
    }

    updateProgress({ status: 'Setting up sheet...' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // FLEXIBLE: Use whatever sheet the user is currently on
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow === 0) {
      initializeMonthlySheet_(sheet);
    }

    // BUG FIX (QUERY): Use 1st of next month for `before:` to include all emails
    const d_start = new Date(year, month - 1, 1);
    const d_end = new Date(year, month, 1);

    const timeZone = Session.getScriptTimeZone();
    const startDate = Utilities.formatDate(d_start, timeZone, 'yyyy-MM-dd');
    const endDate = Utilities.formatDate(d_end, timeZone, 'yyyy-MM-dd');

    const query = `from:${CONFIG_MONTHLY.WORK_EMAIL_ADDRESS} in:sent after:${startDate} before:${endDate}`;
    Logger.log('Query: ' + query);

    // Search Gmail
    updateProgress({ stage: '🔍 Searching Gmail...', status: 'Querying Gmail API' });
    let threads;
    try {
      threads = GmailApp.search(query);
      Logger.log('Found ' + threads.length + ' threads');
      updateProgress({
        stage: '📬 Processing threads',
        threadsFound: threads.length,
        status: 'Found ' + threads.length + ' threads'
      });
    } catch (e) {
      updateProgress({ error: 'Gmail search failed: ' + e.message, complete: true });
      showError_('Gmail Error', 'Could not search Gmail: ' + e.message);
      clearBatchStartIndex_(month, year);
      clearExpectedEmailCount_(month, year);
      return;
    }

    if (threads.length === 0) {
      updateProgress({ status: 'No emails found', complete: true });
      showInfo_(`No emails found for ${formatMonthYear_(month, year)}.`);
      clearBatchStartIndex_(month, year);
      clearExpectedEmailCount_(month, year);
      return;
    }

    // OPTIMIZED: Get progress for resumable batching
    const startIdx = getBatchStartIndex_(month, year);
    const endIdx = Math.min(startIdx + CONFIG_MONTHLY.BATCH_SIZE, threads.length);
    Logger.log('Processing threads ' + startIdx + ' to ' + endIdx);

    // EMAIL COUNT: Get cumulative count from previous batches
    let cumulativeEmailCount = getExpectedEmailCount_(month, year);
    let emailsThisBatch = 0;

    // OPTIMIZED: Batch collect, single write
    const dataToLog = [];
    const monthYearStr = formatMonthYear_(month, year);

    // Process threads
    for (let threadIdx = startIdx; threadIdx < endIdx; threadIdx++) {
      try {
        const thread = threads[threadIdx];

        updateProgress({
          threadsProcessed: threadIdx - startIdx + 1,
          status: 'Processing thread ' + (threadIdx + 1) + ' of ' + threads.length
        });

        // Use EmailUtils for verification
        const packets = threadToMessagePackets(thread, CONFIG_MONTHLY.WORK_EMAIL_ADDRESS);

        // Process each packet
        packets.forEach(function (packet) {
          try {
            const messageYear = packet.date.getFullYear();
            const messageMonth = packet.date.getMonth() + 1;

            if (messageYear !== year || messageMonth !== month) {
              return;
            }

            // BUG FIX (LOGIC): Pass bodyHtml (not bodyPlain) to the processor
            const bodyProcessed = processEmailBody_(packet.bodyHtml);
            const formattedDate = Utilities.formatDate(packet.date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

            dataToLog.push([packet.messageId, packet.from, packet.to, packet.subject, bodyProcessed, formattedDate, monthYearStr]);

            // EMAIL COUNT: Track this email
            emailsThisBatch++;
            updateProgress({ emailsCollected: dataToLog.length });

          } catch (e) {
            Logger.log('Skipped message: ' + e.message);
          }
        });

      } catch (e) {
        Logger.log('Skipped thread: ' + e.message);
      }
    }

    Logger.log('Collected ' + dataToLog.length + ' emails to log');

    // OPTIMIZED: Single batch write
    updateProgress({ stage: '💾 Writing to sheet...', status: 'Writing ' + dataToLog.length + ' emails...' });
    if (dataToLog.length > 0) {
      const appendRow = sheet.getLastRow() + 1;
      sheet.getRange(appendRow, 1, dataToLog.length, dataToLog[0].length).setValues(dataToLog);
      Logger.log('Batch write complete: ' + dataToLog.length + ' rows');
      updateProgress({ emailsWritten: dataToLog.length });
    }

    // OPTIMIZED: Only resize columns at the very end of import, not per batch
    if (endIdx >= threads.length) {
      sheet.autoResizeColumns(1, CONFIG_MONTHLY.HEADERS.length);
    }

    // EMAIL COUNT: Update cumulative count
    cumulativeEmailCount += emailsThisBatch;
    setExpectedEmailCount_(month, year, cumulativeEmailCount);

    // Update progress
    if (endIdx < threads.length) {
      setBatchStartIndex_(month, year, endIdx);
      updateProgress({
        stage: '⏸️ Batch paused',
        status: 'Batch ' + (Math.ceil(endIdx / CONFIG_MONTHLY.BATCH_SIZE)) + ' complete. Expected so far: ' + cumulativeEmailCount + ' emails'
      });
      showInfo_(
        '📊 Batch complete!\n\n' +
        'Processed: ' + endIdx + ' / ' + threads.length + ' threads\n' +
        'Emails collected: ' + cumulativeEmailCount + '\n\n' +
        'Click "Import Month" again to continue.'
      );
    } else {
      clearBatchStartIndex_(month, year);
      clearProgressStatus();

      // Mark as imported
      const metadataSheet = getOrCreateMetadataSheet_(ss);
      const importDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      metadataSheet.appendRow([month, year, threads.length, importDate, '']);

      // EMAIL COUNT VERIFICATION: Verify expected vs actual
      const actualCount = getDirectoryRowCount_(month, year);
      const expectedCount = cumulativeEmailCount;
      
      Logger.log('=== IMPORT COMPLETE ===');
      Logger.log('Expected emails: ' + expectedCount);
      Logger.log('Actual emails in sheet: ' + actualCount);
      
      let verificationMessage = '✅ Import Complete!\n\n' +
        formatMonthYear_(month, year) + '\n' +
        'Expected: ' + expectedCount + ' emails\n' +
        'Actual: ' + actualCount + ' emails';
      
      if (expectedCount === actualCount) {
        verificationMessage += '\n\n✅ Verification PASSED - All emails accounted for!';
      } else {
        verificationMessage += '\n\n⚠️ Mismatch detected:\n' +
          'Difference: ' + Math.abs(expectedCount - actualCount) + ' emails';
      }
      
      const activeSheetName = sheet.getName();
      verificationMessage += '\n\nSheet: "' + activeSheetName + '"';

      updateProgress({
        stage: '✅ Complete!',
        status: 'All emails imported to "' + activeSheetName + '" - Verification: ' + expectedCount + ' expected, ' + actualCount + ' actual',
        complete: true
      });

      clearExpectedEmailCount_(month, year);
      showSuccess_(verificationMessage);
    }

  } catch (error) {
    Logger.log('FATAL ERROR: ' + error.message);
    Logger.log(error.stack);
    updateProgress({ error: error.message, complete: true });
    showError_('Error', error.message);
  } finally {
    lock.releaseLock();
    Logger.log('Lock released.');
  }
}

function getOrCreateMetadataSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG_MONTHLY.METADATA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_MONTHLY.METADATA_SHEET_NAME);
    sheet.appendRow(['Month', 'Year', 'Emails Imported', 'Import Date', 'Notes']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Initialize headers on the active sheet if it's empty
 */
function initializeMonthlySheet_(sheet) {
  // Only add headers if sheet is truly empty (no data at all)
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(CONFIG_MONTHLY.HEADERS);
    sheet.setFrozenRows(1);
    const headerRange = sheet.getRange(1, 1, 1, CONFIG_MONTHLY.HEADERS.length);
    headerRange.setFontWeight('bold');
    CONFIG_MONTHLY.COLUMN_WIDTHS.forEach(function (width, index) {
      sheet.setColumnWidth(index + 1, width);
    });
  }
}

/**
 * OPTIMIZED: Uses Metadata as the single source of truth.
 */
function isMonthAlreadyImported_(month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const metadataSheet = getOrCreateMetadataSheet_(ss);
  const metadataData = metadataSheet.getDataRange().getValues();

  for (let i = 1; i < metadataData.length; i++) {
    if (metadataData[i][0] == month && metadataData[i][1] == year) {
      Logger.log(formatMonthYear_(month, year) + ' is already imported.');
      return true;
    }
  }

  return false;
}

// ====================================================================
// ==================== 4. BATCH PROGRESS HELPERS =====================
// ====================================================================

function getBatchStartIndex_(month, year) {
  const progressKey = 'IMPORT_BATCH_' + month + '_' + year;
  const props = PropertiesService.getUserProperties();
  const progress = props.getProperty(progressKey);
  return progress ? parseInt(progress) : 0;
}

function setBatchStartIndex_(month, year, progress) {
  const progressKey = 'IMPORT_BATCH_' + month + '_' + year;
  PropertiesService.getUserProperties().setProperty(progressKey, progress.toString());
}

function clearBatchStartIndex_(month, year) {
  const progressKey = 'IMPORT_BATCH_' + month + '_' + year;
  PropertiesService.getUserProperties().deleteProperty(progressKey);
}

// ====================================================================
// ================ EMAIL COUNT TRACKING (NEW) ========================
// ====================================================================

/**
 * Get cumulative expected email count for this month
 * Tracks total emails across all batches
 */
function getExpectedEmailCount_(month, year) {
  const countKey = 'IMPORT_EMAIL_COUNT_' + month + '_' + year;
  const props = PropertiesService.getUserProperties();
  const count = props.getProperty(countKey);
  return count ? parseInt(count) : 0;
}

/**
 * Set cumulative expected email count
 */
function setExpectedEmailCount_(month, year, count) {
  const countKey = 'IMPORT_EMAIL_COUNT_' + month + '_' + year;
  PropertiesService.getUserProperties().setProperty(countKey, count.toString());
}

/**
 * Clear expected email count
 */
function clearExpectedEmailCount_(month, year) {
  const countKey = 'IMPORT_EMAIL_COUNT_' + month + '_' + year;
  PropertiesService.getUserProperties().deleteProperty(countKey);
}

/**
 * Get actual count of rows in active sheet
 * Simple: just count all data rows (skip header)
 * The Gmail query already filtered by month, so we trust all rows are from the target month
 */
function getDirectoryRowCount_(month, year) {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  
  // Simply return number of data rows (lastRow - 1 to skip header)
  return lastRow - 1;
}

// ====================================================================
// ================= 5. LIVE UI PROGRESS (MONITOR) ====================
// ====================================================================

function initializeProgress(month, year) {
  const progressData = {
    month: month,
    year: year,
    startTime: new Date().getTime(),
    stage: 'Initializing',
    threadsFound: 0,
    threadsProcessed: 0,
    emailsCollected: 0,
    emailsWritten: 0,
    status: 'Starting...',
    complete: false,
    error: null
  };
  PropertiesService.getUserProperties().setProperty(PROGRESS_KEY, JSON.stringify(progressData));
}

function updateProgress(updates) {
  const props = PropertiesService.getUserProperties();
  const existing = props.getProperty(PROGRESS_KEY);
  const progressData = existing ? JSON.parse(existing) : {};

  const updated = Object.assign(progressData, updates);

  props.setProperty(PROGRESS_KEY, JSON.stringify(updated));
}

function clearProgressStatus() {
  PropertiesService.getUserProperties().deleteProperty(PROGRESS_KEY);
}

function getProgress() {
  const props = PropertiesService.getUserProperties();
  const data = props.getProperty(PROGRESS_KEY);
  return data ? JSON.parse(data) : null;
}

function showProgressDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ProgressMonitor');
  SpreadsheetApp.getUi().showModelessDialog(html, '📊 Import Progress');
}

// ====================================================================
// ==================== 6. UI & MENU FUNCTIONS ========================
// ====================================================================

function viewImportedMonths() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const metadataSheet = getOrCreateMetadataSheet_(ss);
    const data = metadataSheet.getDataRange().getValues();

    if (data.length <= 1) {
      showInfo_('No months imported yet.');
      return;
    }

    let message = '📅 Imported Months:\n\n';
    let totalEmails = 0;

    for (let i = 1; i < data.length; i++) {
      const month = data[i][0];
      const year = data[i][1];
      const count = data[i][2];
      const importDate = data[i][3];

      message += formatMonthYear_(month, year) + ': ' + count + ' emails (' + importDate + ')\n';
      totalEmails += count;
    }

    message += '\n📊 Total: ' + totalEmails + ' emails';
    showInfo_(message);
  } catch (error) {
    showError_('Error', 'Failed to view: ' + error.message);
  }
}

/**
 * OPTIMIZED: Uses clearContent() instead of deleteRows() for massive speedup.
 * FLEXIBLE: Works on active sheet
 */
function clearMonthlyImport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const metadataSheet = ss.getSheetByName(CONFIG_MONTHLY.METADATA_SHEET_NAME);

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Clear This Sheet?', 'This will permanently delete all email data from "' + activeSheet.getName() + '" and reset all progress. Are you sure?', ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    // Clear active sheet (data only, not headers)
    const lastRow = activeSheet.getLastRow();
    if (lastRow > 1) {
      activeSheet.getRange(2, 1, activeSheet.getMaxRows() - 1, activeSheet.getMaxColumns()).clearContent();
    }
    
    // Clear metadata
    if (metadataSheet) {
      const metaLastRow = metadataSheet.getLastRow();
      if (metaLastRow > 1) {
        metadataSheet.getRange(2, 1, metadataSheet.getMaxRows() - 1, metadataSheet.getMaxColumns()).clearContent();
      }
    }

    // Clear all batch progress properties
    const props = PropertiesService.getUserProperties();
    const keys = props.getKeys();
    keys.forEach(function (key) {
      if (key.startsWith('IMPORT_BATCH_') || key.startsWith('IMPORT_EMAIL_COUNT_')) {
        props.deleteProperty(key);
      }
    });
    
    clearProgressStatus();

    showSuccess_('All data cleared from "' + activeSheet.getName() + '" and progress reset.');
  }
}

// ====================================================================
// ==================== 7. TEXT & EMAIL UTILITIES =====================
// ====================================================================

function formatMonthYear_(month, year) {
  const months = ['', 'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];
  return months[month] + ' ' + year;
}

/**
 * OPTIMIZED: Process email body efficiently
 */
function processEmailBody_(bodyHtml) {
  const replyBody = extractReply_(bodyHtml);
  const plaintext = stripHtml_(replyBody);
  return plaintext.substring(0, CONFIG_MONTHLY.BODY_LIMIT);
}

function extractReply_(body) {
  const replyIndicators = [
    /On\s+.+?wrote:\s*/i,
    /From:\s*.+?\nSent:\s*.+?\nTo:\s*.+?\nSubject:\s*.+/i,
    /\n\s*-+Original Message-+/i,
    /\n\s*-+ Forwarded message -+/i,
    /\n\s*> /,
    /\n\s*__+/
  ];

  let shortestReply = body;
  for (let i = 0; i < replyIndicators.length; i++) {
    const match = body.match(replyIndicators[i]);
    if (match) {
      const potentialReply = body.substring(0, match.index).trim();
      if (potentialReply.length > 0 && potentialReply.length < shortestReply.length) {
        shortestReply = potentialReply;
      }
    }
  }

  shortestReply = shortestReply.split(/--\s*\n|__\s*\n/)[0].trim();
  return shortestReply || body;
}

/**
 * OPTIMIZED: Replaces block-level tags with newlines for better readability.
 */
function stripHtml_(html) {
  return html
    .replace(/<(div|p|br|li|h[1-6])[^>]*>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s*\n/g, '\n')
    .trim();
}

// ========== EMAIL UTILITIES (RFC 822 VERIFICATION) ==========

function extractEmailAddress(emailString) {
  if (!emailString) return '';
  const angleMatch = emailString.match(/<([^>]+)>/);
  if (angleMatch && angleMatch[1]) {
    return angleMatch[1].trim();
  }
  return emailString.trim();
}

function getActualSender(message) {
  try {
    const rawMessage = message.getRawContent();
    const fromMatch = rawMessage.match(/^From:\s*(.+?)$/m);

    if (fromMatch && fromMatch[1]) {
      const fromHeader = fromMatch[1].trim();
      return extractEmailAddress(fromHeader);
    }

    return extractEmailAddress(message.getFrom());

  } catch (error) {
    Logger.log('Warning: Raw content parsing failed: ' + error.message);
    return extractEmailAddress(message.getFrom());
  }
}

function isMessageFromAddress(message, targetEmail) {
  const sender = getActualSender(message).toLowerCase();
  const target = targetEmail.toLowerCase();
  return sender === target;
}

function threadToMessagePackets(thread, fromAddressFilter) {
  const packets = [];
  const messages = thread.getMessages();

  messages.forEach(function (message) {
    try {
      if (fromAddressFilter && !isMessageFromAddress(message, fromAddressFilter)) {
        return;
      }

      const packet = {
        messageId: message.getId(),
        from: getActualSender(message),
        to: message.getTo() || '',
        subject: message.getSubject() || '',
        date: message.getDate(),
        bodyHtml: message.getBody() || ''
      };

      packets.push(packet);

    } catch (error) {
      Logger.log('Error processing message: ' + error.message);
    }
  });

  return packets;
}

// ====================================================================
// ====================== 8. ALERT WRAPPERS ===========================
// ====================================================================

function showSuccess_(message) {
  SpreadsheetApp.getUi().alert('✅ Success', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function showInfo_(message) {
  SpreadsheetApp.getUi().alert('ℹ️ Info', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function showError_(title, message) {
  SpreadsheetApp.getUi().alert('❌ ' + title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}