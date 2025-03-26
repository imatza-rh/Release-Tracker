/**
 * Z_Main.gs - Main entry point and initialization
 * 
 * This file contains the main entry point for the application.
 * It should be loaded last (hence the 'Z' prefix).
 * 
 * @requires log, handleError, safeExecute from B_Logging.gs
 * @requires createSidebar, createDialog from I_TemplateLoader.gs
 */

/**
 * Adds a "Release Tracker" menu with emojis, a separator, and items
 * 
 * @interface Triggered automatically when spreadsheet is opened
 */
function onOpen() {
  return safeExecute('onOpen', () => {
    const ui = SpreadsheetApp.getUi();
    
    // Create the main Release Tracker menu
    ui.createMenu('🚀 Release Tracker')
      .addSubMenu(ui.createMenu('📜 Control Jobs')
        .addItem('🛠 Manage Jobs', 'showManageSidebar')
        .addItem('👀 View Jobs', 'showViewSidebar'))
      .addItem('📝 Create Jobs Tracking Sheet', 'showCreateTrackingSheetDialog')
      .addSeparator()
      .addItem('🔔 Phase Notifications', 'showNotificationSettings')
      .addToUi();
    
    log('Menu initialized successfully', LOG_LEVELS.INFO);
  }, false);
}

/**
 * Opens the management sidebar
 * 
 * @interface Called from menu and other functions
 */
function showManageSidebar() {
  return safeExecute('showManageSidebar', () => {
    const activeSheetName = getActiveJobSheet();
    log(`Opening ManageSidebar with sheet: ${activeSheetName}`, LOG_LEVELS.INFO);
    
    const html = createSidebar('ManageSidebar', 'Manage Jobs', { activeSheetName });
    SpreadsheetApp.getUi().showSidebar(html);
    
    log('ManageSidebar opened for sheet: ' + activeSheetName, LOG_LEVELS.INFO);
  }, true);
}

/**
 * Opens the view sidebar
 * 
 * @interface Called from menu and other functions
 */
function showViewSidebar() {
  return safeExecute('showViewSidebar', () => {
    const activeSheetName = getActiveJobSheet();
    log(`Opening ViewSidebar with sheet: ${activeSheetName}`, LOG_LEVELS.INFO);
    
    const html = createSidebar('ViewSidebar', 'View Jobs', { activeSheetName });
    SpreadsheetApp.getUi().showSidebar(html);
    
    log('ViewSidebar opened for sheet: ' + activeSheetName, LOG_LEVELS.INFO);
  }, true);
}

/**
 * Opens the create tracking sheet dialog
 * 
 * @interface Called from menu and other functions
 */
function showCreateTrackingSheetDialog() {
  return safeExecute('showCreateTrackingSheetDialog', () => {
    const html = createDialog('CreateTrackingSheet', 'Create Tracking Sheet', {}, {
      width: 500,
      height: 520
    });
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Create Tracking Sheet');
    log('Create Tracking Sheet dialog opened', LOG_LEVELS.INFO);
  }, true);
}

/**
 * Gets the name of the active sheet or default if not a valid job sheet
 * 
 * @returns {string} Name of active sheet or JOBS_SHEET_NAME if not valid
 */
function getActiveJobSheet() {
  return safeExecute('getActiveJobSheet', () => {
    // Get active sheet
    const activeSheet = SpreadsheetApp.getActiveSheet();
    if (!activeSheet) {
      log(`No active sheet found, using default: ${JOBS_SHEET_NAME}`, LOG_LEVELS.INFO);
      return JOBS_SHEET_NAME;
    }
    
    const activeSheetName = activeSheet.getName();
    
    // Log for debugging
    log(`Checking active sheet: ${activeSheetName}`, LOG_LEVELS.INFO);
    
    // Check if the active sheet is the DEFAULT_JOBS_SHEET_NAME
    if (activeSheetName === DEFAULT_JOBS_SHEET_NAME) {
      log(`Active sheet is DefaultJobs sheet, using ${JOBS_SHEET_NAME} instead`, LOG_LEVELS.INFO);
      return JOBS_SHEET_NAME;
    }
    
    // Check if the active sheet has the required job columns (at minimum Job Name and Status)
    const headerRow = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues()[0];
    if (headerRow.includes("Job Name") && headerRow.includes("Status")) {
      log(`Valid job sheet found: ${activeSheetName}`, LOG_LEVELS.INFO);
      return activeSheetName;
    }
    
    // If active sheet is not a job sheet, get available job sheets
    const jobSheets = listJobSheets();
    
    log(`Available job sheets: ${jobSheets.join(', ')}`, LOG_LEVELS.INFO);
    
    if (jobSheets.length > 0) {
      log(`Active sheet not a job sheet. Using first available: ${jobSheets[0]}`, LOG_LEVELS.INFO);
      return jobSheets[0]; // Return first available job sheet
    }
    
    log(`No valid job sheets found, using default: ${JOBS_SHEET_NAME}`, LOG_LEVELS.INFO);
    return JOBS_SHEET_NAME; // Fall back to default
  }, false) || JOBS_SHEET_NAME; // Return default on error
}

/**
 * Manually checks for phases that are active today and runs their notifications
 * This is mainly for testing purposes.
 */
function manuallyCheckPhases() {
  return safeExecute('manuallyCheckPhases', () => {
    // This function remains accessible for manual testing but is no longer in the menu
    const result = sendPhaseNotifications();
    return result;
  }, true);
}

/**
 * Activates a specified sheet by name
 * 
 * @param {string} sheetName - Name of the sheet to activate
 * @returns {Object} Result object indicating success or failure
 */
function activateSheet(sheetName) {
  return safeExecute('activateSheet', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return {
        status: 'error',
        message: `Sheet "${sheetName}" not found`
      };
    }
    
    ss.setActiveSheet(sheet);
    
    return {
      status: 'success',
      message: `Sheet "${sheetName}" activated`
    };
  }, false);
}

/**
 * Initializes the application with default data if needed
 * This is useful when the application is first installed
 * 
 * @returns {Object} Result of initialization
 */
function initializeApplication() {
  return safeExecute('initializeApplication', () => {
    log('Initializing application', LOG_LEVELS.INFO);
    
    // Load config to ensure defaults are created
    const config = getConfig();
    
    // Create the default Jobs sheet if it doesn't exist
    const jobsSheet = getOrCreateSheet(JOBS_SHEET_NAME, null, true);
    
    // Set up proper data validation on the Jobs sheet
    setupDataValidation(jobsSheet);
    
    // Apply formatting
    formatSheet(jobsSheet);
    
    return {
      status: 'success',
      message: 'Application initialized successfully'
    };
  }, true);
}