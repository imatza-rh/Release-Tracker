/**
 * H_UI.gs - User interface elements
 * 
 * This file contains functions for showing UI elements like
 * sidebars and dialogs.
 * 
 * @requires log, handleError from B_Logging.gs
 */

/**
 * Gets the name of the active sheet or default if not a valid job sheet
 * 
 * @returns {string} Name of active sheet or JOBS_SHEET_NAME if not valid
 */
function getActiveJobSheet() {
  try {
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
  } catch (error) {
    handleError('getActiveJobSheet', error, false);
    return JOBS_SHEET_NAME; // Return default on error
  }
}

/**
 * Opens the management sidebar
 * 
 * @interface Called from menu and other functions
 */
function showManageSidebar() {
  try {
    const activeSheetName = getActiveJobSheet();
    log(`Opening ManageSidebar with sheet: ${activeSheetName}`, LOG_LEVELS.INFO);
    
    const template = HtmlService.createTemplateFromFile('ManageSidebar');
    template.activeSheetName = activeSheetName; // Pass active sheet name to the sidebar
    
    const html = template
      .evaluate()
      .setTitle('Manage Jobs');
    
    SpreadsheetApp.getUi().showSidebar(html);
    log('ManageSidebar opened for sheet: ' + activeSheetName, LOG_LEVELS.INFO);
  } catch (error) {
    handleError('showManageSidebar', error, true);
  }
}

/**
 * Opens the view sidebar
 * 
 * @interface Called from menu and other functions
 */
function showViewSidebar() {
  try {
    const activeSheetName = getActiveJobSheet();
    log(`Opening ViewSidebar with sheet: ${activeSheetName}`, LOG_LEVELS.INFO);
    
    const template = HtmlService.createTemplateFromFile('ViewSidebar');
    template.activeSheetName = activeSheetName; // Pass active sheet name to the sidebar
    
    const html = template
      .evaluate()
      .setTitle('View Jobs');
    
    SpreadsheetApp.getUi().showSidebar(html);
    log('ViewSidebar opened for sheet: ' + activeSheetName, LOG_LEVELS.INFO);
  } catch (error) {
    handleError('showViewSidebar', error, true);
  }
}

/**
 * Opens the create tracking sheet dialog
 * 
 * @interface Called from menu and other functions
 */
function showCreateTrackingSheetDialog() {
  try {
    const html = HtmlService.createTemplateFromFile('CreateTrackingSheet')
      .evaluate()
      .setWidth(500)
      .setHeight(520);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Create Tracking Sheet');
    log('Create Tracking Sheet dialog opened', LOG_LEVELS.INFO);
  } catch (error) {
    handleError('showCreateTrackingSheetDialog', error, true);
  }
}