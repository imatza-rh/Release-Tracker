/**
 * H_UI.gs - User interface elements
 * 
 * This file contains functions for showing UI elements like
 * sidebars and dialogs.
 * 
 * @requires log, handleError from B_Logging.gs
 */

/**
 * Opens the management sidebar
 * 
 * @interface Called from menu and other functions
 */
function showManageSidebar() {
  try {
    const html = HtmlService.createTemplateFromFile('ManageSidebar')
      .evaluate()
      .setTitle('Manage Jobs');
    
    SpreadsheetApp.getUi().showSidebar(html);
    log('ManageSidebar opened', LOG_LEVELS.INFO);
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
    const html = HtmlService.createTemplateFromFile('ViewSidebar')
      .evaluate()
      .setTitle('View Jobs');
    
    SpreadsheetApp.getUi().showSidebar(html);
    log('ViewSidebar opened', LOG_LEVELS.INFO);
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