/**
 * Z_Main.gs - Main entry point and initialization
 * 
 * This file contains the main entry point for the application.
 * It should be loaded last (hence the 'Z' prefix).
 * 
 * @requires log, handleError from B_Logging.gs
 * @requires showManageSidebar, showViewSidebar, showCreateTrackingSheetDialog from H_UI.gs
 */

/**
 * Adds a "Release Tracker" menu with emojis, a separator, and items
 * 
 * @interface Triggered automatically when spreadsheet is opened
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🚀 Release Tracker')
      .addItem('🛠 Manage Jobs', 'showManageSidebar')
      .addItem('👀 View Jobs', 'showViewSidebar')
      .addSeparator()
      .addItem('📝 Create Tracking Sheet', 'showCreateTrackingSheetDialog')
      .addToUi();
    log('Menu initialized successfully', LOG_LEVELS.INFO);
  } catch (error) {
    handleError('onOpen', error, false);
  }
}
