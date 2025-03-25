/**
 * Z_Main.gs - Main entry point and initialization
 * 
 * This file contains the main entry point for the application.
 * It should be loaded last (hence the 'Z' prefix).
 * 
 * @requires log, handleError from B_Logging.gs
 * @requires showManageSidebar, showViewSidebar, showCreateTrackingSheetDialog from H_UI.gs
 * @requires showNotificationSettings from N_Notifications.gs
 */

/**
 * Adds a "Release Tracker" menu with emojis, a separator, and items
 * 
 * @interface Triggered automatically when spreadsheet is opened
 */
function onOpen() {
  try {
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
  } catch (error) {
    handleError('onOpen', error, false);
  }
}

/**
 * Manually checks for phases that are active today and runs their notifications
 * This is mainly for testing purposes.
 */
function manuallyCheckPhases() {
  try {
    // This function remains accessible for manual testing but is no longer in the menu
    const result = sendPhaseNotifications();
    return result;
  } catch (error) {
    handleError('manuallyCheckPhases', error, true);
    return { error: error.message || String(error) };
  }
}