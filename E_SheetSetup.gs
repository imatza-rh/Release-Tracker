/**
 * E_SheetSetup.gs - Sheet creation and formatting
 * 
 * This file contains functions for setting up and managing spreadsheet
 * sheets for the Release Tracker application.
 * 
 * @requires JOBS_SHEET_NAME from A_Constants.gs
 * @requires log, handleError from B_Logging.gs
 * @requires sanitizeInput from C_Utils.gs
 * @requires getConfig from D_Config.gs
 */

/**
 * Gets or creates a sheet with the given name and specified fields
 * 
 * @param {string} sheetName - Name of sheet to get or create
 * @param {Array} [customFields] - Optional array of custom fields to use for header
 * @returns {Sheet} The sheet object
 * @throws {Error} If sheet cannot be created
 */
function getOrCreateSheet(sheetName, customFields) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      log(`Creating new sheet: ${sheetName}`, LOG_LEVELS.INFO);
      sheet = ss.insertSheet(sheetName);
      
      // Use user-configurable fields as header
      let fields;
      if (Array.isArray(customFields) && customFields.length > 0) {
        // Use provided custom fields if available
        fields = customFields;
        log(`Using custom fields for sheet: ${customFields.join(', ')}`, LOG_LEVELS.INFO);
      } else {
        // Otherwise use default fields from config
        fields = getConfig().fields || ['Job Name', 'Type', 'Status', 'Priority', 'Notes', 'Job Link', 'Jira Ticket'];
      }
      
      // Ensure 'Job Name' is always included as the first field
      if (!fields.includes('Job Name')) {
        fields.unshift('Job Name');
      } else if (fields.indexOf('Job Name') !== 0) {
        // If Job Name exists but not first, rearrange it
        fields = ['Job Name'].concat(fields.filter(f => f !== 'Job Name'));
      }
      
      sheet.appendRow(fields);
      
      // Format header row
      sheet.getRange(1, 1, 1, fields.length).setFontWeight('bold').setBackground('#f3f3f3');
      
      // Setup validation
      setupDataValidation(sheet);
    }
    
    return sheet;
  } catch (error) {
    handleError('getOrCreateSheet', error, false);
    throw error; // Re-throw to handle where this is called
  }
}

/**
 * Applies data validation for dropdown columns
 * 
 * @param {Sheet} sheet - The sheet to set up validation on
 */
function setupDataValidation(sheet) {
  try {
    const config = getConfig();
    const fields = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const lastRow = Math.max(2, sheet.getMaxRows()); // Ensure at least one data row
    
    // Find indices for validation columns
    const statusesColIndex = fields.indexOf('Status');
    const typesColIndex = fields.indexOf('Type');
    const prioColIndex = fields.indexOf('Priority');
    
    // Create validation rules
    if (statusesColIndex >= 0 && config.statuses && config.statuses.length > 0) {
      const ruleStatus = SpreadsheetApp.newDataValidation()
        .requireValueInList(config.statuses, true)
        .setAllowInvalid(false)
        .build();
      
      // Apply validation to all rows
      sheet.getRange(2, statusesColIndex+1, lastRow-1, 1).setDataValidation(ruleStatus);
    }
    
    if (typesColIndex >= 0 && config.types && config.types.length > 0) {
      const ruleType = SpreadsheetApp.newDataValidation()
        .requireValueInList(config.types, true)
        .setAllowInvalid(false)
        .build();
      
      // Apply validation to all rows
      sheet.getRange(2, typesColIndex+1, lastRow-1, 1).setDataValidation(ruleType);
    }
    
    if (prioColIndex >= 0 && config.priorities && config.priorities.length > 0) {
      const rulePrio = SpreadsheetApp.newDataValidation()
        .requireValueInList(config.priorities, true)
        .setAllowInvalid(false)
        .build();
      
      // Apply validation to all rows
      sheet.getRange(2, prioColIndex+1, lastRow-1, 1).setDataValidation(rulePrio);
    }
    
    log('Data validation set up successfully', LOG_LEVELS.DEBUG);
  } catch (error) {
    handleError('setupDataValidation', error, false);
  }
}

/**
 * Helper function to find a row by job name
 * 
 * @param {string} jobName - Name of the job to find
 * @param {string} [sheetName=JOBS_SHEET_NAME] - Name of sheet to search in
 * @returns {Object} Result object with job data or error information
 */
function findJobRow(jobName, sheetName = JOBS_SHEET_NAME) {
  try {
    const sheet = getOrCreateSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      return { found: false, error: `No rows in sheet: ${sheetName}` };
    }
    
    const header = data[0]; // First row = column names
    const nameColIndex = header.indexOf("Job Name");
    
    if (nameColIndex < 0) {
      return { found: false, error: `No 'Job Name' column found in: ${sheetName}` };
    }
    
    jobName = sanitizeInput(jobName);
    
    for (let r = 1; r < data.length; r++) {
      const rowName = sanitizeInput(data[r][nameColIndex]);
      if (rowName === jobName) {
        return { 
          found: true, 
          rowIndex: r, 
          rowNumber: r+1,
          sheet: sheet,
          header: header,
          data: data[r]
        };
      }
    }
    
    return { found: false, error: `Job not found: ${jobName}` };
  } catch (error) {
    handleError('findJobRow', error, false);
    return { found: false, error: error.message || error };
  }
}