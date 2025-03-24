/**
 * K_CreateTrackingSheet.gs - Functions for creating new tracking sheets
 * 
 * This file contains functions specific to creating new tracking sheets
 * with user-defined options.
 * 
 * @requires log, handleError from B_Logging.gs
 * @requires getConfig, saveConfig from D_Config.gs
 * @requires getOrCreateSheet, setupDataValidation from E_SheetSetup.gs
 * @requires addJobsFromDefaultSheet from G_DefaultJobs.gs
 */

/**
 * Creates a new tracking sheet with custom field options
 * 
 * @param {string} sheetName - Name of the sheet to create
 * @param {string} source - Source for jobs ('empty', 'DefaultJobs', or 'manual')
 * @param {Object} fieldOptions - Custom field options (statuses, types, priorities)
 * @param {string[]} [manualJobs=[]] - Array of job names when source is 'manual'
 * @returns {Object} Result object indicating success or failure
 */
function createTrackingSheetWithOptions(sheetName, source, fieldOptions, manualJobs = []) {
  try {
    if (!sheetName) {
      return {
        status: 'error',
        message: 'Sheet name is required'
      };
    }
    
    log(`Starting creation of tracking sheet: ${sheetName}`, LOG_LEVELS.INFO);
    
    // First, update the configuration with the custom field options
    const currentConfig = getConfig(true);
    
    // Only update if options are provided
    if (fieldOptions) {
      if (fieldOptions.statuses && fieldOptions.statuses.length > 0) {
        currentConfig.statuses = fieldOptions.statuses;
      }
      
      if (fieldOptions.types && fieldOptions.types.length > 0) {
        currentConfig.types = fieldOptions.types;
      }
      
      if (fieldOptions.priorities && fieldOptions.priorities.length > 0) {
        currentConfig.priorities = fieldOptions.priorities;
      }
      
      // Update colors if provided (we'll store these but not use them directly)
      if (fieldOptions.statusColors && fieldOptions.statusColors.length === fieldOptions.statuses.length) {
        currentConfig.statusColors = fieldOptions.statusColors;
      }
      
      if (fieldOptions.typeColors && fieldOptions.typeColors.length === fieldOptions.types.length) {
        currentConfig.typeColors = fieldOptions.typeColors;
      }
      
      if (fieldOptions.priorityColors && fieldOptions.priorityColors.length === fieldOptions.priorities.length) {
        currentConfig.priorityColors = fieldOptions.priorityColors;
      }
      
      // Update metadata
      currentConfig.metadata = {
        version: (currentConfig.metadata?.version || 0) + 1,
        created: currentConfig.metadata?.created || new Date().toISOString(),
        lastModified: new Date().toISOString()
      };
      
      // Save the updated config
      const props = PropertiesService.getScriptProperties();
      props.setProperty(CONFIG_PROPERTY_KEY, JSON.stringify(currentConfig));
      
      // Update cache
      _configCache = currentConfig;
      
      log('Configuration updated with custom field options', LOG_LEVELS.INFO);
    }
    
    // Now create the tracking sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let existingSheet = ss.getSheetByName(sheetName);
    
    // Check if sheet already exists
    if (existingSheet) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        `Sheet "${sheetName}" already exists. Do you want to keep using it?`,
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.NO) {
        return {
          status: 'info',
          message: 'Operation cancelled'
        };
      }
    }
    
    // Create or get the sheet with proper headers and formatting
    log('Getting or creating sheet', LOG_LEVELS.INFO);
    const sheet = getOrCreateSheet(sheetName);
    
    // Add Jira and Job Link columns if they don't exist
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let headerUpdated = false;
    
    // Add Job Link column if it doesn't exist
    if (!header.includes('Job Link')) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue('Job Link');
      headerUpdated = true;
    }
    
    // Add Jira Ticket column if it doesn't exist
    if (!header.includes('Jira Ticket')) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue('Jira Ticket');
      headerUpdated = true;
    }
    
    // Re-apply formatting to header row if we updated it
    if (headerUpdated) {
      sheet.getRange(1, 1, 1, sheet.getLastColumn())
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
      
      log('Added additional columns to sheet', LOG_LEVELS.INFO);
    }
      
    // Apply validation based on updated config
    setupDataValidation(sheet);
    log('Data validation applied', LOG_LEVELS.INFO);
    
    // If manual jobs were provided, add them to the sheet
    if (source === 'manual' && Array.isArray(manualJobs) && manualJobs.length > 0) {
      return addManualJobs(sheetName, manualJobs);
    }
    
    // If source is 'empty', we're done
    if (source === 'empty') {
      return {
        status: 'success',
        message: `Sheet "${sheetName}" created successfully with empty job list`
      };
    }
    
    // For DefaultJobs source, check if it exists first
    if (source === 'DefaultJobs') {
      const defaultSheet = ss.getSheetByName(DEFAULT_JOBS_SHEET_NAME);
      if (!defaultSheet) {
        log('DefaultJobs sheet does not exist, creating it', LOG_LEVELS.INFO);
        createDefaultJobsTemplate();
      }
    }
    
    // Add default jobs from DefaultJobs sheet
    return addJobsFromDefaultSheet(sheetName);
  } catch (error) {
    return handleError('createTrackingSheetWithOptions', error, true);
  }
}

/**
 * Adds manually entered jobs to a tracking sheet
 * 
 * @param {string} sheetName - Name of the sheet to add jobs to
 * @param {string[]} jobNames - Array of job names to add
 * @returns {Object} Result object indicating success or failure
 */
function addManualJobs(sheetName, jobNames) {
  try {
    if (!Array.isArray(jobNames) || jobNames.length === 0) {
      return {
        status: 'info',
        message: `No jobs provided to add to ${sheetName}`
      };
    }
    
    log(`Adding ${jobNames.length} manual jobs to ${sheetName}`, LOG_LEVELS.INFO);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return {
        status: 'error',
        message: `Sheet ${sheetName} not found`
      };
    }
    
    // Get config for default values
    const config = getConfig();
    
    // Get header row
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Find column indices
    const nameIndex = header.indexOf('Job Name');
    const typeIndex = header.indexOf('Type');
    const statusIndex = header.indexOf('Status');
    const priorityIndex = header.indexOf('Priority');
    
    // Prepare data for batch insert
    const rowsToAdd = [];
    
    // Process each job name
    jobNames.forEach(jobName => {
      jobName = jobName.trim();
      if (!jobName) return;
      
      // Create a new row with empty cells
      const newRow = Array(header.length).fill('');
      
      // Add the job name
      if (nameIndex !== -1) {
        newRow[nameIndex] = jobName;
      }
      
      // Add default values
      if (typeIndex !== -1) {
        newRow[typeIndex] = config.types[0] || 'Build';
      }
      
      if (statusIndex !== -1) {
        newRow[statusIndex] = config.statuses[0] || 'Pending';
      }
      
      if (priorityIndex !== -1) {
        newRow[priorityIndex] = config.priorities[1] || 'Medium'; // Default to Medium priority
      }
      
      rowsToAdd.push(newRow);
    });
    
    // Add all rows at once for better performance
    if (rowsToAdd.length > 0) {
      log(`Adding ${rowsToAdd.length} job rows in batch`, LOG_LEVELS.INFO);
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, header.length)
        .setValues(rowsToAdd);
    }
    
    // Ensure data validation is applied
    setupDataValidation(sheet);
    
    return {
      status: 'success',
      message: `Sheet "${sheetName}" created successfully with ${rowsToAdd.length} jobs`
    };
  } catch (error) {
    return handleError('addManualJobs', error, true);
  }
}
