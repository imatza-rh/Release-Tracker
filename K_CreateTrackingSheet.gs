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
 * @param {Object} [formatOptions={}] - Formatting options for the sheet
 * @returns {Object} Result object indicating success or failure
 */
function createTrackingSheetWithOptions(sheetName, source, fieldOptions, manualJobs = [], formatOptions = {}) {
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
      
      // Update colors if provided
      if (fieldOptions.statusColors && fieldOptions.statusColors.length === fieldOptions.statuses.length) {
        currentConfig.statusColors = fieldOptions.statusColors;
      }
      
      if (fieldOptions.typeColors && fieldOptions.typeColors.length === fieldOptions.types.length) {
        currentConfig.typeColors = fieldOptions.typeColors;
      }
      
      if (fieldOptions.priorityColors && fieldOptions.priorityColors.length === fieldOptions.priorities.length) {
        currentConfig.priorityColors = fieldOptions.priorityColors;
      }
      
      // Update selected fields if provided
      if (fieldOptions.selectedFields && Array.isArray(fieldOptions.selectedFields)) {
        currentConfig.fields = fieldOptions.selectedFields;
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
    
    // Set default format options if not provided
    const formatOpts = {
      formatAsTable: formatOptions.formatAsTable !== false,
      freezeHeaders: formatOptions.freezeHeaders !== false
    };
    
    // Create or get the sheet with proper headers and formatting
    log('Getting or creating sheet', LOG_LEVELS.INFO);
    const sheet = getOrCreateSheet(sheetName, fieldOptions.selectedFields);
    
    // If manual jobs were provided, add them to the sheet
    let result;
    if (source === 'manual' && Array.isArray(manualJobs) && manualJobs.length > 0) {
      result = addManualJobs(sheetName, manualJobs);
    } else if (source === 'empty') {
      result = {
        status: 'success',
        message: `Sheet "${sheetName}" created successfully with empty job list`
      };
    } else if (source === 'DefaultJobs') {
      // For DefaultJobs source, check if it exists first
      const defaultSheet = ss.getSheetByName(DEFAULT_JOBS_SHEET_NAME);
      if (!defaultSheet) {
        log('DefaultJobs sheet does not exist, creating it', LOG_LEVELS.INFO);
        createDefaultJobsTemplate();
      }
      
      // Add default jobs from DefaultJobs sheet
      result = addJobsFromDefaultSheet(sheetName);
    }
    
    // Apply table formatting
    applyTableFormatting(sheet, formatOpts);
    
    return result || {
      status: 'success',
      message: `Sheet "${sheetName}" created successfully`
    };
  } catch (error) {
    return handleError('createTrackingSheetWithOptions', error, true);
  }
}

/**
 * Applies table formatting to the sheet
 * 
 * @param {Sheet} sheet - The sheet to format
 * @param {Object} options - Formatting options
 */
function applyTableFormatting(sheet, options) {
  try {
    if (!sheet) return;
    
    // Get the data range
    const lastRow = Math.max(sheet.getLastRow(), 2); // Ensure at least header + 1 row
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol === 0) return; // No data to format
    
    // Apply table formatting
    if (options.formatAsTable) {
      // Format header row first
      const headerRange = sheet.getRange(1, 1, 1, lastCol);
      headerRange.setBackground("#f3f3f3");
      headerRange.setFontWeight("bold");
      headerRange.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
      
      // Apply zebra striping (alternating row colors) manually for each row
      for (let i = 2; i <= lastRow; i++) {
        // Use modulo to determine even/odd rows - add +1 if needed to adjust pattern
        const isEvenRow = (i % 2 === 0);
        const rowColor = isEvenRow ? "#f9f9f9" : "#ffffff"; // Light gray for even rows
        
        // Apply background color to this row
        const rowRange = sheet.getRange(i, 1, 1, lastCol);
        rowRange.setBackground(rowColor);
        
        // Add consistent borders to every cell in the row
        rowRange.setBorder(
          true,   // top
          true,   // left
          i === lastRow, // bottom - only add bottom border to last row
          true,   // right
          false,  // vertical
          false,  // horizontal
          "#e0e0e0", 
          SpreadsheetApp.BorderStyle.SOLID
        );
      }
      
      log("Applied table formatting with alternating rows", LOG_LEVELS.INFO);
    }
    
    // Freeze header row if requested
    if (options.freezeHeaders) {
      sheet.setFrozenRows(1);
      log("Froze header row", LOG_LEVELS.INFO);
    }
    
    // Auto-resize columns for better readability
    for (let i = 1; i <= lastCol; i++) {
      sheet.autoResizeColumn(i);
    }
    
    log("Auto-resized columns", LOG_LEVELS.INFO);
  } catch (error) {
    handleError('applyTableFormatting', error, false);
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
    const jobLinkIndex = header.indexOf('Job Link');
    const jiraTicketIndex = header.indexOf('Jira Ticket');
    
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
      
      // Add default values for optional fields if they exist in the header
      if (typeIndex !== -1 && config.types && config.types.length > 0) {
        newRow[typeIndex] = config.types[0] || 'Build';
      }
      
      if (statusIndex !== -1 && config.statuses && config.statuses.length > 0) {
        newRow[statusIndex] = config.statuses[0] || 'Pending';
      }
      
      if (priorityIndex !== -1 && config.priorities && config.priorities.length > 0) {
        newRow[priorityIndex] = config.priorities[1] || 'Medium'; // Default to Medium priority
      }
      
      // Handle Job Link column
      if (jobLinkIndex !== -1) {
        // Generate a link ID from the job name
        const linkId = jobName.toLowerCase().replace(/[^a-z0-9]+/g, '-');
        // Create the hyperlink formula
        newRow[jobLinkIndex] = createJenkinsJobLink(linkId);
      }
      
      // Handle Jira Ticket column
      if (jiraTicketIndex !== -1) {
        // Generate a random Jira-like ticket ID
        const ticketPrefix = "PROJ-";
        const ticketNumber = Math.floor(1000 + Math.random() * 9000); // 4-digit number
        const ticketId = ticketPrefix + ticketNumber;
        // Create the hyperlink formula
        newRow[jiraTicketIndex] = createJiraTicketLink(ticketId);
      }
      
      rowsToAdd.push(newRow);
    });
    
    // Add all rows at once for better performance
    if (rowsToAdd.length > 0) {
      log(`Adding ${rowsToAdd.length} job rows in batch`, LOG_LEVELS.INFO);
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, header.length)
        .setValues(rowsToAdd);
    }
    
    return {
      status: 'success',
      message: `Sheet "${sheetName}" created successfully with ${rowsToAdd.length} jobs`
    };
  } catch (error) {
    return handleError('addManualJobs', error, true);
  }
}