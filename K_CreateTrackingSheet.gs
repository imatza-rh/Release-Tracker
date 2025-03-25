/**
 * K_CreateTrackingSheet.gs - Functions for creating new tracking sheets
 * 
 * This file contains functions specific to creating new tracking sheets
 * with user-defined options and enhanced visual styling.
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
    
    // Apply enhanced table formatting
    applyEnhancedTableFormatting(sheet, formatOpts);
    
    return result || {
      status: 'success',
      message: `Sheet "${sheetName}" created successfully`
    };
  } catch (error) {
    return handleError('createTrackingSheetWithOptions', error, true);
  }
}

/**
 * Applies enhanced table formatting to the sheet
 * 
 * @param {Sheet} sheet - The sheet to format
 * @param {Object} options - Formatting options
 */
function applyEnhancedTableFormatting(sheet, options) {
  try {
    if (!sheet) return;
    
    // Get the data range
    const lastRow = Math.max(sheet.getLastRow(), 2); // Ensure at least header + 1 row
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol === 0) return; // No data to format
    
    // Apply table formatting
    if (options.formatAsTable) {
      // Format header row with a modern look
      const headerRange = sheet.getRange(1, 1, 1, lastCol);
      headerRange.setBackground("#37474f"); // Dark blue-gray
      headerRange.setFontColor("#ffffff"); // White text
      headerRange.setFontWeight("bold");
      headerRange.setVerticalAlignment("middle");
      headerRange.setHorizontalAlignment("center");
      
      // Set proper row height for header
      sheet.setRowHeight(1, 32);
      
      // Apply consistent font to all cells
      const allRange = sheet.getRange(1, 1, lastRow, lastCol);
      allRange.setFontFamily("Arial");
      
      // Format data rows
      const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      dataRange.setFontSize(11);
      dataRange.setVerticalAlignment("middle");
      
      // Apply zebra striping (alternating row colors) with subtle colors
      for (let i = 2; i <= lastRow; i++) {
        // Use modulo to determine even/odd rows
        const isEvenRow = (i % 2 === 0);
        const rowColor = isEvenRow ? "#f9f9f9" : "#ffffff"; // Light gray for even rows
        
        // Get the row range
        const rowRange = sheet.getRange(i, 1, 1, lastCol);
        
        // Set row height for better spacing
        sheet.setRowHeight(i, 28);
        
        // Apply background color
        rowRange.setBackground(rowColor);
        
        // Add bottom border for cleaner look
        rowRange.setBorder(
          false,  // top
          false,  // left
          true,   // bottom
          false,  // right
          false,  // vertical
          false,  // horizontal
          "#e0e0e0", 
          SpreadsheetApp.BorderStyle.SOLID
        );
      }
      
      // Add column borders for better readability
      for (let i = 1; i <= lastCol; i++) {
        const colRange = sheet.getRange(1, i, lastRow, 1);
        colRange.setBorder(
          false,  // top
          false,  // left
          false,  // bottom
          true,   // right
          false,  // vertical
          false,  // horizontal
          "#e0e0e0", 
          SpreadsheetApp.BorderStyle.SOLID
        );
      }
      
      // Apply centered alignment to specific columns
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      
      // Find and center columns by their header names
      const columnsToCenter = ['Status', 'Type', 'Priority', 'Jira Ticket', 'Job Link'];
      columnsToCenter.forEach(columnName => {
        const colIndex = headers.indexOf(columnName);
        if (colIndex >= 0) {
          const centerRange = sheet.getRange(2, colIndex + 1, lastRow - 1, 1);
          centerRange.setHorizontalAlignment('center');
        }
      });
      
      log("Applied enhanced table formatting", LOG_LEVELS.INFO);
    }
    
    // Freeze header row if requested
    if (options.freezeHeaders) {
      sheet.setFrozenRows(1);
      log("Froze header row", LOG_LEVELS.INFO);
    }
    
    // Auto-resize columns for better readability
    for (let i = 1; i <= lastCol; i++) {
      sheet.autoResizeColumn(i);
      
      // Ensure minimum column width for readability
      if (sheet.getColumnWidth(i) < 100) {
        // Only widen if it's not the Job Name column (which might need to be wider)
        if (i !== 1) { // Assuming Job Name is first column
          sheet.setColumnWidth(i, 100);
        }
      }
    }
    
    // Make sure the Job Name column is wide enough
    const jobNameIndex = 1; // Assuming Job Name is the first column
    if (sheet.getColumnWidth(jobNameIndex) < 200) {
      sheet.setColumnWidth(jobNameIndex, 200);
    }
    
    log("Optimized column widths for better readability", LOG_LEVELS.INFO);
    
    // Setup conditional formatting for colors
    setupDataValidation(sheet);
    
  } catch (error) {
    handleError('applyEnhancedTableFormatting', error, false);
  }
}

/**
 * Adds manually entered jobs to a tracking sheet with improved styling
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
    const notesIndex = header.indexOf('Notes');
    
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
      
      // Add some sample notes based on the job name
      if (notesIndex !== -1) {
        if (jobName.toLowerCase().includes('database') || jobName.toLowerCase().includes('db')) {
          newRow[notesIndex] = 'Database-related task';
        } else if (jobName.toLowerCase().includes('deploy')) {
          newRow[notesIndex] = 'Deployment to production environment';
        } else if (jobName.toLowerCase().includes('test')) {
          newRow[notesIndex] = 'Testing functionality and regression';
        } else if (jobName.toLowerCase().includes('review')) {
          newRow[notesIndex] = 'Code review for recent changes';
        }
      }
      
      rowsToAdd.push(newRow);
    });
    
    // Add all rows at once for better performance
    if (rowsToAdd.length > 0) {
      log(`Adding ${rowsToAdd.length} job rows in batch`, LOG_LEVELS.INFO);
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, header.length)
        .setValues(rowsToAdd);
    }
    
    // Apply enhanced table formatting
    applyEnhancedTableFormatting(sheet, {
      formatAsTable: true,
      freezeHeaders: true
    });
    
    return {
      status: 'success',
      message: `Sheet "${sheetName}" created successfully with ${rowsToAdd.length} jobs`
    };
  } catch (error) {
    return handleError('addManualJobs', error, true);
  }
}