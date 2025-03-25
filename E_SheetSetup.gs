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
 * @param {boolean} [createIfMissing=true] - Whether to create the sheet if it doesn't exist
 * @returns {Sheet|null} The sheet object or null if sheet doesn't exist and createIfMissing is false
 */
function getOrCreateSheet(sheetName, customFields, createIfMissing = true) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    // Return null if sheet doesn't exist and we're not supposed to create it
    if (!sheet && !createIfMissing) {
      log(`Sheet ${sheetName} not found and not creating it`, LOG_LEVELS.INFO);
      return null;
    }
    
    // Create sheet if it doesn't exist and createIfMissing is true
    if (!sheet && createIfMissing) {
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
      
      // Format header row with improved styling
      const headerRange = sheet.getRange(1, 1, 1, fields.length);
      headerRange.setBackground('#37474f');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setVerticalAlignment('middle');
      headerRange.setHorizontalAlignment('center');
      headerRange.setFontSize(12);
      
      // Add proper padding to header cells
      sheet.setRowHeight(1, 32);
      
      // Setup validation
      setupDataValidation(sheet);
    }
    
    return sheet;
  } catch (error) {
    handleError('getOrCreateSheet', error, false);
    if (createIfMissing) {
      throw error; // Re-throw only if we were trying to create the sheet
    }
    return null;
  }
}

/**
 * Applies data validation for dropdown columns and sets up conditional formatting
 * for cell colors based on values
 * 
 * @param {Sheet} sheet - The sheet to set up validation on
 */
function setupDataValidation(sheet) {
  try {
    const config = getConfig();
    const fields = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Calculate a reasonable number of rows for validation
    // Limit to a reasonable number to avoid performance issues
    const dataRows = Math.max(1, sheet.getLastRow() - 1); // Existing data rows
    const additionalRows = 500; // Buffer for future rows
    const rowsToValidate = Math.min(dataRows + additionalRows, 1000); // Cap at 1000 rows
    
    // Find indices for validation columns
    const statusColIndex = fields.indexOf('Status');
    const typesColIndex = fields.indexOf('Type');
    const prioColIndex = fields.indexOf('Priority');
    
    // Create validation rules using Google Sheets' built-in validation
    if (statusColIndex >= 0 && config.statuses && config.statuses.length > 0) {
      const ruleStatus = SpreadsheetApp.newDataValidation()
        .requireValueInList(config.statuses, true) // Show dropdown with values
        .setAllowInvalid(false) // Require selection from list
        .build();
      
      // Apply validation to data rows + buffer
      sheet.getRange(2, statusColIndex+1, rowsToValidate, 1).setDataValidation(ruleStatus);
    }
    
    if (typesColIndex >= 0 && config.types && config.types.length > 0) {
      const ruleType = SpreadsheetApp.newDataValidation()
        .requireValueInList(config.types, true)
        .setAllowInvalid(false)
        .build();
      
      sheet.getRange(2, typesColIndex+1, rowsToValidate, 1).setDataValidation(ruleType);
    }
    
    if (prioColIndex >= 0 && config.priorities && config.priorities.length > 0) {
      const rulePrio = SpreadsheetApp.newDataValidation()
        .requireValueInList(config.priorities, true)
        .setAllowInvalid(false)
        .build();
      
      sheet.getRange(2, prioColIndex+1, rowsToValidate, 1).setDataValidation(rulePrio);
    }
    
    // Now set up conditional formatting for colors
    setupConditionalFormatting(sheet, config);
    
    log('Data validation set up successfully', LOG_LEVELS.DEBUG);
  } catch (error) {
    handleError('setupDataValidation', error, false);
  }
}

/**
 * Applies consistent table formatting to the sheet with improved row formatting
 * 
 * @param {Sheet} sheet - The sheet to format
 */
function applyTableFormatting(sheet) {
  try {
    if (!sheet) return;
    
    const lastRow = Math.max(sheet.getLastRow(), 2);
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol === 0) return;
    
    // Format data rows with better styling
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    
    // Apply font family and size to all data cells
    dataRange.setFontFamily('Arial');
    dataRange.setFontSize(11);
    dataRange.setVerticalAlignment('middle');
    
    // Important fix: Apply consistent formatting to ALL rows, not just existing ones
    // This ensures new rows added later will have the same formatting
    const maxRows = Math.max(lastRow + 100, 500); // Format extra rows for future additions
    
    // Alternate row colors and set borders for all rows including buffer
    for (let i = 2; i <= maxRows; i++) {
      const isEvenRow = (i % 2 === 0);
      const rowColor = isEvenRow ? '#f9f9f9' : '#ffffff';
      const rowRange = sheet.getRange(i, 1, 1, lastCol);
      
      // Set proper row height for better spacing
      sheet.setRowHeight(i, 28);
      
      // Set background color for alternating rows
      rowRange.setBackground(rowColor);
      
      // Add border only to the bottom of each row for cleaner look
      rowRange.setBorder(
        false, // top
        false, // left
        true,  // bottom
        false, // right
        false, // vertical
        false, // horizontal
        '#e0e0e0', 
        SpreadsheetApp.BorderStyle.SOLID
      );
    }
    
    // Add vertical borders to columns
    for (let i = 1; i <= lastCol; i++) {
      const colRange = sheet.getRange(1, i, maxRows, 1);
      colRange.setBorder(
        false, // top
        false, // left
        false, // bottom
        true,  // right
        false, // vertical
        false, // horizontal
        '#e0e0e0',
        SpreadsheetApp.BorderStyle.SOLID
      );
    }
    
    // Center-align specific columns (Status, Type, Priority)
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // Find and center columns by their header names
    const columnsToCenter = ['Status', 'Type', 'Priority', 'Jira Ticket', 'Job Link'];
    columnsToCenter.forEach(columnName => {
      const colIndex = headers.indexOf(columnName);
      if (colIndex >= 0) {
        const centerRange = sheet.getRange(2, colIndex + 1, maxRows - 1, 1);
        centerRange.setHorizontalAlignment('center');
      }
    });
    
    // Auto-resize columns for better readability
    for (let i = 1; i <= lastCol; i++) {
      sheet.autoResizeColumn(i);
    }
    
    log('Applied enhanced table formatting', LOG_LEVELS.INFO);
  } catch (error) {
    handleError('applyTableFormatting', error, false);
  }
}

/**
 * Sets up conditional formatting rules to color cells based on their values
 * with improved visual styling
 * 
 * @param {Sheet} sheet - The sheet to set up formatting on
 * @param {Object} config - Configuration object with field values and colors
 */
function setupConditionalFormatting(sheet, config) {
  try {
    // First check if we have colors defined
    const hasStatusColors = config.statuses && config.statusColors && 
                           config.statuses.length === config.statusColors.length;
    const hasTypeColors = config.types && config.typeColors && 
                         config.types.length === config.typeColors.length;
    const hasPriorityColors = config.priorities && config.priorityColors && 
                             config.priorities.length === config.priorityColors.length;
    
    // If no colors defined, don't bother with conditional formatting
    if (!hasStatusColors && !hasTypeColors && !hasPriorityColors) {
      return;
    }
    
    // Get header row
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Clear existing rules to avoid duplication
    sheet.clearConditionalFormatRules();
    
    // Initialize rules array
    const rules = [];
    
    // Status column formatting
    if (hasStatusColors) {
      const statusCol = headerRow.indexOf('Status');
      if (statusCol >= 0) {
        const statusRange = sheet.getRange(2, statusCol + 1, sheet.getMaxRows() - 1, 1);
        
        config.statuses.forEach((status, index) => {
          const color = config.statusColors[index];
          // Calculate a darker text color for better contrast
          const textColor = getDarkerShade(color);
          
          const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(status)
            .setBackground(color)
            .setFontColor(textColor)
            .setBold(true)
            .setRanges([statusRange])
            .build();
          
          rules.push(rule);
        });
      }
    }
    
    // Type column formatting
    if (hasTypeColors) {
      const typeCol = headerRow.indexOf('Type');
      if (typeCol >= 0) {
        const typeRange = sheet.getRange(2, typeCol + 1, sheet.getMaxRows() - 1, 1);
        
        config.types.forEach((type, index) => {
          const color = config.typeColors[index];
          const textColor = getDarkerShade(color);
          
          const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(type)
            .setBackground(color)
            .setFontColor(textColor)
            .setBold(false) // We're already making text darker
            .setRanges([typeRange])
            .build();
          
          rules.push(rule);
        });
      }
    }
    
    // Priority column formatting
    if (hasPriorityColors) {
      const priorityCol = headerRow.indexOf('Priority');
      if (priorityCol >= 0) {
        const priorityRange = sheet.getRange(2, priorityCol + 1, sheet.getMaxRows() - 1, 1);
        
        config.priorities.forEach((priority, index) => {
          const color = config.priorityColors[index];
          const textColor = getDarkerShade(color);
          
          const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(priority)
            .setBackground(color)
            .setFontColor(textColor)
            .setBold(true)
            .setRanges([priorityRange])
            .build();
          
          rules.push(rule);
        });
      }
    }
    
    // Apply all rules at once
    if (rules.length > 0) {
      sheet.setConditionalFormatRules(rules);
      log(`Applied ${rules.length} conditional formatting rules for colors`, LOG_LEVELS.INFO);
    }
    
    // Apply general table formatting
    applyTableFormatting(sheet);
    
  } catch (error) {
    handleError('setupConditionalFormatting', error, false);
  }
}

/**
 * Calculates a darker shade of a color for better text contrast
 * 
 * @param {string} hexColor - Hex color code (e.g., "#RRGGBB")
 * @param {number} [factor=0.6] - Darkening factor (lower = darker)
 * @returns {string} Darker hex color code
 */
function getDarkerShade(hexColor, factor = 0.6) {
  try {
    // Remove the # if it exists
    hexColor = hexColor.replace('#', '');
    
    // Convert to RGB
    const r = parseInt(hexColor.substring(0, 2), 16);
    const g = parseInt(hexColor.substring(2, 4), 16);
    const b = parseInt(hexColor.substring(4, 6), 16);
    
    // Darken the color
    const darkerR = Math.floor(r * factor);
    const darkerG = Math.floor(g * factor);
    const darkerB = Math.floor(b * factor);
    
    // Convert back to hex
    return '#' + 
      darkerR.toString(16).padStart(2, '0') + 
      darkerG.toString(16).padStart(2, '0') + 
      darkerB.toString(16).padStart(2, '0');
  } catch (e) {
    // In case of any error, return a default dark color
    return '#333333';
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
    // Get the sheet without creating it if it doesn't exist
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      log(`Sheet not found in findJobRow: ${sheetName}`, LOG_LEVELS.WARN);
      return { found: false, error: `Sheet not found: ${sheetName}` };
    }
    
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