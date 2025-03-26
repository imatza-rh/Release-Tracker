/**
 * E_SheetSetup.gs - Sheet creation and formatting
 * 
 * This file contains functions for setting up and managing spreadsheet
 * sheets for the Release Tracker application.
 * 
 * @requires JOBS_SHEET_NAME from A_Constants.gs
 * @requires log, handleError, safeExecute from B_Logging.gs
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
  return safeExecute('getOrCreateSheet', () => {
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
      
      // Apply formatting with default options
      formatSheet(sheet, {
        formatAsTable: true,
        freezeHeaders: true,
        rowHeight: 28,
        headerRowHeight: 32
      });
      
      // Setup validation
      setupDataValidation(sheet);
    }
    
    return sheet;
  }, false);
}

/**
 * Applies comprehensive formatting to a sheet with configurable options
 * 
 * @param {Sheet} sheet - The sheet to format
 * @param {Object} [options={}] - Formatting options
 */
function formatSheet(sheet, options = {}) {
  return safeExecute('formatSheet', () => {
    if (!sheet) return;
    
    // Default options
    const opts = {
      formatAsTable: true,
      freezeHeaders: true,
      headerBackground: "#37474f",
      headerFontColor: "#ffffff",
      evenRowColor: "#f9f9f9",
      oddRowColor: "#ffffff",
      borderColor: "#e0e0e0",
      rowHeight: 28,
      headerRowHeight: 32,
      fontFamily: "Arial",
      fontSize: 11,
      centerColumns: ['Status', 'Type', 'Priority', 'Jira Ticket', 'Job Link']
    };
    
    // Override defaults with provided options
    Object.keys(options).forEach(key => {
      if (options[key] !== undefined) {
        opts[key] = options[key];
      }
    });
    
    // Get the data range
    const lastRow = Math.max(sheet.getLastRow(), 2); // Ensure at least header + 1 row
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol === 0) return; // No data to format
    
    // Apply formatting
    if (opts.formatAsTable) {
      // Format header row with a modern look
      const headerRange = sheet.getRange(1, 1, 1, lastCol);
      headerRange.setBackground(opts.headerBackground);
      headerRange.setFontColor(opts.headerFontColor);
      headerRange.setFontWeight("bold");
      headerRange.setVerticalAlignment("middle");
      headerRange.setHorizontalAlignment("center");
      headerRange.setFontFamily(opts.fontFamily);
      
      // Set proper row height for header
      sheet.setRowHeight(1, opts.headerRowHeight);
      
      // Apply consistent font to all cells
      const allRange = sheet.getRange(1, 1, lastRow, lastCol);
      allRange.setFontFamily(opts.fontFamily);
      allRange.setFontSize(opts.fontSize);
      
      // Format data rows - Apply alternating colors and borders
      formatDataRows(sheet, lastRow, lastCol, opts);
      
      // Apply centered alignment to specific columns
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      
      // Apply centered alignment to columns
      centerColumnsByName(sheet, headers, opts.centerColumns, lastRow);
      
      log("Applied enhanced table formatting", LOG_LEVELS.INFO);
    }
    
    // Freeze header row if requested
    if (opts.freezeHeaders) {
      sheet.setFrozenRows(1);
      log("Froze header row", LOG_LEVELS.INFO);
    }
    
    // Auto-resize columns for better readability
    for (let i = 1; i <= lastCol; i++) {
      sheet.autoResizeColumn(i);
      
      // Ensure minimum column width for readability (except for Job Name which needs to be wider)
      if (sheet.getColumnWidth(i) < 100 && i !== 1) {
        sheet.setColumnWidth(i, 100);
      }
    }
    
    // Make sure the Job Name column is wide enough
    const jobNameIndex = headers ? headers.indexOf("Job Name") + 1 : 1;
    if (jobNameIndex && sheet.getColumnWidth(jobNameIndex) < 200) {
      sheet.setColumnWidth(jobNameIndex, 200);
    }
    
    log("Optimized column widths for better readability", LOG_LEVELS.INFO);
    
    // Setup conditional formatting for colors
    setupDataValidation(sheet);
  }, false);
}

/**
 * Format data rows with alternating colors and proper borders
 * 
 * @param {Sheet} sheet - The sheet to format
 * @param {number} lastRow - Last row number with data
 * @param {number} lastCol - Last column number with data
 * @param {Object} opts - Formatting options
 */
function formatDataRows(sheet, lastRow, lastCol, opts) {
  // Common maximum row limit to prevent excessive operations
  const maxRows = Math.max(lastRow + 50, 500); // Format extra rows for future additions
  
  // Process EACH row individually for consistency
  for (let i = 2; i <= maxRows; i++) {
    // Use modulo to determine even/odd rows
    const isEvenRow = (i % 2 === 0);
    const rowColor = isEvenRow ? opts.evenRowColor : opts.oddRowColor;
    
    // Get the row range
    const rowRange = sheet.getRange(i, 1, 1, lastCol);
    
    // Set row height for better spacing
    sheet.setRowHeight(i, opts.rowHeight);
    
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
      opts.borderColor, 
      SpreadsheetApp.BorderStyle.SOLID
    );
  }
  
  // Add column borders for better readability
  for (let i = 1; i <= lastCol; i++) {
    const colRange = sheet.getRange(1, i, maxRows, 1);
    colRange.setBorder(
      false,  // top
      false,  // left
      false,  // bottom
      true,   // right
      false,  // vertical
      false,  // horizontal
      opts.borderColor, 
      SpreadsheetApp.BorderStyle.SOLID
    );
  }
}

/**
 * Centers specified columns by their names
 * 
 * @param {Sheet} sheet - The sheet to format
 * @param {Array} headers - Array of header names
 * @param {Array} columnsToCenter - Array of column names to center
 * @param {number} lastRow - Last row number with data
 */
function centerColumnsByName(sheet, headers, columnsToCenter, lastRow) {
  if (!Array.isArray(columnsToCenter) || !Array.isArray(headers)) return;
  
  columnsToCenter.forEach(columnName => {
    const colIndex = headers.indexOf(columnName);
    if (colIndex >= 0) {
      const centerRange = sheet.getRange(2, colIndex + 1, lastRow - 1, 1);
      centerRange.setHorizontalAlignment('center');
    }
  });
}

/**
 * Applies data validation for dropdown columns and sets up conditional formatting
 * for cell colors based on values
 * 
 * @param {Sheet} sheet - The sheet to set up validation on
 */
function setupDataValidation(sheet) {
  return safeExecute('setupDataValidation', () => {
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
  }, false);
}

/**
 * Sets up conditional formatting rules to color cells based on their values
 * with improved visual styling
 * 
 * @param {Sheet} sheet - The sheet to set up formatting on
 * @param {Object} config - Configuration object with field values and colors
 */
function setupConditionalFormatting(sheet, config) {
  return safeExecute('setupConditionalFormatting', () => {
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
  }, false);
}

/**
 * Calculates a darker shade of a color for better text contrast
 * 
 * @param {string} hexColor - Hex color code (e.g., "#RRGGBB")
 * @param {number} [factor=0.6] - Darkening factor (lower = darker)
 * @returns {string} Darker hex color code
 */
function getDarkerShade(hexColor, factor = 0.6) {
  return safeExecute('getDarkerShade', () => {
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
  }, false) || '#333333'; // Default to a dark color if an error occurs
}

/**
 * Helper function to find a row by job name
 * 
 * @param {string} jobName - Name of the job to find
 * @param {string} [sheetName=JOBS_SHEET_NAME] - Name of sheet to search in
 * @returns {Object} Result object with job data or error information
 */
function findJobRow(jobName, sheetName = JOBS_SHEET_NAME) {
  return safeExecute('findJobRow', () => {
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
  }, false) || { found: false, error: 'Error in findJobRow function' };
}