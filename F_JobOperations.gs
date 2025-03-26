/**
 * F_JobOperations.gs - Job CRUD operations
 * 
 * This file contains functions for managing jobs in the Release Tracker.
 * 
 * @requires JOBS_SHEET_NAME from A_Constants.gs
 * @requires log, handleError, safeExecute from B_Logging.gs
 * @requires sanitizeInput from C_Utils.gs
 * @requires getConfig from D_Config.gs
 * @requires getOrCreateSheet, findJobRow from E_SheetSetup.gs
 * @requires createJenkinsJobLink, createJiraTicketLink from J_Integrations.gs
 */

/**
 * Get job schema - centralized definition of job structure
 * 
 * @returns {Object} Schema definition with default values
 */
function getJobSchema() {
  const config = getConfig();
  return {
    "Job Name": "",
    "Type": config.types && config.types.length > 0 ? config.types[0] : "Build",
    "Status": config.statuses && config.statuses.length > 0 ? config.statuses[0] : "Pending",
    "Priority": config.priorities && config.priorities.length > 0 ? config.priorities[1] : "Medium",
    "Notes": "",
    "Job Link": "",
    "Jira Ticket": ""
  };
}

/**
 * Validate job object against schema and apply defaults
 * 
 * @param {Object} jobObj - Job object to validate
 * @param {boolean} [applyDefaults=true] - Whether to apply default values
 * @returns {Object} Validated job object with defaults applied
 */
function validateJobObject(jobObj, applyDefaults = true) {
  return safeExecute('validateJobObject', () => {
    const schema = getJobSchema();
    const validJob = {};
    
    // Apply schema defaults and validation
    for (const field in schema) {
      if (field in jobObj && jobObj[field] !== undefined && jobObj[field] !== null) {
        validJob[field] = sanitizeInput(jobObj[field]);
      } else if (applyDefaults) {
        validJob[field] = schema[field];
      }
    }
    
    // Ensure Job Name is sanitized
    if ("Job Name" in validJob) {
      validJob["Job Name"] = sanitizeInput(validJob["Job Name"]);
    }
    
    return validJob;
  }, false);
}

/**
 * Format job object for display (handle formulas, etc.)
 * 
 * @param {Object} jobObj - Raw job object from sheet
 * @returns {Object} Formatted job object for display
 */
function formatJobForDisplay(jobObj) {
  return safeExecute('formatJobForDisplay', () => {
    const formattedJob = {};
    
    for (const key in jobObj) {
      let value = jobObj[key];
      
      // Handle hyperlink formulas
      if (typeof value === 'string' && value.startsWith('=HYPERLINK')) {
        const match = value.match(/=HYPERLINK\("[^"]*","([^"]*)"\)/);
        if (match && match[1]) {
          value = match[1]; // Extract display text
        }
      }
      
      formattedJob[key] = value;
    }
    
    return formattedJob;
  }, false);
}

/**
 * Insert one job row
 * 
 * @interface Used by ManageSidebar.html
 * @param {Object} jobObj - Job data to add
 * @param {string} [targetSheetName=JOBS_SHEET_NAME] - Name of sheet to add job to
 * @returns {Object} Result object indicating success or failure
 */
function addJob(jobObj, targetSheetName = JOBS_SHEET_NAME) {
  return safeExecute('addJob', () => {
    // Make sure the sheet name is properly sanitized
    targetSheetName = String(targetSheetName || JOBS_SHEET_NAME).trim();
    log(`Adding job to sheet: "${targetSheetName}"`, LOG_LEVELS.INFO);
    
    // Get the sheet - allow creation if needed
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(targetSheetName);
    
    if (!sheet) {
      log(`Sheet "${targetSheetName}" not found in addJob, creating it`, LOG_LEVELS.WARN);
      // Create the sheet with default fields if it doesn't exist
      sheet = getOrCreateSheet(targetSheetName, null, true);
      
      if (!sheet) {
        return {
          status: 'error',
          message: `Failed to create or access sheet "${targetSheetName}"`
        };
      }
    }
    
    // Validate and sanitize the job object
    jobObj = validateJobObject(jobObj);
    
    // Handle integration features - create hyperlinks for Job Link and Jira Ticket
    if (jobObj["Job Link"] && typeof jobObj["Job Link"] === 'string') {
      jobObj["Job Link"] = createJenkinsJobLink(jobObj["Job Link"]);
    }
    
    if (jobObj["Jira Ticket"] && typeof jobObj["Jira Ticket"] === 'string') {
      jobObj["Jira Ticket"] = createJiraTicketLink(jobObj["Jira Ticket"]);
    }
    
    // Read the existing header
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Build a new row in correct order
    const newRow = header.map(colName => {
      if (colName in jobObj) {
        return jobObj[colName];
      }
      return "";
    });
    
    // Add the row
    sheet.appendRow(newRow);
    
    log('Job added successfully', LOG_LEVELS.INFO, jobObj);
    return { 
      status: 'success', 
      message: `Job added: ${jobObj["Job Name"] || "(unnamed)"}`,
      jobName: jobObj["Job Name"] || "(unnamed)"
    };
  }, false);
}

/**
 * Returns an array of job objects, with optional pagination
 * 
 * @interface Used by ViewSidebar.html
 * @param {number} [page=1] - Page number for pagination
 * @param {number} [pageSize=0] - Items per page (0 for all)
 * @param {string} [sheetName=JOBS_SHEET_NAME] - Name of sheet to get jobs from
 * @returns {Object} Result object with jobs array
 */
function getJobs(page = 1, pageSize = 0, sheetName = JOBS_SHEET_NAME) {
  return safeExecute('getJobs', () => {
    // Make sure we always have a valid string for the sheet name, not undefined, null, or other types
    sheetName = String(sheetName || JOBS_SHEET_NAME).trim();
    
    log(`Getting jobs from sheet: "${sheetName}"`, LOG_LEVELS.INFO);
    
    // Check if the sheet exists 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      log(`Sheet "${sheetName}" not found in getJobs`, LOG_LEVELS.WARN);
      return { 
        status: 'error', 
        message: `Sheet "${sheetName}" not found`,
        jobs: [] 
      };
    }
    
    // Get all values from the sheet - we want everything from column A to the last column
    // and from row 1 to the last row with data
    const values = sheet.getDataRange().getValues();
    
    if (values.length < 2) {
      log(`No job data found in ${sheetName}`, LOG_LEVELS.INFO);
      return { 
        status: 'success', 
        message: 'No jobs found',
        jobs: [] 
      };
    }
    
    const header = values[0]; // First row = column names
    const allJobs = values.slice(1).map(row => {
      let jobObj = {};
      header.forEach((col, index) => {
        jobObj[col] = row[index];
      });
      return formatJobForDisplay(jobObj);
    });
    
    // Apply pagination if requested
    if (pageSize > 0) {
      const startIndex = (page - 1) * pageSize;
      const endIndex = startIndex + pageSize;
      const paginatedJobs = allJobs.slice(startIndex, Math.min(endIndex, allJobs.length));
      
      return {
        status: 'success',
        jobs: paginatedJobs,
        totalJobs: allJobs.length,
        currentPage: page,
        totalPages: Math.ceil(allJobs.length / pageSize)
      };
    }
    
    // Return all jobs if no pagination
    log(`Fetched ${allJobs.length} jobs from ${sheetName}`, LOG_LEVELS.DEBUG);
    return { 
      status: 'success', 
      jobs: allJobs 
    };
  }, false);
}

/**
 * Updates specified fields for a job by name
 * 
 * @interface Used by ManageSidebar.html
 * @param {string} jobName - Name of job to update
 * @param {Object} updates - Object with fields to update
 * @param {string} [sheetName=JOBS_SHEET_NAME] - Name of sheet to update job in
 * @returns {Object} Result object indicating success or failure
 */
function updateJob(jobName, updates, sheetName = JOBS_SHEET_NAME) {
  return safeExecute('updateJob', () => {
    // Make sure we have a valid sheet name
    sheetName = String(sheetName || JOBS_SHEET_NAME).trim();
    log(`Updating job "${jobName}" in sheet: "${sheetName}"`, LOG_LEVELS.INFO);
    
    // Use findJobRow which has been fixed to properly check for sheet existence
    const result = findJobRow(jobName, sheetName);
    
    if (!result.found) {
      return {
        status: 'error',
        message: result.error || `Could not find job "${jobName}" in sheet "${sheetName}"`
      };
    }
    
    const { rowNumber, sheet, header } = result;
    
    // Validate the updates
    updates = validateJobObject(updates, false);
    
    // Process special fields that need formatting
    if ('Job Link' in updates && updates['Job Link']) {
      updates['Job Link'] = createJenkinsJobLink(updates['Job Link']);
    }
    
    if ('Jira Ticket' in updates && updates['Jira Ticket']) {
      updates['Jira Ticket'] = createJiraTicketLink(updates['Jira Ticket']);
    }
    
    // Batch update to minimize API calls
    const updateRanges = [];
    const updateValues = [];
    
    for (let field in updates) {
      const colIndex = header.indexOf(field);
      if (colIndex >= 0) {
        updateRanges.push(sheet.getRange(rowNumber, colIndex+1));
        updateValues.push(updates[field]);
      }
    }
    
    // Perform the updates if any
    if (updateRanges.length > 0) {
      const batchSize = 10; // Google's limit for batch operations
      
      // Update in batches
      for (let i = 0; i < updateRanges.length; i += batchSize) {
        const batchRanges = updateRanges.slice(i, i + batchSize);
        const batchValues = updateValues.slice(i, i + batchSize);
        
        batchRanges.forEach((range, index) => {
          range.setValue(batchValues[index]);
        });
      }
    }
    
    log('Job updated successfully', LOG_LEVELS.INFO, { jobName, updates });
    return { 
      status: 'success', 
      message: `Job updated: ${jobName}` 
    };
  }, false);
}

/**
 * Removes a row by job name
 * 
 * @interface Used by ManageSidebar.html
 * @param {string} jobName - Name of job to delete
 * @param {string} [sheetName=JOBS_SHEET_NAME] - Name of sheet to delete job from
 * @returns {Object} Result object indicating success or failure
 */
function deleteJobByName(jobName, sheetName = JOBS_SHEET_NAME) {
  return safeExecute('deleteJobByName', () => {
    // Make sure we have a valid sheet name
    sheetName = String(sheetName || JOBS_SHEET_NAME).trim();
    log(`Deleting job "${jobName}" from sheet: "${sheetName}"`, LOG_LEVELS.INFO);
    
    const result = findJobRow(jobName, sheetName);
    
    if (!result.found) {
      return {
        status: 'error',
        message: result.error || `Could not find job "${jobName}" in sheet "${sheetName}"`
      };
    }
    
    const { rowNumber, sheet } = result;
    sheet.deleteRow(rowNumber);
    
    log('Job deleted successfully', LOG_LEVELS.INFO, { jobName });
    return { 
      status: 'success', 
      message: `Deleted job: ${jobName}` 
    };
  }, false);
}

/**
 * Sets "Status" to "Done" or "Completed"
 * 
 * @param {string} jobName - Name of job to mark as done
 * @param {string} [sheetName=JOBS_SHEET_NAME] - Name of sheet to update job in
 * @returns {Object} Result object indicating success or failure
 */
function markJobDone(jobName, sheetName = JOBS_SHEET_NAME) {
  return safeExecute('markJobDone', () => {
    const config = getConfig();
    let doneStatus = "Done";
    
    if (config.statuses.includes("Completed")) {
      doneStatus = "Completed";
    }
    
    return updateJob(jobName, { "Status": doneStatus }, sheetName);
  }, false);
}

/**
 * Finds and activates the row with the given job name
 * 
 * @interface Used by ViewSidebar.html
 * @param {string} jobName - Name of job to locate
 * @param {string} [sheetName=JOBS_SHEET_NAME] - Name of sheet to search in
 * @returns {Object} Result object indicating success or failure
 */
function locateInSheet(jobName, sheetName = JOBS_SHEET_NAME) {
  return safeExecute('locateInSheet', () => {
    // Validate parameters
    if (!jobName) {
      return {
        status: 'error',
        message: "No job name specified"
      };
    }
    
    // Make sure we have a valid sheet name
    sheetName = String(sheetName || JOBS_SHEET_NAME).trim();
    log(`Locating job "${jobName}" in sheet: "${sheetName}"`, LOG_LEVELS.INFO);
    
    // Check if the sheet exists first
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      log(`Sheet "${sheetName}" not found in locateInSheet`, LOG_LEVELS.WARN);
      return {
        status: 'error',
        message: `Sheet "${sheetName}" not found`
      };
    }
    
    // Find the job row
    const result = findJobRow(jobName, sheetName);
    
    if (!result.found) {
      return {
        status: 'error',
        message: result.error || `Could not find job "${jobName}" in sheet "${sheetName}"`
      };
    }
    
    const { rowNumber, sheet: jobSheet, header } = result;
    
    // Highlight the row
    ss.setActiveSheet(jobSheet);
    
    // We can highlight from column 1 to the last column of the header
    const lastCol = header.length;
    jobSheet.setActiveRange(jobSheet.getRange(rowNumber, 1, 1, lastCol));
    
    log('Job located successfully', LOG_LEVELS.INFO, { jobName, rowNumber });
    return { 
      status: 'success', 
      message: `Located job "${jobName}" at row ${rowNumber} in "${sheetName}"` 
    };
  }, false);
}

/**
 * For "ManageSidebar" to populate UI elements
 * 
 * @interface Used by ManageSidebar.html
 * @param {string} [sheetName=JOBS_SHEET_NAME] - Name of sheet to get data from
 * @returns {Object} Data object with job configuration
 */
function getInitData(sheetName = JOBS_SHEET_NAME) {
  return safeExecute('getInitData', () => {
    // Make sure we have a valid sheet name
    sheetName = String(sheetName || JOBS_SHEET_NAME).trim();
    log(`Getting initialization data for sheet: "${sheetName}"`, LOG_LEVELS.INFO);
    
    const config = getConfig();
    
    // Get jobs only if the sheet exists
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    let jobList = [];
    if (sheet) {
      jobList = getJobs(1, 0, sheetName).jobs
        .map(j => sanitizeInput(j["Job Name"]))
        .filter(n => n)
        .sort();
    }
    
    return {
      status: 'success',
      statuses:   config.statuses,
      jobTypes:   config.types,
      priorities: config.priorities,
      jobNames:   jobList
    };
  }, false);
}

/**
 * Lists all job sheets in the spreadsheet
 * 
 * @returns {Array} List of sheet names
 */
function listJobSheets() {
  return safeExecute('listJobSheets', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    // Filter to sheets that look like job tracking sheets
    // They should have at least "Job Name", "Status" columns
    const jobSheets = sheets
      .filter(sheet => {
        if (sheet.getName() === DEFAULT_JOBS_SHEET_NAME) return false;
        
        const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        return headerRow.includes("Job Name") && headerRow.includes("Status");
      })
      .map(sheet => sheet.getName());
    
    log(`Found job sheets: ${jobSheets.join(', ')}`, LOG_LEVELS.INFO);
    return jobSheets;
  }, false);
}