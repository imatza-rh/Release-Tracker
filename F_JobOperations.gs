/**
 * F_JobOperations.gs - Job CRUD operations
 * 
 * This file contains functions for managing jobs in the Release Tracker.
 * 
 * @requires JOBS_SHEET_NAME from A_Constants.gs
 * @requires log, handleError from B_Logging.gs
 * @requires sanitizeInput from C_Utils.gs
 * @requires getConfig from D_Config.gs
 * @requires getOrCreateSheet, findJobRow from E_SheetSetup.gs
 * @requires createJenkinsJobLink, createJiraTicketLink from J_Integrations.gs
 */

/**
 * Insert one job row
 * 
 * @interface Used by ManageSidebar.html
 * @param {Object} jobObj - Job data to add
 * @param {string} [targetSheetName=JOBS_SHEET_NAME] - Name of sheet to add job to
 * @returns {Object} Result object indicating success or failure
 */
function addJob(jobObj, targetSheetName = JOBS_SHEET_NAME) {
  try {
    const sheet = getOrCreateSheet(targetSheetName);
    
    // Sanitize input
    if (jobObj["Job Name"]) {
      jobObj["Job Name"] = sanitizeInput(jobObj["Job Name"]);
    }
    
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
  } catch (error) {
    return handleError('addJob', error, false);
  }
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
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet) {
      log(`Sheet ${sheetName} not found`, LOG_LEVELS.WARN);
      return { 
        status: 'error', 
        message: `Sheet ${sheetName} not found`,
        jobs: [] 
      };
    }
    
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
        // Skip formula cells for display purposes (hyperlinks)
        if (row[index] && typeof row[index] === 'string' && row[index].startsWith('=HYPERLINK')) {
          // Extract the display text from the hyperlink formula
          const match = row[index].match(/=HYPERLINK\("[^"]*","([^"]*)"\)/);
          if (match && match[1]) {
            jobObj[col] = match[1]; // Use the display text
          } else {
            jobObj[col] = row[index]; // Fallback to the formula if we can't parse it
          }
        } else {
          jobObj[col] = row[index];
        }
      });
      return jobObj;
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
  } catch (error) {
    return handleError('getJobs', error, false);
  }
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
  try {
    const result = findJobRow(jobName, sheetName);
    
    if (!result.found) {
      throw new Error(result.error);
    }
    
    const { rowNumber, sheet, header } = result;
    
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
  } catch (error) {
    return handleError('updateJob', error, false);
  }
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
  try {
    const result = findJobRow(jobName, sheetName);
    
    if (!result.found) {
      throw new Error(result.error);
    }
    
    const { rowNumber, sheet } = result;
    sheet.deleteRow(rowNumber);
    
    log('Job deleted successfully', LOG_LEVELS.INFO, { jobName });
    return { 
      status: 'success', 
      message: `Deleted job: ${jobName}` 
    };
  } catch (error) {
    return handleError('deleteJobByName', error, false);
  }
}

/**
 * Sets "Status" to "Done" or "Completed"
 * 
 * @param {string} jobName - Name of job to mark as done
 * @param {string} [sheetName=JOBS_SHEET_NAME] - Name of sheet to update job in
 * @returns {Object} Result object indicating success or failure
 */
function markJobDone(jobName, sheetName = JOBS_SHEET_NAME) {
  try {
    const config = getConfig();
    let doneStatus = "Done";
    
    if (config.statuses.includes("Completed")) {
      doneStatus = "Completed";
    }
    
    return updateJob(jobName, { "Status": doneStatus }, sheetName);
  } catch (error) {
    return handleError('markJobDone', error, false);
  }
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
  try {
    if (!jobName) {
      throw new Error("No job name specified");
    }
    
    const result = findJobRow(jobName, sheetName);
    
    if (!result.found) {
      throw new Error(result.error);
    }
    
    const { rowNumber, sheet, header } = result;
    
    // Highlight the row
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setActiveSheet(sheet);
    
    // We can highlight from column 1 to the last column of the header
    const lastCol = header.length;
    sheet.setActiveRange(sheet.getRange(rowNumber, 1, 1, lastCol));
    
    log('Job located successfully', LOG_LEVELS.INFO, { jobName, rowNumber });
    return { 
      status: 'success', 
      message: `Located job "${jobName}" at row ${rowNumber} in "${sheetName}"` 
    };
  } catch (error) {
    return handleError('locateInSheet', error, false);
  }
}

/**
 * For "ManageSidebar" to populate UI elements
 * 
 * @interface Used by ManageSidebar.html
 * @param {string} [sheetName=JOBS_SHEET_NAME] - Name of sheet to get data from
 * @returns {Object} Data object with job configuration
 */
function getInitData(sheetName = JOBS_SHEET_NAME) {
  try {
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
  } catch (error) {
    return handleError('getInitData', error, false);
  }
}

/**
 * Lists all job sheets in the spreadsheet
 * 
 * @returns {Array} List of sheet names
 */
function listJobSheets() {
  try {
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
    
    return jobSheets;
  } catch (error) {
    log('Error listing job sheets', LOG_LEVELS.ERROR, error);
    return [JOBS_SHEET_NAME]; // Return default as fallback
  }
}
