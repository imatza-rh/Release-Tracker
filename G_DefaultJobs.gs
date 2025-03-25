/**
 * G_DefaultJobs.gs - Job initialization functionality
 * 
 * @requires LOG_LEVELS from A_Constants.gs
 * @requires JOBS_SHEET_NAME, DEFAULT_JOBS_SHEET_NAME from A_Constants.gs
 * @requires log, handleError from B_Logging.gs
 * @requires sanitizeInput from C_Utils.gs
 * @requires getConfig from D_Config.gs
 * @requires getOrCreateSheet, setupDataValidation from E_SheetSetup.gs
 * @requires createJenkinsJobLink, createJiraTicketLink from J_Integrations.gs
 */

/**
 * Adds jobs from DefaultJobs sheet to the specified sheet
 * 
 * @param {string} targetSheetName - Name of sheet to add jobs to
 * @returns {Object} Result object indicating success, info, or failure
 */
function addJobsFromDefaultSheet(targetSheetName = JOBS_SHEET_NAME) {
  try {
    // Get or create the DefaultJobs sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let defSheet = ss.getSheetByName(DEFAULT_JOBS_SHEET_NAME);
    
    if (!defSheet) {
      return {
        status: 'info',
        message: `DefaultJobs sheet not found. Use the Create DefaultJobs option first.`
      };
    }
    
    // Get target sheet
    let targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) {
      // Create the target sheet if it doesn't exist
      targetSheet = getOrCreateSheet(targetSheetName);
    }
    
    // Get existing job names to check for duplicates
    const existingJobsRange = targetSheet.getRange(2, 1, Math.max(1, targetSheet.getLastRow()-1), 1);
    const existingJobsData = existingJobsRange.getValues();
    const existingJobs = existingJobsData.map(row => sanitizeInput(row[0]).toLowerCase());
    const existingSet = new Set(existingJobs);
    
    // Read default jobs sheet data
    const defData = defSheet.getDataRange().getValues();
    
    if (defData.length < 2) {
      // No jobs in DefaultJobs
      return { 
        status: 'info', 
        message: 'No jobs found in the DefaultJobs sheet' 
      };
    }
    
    // Get target sheet header
    const targetHeader = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    
    // Get source sheet header (first row)
    const defHeader = defData[0];
    
    // Find column indices in source (DefaultJobs)
    const sourceIndices = {
      jobName: defHeader.indexOf('Job Name'),
      jobLink: defHeader.indexOf('Job Link'),
      jiraTicket: defHeader.indexOf('Jira Ticket'),
      type: defHeader.indexOf('Type'),
      status: defHeader.indexOf('Status'),
      priority: defHeader.indexOf('Priority'),
      notes: defHeader.indexOf('Notes')
    };
    
    // Find column indices in target
    const targetIndices = {
      jobName: targetHeader.indexOf('Job Name'),
      jobLink: targetHeader.indexOf('Job Link'),
      jiraTicket: targetHeader.indexOf('Jira Ticket'),
      type: targetHeader.indexOf('Type'),
      status: targetHeader.indexOf('Status'),
      priority: targetHeader.indexOf('Priority'),
      notes: targetHeader.indexOf('Notes')
    };
    
    // Ensure Job Name column exists in source
    if (sourceIndices.jobName === -1) {
      return {
        status: 'error',
        message: `"Job Name" column not found in ${DEFAULT_JOBS_SHEET_NAME} sheet`
      };
    }
    
    // Ensure Job Name column exists in target
    if (targetIndices.jobName === -1) {
      return {
        status: 'error',
        message: `"Job Name" column not found in ${targetSheetName} sheet`
      };
    }
    
    // Get config for default values
    const config = getConfig();
    
    // Process each job in DefaultJobs
    const rowsToAdd = [];
    let addedCount = 0;
    
    // Start from row 1 (skipping header)
    for (let r = 1; r < defData.length; r++) {
      const row = defData[r];
      if (row.join('').trim() === '') continue; // Skip blank rows
      
      // Get job name
      const jobName = sanitizeInput(row[sourceIndices.jobName]);
      if (!jobName || existingSet.has(jobName.toLowerCase())) {
        continue; // Skip if empty or already exists
      }
      
      // Create a new row
      const newRow = Array(targetHeader.length).fill('');
      
      // Add Job Name
      newRow[targetIndices.jobName] = jobName;
      
      // Add Job Link if both columns exist
      if (sourceIndices.jobLink !== -1 && targetIndices.jobLink !== -1) {
        const jobLink = row[sourceIndices.jobLink];
        if (jobLink) {
          newRow[targetIndices.jobLink] = jobLink;
        }
      }
      
      // Add Jira Ticket if both columns exist
      if (sourceIndices.jiraTicket !== -1 && targetIndices.jiraTicket !== -1) {
        const jiraTicket = row[sourceIndices.jiraTicket];
        if (jiraTicket) {
          newRow[targetIndices.jiraTicket] = jiraTicket;
        }
      }
      
      // Add Type if both columns exist
      if (sourceIndices.type !== -1 && targetIndices.type !== -1) {
        newRow[targetIndices.type] = row[sourceIndices.type] || config.types[0] || 'Build';
      } else if (targetIndices.type !== -1) {
        // Set default type if target has column but source doesn't
        newRow[targetIndices.type] = config.types[0] || 'Build';
      }
      
      // Add Status if both columns exist
      if (sourceIndices.status !== -1 && targetIndices.status !== -1) {
        newRow[targetIndices.status] = row[sourceIndices.status] || config.statuses[0] || 'Pending';
      } else if (targetIndices.status !== -1) {
        // Set default status if target has column but source doesn't
        newRow[targetIndices.status] = config.statuses[0] || 'Pending';
      }
      
      // Add Priority if both columns exist
      if (sourceIndices.priority !== -1 && targetIndices.priority !== -1) {
        newRow[targetIndices.priority] = row[sourceIndices.priority] || config.priorities[1] || 'Medium';
      } else if (targetIndices.priority !== -1) {
        // Set default priority if target has column but source doesn't
        newRow[targetIndices.priority] = config.priorities[1] || 'Medium';
      }
      
      // Add Notes if both columns exist
      if (sourceIndices.notes !== -1 && targetIndices.notes !== -1) {
        newRow[targetIndices.notes] = row[sourceIndices.notes] || '';
      }
      
      // Add the row
      rowsToAdd.push(newRow);
      existingSet.add(jobName.toLowerCase());
      addedCount++;
    }
    
    // Add all rows at once for performance
    if (rowsToAdd.length > 0) {
      targetSheet.getRange(targetSheet.getLastRow() + 1, 1, rowsToAdd.length, targetHeader.length)
        .setValues(rowsToAdd);
    }
    
    // Ensure data validation is applied
    setupDataValidation(targetSheet);
    
    // Return success
    return { 
      status: 'success', 
      message: `Added ${addedCount} jobs from ${DEFAULT_JOBS_SHEET_NAME} to ${targetSheetName}` 
    };
  } catch (error) {
    return handleError('addJobsFromDefaultSheet', error, true);
  }
}