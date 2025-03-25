/**
 * J_Integrations.gs - Integration settings and functionality
 * 
 * This file contains functions for integrating with external systems
 * like Jenkins and Jira.
 * 
 * @requires log, handleError from B_Logging.gs
 */

// Property key for storing integration settings
const INTEGRATION_SETTINGS_KEY = 'RELEASE_TRACKER_INTEGRATIONS';

/**
 * Retrieves integration settings
 * 
 * @returns {Object} Integration settings
 */
function getIntegrationSettings() {
  try {
    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty(INTEGRATION_SETTINGS_KEY);
    
    if (stored) {
      return JSON.parse(stored);
    }
    
    // Default settings
    return {
      jenkinsUrl: '',
      jiraUrl: ''
    };
  } catch (error) {
    handleError('getIntegrationSettings', error, false);
    return { jenkinsUrl: '', jiraUrl: '' };
  }
}

/**
 * Saves integration settings
 * 
 * @param {Object} settings - Integration settings to save
 * @returns {Object} Result indicating success or failure
 */
function saveIntegrationSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(INTEGRATION_SETTINGS_KEY, JSON.stringify(settings));
    
    log('Integration settings saved', LOG_LEVELS.INFO, settings);
    return { 
      status: 'success', 
      message: 'Integration settings saved successfully' 
    };
  } catch (error) {
    return handleError('saveIntegrationSettings', error, false);
  }
}

/**
 * Creates a hyperlink for a Jenkins job
 * 
 * @param {string} jobName - Name of the Jenkins job
 * @returns {string} Hyperlink formula or original job name if no URL is configured
 */
function createJenkinsJobLink(jobName) {
  try {
    if (!jobName) return '';
    
    const settings = getIntegrationSettings();
    let baseUrl = settings.jenkinsUrl || '';
    
    if (!baseUrl) return jobName;
    
    // Clean and normalize the base URL first
    baseUrl = baseUrl.trim();
    if (!baseUrl.startsWith('http://') && !baseUrl.startsWith('https://')) {
      baseUrl = 'https://' + baseUrl;
    }
    
    // Clean trailing slashes
    while (baseUrl.endsWith('/')) {
      baseUrl = baseUrl.slice(0, -1);
    }
    
    // Determine the correct path to add
    let finalUrl = baseUrl;
    const lowerUrl = baseUrl.toLowerCase();
    
    // Check if URL already contains "/job" pattern
    if (lowerUrl.includes('/job/') || lowerUrl.endsWith('/job')) {
      // If URL ends with /job, add another slash
      if (lowerUrl.endsWith('/job')) {
        finalUrl += '/';
      }
    } else {
      // Otherwise add /job/ path
      finalUrl += '/job/';
    }
    
    // Add encoded job name
    finalUrl += encodeURIComponent(jobName);
    
    // DEBUG log the final URL to ensure it's correct
    log(`Jenkins URL generated: ${finalUrl}`, LOG_LEVELS.DEBUG);
    
    // Create a hyperlink formula with proper quote escaping for the URL
    // Use double quotes for the formula and escaped double quotes for attributes
    const escapedUrl = finalUrl.replace(/"/g, '""');
    const escapedJobName = jobName.replace(/"/g, '""');
    
    return `=HYPERLINK("${escapedUrl}","${escapedJobName}")`;
  } catch (error) {
    log('Error creating Jenkins link', LOG_LEVELS.ERROR, error);
    return jobName;
  }
}

/**
 * Creates a hyperlink for a Jira ticket
 * 
 * @param {string} ticketId - Jira ticket ID
 * @returns {string} Hyperlink formula or original ticket ID if no URL is configured
 */
function createJiraTicketLink(ticketId) {
  try {
    if (!ticketId) return '';
    
    const settings = getIntegrationSettings();
    let baseUrl = settings.jiraUrl || '';
    
    if (!baseUrl) return ticketId;
    
    // Clean and normalize the base URL first
    baseUrl = baseUrl.trim();
    if (!baseUrl.startsWith('http://') && !baseUrl.startsWith('https://')) {
      baseUrl = 'https://' + baseUrl;
    }
    
    // Clean trailing slashes
    while (baseUrl.endsWith('/')) {
      baseUrl = baseUrl.slice(0, -1);
    }
    
    // Determine the correct path to add
    let finalUrl = baseUrl;
    const lowerUrl = baseUrl.toLowerCase();
    
    // Check if URL already contains "/browse" pattern
    if (lowerUrl.includes('/browse/') || lowerUrl.endsWith('/browse')) {
      // If URL ends with /browse, add another slash
      if (lowerUrl.endsWith('/browse')) {
        finalUrl += '/';
      }
    } else {
      // Otherwise add /browse/ path
      finalUrl += '/browse/';
    }
    
    // Add encoded ticket ID
    finalUrl += encodeURIComponent(ticketId);
    
    // DEBUG log the final URL to ensure it's correct
    log(`Jira URL generated: ${finalUrl}`, LOG_LEVELS.DEBUG);
    
    // Create a hyperlink formula with proper quote escaping for the URL
    // Use double quotes for the formula and escaped double quotes for attributes
    const escapedUrl = finalUrl.replace(/"/g, '""');
    const escapedTicketId = ticketId.replace(/"/g, '""');
    
    return `=HYPERLINK("${escapedUrl}","${escapedTicketId}")`;
  } catch (error) {
    log('Error creating Jira link', LOG_LEVELS.ERROR, error);
    return ticketId;
  }
}

/**
 * Checks if the DefaultJobs sheet exists
 * 
 * @returns {boolean} True if the DefaultJobs sheet exists, false otherwise
 */
function checkDefaultJobsSheetExists() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(DEFAULT_JOBS_SHEET_NAME);
    return sheet !== null;
  } catch (error) {
    log('Error checking DefaultJobs sheet', LOG_LEVELS.ERROR, error);
    return false;
  }
}

/**
 * Creates a DefaultJobs template sheet with example jobs
 * 
 * @returns {Object} Result indicating success or failure
 */
function createDefaultJobsTemplate() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DEFAULT_JOBS_SHEET_NAME);
    
    // If sheet already exists, return success
    if (sheet) {
      return {
        status: 'success',
        message: 'DefaultJobs sheet already exists'
      };
    }
    
    // Create new sheet
    sheet = ss.insertSheet(DEFAULT_JOBS_SHEET_NAME);
    
    // Add headers
    const headers = ['Job Name', 'Job Link', 'Jira Ticket', 'Type', 'Status', 'Priority', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Add example jobs
    const exampleJobs = [
      ['Update database', 'update-db', 'PROJ-123', 'Build', 'Pending', 'High', 'Monthly database update'],
      ['Deploy application', 'deploy-app', 'PROJ-124', 'Deploy', 'Pending', 'Medium', 'Production deployment'],
      ['Run tests', 'run-tests', 'PROJ-125', 'Test', 'Pending', 'Medium', 'Integration tests'],
      ['Review code changes', '', 'PROJ-126', 'Build', 'Pending', 'Low', 'Code review for recent changes']
    ];
    
    sheet.getRange(2, 1, exampleJobs.length, headers.length).setValues(exampleJobs);
    
    // Format as table (alternating row colors)
    const dataRange = sheet.getRange(2, 1, exampleJobs.length, headers.length);
    for (let i = 2; i <= exampleJobs.length + 1; i++) {
      const rowColor = i % 2 === 0 ? "#f9f9f9" : "#ffffff";
      sheet.getRange(i, 1, 1, headers.length).setBackground(rowColor);
    }
    
    // Add borders
    dataRange.setBorder(true, true, true, true, true, true, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    // Log success
    log('Created DefaultJobs template sheet successfully', LOG_LEVELS.INFO);
    
    return {
      status: 'success',
      message: 'DefaultJobs template created successfully'
    };
  } catch (error) {
    // Improved error logging
    log('Error creating DefaultJobs template', LOG_LEVELS.ERROR, error);
    
    // Return detailed error information
    return {
      status: 'error',
      message: 'Failed to create DefaultJobs sheet: ' + (error.message || String(error))
    };
  }
}