/**
 * J_Integrations.gs - Integration settings and functionality
 * 
 * This file contains functions for integrating with external systems
 * like Jenkins and Jira.
 * 
 * @requires log, handleError, safeExecute from B_Logging.gs
 * @requires normalizeBaseUrl, createHyperlink from C_Utils.gs
 */

// Property key for storing integration settings
const INTEGRATION_SETTINGS_KEY = 'RELEASE_TRACKER_INTEGRATIONS';

/**
 * Retrieves integration settings
 * 
 * @returns {Object} Integration settings
 */
function getIntegrationSettings() {
  return safeExecute('getIntegrationSettings', () => {
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
  }, false);
}

/**
 * Saves integration settings
 * 
 * @param {Object} settings - Integration settings to save
 * @returns {Object} Result indicating success or failure
 */
function saveIntegrationSettings(settings) {
  return safeExecute('saveIntegrationSettings', () => {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(INTEGRATION_SETTINGS_KEY, JSON.stringify(settings));
    
    log('Integration settings saved', LOG_LEVELS.INFO, settings);
    return { 
      status: 'success', 
      message: 'Integration settings saved successfully' 
    };
  }, false);
}

/**
 * Creates a hyperlink for a Jenkins job
 * 
 * @param {string} jobName - Name of the Jenkins job
 * @returns {string} Hyperlink formula or original job name if no URL is configured
 */
function createJenkinsJobLink(jobName) {
  return safeExecute('createJenkinsJobLink', () => {
    if (!jobName) return '';
    
    const settings = getIntegrationSettings();
    let baseUrl = settings.jenkinsUrl || '';
    
    if (!baseUrl) return jobName;
    
    // Use the utility function to normalize the URL
    baseUrl = normalizeBaseUrl(baseUrl, '/job/');
    
    // Create the final URL
    const finalUrl = baseUrl + encodeURIComponent(jobName);
    
    // Create a direct hyperlink formula
    return createHyperlink(finalUrl, jobName);
  }, false) || jobName;
}

/**
 * Creates a hyperlink for a Jira ticket
 * 
 * @param {string} ticketId - Jira ticket ID
 * @returns {string} Hyperlink formula or original ticket ID if no URL is configured
 */
function createJiraTicketLink(ticketId) {
  return safeExecute('createJiraTicketLink', () => {
    if (!ticketId) return '';
    
    const settings = getIntegrationSettings();
    let baseUrl = settings.jiraUrl || '';
    
    if (!baseUrl) return ticketId;
    
    // Use the utility function to normalize the URL
    baseUrl = normalizeBaseUrl(baseUrl, '/browse/');
    
    // Create the final URL
    const finalUrl = baseUrl + encodeURIComponent(ticketId);
    
    // Create a direct hyperlink formula
    return createHyperlink(finalUrl, ticketId);
  }, false) || ticketId;
}

/**
 * Checks if the DefaultJobs sheet exists
 * 
 * @returns {boolean} True if the DefaultJobs sheet exists, false otherwise
 */
function checkDefaultJobsSheetExists() {
  return safeExecute('checkDefaultJobsSheetExists', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(DEFAULT_JOBS_SHEET_NAME);
    return sheet !== null;
  }, false) || false;
}

/**
 * Creates a DefaultJobs template sheet with example jobs
 * 
 * @returns {Object} Result indicating success or failure
 */
function createDefaultJobsTemplate() {
  return safeExecute('createDefaultJobsTemplate', () => {
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
  }, false);
}