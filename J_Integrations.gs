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
    let baseUrl = settings.jenkinsUrl;
    
    if (!baseUrl) return jobName;
    
    // Normalize URL by removing trailing slash
    if (baseUrl.endsWith('/')) {
      baseUrl = baseUrl.slice(0, -1);
    }
    
    // Add /job/ between base URL and job name
    const jobUrl = `${baseUrl}/job/${encodeURIComponent(jobName)}`;
    
    // Create a hyperlink formula: =HYPERLINK("url", "display text")
    return `=HYPERLINK("${jobUrl}", "${jobName}")`;
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
    let baseUrl = settings.jiraUrl;
    
    if (!baseUrl) return ticketId;
    
    // Normalize URL by removing trailing slash
    if (baseUrl.endsWith('/')) {
      baseUrl = baseUrl.slice(0, -1);
    }
    
    // Create Jira URL
    const ticketUrl = `${baseUrl}/${encodeURIComponent(ticketId)}`;
    
    // Create a hyperlink formula: =HYPERLINK("url", "display text")
    return `=HYPERLINK("${ticketUrl}", "${ticketId}")`;
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
    
    // If sheet already exists, return
    if (sheet) {
      return {
        status: 'info',
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
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    return {
      status: 'success',
      message: 'DefaultJobs template created successfully'
    };
  } catch (error) {
    return handleError('createDefaultJobsTemplate', error, true);
  }
}
