/**
 * B_Logging.gs - Logging and error handling
 * 
 * This file contains logging utilities and standardized error handling
 * for the Release Tracker application.
 * 
 * @requires LOG_LEVELS from A_Constants.gs
 * @requires currentLogLevel from A_Constants.gs
 */

/**
 * Enhanced logging with levels
 * 
 * @param {string} message - The message to log
 * @param {number} [level=LOG_LEVELS.INFO] - The log level
 * @param {Object} [data=null] - Optional data to include in the log
 */
function log(message, level = LOG_LEVELS.INFO, data = null) {
  if (level >= currentLogLevel) {
    const levelStr = Object.keys(LOG_LEVELS).find(key => LOG_LEVELS[key] === level);
    const logMsg = `[${levelStr}] ${message}`;
    
    if (data) {
      console.log(`${logMsg}: ${JSON.stringify(data)}`);
      Logger.log(`${logMsg}: ${JSON.stringify(data)}`);
    } else {
      console.log(logMsg);
      Logger.log(logMsg);
    }
  }
}

/**
 * Standardized error handling
 * 
 * @param {string} functionName - Name of function where error occurred
 * @param {Error|string} error - The error object or message
 * @param {boolean} [showUi=true] - Whether to show UI alert
 * @returns {Object} Error result object
 */
function handleError(functionName, error, showUi = true) {
  const errorMsg = `Error in ${functionName}: ${error.message || error}`;
  log(errorMsg, LOG_LEVELS.ERROR);
  
  if (showUi) {
    try {
      SpreadsheetApp.getUi().alert(errorMsg);
    } catch (e) {
      log(`Couldn't show UI alert: ${e}`, LOG_LEVELS.WARN);
    }
  }
  
  return { status: 'error', message: errorMsg };
}

/**
 * Sets current logging level
 * 
 * @param {string} level - Level name from LOG_LEVELS
 * @returns {Object} Result object
 */
function setLogLevel(level) {
  if (level in LOG_LEVELS) {
    currentLogLevel = LOG_LEVELS[level];
    return { 
      status: 'success', 
      message: `Log level set to ${level}` 
    };
  }
  return { 
    status: 'error', 
    message: `Invalid log level: ${level}` 
  };
}
