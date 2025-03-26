/**
 * C_Utils.gs - Utility functions
 * 
 * This file contains general utility functions used throughout
 * the Release Tracker application.
 */

/**
 * Input sanitization helper
 * 
 * @param {*} value - The value to sanitize
 * @param {string} [type='string'] - The type to convert to
 * @returns {string|number} Sanitized value
 */
function sanitizeInput(value, type = 'string') {
  if (value === null || value === undefined) {
    return '';
  }
  
  if (type === 'string') {
    return String(value).trim();
  } else if (type === 'number') {
    const num = Number(value);
    return isNaN(num) ? 0 : num;
  }
  
  return value;
}

/**
 * Normalizes and formats a base URL
 * 
 * @param {string} url - URL to normalize
 * @param {string} [pathSuffix] - Optional path suffix to ensure
 * @returns {string} Normalized URL
 */
function normalizeBaseUrl(url, pathSuffix) {
  if (!url) return '';
  
  // Clean and normalize the URL
  url = url.trim();
  if (!url.startsWith('http://') && !url.startsWith('https://')) {
    url = 'https://' + url;
  }
  
  // Clean trailing slashes
  while (url.endsWith('/')) {
    url = url.slice(0, -1);
  }
  
  // Add path suffix if needed
  if (pathSuffix) {
    // Clean up the suffix
    const suffix = pathSuffix.replace(/^\/|\/$/g, '');
    
    // Check different cases
    if (url.toLowerCase().endsWith('/' + suffix.toLowerCase())) {
      // URL ends with "/suffix" but no trailing slash
      url += '/';
    } else if (!url.toLowerCase().includes('/' + suffix.toLowerCase() + '/')) {
      // URL doesn't contain "/suffix/" anywhere
      url += '/' + suffix + '/';
    }
  }
  
  return url;
}

/**
 * Creates a direct hyperlink for use in Google Sheets
 * 
 * @param {string} url - URL to link to
 * @param {string} displayText - Text to display for the link
 * @returns {string} Google Sheets HYPERLINK formula or original text if URL is empty
 */
function createHyperlink(url, displayText) {
  if (!url || !displayText) return displayText || '';
  
  // Escape double quotes (they need to be doubled in Sheets formulas)
  const escapedUrl = String(url).replace(/"/g, '""');
  const escapedText = String(displayText).replace(/"/g, '""');
  
  // Create the formula
  return `=HYPERLINK("${escapedUrl}","${escapedText}")`;
}

/**
 * Deep clones an object to avoid reference issues
 * 
 * @param {Object} obj - Object to clone
 * @returns {Object} Deep copy of the object
 */
function deepClone(obj) {
  if (obj === null || typeof obj !== 'object') {
    return obj;
  }
  
  // Handle Date
  if (obj instanceof Date) {
    return new Date(obj.getTime());
  }
  
  // Handle Array
  if (Array.isArray(obj)) {
    return obj.map(item => deepClone(item));
  }
  
  // Handle Object
  const result = {};
  for (const key in obj) {
    if (obj.hasOwnProperty(key)) {
      result[key] = deepClone(obj[key]);
    }
  }
  
  return result;
}

/**
 * Checks if an object is empty (has no properties)
 * 
 * @param {Object} obj - Object to check
 * @returns {boolean} True if the object is empty
 */
function isEmptyObject(obj) {
  if (obj === null || typeof obj !== 'object') {
    return true;
  }
  return Object.keys(obj).length === 0;
}

/**
 * Generates a unique ID for use in various elements
 * 
 * @returns {string} Unique ID
 */
function generateUniqueId() {
  return 'id_' + Math.random().toString(36).substring(2, 11);
}

/**
 * Safely parses JSON with error handling
 * 
 * @param {string} jsonString - JSON string to parse
 * @param {*} defaultValue - Default value to return on error
 * @returns {*} Parsed object or default value on error
 */
function safeParseJson(jsonString, defaultValue = {}) {
  try {
    return JSON.parse(jsonString);
  } catch (e) {
    log(`JSON parse error: ${e}`, LOG_LEVELS.WARN);
    return defaultValue;
  }
}