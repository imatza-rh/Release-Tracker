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
