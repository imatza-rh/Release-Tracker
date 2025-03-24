/**
 * A_Constants.gs - Global constants and shared variables
 * 
 * This file contains all constants and shared state variables
 * used throughout the Release Tracker application.
 */

/**
 * Sheet names used in the application
 * @constant {string}
 */
const JOBS_SHEET_NAME = 'Jobs';

/**
 * Sheet name for default jobs
 * @constant {string}
 */
const DEFAULT_JOBS_SHEET_NAME = 'DefaultJobs';

/**
 * Property key for storing configuration
 * @constant {string}
 */
const CONFIG_PROPERTY_KEY = 'RELEASE_TRACKER_CONFIG';

/**
 * Log level constants
 * @constant {Object}
 */
const LOG_LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3
};

/**
 * Global cache for configuration to avoid repeated property reads
 * @type {Object|null}
 */
let _configCache = null;

/**
 * Current log level setting
 * @type {number}
 */
let currentLogLevel = LOG_LEVELS.INFO;

