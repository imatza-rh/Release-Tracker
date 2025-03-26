/**
 * D_Config.gs - Enhanced configuration management
 * 
 * This file contains functions for managing application configuration
 * with advanced features like versioning, validation, and backup/restore.
 * 
 * @requires CONFIG_PROPERTY_KEY from A_Constants.gs
 * @requires _configCache from A_Constants.gs
 * @requires log, handleError, safeExecute from B_Logging.gs
 * @requires safeParseJson from C_Utils.gs
 */

/**
 * Default config constants for fallback and initialization
 */
const DEFAULT_CONFIG = {
  statuses: ['Pending', 'In-Progress', 'Done'],
  statusColors: ['#e3f2fd', '#fff8e1', '#e8f5e9'], // Light blue, amber, green
  types: ['Build', 'Test', 'Deploy'],
  typeColors: ['#ede7f6', '#e0f7fa', '#f3e5f5'], // Light purple, cyan, pink
  priorities: ['High', 'Medium', 'Low'],
  priorityColors: ['#ffebee', '#fff8e1', '#e8f5e9'], // Light red, amber, green
  fields: ['Job Name', 'Type', 'Status', 'Priority', 'Notes', 'Job Link', 'Jira Ticket'],
  metadata: {
    version: 1,
    created: new Date().toISOString(),
    lastModified: new Date().toISOString()
  }
};

/**
 * Reads user config with caching and version tracking
 * 
 * @param {boolean} [bypassCache=false] - Whether to bypass the cache
 * @returns {Object} Configuration object
 */
function getConfig(bypassCache = false) {
  return safeExecute('getConfig', () => {
    // Return cached config if available and not bypassing
    if (_configCache && !bypassCache) {
      return _configCache;
    }
    
    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty(CONFIG_PROPERTY_KEY);
    
    if (stored) {
      const parsed = safeParseJson(stored, DEFAULT_CONFIG);
      
      // Ensure necessary properties exist with default fallbacks
      const config = ensureCompleteConfig(parsed);
      
      _configCache = config;
      return _configCache;
    }
    
    // Default if not yet saved - includes improved default colors
    _configCache = DEFAULT_CONFIG;
    
    // Save default config to script properties
    props.setProperty(CONFIG_PROPERTY_KEY, JSON.stringify(DEFAULT_CONFIG));
    
    return _configCache;
  }, false) || DEFAULT_CONFIG; // Always return at least the default config
}

/**
 * Makes sure the configuration has all required properties with fallbacks to defaults
 * 
 * @param {Object} config - Configuration object to validate
 * @returns {Object} Complete configuration with all required properties
 */
function ensureCompleteConfig(config) {
  // Add metadata if missing
  if (!config.metadata) {
    config.metadata = DEFAULT_CONFIG.metadata;
  }
  
  // Ensure required fields exist
  const requiredFields = ['statuses', 'types', 'priorities', 'fields'];
  requiredFields.forEach(field => {
    if (!config[field] || !Array.isArray(config[field]) || config[field].length === 0) {
      config[field] = DEFAULT_CONFIG[field];
    }
  });
  
  // Ensure status colors exist (important for ViewSidebar)
  if (config.statuses && Array.isArray(config.statuses)) {
    if (!config.statusColors || !Array.isArray(config.statusColors) || 
        config.statusColors.length !== config.statuses.length) {
      // Generate default colors for statuses
      config.statusColors = generateDefaultColors(config.statuses, 'status');
      log('Generated default status colors', LOG_LEVELS.INFO);
    }
  }
  
  // Ensure type colors exist
  if (config.types && Array.isArray(config.types)) {
    if (!config.typeColors || !Array.isArray(config.typeColors) ||
        config.typeColors.length !== config.types.length) {
      // Generate default colors for types
      config.typeColors = generateDefaultColors(config.types, 'type');
      log('Generated default type colors', LOG_LEVELS.INFO);
    }
  }
  
  // Ensure priority colors exist
  if (config.priorities && Array.isArray(config.priorities)) {
    if (!config.priorityColors || !Array.isArray(config.priorityColors) ||
        config.priorityColors.length !== config.priorities.length) {
      // Generate default colors for priorities
      config.priorityColors = generateDefaultColors(config.priorities, 'priority');
      log('Generated default priority colors', LOG_LEVELS.INFO);
    }
  }
  
  return config;
}

/**
 * Generates default colors for categories with improved color schemes
 * 
 * @param {Array} items - Array of items to generate colors for
 * @param {string} category - Category type ('status', 'type', or 'priority')
 * @returns {Array} Array of color hex codes
 */
function generateDefaultColors(items, category) {
  return safeExecute('generateDefaultColors', () => {
    if (!items || !Array.isArray(items) || items.length === 0) {
      return [];
    }
    
    // Modern color palettes with improved visual aesthetics
    const palettes = {
      status: {
        // Status colors - more subtle and professional
        'pending': '#e3f2fd',     // Light blue
        'in-progress': '#fff8e1', // Light amber
        'inprogress': '#fff8e1',  // Light amber
        'in progress': '#fff8e1', // Light amber
        'done': '#e8f5e9',        // Light green
        'completed': '#e8f5e9',   // Light green
        'blocked': '#ffebee',     // Light red
        'canceled': '#f5f5f5',    // Light gray
        'cancelled': '#f5f5f5',   // Light gray
        'postponed': '#f3e5f5',   // Light purple
        'default': '#e3f2fd'      // Default light blue
      },
      type: {
        // Type colors - visually distinct but cohesive
        'build': '#ede7f6',      // Light deep purple
        'test': '#e0f7fa',       // Light cyan
        'deploy': '#f3e5f5',     // Light purple
        'development': '#ede7f6', // Light deep purple
        'design': '#f3e5f5',     // Light purple
        'documentation': '#e1f5fe', // Light blue
        'planning': '#e8f5e9',   // Light green
        'review': '#fff3e0',     // Light orange
        'research': '#e0f2f1',   // Light teal
        'default': '#ede7f6'     // Default light purple
      },
      priority: {
        // Priority colors - intuitive color coding
        'high': '#ffebee',       // Light red
        'critical': '#ffcdd2',   // Lighter red
        'urgent': '#ffcdd2',     // Lighter red
        'medium': '#fff8e1',     // Light amber
        'normal': '#fff8e1',     // Light amber
        'low': '#e8f5e9',        // Light green
        'trivial': '#f5f5f5',    // Light gray
        'default': '#fff8e1'     // Default light amber
      }
    };
    
    // Get the appropriate palette
    const palette = palettes[category] || palettes.status;
    
    // Generate colors for each item
    return items.map(item => {
      const key = String(item).toLowerCase();
      return palette[key] || palette.default;
    });
  }, false) || [];
}

/**
 * Persists config and updates the cache with versioning
 * 
 * @param {Object} updatedConfig - Configuration object to save
 * @returns {Object} Result object indicating success or failure
 */
function saveConfig(updatedConfig) {
  return safeExecute('saveConfig', () => {
    // Get current config to preserve metadata
    const currentConfig = getConfig(true);
    
    // Update metadata
    updatedConfig.metadata = {
      version: (currentConfig.metadata?.version || 0) + 1,
      created: currentConfig.metadata?.created || new Date().toISOString(),
      lastModified: new Date().toISOString()
    };
    
    // Ensure all required properties with defaults
    const completeConfig = ensureCompleteConfig(updatedConfig);
    
    // Store in script properties
    const props = PropertiesService.getScriptProperties();
    props.setProperty(CONFIG_PROPERTY_KEY, JSON.stringify(completeConfig));
    
    // Update cache
    _configCache = completeConfig;
    
    log('Configuration saved successfully', LOG_LEVELS.INFO, completeConfig);
    return { 
      status: 'success',
      message: 'Settings saved successfully (version ' + completeConfig.metadata.version + ')'
    };
  }, false);
}

/**
 * Creates a backup of the current configuration
 * 
 * @param {Object} config - Current configuration to backup
 */
function createBackup(config) {
  return safeExecute('createBackup', () => {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupKey = `${CONFIG_PROPERTY_KEY}_BACKUP_${timestamp}`;
    
    // Store backup in script properties
    const props = PropertiesService.getScriptProperties();
    props.setProperty(backupKey, JSON.stringify(config));
    
    // Keep only the last 5 backups to avoid hitting storage limits
    cleanupOldBackups(5);
    
    log('Configuration backup created', LOG_LEVELS.INFO, { backupKey });
    return true;
  }, false);
}

/**
 * Cleans up old backups to prevent hitting storage limits
 * 
 * @param {number} keepCount - Number of most recent backups to keep
 */
function cleanupOldBackups(keepCount = 5) {
  return safeExecute('cleanupOldBackups', () => {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    
    // Find all backups
    const backupKeys = Object.keys(allProps)
      .filter(key => key.startsWith(`${CONFIG_PROPERTY_KEY}_BACKUP_`))
      .sort()  // Sort chronologically (oldest first)
      .reverse(); // Reverse to have newest first
    
    // If we have more than keepCount, delete the oldest ones
    if (backupKeys.length > keepCount) {
      const keysToDelete = backupKeys.slice(keepCount);
      
      for (const key of keysToDelete) {
        props.deleteProperty(key);
        log('Deleted old config backup', LOG_LEVELS.DEBUG, { key });
      }
    }
    
    return true;
  }, false);
}

/**
 * Clears the config cache for testing
 * 
 * @returns {Object} Result object indicating success
 */
function clearConfigCache() {
  _configCache = null;
  return { 
    status: 'success', 
    message: 'Config cache cleared' 
  };
}