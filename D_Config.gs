/**
 * D_Config.gs - Enhanced configuration management
 * 
 * This file contains functions for managing application configuration
 * with advanced features like versioning, validation, and backup/restore.
 * 
 * @requires CONFIG_PROPERTY_KEY from A_Constants.gs
 * @requires _configCache from A_Constants.gs
 * @requires log, handleError from B_Logging.gs
 */

/**
 * Reads user config with caching and version tracking
 * 
 * @param {boolean} [bypassCache=false] - Whether to bypass the cache
 * @returns {Object} Configuration object
 */
function getConfig(bypassCache = false) {
  // Return cached config if available and not bypassing
  if (_configCache && !bypassCache) {
    return _configCache;
  }
  
  try {
    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty(CONFIG_PROPERTY_KEY);
    
    if (stored) {
      const parsed = JSON.parse(stored);
      
      // Add metadata if missing
      if (!parsed.metadata) {
        parsed.metadata = {
          version: 1,
          lastModified: new Date().toISOString(),
          created: new Date().toISOString()
        };
      }
      
      // Ensure status colors exist (important for ViewSidebar)
      if (parsed.statuses && Array.isArray(parsed.statuses)) {
        if (!parsed.statusColors || !Array.isArray(parsed.statusColors) || 
            parsed.statusColors.length !== parsed.statuses.length) {
          // Generate default colors for statuses
          parsed.statusColors = generateDefaultColors(parsed.statuses, 'status');
          log('Generated default status colors', LOG_LEVELS.INFO);
        }
      }
      
      // Ensure type colors exist
      if (parsed.types && Array.isArray(parsed.types)) {
        if (!parsed.typeColors || !Array.isArray(parsed.typeColors) ||
            parsed.typeColors.length !== parsed.types.length) {
          // Generate default colors for types
          parsed.typeColors = generateDefaultColors(parsed.types, 'type');
          log('Generated default type colors', LOG_LEVELS.INFO);
        }
      }
      
      // Ensure priority colors exist
      if (parsed.priorities && Array.isArray(parsed.priorities)) {
        if (!parsed.priorityColors || !Array.isArray(parsed.priorityColors) ||
            parsed.priorityColors.length !== parsed.priorities.length) {
          // Generate default colors for priorities
          parsed.priorityColors = generateDefaultColors(parsed.priorities, 'priority');
          log('Generated default priority colors', LOG_LEVELS.INFO);
        }
      }
      
      _configCache = parsed;
      return _configCache;
    }
    
    // Default if not yet saved - includes default colors
    const defaultConfig = {
      statuses:    ['Pending', 'In-Progress', 'Done'],
      statusColors: ['#1976d2', '#ff9800', '#4caf50'], // Blue, Orange, Green
      types:       ['Build', 'Test', 'Deploy'],
      typeColors:  ['#e69138', '#674ea7', '#3d85c6'], // Orange, Purple, Blue
      priorities:  ['High', 'Medium', 'Low'],
      priorityColors: ['#e06666', '#f1c232', '#6aa84f'], // Red, Yellow, Green
      fields:      ['Job Name', 'Type', 'Status', 'Priority', 'Notes', 'Job Link', 'Jira Ticket'],
      metadata: {
        version: 1,
        created: new Date().toISOString(),
        lastModified: new Date().toISOString()
      }
    };
    
    _configCache = defaultConfig;
    
    // Save default config to script properties
    props.setProperty(CONFIG_PROPERTY_KEY, JSON.stringify(defaultConfig));
    
    return _configCache;
  } catch (error) {
    handleError('getConfig', error, false);
    
    // Return default config if error, including default colors
    return {
      statuses:    ['Pending', 'In-Progress', 'Done'],
      statusColors: ['#1976d2', '#ff9800', '#4caf50'], // Blue, Orange, Green
      types:       ['Build', 'Test', 'Deploy'],
      typeColors:  ['#e69138', '#674ea7', '#3d85c6'], // Orange, Purple, Blue
      priorities:  ['High', 'Medium', 'Low'],
      priorityColors: ['#e06666', '#f1c232', '#6aa84f'], // Red, Yellow, Green
      fields:      ['Job Name', 'Type', 'Status', 'Priority', 'Notes', 'Job Link', 'Jira Ticket'],
      metadata: {
        version: 1,
        created: new Date().toISOString(),
        lastModified: new Date().toISOString()
      }
    };
  }
}

/**
 * Generates default colors for categories
 * 
 * @param {Array} items - Array of items to generate colors for
 * @param {string} category - Category type ('status', 'type', or 'priority')
 * @returns {Array} Array of color hex codes
 */
function generateDefaultColors(items, category) {
  if (!items || !Array.isArray(items) || items.length === 0) {
    return [];
  }
  
  // Default color palettes
  const palettes = {
    status: {
      // Common status colors
      'pending': '#1976d2',    // Blue
      'in-progress': '#ff9800', // Orange
      'in progress': '#ff9800', // Orange
      'inprogress': '#ff9800',  // Orange
      'done': '#4caf50',       // Green
      'completed': '#4caf50',  // Green
      'blocked': '#f44336',    // Red
      'canceled': '#9e9e9e',   // Gray
      'cancelled': '#9e9e9e',  // Gray
      'postponed': '#9c27b0',  // Purple
      'default': '#1976d2'     // Default blue
    },
    type: {
      // Common type colors
      'build': '#e69138',      // Orange
      'test': '#674ea7',       // Purple
      'deploy': '#3d85c6',     // Blue
      'development': '#e69138', // Orange
      'design': '#9c27b0',     // Purple
      'documentation': '#607d8b', // Blue Gray
      'planning': '#4caf50',   // Green
      'review': '#ff5722',     // Deep Orange
      'research': '#009688',   // Teal
      'default': '#e69138'     // Default orange
    },
    priority: {
      // Common priority colors
      'high': '#e06666',      // Red
      'critical': '#dc3545',  // Darker Red
      'urgent': '#dc3545',    // Darker Red
      'medium': '#f1c232',    // Yellow
      'normal': '#f1c232',    // Yellow
      'low': '#6aa84f',       // Green
      'trivial': '#9e9e9e',   // Gray
      'default': '#f1c232'    // Default yellow
    }
  };
  
  // Get the appropriate palette
  const palette = palettes[category] || palettes.status;
  
  // Generate colors for each item
  return items.map(item => {
    const key = String(item).toLowerCase();
    return palette[key] || palette.default;
  });
}

/**
 * Validates configuration data before saving
 * 
 * @param {Array} configArray - Array of configuration objects
 * @returns {Object} Validation result {valid: boolean, message: string}
 */
function validateConfig(configArray) {
  if (!Array.isArray(configArray)) {
    return { valid: false, message: "Configuration must be an array" };
  }
  
  // Check for required properties
  const requiredProperties = ["Status", "Job Type", "Priority"];
  const foundProperties = configArray.map(prop => prop.name);
  
  for (const required of requiredProperties) {
    if (!foundProperties.includes(required)) {
      return { valid: false, message: `Required property "${required}" is missing` };
    }
  }
  
  // Validate each property
  for (const prop of configArray) {
    // Check for empty name
    if (!prop.name || prop.name.trim() === "") {
      return { valid: false, message: "Property names cannot be empty" };
    }
    
    // For dropdown types, ensure values exist
    if (prop.type === "dropdown") {
      if (!Array.isArray(prop.values) || prop.values.length === 0) {
        return { valid: false, message: `Dropdown "${prop.name}" must have at least one value` };
      }
      
      // Check for empty values
      const emptyValues = prop.values.filter(v => v === "");
      if (emptyValues.length > 0) {
        return { valid: false, message: `Dropdown "${prop.name}" contains empty values` };
      }
    }
  }
  
  return { valid: true, message: "Configuration is valid" };
}

/**
   * Persists config and updates the cache with versioning
   * 
   * @param {Array} configArray - Array of configuration objects from Settings.html
   * @returns {Object} Result object indicating success or failure
   */
  function saveConfig(configArray) {
    try {
      // Log the incoming data
      log("Saving configuration", LOG_LEVELS.DEBUG, configArray);
      
      // Validate configuration
      const validation = validateConfig(configArray);
      if (!validation.valid) {
        return { 
          status: 'error', 
          message: validation.message 
        };
      }
      
      // Get current config to preserve fields and metadata
      const currentConfig = getConfig(true);
      
      // Extract dropdown values to save in our config structure
      const updatedConfig = {
        statuses: [],
        types: [],
        priorities: [],
        statusColors: [],
        typeColors: [],
        priorityColors: [],
        fields: currentConfig.fields || ['Job Name', 'Type', 'Status', 'Priority', 'Notes', 'Job Link', 'Jira Ticket'],
        metadata: {
          version: (currentConfig.metadata?.version || 0) + 1,
          created: currentConfig.metadata?.created || new Date().toISOString(),
          lastModified: new Date().toISOString()
        }
      };
      
      // Process each property from the settings dialog
      for (const prop of configArray) {
        if (prop.name === "Status" && Array.isArray(prop.values)) {
          updatedConfig.statuses = prop.values.filter(v => v !== "");
          // Store colors if available
          if (Array.isArray(prop.colors)) {
            updatedConfig.statusColors = prop.colors.slice(0, updatedConfig.statuses.length);
          }
        } else if (prop.name === "Job Type" && Array.isArray(prop.values)) {
          updatedConfig.types = prop.values.filter(v => v !== "");
          // Store colors if available
          if (Array.isArray(prop.colors)) {
            updatedConfig.typeColors = prop.colors.slice(0, updatedConfig.types.length);
          }
        } else if (prop.name === "Priority" && Array.isArray(prop.values)) {
          updatedConfig.priorities = prop.values.filter(v => v !== "");
          // Store colors if available
          if (Array.isArray(prop.colors)) {
            updatedConfig.priorityColors = prop.colors.slice(0, updatedConfig.priorities.length);
          }
        }
      }
      
      // Ensure colors are defined for all options
      if (updatedConfig.statuses && updatedConfig.statuses.length > 0) {
        if (!updatedConfig.statusColors || updatedConfig.statusColors.length !== updatedConfig.statuses.length) {
          updatedConfig.statusColors = generateDefaultColors(updatedConfig.statuses, 'status');
        }
      }
      
      if (updatedConfig.types && updatedConfig.types.length > 0) {
        if (!updatedConfig.typeColors || updatedConfig.typeColors.length !== updatedConfig.types.length) {
          updatedConfig.typeColors = generateDefaultColors(updatedConfig.types, 'type');
        }
      }
      
      if (updatedConfig.priorities && updatedConfig.priorities.length > 0) {
        if (!updatedConfig.priorityColors || updatedConfig.priorityColors.length !== updatedConfig.priorities.length) {
          updatedConfig.priorityColors = generateDefaultColors(updatedConfig.priorities, 'priority');
        }
      }
      
      // Create backup before saving
      createBackup(currentConfig);
      
      // Store in script properties
      const props = PropertiesService.getScriptProperties();
      props.setProperty(CONFIG_PROPERTY_KEY, JSON.stringify(updatedConfig));
      
      // Update cache
      _configCache = updatedConfig;
      
      log('Configuration saved successfully', LOG_LEVELS.INFO, updatedConfig);
      return { 
        status: 'success',
        message: 'Settings saved successfully (version ' + updatedConfig.metadata.version + ')'
      };
    } catch (error) {
      return handleError('saveConfig', error, false);
    }
  }

/**
 * Creates a backup of the current configuration
 * 
 * @param {Object} config - Current configuration to backup
 */
function createBackup(config) {
  try {
    const timestamp = new Date().toISOString().replace(/:/g, '-');
    const backupKey = `${CONFIG_PROPERTY_KEY}_BACKUP_${timestamp}`;
    
    // Store backup in script properties
    const props = PropertiesService.getScriptProperties();
    props.setProperty(backupKey, JSON.stringify(config));
    
    // Keep only the last 5 backups to avoid hitting storage limits
    cleanupOldBackups(5);
    
    log('Configuration backup created', LOG_LEVELS.INFO, { backupKey });
  } catch (error) {
    log('Error creating config backup', LOG_LEVELS.ERROR, error);
  }
}

/**
 * Cleans up old backups to prevent hitting storage limits
 * 
 * @param {number} keepCount - Number of most recent backups to keep
 */
function cleanupOldBackups(keepCount = 5) {
  try {
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
  } catch (error) {
    log('Error cleaning up old backups', LOG_LEVELS.ERROR, error);
  }
}

/**
 * Gets the list of available backups
 * 
 * @returns {Array} List of backup objects with key and timestamp
 */
function getBackupsList() {
  try {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    
    // Find all backups and format them for display
    return Object.keys(allProps)
      .filter(key => key.startsWith(`${CONFIG_PROPERTY_KEY}_BACKUP_`))
      .map(key => {
        const timestamp = key.replace(`${CONFIG_PROPERTY_KEY}_BACKUP_`, '');
        const readableDate = timestamp.replace('T', ' ').substring(0, 19).replace(/-/g, ':');
        
        return {
          key: key,
          timestamp: timestamp,
          readableDate: readableDate
        };
      })
      .sort((a, b) => b.timestamp.localeCompare(a.timestamp)); // Sort newest first
  } catch (error) {
    log('Error listing backups', LOG_LEVELS.ERROR, error);
    return [];
  }
}

/**
 * Restores configuration from a backup
 * 
 * @param {string} backupKey - Key of the backup to restore
 * @returns {Object} Result object indicating success or failure
 */
function restoreBackup(backupKey) {
  try {
    const props = PropertiesService.getScriptProperties();
    const backupData = props.getProperty(backupKey);
    
    if (!backupData) {
      return { 
        status: 'error', 
        message: 'Backup not found' 
      };
    }
    
    // Parse backup data
    const config = JSON.parse(backupData);
    
    // Update metadata for restored version
    if (!config.metadata) {
      config.metadata = {};
    }
    
    config.metadata.lastModified = new Date().toISOString();
    config.metadata.restoredFrom = backupKey;
    
    // Ensure all color arrays exist and are complete
    if (config.statuses && Array.isArray(config.statuses)) {
      if (!config.statusColors || config.statusColors.length !== config.statuses.length) {
        config.statusColors = generateDefaultColors(config.statuses, 'status');
      }
    }
    
    if (config.types && Array.isArray(config.types)) {
      if (!config.typeColors || config.typeColors.length !== config.types.length) {
        config.typeColors = generateDefaultColors(config.types, 'type');
      }
    }
    
    if (config.priorities && Array.isArray(config.priorities)) {
      if (!config.priorityColors || config.priorityColors.length !== config.priorities.length) {
        config.priorityColors = generateDefaultColors(config.priorities, 'priority');
      }
    }
    
    // Save as current config
    props.setProperty(CONFIG_PROPERTY_KEY, JSON.stringify(config));
    
    // Update cache
    _configCache = config;
    
    log('Configuration restored from backup', LOG_LEVELS.INFO, { backupKey });
    return { 
      status: 'success', 
      message: 'Configuration restored successfully' 
    };
  } catch (error) {
    return handleError('restoreBackup', error, false);
  }
}

/**
 * Exports configuration as JSON text
 * 
 * @returns {Object} Result with the exported config JSON
 */
function exportConfig() {
  try {
    const config = getConfig();
    return { 
      status: 'success', 
      data: JSON.stringify(config, null, 2) 
    };
  } catch (error) {
    return handleError('exportConfig', error, false);
  }
}

/**
 * Imports configuration from JSON text
 * 
 * @param {string} jsonConfig - JSON string with configuration
 * @returns {Object} Result object indicating success or failure
 */
function importConfig(jsonConfig) {
  try {
    // Parse and validate the JSON
    const config = JSON.parse(jsonConfig);
    
    // Basic validation
    if (!config.statuses || !config.types || !config.priorities || !config.fields) {
      return { 
        status: 'error', 
        message: 'Invalid configuration format' 
      };
    }
    
    // Create backup before replacing
    createBackup(getConfig());
    
    // Update metadata
    if (!config.metadata) {
      config.metadata = {};
    }
    
    config.metadata.lastModified = new Date().toISOString();
    config.metadata.importedAt = new Date().toISOString();
    
    // Ensure all color arrays exist and are complete
    if (config.statuses && Array.isArray(config.statuses)) {
      if (!config.statusColors || config.statusColors.length !== config.statuses.length) {
        config.statusColors = generateDefaultColors(config.statuses, 'status');
      }
    }
    
    if (config.types && Array.isArray(config.types)) {
      if (!config.typeColors || config.typeColors.length !== config.types.length) {
        config.typeColors = generateDefaultColors(config.types, 'type');
      }
    }
    
    if (config.priorities && Array.isArray(config.priorities)) {
      if (!config.priorityColors || config.priorityColors.length !== config.priorities.length) {
        config.priorityColors = generateDefaultColors(config.priorities, 'priority');
      }
    }
    
    // Save the imported config
    const props = PropertiesService.getScriptProperties();
    props.setProperty(CONFIG_PROPERTY_KEY, JSON.stringify(config));
    
    // Update cache
    _configCache = config;
    
    log('Configuration imported', LOG_LEVELS.INFO);
    return { 
      status: 'success', 
      message: 'Configuration imported successfully' 
    };
  } catch (error) {
    return handleError('importConfig', error, false);
  }
}

/**
 * Loads configuration in a format suitable for the settings dialog
 * 
 * @interface Used by InitSettings.html
 * @returns {Array} Configuration objects for the settings panel
 */
function loadConfig() {
  try {
    const config = getConfig();
    const result = [];
    
    // Default colors if none are defined
    const defaultColors = {
      status: '#1976d2',   // Blue
      type: '#ff9800',     // Orange
      priority: '#f44336'  // Red
    };
    
    // Status dropdown
    if (config.statuses && Array.isArray(config.statuses)) {
      // Use stored colors if available, otherwise use defaults
      let colors;
      if (config.statusColors && Array.isArray(config.statusColors) && 
          config.statusColors.length === config.statuses.length) {
        colors = config.statusColors;
      } else {
        colors = Array(config.statuses.length).fill(defaultColors.status);
      }
      
      result.push({
        name: "Status",
        type: "dropdown", 
        values: config.statuses,
        colors: colors
      });
    } else {
      // Fallback if statuses is missing
      result.push({
        name: "Status",
        type: "dropdown",
        values: ['Pending', 'In-Progress', 'Done'],
        colors: ['#1976d2', '#1976d2', '#1976d2']
      });
    }
    
    // Job Type dropdown
    if (config.types && Array.isArray(config.types)) {
      // Use stored colors if available, otherwise use defaults
      let colors;
      if (config.typeColors && Array.isArray(config.typeColors) && 
          config.typeColors.length === config.types.length) {
        colors = config.typeColors;
      } else {
        colors = Array(config.types.length).fill(defaultColors.type);
      }
      
      result.push({
        name: "Job Type",
        type: "dropdown",
        values: config.types,
        colors: colors
      });
    } else {
      // Fallback if types is missing
      result.push({
        name: "Job Type",
        type: "dropdown",
        values: ['Build', 'Test', 'Deploy'],
        colors: ['#ff9800', '#ff9800', '#ff9800']
      });
    }
    
    // Priority dropdown
    if (config.priorities && Array.isArray(config.priorities)) {
      // Use stored colors if available, otherwise use defaults
      let colors;
      if (config.priorityColors && Array.isArray(config.priorityColors) && 
          config.priorityColors.length === config.priorities.length) {
        colors = config.priorityColors;
      } else {
        colors = Array(config.priorities.length).fill(defaultColors.priority);
      }
      
      result.push({
        name: "Priority",
        type: "dropdown",
        values: config.priorities, 
        colors: colors
      });
    } else {
      // Fallback if priorities is missing
      result.push({
        name: "Priority",
        type: "dropdown",
        values: ['High', 'Medium', 'Low'],
        colors: ['#f44336', '#f44336', '#f44336']
      });
    }
    
    log('Loaded configuration for settings', LOG_LEVELS.DEBUG, result);
    return result;
  } catch (error) {
    handleError('loadConfig', error, false);
    
    // Return fallback values if there's an error
    return [
      {
        name: "Status",
        type: "dropdown",
        values: ['Pending', 'In-Progress', 'Done'],
        colors: ['#1976d2', '#1976d2', '#1976d2']
      },
      {
        name: "Job Type",
        type: "dropdown",
        values: ['Build', 'Test', 'Deploy'],
        colors: ['#ff9800', '#ff9800', '#ff9800']
      },
      {
        name: "Priority",
        type: "dropdown",
        values: ['High', 'Medium', 'Low'],
        colors: ['#f44336', '#f44336', '#f44336']
      }
    ];
  }
}

/**
 * Gets metadata about the current configuration
 *
 * @returns {Object} Configuration metadata
 */
function getConfigMetadata() {
  const config = getConfig();
  return {
    version: config.metadata?.version || 1,
    created: config.metadata?.created || "Unknown",
    lastModified: config.metadata?.lastModified || "Unknown",
    statusCount: config.statuses?.length || 0,
    typeCount: config.types?.length || 0,
    priorityCount: config.priorities?.length || 0
  };
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