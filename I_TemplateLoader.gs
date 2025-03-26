/**
 * I_TemplateLoader.gs - HTML template loader with includes
 * 
 * This file provides functions for loading HTML templates with the ability
 * to include shared components.
 * 
 * @requires log, handleError, safeExecute from B_Logging.gs
 */

/**
 * Includes an HTML file in another HTML file as a template.
 * Used in HTML files with the <?!= include('filename') ?> syntax.
 * 
 * @param {string} fileName - Name of the HTML file to include (without .html)
 * @returns {string} Content of the included file
 */
function include(fileName) {
  return safeExecute('include', () => {
    return HtmlService.createHtmlOutputFromFile(fileName).getContent();
  }, false) || `Error loading ${fileName}`;
}

/**
 * Loads an HTML template with data and evaluates it
 * 
 * @param {string} templateName - Name of the HTML template without .html extension
 * @param {Object} [data={}] - Data to pass to the template
 * @param {boolean} [asHtml=false] - Whether to return as HTML output (true) or raw string (false)
 * @returns {HtmlOutput|string} Evaluated template as HtmlOutput or string
 */
function loadTemplate(templateName, data = {}, asHtml = false) {
  return safeExecute('loadTemplate', () => {
    const template = HtmlService.createTemplateFromFile(templateName);
    
    // Add each property from data to the template
    Object.keys(data).forEach(key => {
      template[key] = data[key];
    });
    
    // Evaluate the template
    const evaluated = template.evaluate();
    
    if (asHtml) {
      return evaluated;
    } else {
      return evaluated.getContent();
    }
  }, false);
}

/**
 * Creates a sidebar from an HTML template with data
 * 
 * @param {string} templateName - Name of the HTML template without .html extension
 * @param {string} title - Title for the sidebar
 * @param {Object} [data={}] - Data to pass to the template
 * @returns {HtmlOutput} Sidebar HTML output
 */
function createSidebar(templateName, title, data = {}) {
  return safeExecute('createSidebar', () => {
    const template = HtmlService.createTemplateFromFile(templateName);
    
    // Add each property from data to the template
    Object.keys(data).forEach(key => {
      template[key] = data[key];
    });
    
    return template.evaluate()
      .setTitle(title)
      .setWidth(300);
  }, false);
}

/**
 * Creates a modal dialog from an HTML template with data
 * 
 * @param {string} templateName - Name of the HTML template without .html extension
 * @param {string} title - Title for the dialog
 * @param {Object} [data={}] - Data to pass to the template
 * @param {Object} [options={}] - Dialog options (width, height)
 * @returns {HtmlOutput} Dialog HTML output
 */
function createDialog(templateName, title, data = {}, options = {}) {
  return safeExecute('createDialog', () => {
    const template = HtmlService.createTemplateFromFile(templateName);
    
    // Add each property from data to the template
    Object.keys(data).forEach(key => {
      template[key] = data[key];
    });
    
    const dialog = template.evaluate()
      .setTitle(title);
    
    // Set width if provided
    if (options.width) {
      dialog.setWidth(options.width);
    }
    
    // Set height if provided
    if (options.height) {
      dialog.setHeight(options.height);
    }
    
    return dialog;
  }, false);
}