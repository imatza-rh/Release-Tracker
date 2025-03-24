<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    :root {
      --primary-color: #1976d2;
      --hover-color: #1565c0;
      --text-color: #444;
      --light-bg: #f9f9f9;
      --border-color: #e0e0e0;
      --danger-color: #dc3545;
      --success-color: #28a745;
      --warning-color: #ffc107;
      --info-color: #17a2b8;
    }

    body {
      font-family: Arial, sans-serif;
      margin: 10px;
      font-size: 14px;
      color: var(--text-color);
      background-color: #fff;
    }

    h3 {
      margin-top: 0;
      margin-bottom: 16px;
      color: var(--primary-color);
      font-size: 18px;
      padding-bottom: 8px;
      border-bottom: 2px solid var(--primary-color);
    }
    
    h4 {
      color: var(--primary-color);
      margin-bottom: 10px;
      border-bottom: 1px solid var(--primary-color);
      padding-bottom: 5px;
    }

    p {
      margin-bottom: 12px;
    }

    /* Tab Control Styles */
    .tab-container {
      margin-top: 20px;
    }
    
    .tab-buttons {
      display: flex;
      border-bottom: 1px solid var(--border-color);
    }
    
    .tab-button {
      padding: 10px 20px;
      cursor: pointer;
      background: #f1f1f1;
      margin-right: 2px;
      border: 1px solid var(--border-color);
      border-bottom: none;
      border-radius: 4px 4px 0 0;
    }
    
    .tab-button.active {
      background: white;
      font-weight: bold;
      color: var(--primary-color);
      border-bottom: 3px solid white;
      margin-bottom: -2px;
    }
    
    .tab-content {
      display: none;
      padding: 20px;
      border: 1px solid var(--border-color);
      border-top: none;
      background: white;
    }
    
    .tab-content.active {
      display: block;
    }

    /* Form controls */
    label {
      display: block;
      margin-top: 10px;
      margin-bottom: 5px;
      font-weight: bold;
      font-size: 13px;
    }
    
    select, input, textarea {
      width: 100%;
      box-sizing: border-box;
      margin-bottom: 12px;
      font-size: 13px;
      padding: 8px;
      border: 1px solid var(--border-color);
      border-radius: 4px;
    }
    
    select:focus, input:focus, textarea:focus {
      outline: none;
      border-color: var(--primary-color);
      box-shadow: 0 0 0 2px rgba(25, 118, 210, 0.2);
    }

    /* Job list textarea */
    #manualJobsList {
      height: 120px;
      font-family: monospace;
      resize: vertical;
    }

    /* Field Options styling to match Google Sheets */
    .field-section {
      margin-bottom: 30px;
    }

    .status-header, .type-header, .priority-header {
      color: var(--primary-color);
      font-size: 16px;
      font-weight: bold;
      border-bottom: 1px solid var(--primary-color);
      padding-bottom: 8px;
      margin-bottom: 16px;
    }

    .field-description {
      margin-bottom: 15px;
    }

    .field-item {
      display: flex;
      align-items: center;
      margin-bottom: 10px;
    }

    .check-input {
      margin-right: 10px;
      width: 20px;
      height: 20px;
    }

    .text-input {
      flex-grow: 1;
      padding: 8px 12px;
      border: 1px solid #ccc;
      border-radius: 4px;
      font-size: 14px;
      margin-right: 10px;
    }

    .color-box {
      width: 32px;
      height: 32px;
      border-radius: 4px;
      border: 1px solid #ccc;
      margin-right: 10px;
      cursor: pointer;
    }

    .remove-btn {
      width: 32px;
      height: 32px;
      background: #6c757d;
      color: white;
      border: none;
      border-radius: 4px;
      font-size: 16px;
      cursor: pointer;
    }

    .add-btn {
      background: #6c757d;
      color: white;
      border: none;
      padding: 8px 16px;
      border-radius: 4px;
      cursor: pointer;
      margin-top: 10px;
      display: inline-flex;
      align-items: center;
    }

    /* Color Picker (Google Sheets Style) */
    #color-picker {
      display: none;
      position: absolute;
      background: white;
      border: 1px solid #ccc;
      box-shadow: 0 2px 10px rgba(0,0,0,0.2);
      padding: 10px;
      z-index: 1000;
      width: 240px;
    }

    .color-picker-title {
      font-size: 12px;
      font-weight: bold;
      color: #666;
      margin: 4px 0 6px 0;
    }

    .color-grid {
      display: flex;
      flex-wrap: wrap;
      margin-bottom: 10px;
    }

    .color-cell {
      width: 20px;
      height: 20px;
      margin: 2px;
      border-radius: 50%;
      cursor: pointer;
      position: relative;
    }

    .color-cell.selected::after {
      content: "✓";
      position: absolute;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      color: white;
      font-size: 12px;
      text-shadow: 0 0 1px rgba(0,0,0,0.8);
    }

    /* Table Format Options */
    .format-options {
      margin-top: 15px;
      padding: 10px;
      border: 1px solid #eee;
      background-color: #f9f9f9;
      border-radius: 4px;
    }

    .format-title {
      font-weight: bold;
      margin-bottom: 8px;
    }

    .format-checkbox {
      margin-right: 8px;
    }

    /* Action and notice areas */
    .action-area {
      background-color: #f8f9fa;
      padding: 10px;
      border-radius: 4px;
      margin-bottom: 15px;
    }

    .notice {
      background-color: #fff3cd;
      color: #856404;
      padding: 8px 12px;
      border-radius: 4px;
      margin-bottom: 10px;
    }

    .create-default-btn {
      background-color: var(--primary-color);
      color: white;
      border: none;
      padding: 6px 12px;
      border-radius: 4px;
      cursor: pointer;
      margin-top: 10px;
      font-size: 13px;
    }

    /* Success/Error Dialogs */
    .dialog-overlay {
      display: none;
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background-color: rgba(0,0,0,0.5);
      z-index: 2000;
      align-items: center;
      justify-content: center;
    }

    .dialog-box {
      width: 300px;
      background-color: white;
      border-radius: 8px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.2);
      padding: 20px;
      text-align: center;
    }

    .dialog-icon {
      font-size: 48px;
      margin-bottom: 15px;
    }

    .dialog-title {
      font-size: 18px;
      font-weight: bold;
      margin-bottom: 10px;
    }

    .dialog-message {
      margin-bottom: 20px;
    }

    .dialog-button {
      background-color: var(--primary-color);
      color: white;
      border: none;
      padding: 8px 16px;
      border-radius: 4px;
      cursor: pointer;
      font-size: 14px;
    }

    /* Loading Overlay */
    .loading-overlay {
      display: none;
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background-color: rgba(255,255,255,0.8);
      z-index: 2000;
      flex-direction: column;
      align-items: center;
      justify-content: center;
    }

    .spinner {
      width: 40px;
      height: 40px;
      border: 4px solid rgba(25, 118, 210, 0.2);
      border-top: 4px solid var(--primary-color);
      border-radius: 50%;
      animation: spin 1s linear infinite;
      margin-bottom: 20px;
    }

    .loading-message {
      font-size: 16px;
      font-weight: bold;
      color: var(--primary-color);
    }

    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }

    /* Buttons */
    .buttons-container {
      display: flex;
      justify-content: flex-end;
      gap: 10px;
      margin-top: 20px;
      padding-top: 15px;
      border-top: 1px solid #eee;
    }

    button {
      padding: 8px 16px;
      border-radius: 4px;
      border: none;
      cursor: pointer;
      font-size: 14px;
    }
    
    .btn-primary {
      background-color: var(--primary-color);
      color: white;
    }
    
    .btn-primary:hover {
      background-color: var(--hover-color);
    }
    
    .btn-success {
      background-color: var(--success-color);
      color: white;
    }

    .btn-secondary {
      background-color: #6c757d;
      color: white;
    }
    
    /* Helper text */
    .helper-text {
      font-size: 12px;
      color: #757575;
      margin-top: -8px;
      margin-bottom: 12px;
      font-style: italic;
    }
  </style>
</head>
<body>
  <h3>Create Tracking Sheet</h3>
  
  <div class="tab-container">
    <div class="tab-buttons">
      <div id="tab-btn-basics" class="tab-button active" onclick="showTab('basics')">Basics</div>
      <div id="tab-btn-external-links" class="tab-button" onclick="showTab('external-links')">External Links</div>
      <div id="tab-btn-fields" class="tab-button" onclick="showTab('fields')">Field Options</div>
    </div>
    
    <!-- Basics Tab -->
    <div id="tab-basics" class="tab-content active">
      <p>Set up basic information for your new tracking sheet.</p>
      
      <label for="jobSheetName">Sheet Name:</label>
      <input type="text" id="jobSheetName" value="Jobs" placeholder="Enter sheet name">
      <div class="helper-text">Default is "Jobs". Change this to create a new tracking sheet.</div>
      
      <label for="defaultJobsSource">Job Source:</label>
      <select id="defaultJobsSource" onchange="toggleJobSource()">
        <option value="empty">Empty sheet (headers only)</option>
        <option value="DefaultJobs">Use DefaultJobs sheet</option>
        <option value="manual">Enter jobs below</option>
      </select>
      
      <div id="defaultJobsInfo" style="display: none;" class="action-area">
        <div id="defaultJobsNotice" class="notice">
          Note: The DefaultJobs sheet will be used as a template. If it doesn't exist, you'll need to create it first.
        </div>
        <button id="createDefaultJobsBtn" class="create-default-btn" onclick="createDefaultJobsSheet()">
          Create DefaultJobs Sheet
        </button>
      </div>
      
      <div id="manualJobsInput" style="display: none;">
        <label for="manualJobsList">Enter job names (one per line):</label>
        <textarea id="manualJobsList" placeholder="Example:
Update database
Deploy application
Run tests
Review code changes"></textarea>
      </div>
      
      <div class="format-options">
        <div class="format-title">Table Formatting Options:</div>
        <div style="margin-bottom: 8px;">
          <input type="checkbox" id="formatAsTable" class="format-checkbox" checked>
          <label for="formatAsTable" style="display: inline; font-weight: normal;">Format as table (alternating row colors)</label>
        </div>
        <div>
          <input type="checkbox" id="freezeHeaders" class="format-checkbox" checked>
          <label for="freezeHeaders" style="display: inline; font-weight: normal;">Freeze header row</label>
        </div>
      </div>
    </div>
    
    <!-- External Links Tab -->
    <div id="tab-external-links" class="tab-content">
      <p>Configure external system URLs to enable hyperlinks in your tracking sheet.</p>
      
      <label for="jenkinsUrl">Jenkins Base URL:</label>
      <input type="text" id="jenkinsUrl" placeholder="https://jenkins.example.com">
      <div class="helper-text">Do not include trailing slash. We'll add "/job/" between this URL and your job name.</div>
      
      <label for="jiraUrl">Jira Base URL:</label>
      <input type="text" id="jiraUrl" placeholder="https://jira.example.com/browse">
      <div class="helper-text">Do not include trailing slash. Your ticket ID will be appended to this URL.</div>
    </div>
    
    <!-- Field Options Tab -->
    <div id="tab-fields" class="tab-content">
      <p>Customize the drop-down fields for your tracking sheet.</p>
      
      <div class="field-section">
        <div class="status-header">Status Options</div>
        <div class="field-description">Define status options to track job progress.</div>
        <div id="status-items"></div>
        <button onclick="addFieldOption('status')" class="add-btn">+ Add Status</button>
      </div>
      
      <div class="field-section">
        <div class="type-header">Job Type Options</div>
        <div class="field-description">Define different types of jobs you want to track.</div>
        <div id="type-items"></div>
        <button onclick="addFieldOption('type')" class="add-btn">+ Add Type</button>
      </div>
      
      <div class="field-section">
        <div class="priority-header">Priority Options</div>
        <div class="field-description">Define priority levels for your jobs.</div>
        <div id="priority-items"></div>
        <button onclick="addFieldOption('priority')" class="add-btn">+ Add Priority</button>
      </div>
    </div>
  </div>

  <div class="buttons-container">
    <button id="cancelBtn" class="btn-secondary">Cancel</button>
    <button id="createBtn" class="btn-primary">Create Tracking Sheet</button>
  </div>
  
  <!-- Color Picker (Hidden) -->
  <div id="color-picker">
    <div class="color-picker-title">STANDARD</div>
    <div class="color-grid">
      <div class="color-cell" data-color="#000000" style="background-color: #000000"></div>
      <div class="color-cell" data-color="#434343" style="background-color: #434343"></div>
      <div class="color-cell" data-color="#666666" style="background-color: #666666"></div>
      <div class="color-cell" data-color="#999999" style="background-color: #999999"></div>
      <div class="color-cell" data-color="#b7b7b7" style="background-color: #b7b7b7"></div>
      <div class="color-cell" data-color="#cccccc" style="background-color: #cccccc"></div>
      <div class="color-cell" data-color="#d9d9d9" style="background-color: #d9d9d9"></div>
      <div class="color-cell" data-color="#efefef" style="background-color: #efefef"></div>
      <div class="color-cell" data-color="#f3f3f3" style="background-color: #f3f3f3"></div>
      <div class="color-cell" data-color="#ffffff" style="background-color: #ffffff"></div>
      
      <!-- Colors row 1 -->
      <div class="color-cell" data-color="#980000" style="background-color: #980000"></div>
      <div class="color-cell" data-color="#ff0000" style="background-color: #ff0000"></div>
      <div class="color-cell" data-color="#ff9900" style="background-color: #ff9900"></div>
      <div class="color-cell" data-color="#ffff00" style="background-color: #ffff00"></div>
      <div class="color-cell" data-color="#00ff00" style="background-color: #00ff00"></div>
      <div class="color-cell" data-color="#00ffff" style="background-color: #00ffff"></div>
      <div class="color-cell" data-color="#0000ff" style="background-color: #0000ff"></div>
      <div class="color-cell" data-color="#9900ff" style="background-color: #9900ff"></div>
      <div class="color-cell" data-color="#ff00ff" style="background-color: #ff00ff"></div>
      <div class="color-cell" data-color="#ff0099" style="background-color: #ff0099"></div>
      
      <!-- Lighter colors -->
      <div class="color-cell" data-color="#f4cccc" style="background-color: #f4cccc"></div>
      <div class="color-cell" data-color="#fce5cd" style="background-color: #fce5cd"></div>
      <div class="color-cell" data-color="#fff2cc" style="background-color: #fff2cc"></div>
      <div class="color-cell" data-color="#d9ead3" style="background-color: #d9ead3"></div>
      <div class="color-cell" data-color="#d0e0e3" style="background-color: #d0e0e3"></div>
      <div class="color-cell" data-color="#cfe2f3" style="background-color: #cfe2f3"></div>
      <div class="color-cell" data-color="#d9d2e9" style="background-color: #d9d2e9"></div>
      <div class="color-cell" data-color="#ead1dc" style="background-color: #ead1dc"></div>
    </div>
    
    <div class="color-picker-title">COMMON</div>
    <div class="color-grid">
      <div class="color-cell" data-color="#000000" style="background-color: #000000"></div>
      <div class="color-cell" data-color="#ffffff" style="background-color: #ffffff"></div>
      <div class="color-cell" data-color="#3d85c6" style="background-color: #3d85c6"></div>
      <div class="color-cell" data-color="#e06666" style="background-color: #e06666"></div>
      <div class="color-cell" data-color="#f1c232" style="background-color: #f1c232"></div>
      <div class="color-cell" data-color="#6aa84f" style="background-color: #6aa84f"></div>
      <div class="color-cell" data-color="#e69138" style="background-color: #e69138"></div>
      <div class="color-cell" data-color="#45818e" style="background-color: #45818e"></div>
    </div>
  </div>
  
  <!-- Loading Overlay -->
  <div id="loading-overlay" class="loading-overlay">
    <div class="spinner"></div>
    <div class="loading-message">Creating tracking sheet...</div>
  </div>
  
  <!-- Success Dialog -->
  <div id="success-dialog" class="dialog-overlay">
    <div class="dialog-box">
      <div class="dialog-icon" style="color: var(--success-color);">✓</div>
      <div class="dialog-title">Success!</div>
      <div id="success-message" class="dialog-message">Operation completed successfully.</div>
      <button class="dialog-button" onclick="closeDialog('success-dialog')">OK</button>
    </div>
  </div>
  
  <!-- Error Dialog -->
  <div id="error-dialog" class="dialog-overlay">
    <div class="dialog-box">
      <div class="dialog-icon" style="color: var(--danger-color);">✗</div>
      <div class="dialog-title">Error</div>
      <div id="error-message" class="dialog-message">An error occurred.</div>
      <button class="dialog-button" onclick="closeDialog('error-dialog')">OK</button>
    </div>
  </div>

  <script>
    // Default field values
    const defaultValues = {
      status: [
        { name: "Pending", color: "#3d85c6" },
        { name: "In-Progress", color: "#f1c232" },
        { name: "Done", color: "#6aa84f" }
      ],
      type: [
        { name: "Build", color: "#e69138" },
        { name: "Test", color: "#674ea7" },
        { name: "Deploy", color: "#3d85c6" }
      ],
      priority: [
        { name: "High", color: "#e06666" },
        { name: "Medium", color: "#f1c232" },
        { name: "Low", color: "#6aa84f" }
      ]
    };
    
    // Track current active color button
    let activeColorButton = null;
    let defaultJobsSheetExists = false;
    
    document.addEventListener("DOMContentLoaded", function() {
      // Show the default tab
      showTab('basics');
      
      // Initialize field options
      setupFieldOptions();
      
      // Toggle job source initially
      toggleJobSource();
      
      // Check if DefaultJobs sheet exists
      checkDefaultJobsSheet();
      
      // Load integration settings
      loadSettings();
      
      // Set up event listeners
      document.getElementById("defaultJobsSource").addEventListener("change", toggleJobSource);
      document.getElementById("createBtn").addEventListener("click", createTrackingSheet);
      document.getElementById("cancelBtn").addEventListener("click", () => {
        google.script.host.close();
      });
      
      // Set up color cell click handlers
      document.querySelectorAll('.color-cell').forEach(cell => {
        cell.addEventListener('click', function() {
          if (activeColorButton) {
            // Update active button color
            activeColorButton.style.backgroundColor = this.dataset.color;
            activeColorButton.dataset.color = this.dataset.color;
            
            // Hide color picker
            document.getElementById('color-picker').style.display = 'none';
            activeColorButton = null;
          }
        });
      });
      
      // Close color picker when clicking outside
      document.addEventListener('click', function(e) {
        const colorPicker = document.getElementById('color-picker');
        if (colorPicker.style.display === 'block' && 
            !colorPicker.contains(e.target) && 
            !e.target.classList.contains('color-box')) {
          colorPicker.style.display = 'none';
          activeColorButton = null;
        }
      });
    });
    
    function setupFieldOptions() {
      // Setup each field type
      createFieldItems('status');
      createFieldItems('type');
      createFieldItems('priority');
    }
    
    function createFieldItems(type) {
      const container = document.getElementById(`${type}-items`);
      container.innerHTML = ''; // Clear container
      
      // Add default values
      defaultValues[type].forEach((value, index) => {
        addFieldItem(type, value.name, value.color, index === 0);
      });
    }
    
    function addFieldItem(type, name, color, isRequired = false) {
      const container = document.getElementById(`${type}-items`);
      
      // Create field item
      const item = document.createElement('div');
      item.className = 'field-item';
      
      // Create checkbox
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.className = 'check-input';
      checkbox.checked = true;
      checkbox.disabled = isRequired;
      
      // Create text input
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'text-input';
      input.value = name;
      input.readOnly = isRequired;
      
      // Create color box
      const colorBox = document.createElement('div');
      colorBox.className = 'color-box';
      colorBox.style.backgroundColor = color;
      colorBox.dataset.color = color;
      colorBox.addEventListener('click', function(e) {
        showColorPicker(e, this);
      });
      
      // Create remove button
      const removeBtn = document.createElement('button');
      removeBtn.className = 'remove-btn';
      removeBtn.innerHTML = '−';
      removeBtn.style.display = isRequired ? 'none' : 'block';
      removeBtn.addEventListener('click', function() {
        container.removeChild(item);
      });
      
      // Add elements to item
      item.appendChild(checkbox);
      item.appendChild(input);
      item.appendChild(colorBox);
      item.appendChild(removeBtn);
      
      // Add item to container
      container.appendChild(item);
    }
    
    function addFieldOption(type) {
      let defaultColor;
      let defaultName;
      
      // Set default values based on type
      if (type === 'status') {
        defaultColor = '#3d85c6';
        defaultName = 'New Status';
      } else if (type === 'type') {
        defaultColor = '#e69138';
        defaultName = 'New Type';
      } else if (type === 'priority') {
        defaultColor = '#e06666';
        defaultName = 'New Priority';
      }
      
      addFieldItem(type, defaultName, defaultColor, false);
    }
    
    function showColorPicker(event, colorBox) {
      event.stopPropagation();
      
      // Set active color box
      activeColorButton = colorBox;
      
      // Get color picker
      const colorPicker = document.getElementById('color-picker');
      
      // Position color picker
      const rect = colorBox.getBoundingClientRect();
      const pickerHeight = 320; // Approximate picker height
      
      // Determine if picker should appear above or below the color box
      const viewportHeight = window.innerHeight;
      const spaceBelowBox = viewportHeight - rect.bottom;
      
      // If space below is insufficient, place above
      if (spaceBelowBox < pickerHeight && rect.top > pickerHeight) {
        colorPicker.style.top = (rect.top - pickerHeight - 5) + 'px';
      } else {
        colorPicker.style.top = (rect.bottom + 5) + 'px';
      }
      
      // Center horizontally relative to the color box
      colorPicker.style.left = (rect.left) + 'px';
      
      // Update selected color
      const selectedColor = colorBox.dataset.color;
      document.querySelectorAll('.color-cell').forEach(cell => {
        cell.classList.remove('selected');
        if (cell.dataset.color === selectedColor) {
          cell.classList.add('selected');
        }
      });
      
      // Show color picker
      colorPicker.style.display = 'block';
    }
    
    function checkDefaultJobsSheet() {
      google.script.run
        .withSuccessHandler(function(exists) {
          defaultJobsSheetExists = exists;
          updateDefaultJobsNotice();
        })
        .withFailureHandler(function(error) {
          console.error("Failed to check DefaultJobs sheet:", error);
          defaultJobsSheetExists = false;
          updateDefaultJobsNotice();
        })
        .checkDefaultJobsSheetExists();
    }
    
    function createDefaultJobsSheet() {
      showLoading(true);
      
      google.script.run
        .withSuccessHandler(function(result) {
          showLoading(false);
          if (result.status === 'success') {
            showSuccessDialog('DefaultJobs sheet created successfully!');
            defaultJobsSheetExists = true;
            updateDefaultJobsNotice();
          } else {
            showErrorDialog(result.message || 'Failed to create DefaultJobs sheet.');
          }
        })
        .withFailureHandler(function(error) {
          showLoading(false);
          showErrorDialog('Failed to create DefaultJobs sheet: ' + error);
        })
        .createDefaultJobsTemplate();
    }
    
    function updateDefaultJobsNotice() {
      const notice = document.getElementById('defaultJobsNotice');
      const createBtn = document.getElementById('createDefaultJobsBtn');
      
      if (defaultJobsSheetExists) {
        notice.innerHTML = 'The DefaultJobs sheet exists and will be used as a template.';
        notice.style.backgroundColor = '#d4edda';
        notice.style.color = '#155724';
        createBtn.style.display = 'none';
      } else {
        notice.innerHTML = 'The DefaultJobs sheet does not exist. Click the button below to create it.';
        notice.style.backgroundColor = '#fff3cd';
        notice.style.color = '#856404';
        createBtn.style.display = 'block';
      }
    }
    
    function toggleJobSource() {
      const source = document.getElementById("defaultJobsSource").value;
      
      // Hide all divs first
      document.getElementById("defaultJobsInfo").style.display = "none";
      document.getElementById("manualJobsInput").style.display = "none";
      
      // Show the appropriate div based on selection
      if (source === "DefaultJobs") {
        document.getElementById("defaultJobsInfo").style.display = "block";
      } else if (source === "manual") {
        document.getElementById("manualJobsInput").style.display = "block";
      }
    }
    
    function loadSettings() {
      // Load integration settings if available
      google.script.run
          .withSuccessHandler(loadIntegrationSettings)
          .withFailureHandler(error => console.error("Integration settings error:", error))
          .getIntegrationSettings();
    }
    
    function loadIntegrationSettings(settings) {
      if (settings) {
        // Remove any trailing slashes
        let jenkinsUrl = settings.jenkinsUrl || "";
        let jiraUrl = settings.jiraUrl || "";
        
        if (jenkinsUrl.endsWith('/')) {
          jenkinsUrl = jenkinsUrl.slice(0, -1);
        }
        
        if (jiraUrl.endsWith('/')) {
          jiraUrl = jiraUrl.slice(0, -1);
        }
        
        document.getElementById("jenkinsUrl").value = jenkinsUrl;
        document.getElementById("jiraUrl").value = jiraUrl;
      }
    }
    
    function showTab(tabName) {
      // Hide all tabs
      const tabs = document.getElementsByClassName('tab-content');
      for (let i = 0; i < tabs.length; i++) {
        tabs[i].classList.remove('active');
      }
      
      // Deactivate all tab buttons
      const tabButtons = document.getElementsByClassName('tab-button');
      for (let i = 0; i < tabButtons.length; i++) {
        tabButtons[i].classList.remove('active');
      }
      
      // Show selected tab
      document.getElementById(`tab-${tabName}`).classList.add('active');
      document.getElementById(`tab-btn-${tabName}`).classList.add('active');
    }
    
    function createTrackingSheet() {
      const sheetName = document.getElementById("jobSheetName").value.trim();
      if (!sheetName) {
        showErrorDialog('Please enter a sheet name');
        return;
      }
      
      const source = document.getElementById("defaultJobsSource").value;
      
      // Get manual job list if needed
      let manualJobs = [];
      if (source === "manual") {
        const jobsList = document.getElementById("manualJobsList").value;
        if (jobsList.trim()) {
          manualJobs = jobsList.trim().split('\n').map(job => job.trim()).filter(job => job);
        }
      }
      
      // Table formatting options
      const formatAsTable = document.getElementById("formatAsTable").checked;
      const freezeHeaders = document.getElementById("freezeHeaders").checked;
      
      // Collect field options
      const fieldOptions = {
        statuses: getFieldValues('status'),
        types: getFieldValues('type'),
        priorities: getFieldValues('priority'),
        statusColors: getFieldColors('status'),
        typeColors: getFieldColors('type'),
        priorityColors: getFieldColors('priority')
      };
      
      // Collect integration settings
      let jenkinsUrl = document.getElementById("jenkinsUrl").value.trim();
      let jiraUrl = document.getElementById("jiraUrl").value.trim();
      
      // Remove trailing slashes
      if (jenkinsUrl.endsWith('/')) {
        jenkinsUrl = jenkinsUrl.slice(0, -1);
      }
      
      if (jiraUrl.endsWith('/')) {
        jiraUrl = jiraUrl.slice(0, -1);
      }
      
      const integrationSettings = {
        jenkinsUrl: jenkinsUrl,
        jiraUrl: jiraUrl
      };
      
      // Show loading overlay
      showLoading(true);
      
      // Save integration settings first
      google.script.run
        .withSuccessHandler(() => {
          // Then create the tracking sheet with field options
          google.script.run
            .withSuccessHandler(onCreateSuccess)
            .withFailureHandler(onCreateFailure)
            .createTrackingSheetWithOptions(sheetName, source, fieldOptions, manualJobs, {
              formatAsTable: formatAsTable,
              freezeHeaders: freezeHeaders
            });
        })
        .withFailureHandler(error => {
          showLoading(false);
          showErrorDialog('Failed to save external link settings: ' + error);
        })
        .saveIntegrationSettings(integrationSettings);
    }
    
    function getFieldValues(type) {
      const container = document.getElementById(`${type}-items`);
      const items = container.querySelectorAll('.field-item');
      const values = [];
      
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        const checkbox = item.querySelector('.check-input');
        const textInput = item.querySelector('.text-input');
        
        if (checkbox.checked && textInput.value.trim()) {
          values.push(textInput.value.trim());
        }
      }
      
      return values;
    }
    
    function getFieldColors(type) {
      const container = document.getElementById(`${type}-items`);
      const items = container.querySelectorAll('.field-item');
      const colors = [];
      
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        const checkbox = item.querySelector('.check-input');
        const colorBox = item.querySelector('.color-box');
        
        if (checkbox.checked) {
          colors.push(colorBox.dataset.color);
        }
      }
      
      return colors;
    }
    
    function onCreateSuccess(result) {
      showLoading(false);
      showSuccessDialog(result.message || 'Tracking sheet created successfully!', function() {
        google.script.host.close();
      });
    }
    
    function onCreateFailure(error) {
      showLoading(false);
      showErrorDialog('Failed to create tracking sheet: ' + error);
    }
    
    function showLoading(show) {
      document.getElementById('loading-overlay').style.display = show ? 'flex' : 'none';
    }
    
    function showSuccessDialog(message, callback) {
      document.getElementById('success-message').textContent = message;
      document.getElementById('success-dialog').style.display = 'flex';
      
      // Store callback if provided
      if (callback) {
        document.getElementById('success-dialog').dataset.callback = 'true';
      } else {
        delete document.getElementById('success-dialog').dataset.callback;
      }
      
      // Store callback function
      document.getElementById('success-dialog')._callback = callback;
    }
    
    function showErrorDialog(message) {
      document.getElementById('error-message').textContent = message;
      document.getElementById('error-dialog').style.display = 'flex';
    }
    
    function closeDialog(dialogId) {
      const dialog = document.getElementById(dialogId);
      dialog.style.display = 'none';
      
      // Execute callback if exists
      if (dialog.dataset.callback && typeof dialog._callback === 'function') {
        dialog._callback();
      }
    }
  </script>
</body>
</html>
