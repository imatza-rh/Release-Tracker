# Release Tracker

![Release Tracker Logo](https://img.shields.io/badge/Release-Tracker-1976d2?style=for-the-badge)
![Version](https://img.shields.io/badge/version-1.0.0-green.svg)
![Status](https://img.shields.io/badge/status-active-success.svg)
![License](https://img.shields.io/badge/license-MIT-blue.svg)

A comprehensive Google Sheets add-on to manage and track release jobs, tasks, and phases with integrated notifications, external system links, and customizable workflows.

## 📋 Overview

Release Tracker transforms Google Sheets into a powerful release management tool. Track jobs, milestones, and project phases all in a familiar spreadsheet interface with the specialized functionality you need to coordinate releases effectively.

Built with Google Apps Script, Release Tracker adds powerful capabilities to Google Sheets:

- Manage jobs with custom status, type, and priority tracking
- Link to external systems like Jenkins and Jira
- Schedule notifications for upcoming phase milestones
- Customize fields and workflows to match your processes
- All data stays in your Google Sheets - no external storage needed

## ✨ Key Features

- **Job Management**: Create, update, and track jobs with customizable statuses, types, and priorities
- **External Links**: Automatic hyperlinks to Jenkins jobs and Jira tickets
- **Phase Tracking**: Monitor project phases with start/end dates and notifications
- **Email & Slack Notifications**: Get alerts when phases start or end
- **Custom Fields**: Adapt the tracker to your specific workflow needs
- **Multiple Tracking Sheets**: Create separate tracking sheets for different projects or releases
- **Visual Status Indicators**: Color-coded statuses and types for at-a-glance monitoring

## 🚀 Getting Started

### Installation

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Create new script files for each `.gs` file in the project
4. Create new HTML files for each `.html` file in the project
5. Copy the code from the project files into the corresponding Apps Script files
6. Save all files
7. Refresh your Google Sheet

### Initial Setup

1. After installation, you'll see a new menu: **🚀 Release Tracker**
2. Configure Jenkins and Jira URLs in the **🔗 External Links** tab
3. Create your first tracking sheet by selecting **📝 Create Jobs Tracking Sheet**

## 📚 Usage Guide

### Managing Jobs

1. Select **🛠 Manage Jobs** from the Release Tracker menu
2. Use the sidebar to add, update, or remove jobs
3. Fill in the required details for each job:
   - Job Name (required)
   - Type (Build, Deploy, Test, etc.)
   - Status (Pending, In-Progress, Done, etc.)
   - Priority (High, Medium, Low)
   - Notes
   - Job Link (for Jenkins)
   - Jira Ticket

### Viewing Jobs

1. Select **👀 View Jobs** from the Release Tracker menu
2. Use filters to quickly find jobs by status or search terms
3. Click the 🔍 icon to locate a job in the spreadsheet

### Phase Notifications

1. Create a "Phases and Dates" sheet using the menu: **🔔 Phase Notifications**
2. Add phases with start and end dates
3. Configure notification recipients (email addresses and Slack webhooks)
4. Enable notifications to receive alerts when phases start or end

## ⚙️ Customization

### Customizing Fields

1. Use the **📝 Create Jobs Tracking Sheet** dialog
2. Add, remove, or modify statuses, types, and priorities
3. Choose colors for visual status indicators

### Creating Multiple Tracking Sheets

1. Select **📝 Create Jobs Tracking Sheet**
2. Specify a name for your new tracking sheet
3. Choose job sources (manual entry, from DefaultJobs sheet, or empty)
4. Customize available fields and options

## 🛠️ Technical Architecture

Release Tracker is built with modular components:

- **A_Constants**: Global constants and shared variables
- **B_Logging**: Logging utilities and error handling
- **C_Utils**: General utility functions
- **D_Config**: Configuration management with backups
- **E_SheetSetup**: Sheet creation and formatting
- **F_JobOperations**: Job CRUD operations
- **G_DefaultJobs**: Job initialization functionality
- **H_UI**: User interface controls
- **I_TemplateLoader**: HTML template system
- **J_Integrations**: External system integrations
- **K_CreateTrackingSheet**: Sheet creation tools
- **N_Notifications**: Notification system
- **Z_Main**: Main entry point and initialization

## 🔧 Troubleshooting

### Common Issues

1. **Menu not appearing**
   - Refresh the page
   - Check for Apps Script errors in the script editor

2. **External links not working**
   - Verify integration URLs in the settings
   - Ensure proper URL formats with correct prefixes

3. **Notifications not sending**
   - Check that the trigger is properly set up
   - Verify recipient email addresses and webhook URLs
   - Ensure the "Phases and Dates" sheet exists with proper columns

For advanced troubleshooting, check the logs in the Apps Script editor console.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👥 Contributors

- Initial development team

## 🙏 Acknowledgments

- Google Apps Script platform
- All beta testers and early adopters

---

*For support, please open an issue in the repository.*