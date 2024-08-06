# Project Management and Financial Reporting Automation

This collection of VBA modules automates various tasks related to project management and financial reporting in Excel. The modules work together to create, update, and manage projection sheets and financial data for multiple projects.

## Main Components:

1. **workbooksCreate.bas**: 
   - Creates projection sheets for new periods
   - Transfers data from previous periods to new sheets
   - Updates sheets with current financial data

2. **Worksheet_ImportData.bas**:
   - Imports data from various financial reports (e.g., Committed Costs, Job Labor Totals)
   - Populates projection sheets with current financial data
   - Calculates and updates various financial metrics

3. **WorksheetsCreate.bas**:
   - Manages the creation and updating of worksheet tabs
   - Copies existing sheets forward to new periods
   - Ensures proper naming and organization of sheets

4. **FoldersCreate.bas**:
   - Creates and manages the folder structure for organizing financial data
   - Sets up year, quarter, and month folders
   - Creates subfolders for different types of reports and documents

5. **SageReports_FetchData.bas**:
   - Checks for the presence of required Sage reports
   - Processes the Committed Costs report
   - Adds projection sections to worksheets
   - Calculates and formats various financial metrics

## Key Features:

- Automated creation of projection sheets for new periods
- Data transfer from previous periods to new sheets
- Integration with Sage accounting software reports
- Calculation of key financial metrics (e.g., percent complete, final cost projections)
- Folder and file management for organized data storage
- Formatting and structuring of financial reports for easy analysis

This automation suite streamlines the process of creating and updating financial projections, saving time and reducing errors in data entry and calculations. It's designed to work with a specific folder structure and set of Sage reports, making it ideal for organizations using Sage accounting software for project management and financial reporting.
