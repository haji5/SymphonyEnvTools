# Environics Data Tools Suite

**Version:** 1.0  
**Release Date:** August 2025  
**Executable:** EnvironicsDataTools.exe

## Overview

The Environics Data Tools Suite is a comprehensive collection of automation tools designed to streamline data processing, web scraping, and report generation for Environics Analytics workflows. This suite provides five integrated tools accessible through a central hub interface.

## System Requirements

- **Operating System:** Windows 10 or later
- **Memory:** 4GB RAM minimum (8GB recommended)
- **Storage:** 500MB free disk space
- **Internet Connection:** Required for web automation tools
- **Browser:** Google Chrome (automatically managed by the application)

## Installation & Setup

1. **Download** the `EnvironicsDataTools.exe` file
2. **Create a folder** for the application (e.g., `C:\EnvironicsDataTools\`)
3. **Place the .exe file** in your chosen folder
4. **Run as Administrator** (recommended for full functionality)
5. **Allow firewall access** when prompted

### First Run Setup
- The application will create necessary folders and files on first launch
- Chrome browser components are included and managed automatically
- No additional software installation required

## Tool Descriptions

### 1. PRIZM Report Generator
**Purpose:** Generate formatted PDF reports from PRIZM data files

**Features:**
- Processes Excel files containing PRIZM demographic data
- Maps PRIZM segment names to standardized codes (1-67 segments)
- Generates professional PDF reports with charts and visualizations
- Supports batch processing of multiple data files
- Automatic formatting and styling

**Usage:**
1. Click "PRIZM Report Generator" from the main menu
2. Select your input Excel file containing PRIZM data
3. Choose output directory for generated PDF reports
4. Click "Generate Report" to process

**Input Format:** Excel files (.xlsx) with PRIZM segment data  
**Output:** Professional PDF reports with demographic analysis

### 2. PRIZM Excel Merger
**Purpose:** Consolidate multiple Excel files from subfolders into a single merged file

**Features:**
- Scans subfolders for Excel files automatically
- Merges data from multiple sources while preserving structure
- Removes duplicate entries automatically
- Supports large datasets with progress tracking
- Maintains data integrity during merge process

**Usage:**
1. Click "PRIZM Excel Merger" from the main menu
2. Select the parent folder containing subfolders with Excel files
3. Choose output location and filename
4. Select data type (PRIZM, Demographic, etc.)
5. Click "Merge Files" to consolidate

**Input:** Multiple Excel files in separate subfolders  
**Output:** Single consolidated Excel file

### 3. Data Transformation Tool
**Purpose:** Automated data transformation and geocoding using Environics web services

**Features:**
- Web-based data processing automation
- Geocoding and address standardization
- File upload and download automation
- Progress tracking and status updates
- Error handling and retry mechanisms

**Usage:**
1. Click "Data Transformation Tool" from the main menu
2. Specify input file location (default: Downloads folder)
3. Set file position and page parameters
4. Click "Start Transformation" to begin automated process
5. Monitor progress through the status window

**Requirements:** Valid Environics Analytics account credentials  
**Input:** Raw data files for geocoding  
**Output:** Processed and geocoded data files

### 4. PRIZM Data Grabber
**Purpose:** Automated creation of 204 PRIZM dashboards for 51 target sets

**Features:**
- Fully automated dashboard generation
- Processes all 51 PRIZM target sets
- Creates 4 dashboard types per target set (204 total)
- Web automation with Chrome browser
- Progress tracking and error recovery
- Automatic login and session management

**Target Sets Include:**
- The A-List, Wealthy & Wise, Asian Sophisticates
- Turbo Burbs, First-Class Families, Downtown Verve
- Mature & Secure, Multiculture-ish
- And 43 additional demographic segments

**Usage:**
1. Click "PRIZM Data Grabber" from the main menu
2. Ensure stable internet connection
3. Click "Start Automated Process"
4. Monitor progress through real-time status updates
5. Process runs completely automatically (no user intervention required)

**Duration:** Approximately 2-4 hours for complete run  
**Output:** 204 PRIZM dashboard reports

### 5. Subdivision Data Grabber
**Purpose:** Automated subdivision dashboard creation for British Columbia geographic data

**Features:**
- Specialized for BC subdivision analysis
- Geographic boundary processing
- Automated data extraction and dashboard creation
- Custom subdivision mapping and analysis
- Integration with Environics geographic databases

**Usage:**
1. Click "Subdivision Data Grabber" from the main menu
2. Ensure connection to Environics Analytics platform
3. Click "Start Subdivision Process"
4. Monitor progress through status console
5. Review generated subdivision dashboards

**Scope:** British Columbia subdivision data  
**Output:** Subdivision-specific dashboard reports

## Central Hub Interface

The main application window provides:
- **Tool Selection:** Large buttons for each of the 5 tools
- **Status Monitoring:** Real-time status updates for active processes
- **Progress Tracking:** Visual feedback for long-running operations
- **Error Handling:** Clear error messages and recovery options
- **Exit Controls:** Safe application shutdown

## Automated Processes

Several tools in this suite run fully automated processes:

### Automation Features:
- **Web Browser Control:** Automated Chrome browser interactions
- **Login Management:** Automatic credential handling for Environics Analytics
- **File Processing:** Batch processing of multiple files
- **Progress Monitoring:** Real-time status updates
- **Error Recovery:** Automatic retry mechanisms for failed operations

### Long-Running Processes:
- **PRIZM Data Grabber:** 2-4 hours for complete execution
- **Subdivision Data Grabber:** 1-3 hours depending on data volume
- **Large File Merging:** Time varies based on file sizes

## Troubleshooting

### Common Issues:

**Application Won't Start:**
- Run as Administrator
- Check Windows Defender/antivirus settings
- Ensure sufficient disk space

**Web Automation Failures:**
- Verify internet connection
- Close other Chrome browser instances
- Check Environics Analytics site accessibility
- Restart the application

**File Processing Errors:**
- Verify input file formats (.xlsx for Excel files)
- Check file permissions (not read-only)
- Ensure sufficient disk space for output
- Close files in other applications

**Performance Issues:**
- Close unnecessary applications
- Ensure 8GB+ RAM for large datasets
- Use SSD storage for better performance

### Error Messages:
- **"Chrome driver not found":** Restart application (Chrome components are bundled)
- **"File access denied":** Close file in Excel/other programs
- **"Connection timeout":** Check internet connection and retry

## Data Security

- **Local Processing:** Most operations performed locally on your machine
- **Secure Connections:** HTTPS connections for web-based operations
- **No Data Storage:** Application doesn't store sensitive data permanently
- **Credential Handling:** Login credentials managed securely during sessions

## Support & Maintenance

### Regular Maintenance:
- Keep Windows updated for optimal performance
- Periodically clear temporary files
- Monitor disk space usage
- Restart application weekly for optimal performance

### Updates:
- Check for application updates monthly
- New versions will be distributed as updated .exe files
- Backup important data before updating

## Technical Specifications

**Built With:**
- Python 3.13
- Selenium WebDriver for automation
- Pandas for data processing
- Tkinter for user interface
- PyInstaller for executable creation

**Dependencies Included:**
- Chrome WebDriver components
- Excel processing libraries
- PDF generation tools
- Windows COM libraries

## Contact Information

For technical support or feature requests, contact the development team.

---

**© 2025 Symphony Tourism Services**  
*Automated Data Processing and Report Generation*
