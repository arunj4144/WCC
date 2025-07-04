// Google Apps Script for Windows Update Compliance Tracking
// Deploy this as a web app to receive data from PowerShell script

// Configuration
const SHEET_NAME = 'Windows Update Compliance';
const COMPLIANCE_EMAIL = 'your_email';
const NOTIFICATION_EMAIL = 'another_email';

// Main function to handle incoming POST requests
function doPost(e) {
  try {
    // Parse the incoming JSON data
    const data = JSON.parse(e.postData.contents);
    
    // Get or create the spreadsheet
    const sheet = getOrCreateSheet();
    
    // Process and add the data
    addDataToSheet(sheet, data);
    
    // Send notifications if needed
    handleNotifications(data);
    
    // Return success response
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Data received and processed successfully',
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Error processing request:', error);
    
    // Return error response
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: error.toString(),
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Function to get or create the compliance tracking sheet
function getOrCreateSheet() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // If no active spreadsheet, create a new one
  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create('Windows Update Compliance Tracking');
  }
  
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  // If sheet doesn't exist, create it
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    setupSheetHeaders(sheet);
  }
  
  return sheet;
}

// Function to setup sheet headers
function setupSheetHeaders(sheet) {
  const headers = [
    'Timestamp',
    'Computer Name',
    'User Name',
    'Current User',
    'OS Version',
    'OS Build',
    'IP Address',
    'Domain',
    'Available Updates',
    'Updates Installed',
    'Reboot Required',
    'Reboot Action',
    'User Email',
    'Postpone Reason',
    'Compliance Status',
    'Last Boot Time',
    'Script Version',
    'Execution Duration (min)',
    'Manufacturer',
    'Model',
    'Total RAM (GB)',
    'Update Details'
  ];
  
  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

// Function to add data to the sheet
function addDataToSheet(sheet, data) {
  const row = [
    new Date(data.Timestamp || new Date()),
    data.ComputerName || '',
    data.UserName || '',
    data.CurrentUser || '',
    data.SystemInfo?.OSVersion || '',
    data.SystemInfo?.OSBuild || '',
    data.SystemInfo?.IPAddress || '',
    data.SystemInfo?.Domain || '',
    data.AvailableUpdates || 0,
    data.UpdatesInstalled || 0,
    data.RebootRequired || false,
    data.RebootAction || 'None',
    data.UserEmail || '',
    data.PostponeReason || '',
    data.ComplianceStatus || '',
    data.SystemInfo?.LastBootTime || '',
    data.ScriptVersion || '',
    data.ExecutionDuration || 0,
    data.SystemInfo?.Manufacturer || '',
    data.SystemInfo?.Model || '',
    data.SystemInfo?.TotalRAM || 0,
    JSON.stringify(data.InstallDetails || [])
  ];
  
  // Add the row to the sheet
  sheet.appendRow(row);
  
  // Apply conditional formatting based on compliance status
  const lastRow = sheet.getLastRow();
  const complianceCell = sheet.getRange(lastRow, 15); // Compliance Status column
  
  const complianceStatus = data.ComplianceStatus || '';
  if (complianceStatus.includes('Compliant')) {
    complianceCell.setBackground('#d4edda');
    complianceCell.setFontColor('#155724');
  } else if (complianceStatus.includes('Non-Compliant')) {
    complianceCell.setBackground('#f8d7da');
    complianceCell.setFontColor('#721c24');
  } else {
    complianceCell.setBackground('#fff3cd');
    complianceCell.setFontColor('#856404');
  }
  
  // Add border to the new row
  sheet.getRange(lastRow, 1, 1, row.length).setBorder(true, true, true, true, true, true);
}

// Function to handle notifications
function handleNotifications(data) {
  const complianceStatus = data.ComplianceStatus || '';
  
  // Send email for non-compliant systems
  if (complianceStatus.includes('Non-Compliant')) {
    sendComplianceNotification(data);
  }
  
  // Send email for postponed reboots
  if (data.RebootAction === 'Postponed') {
    sendPostponeNotification(data);
  }
  
  // Send summary email if it's end of day (optional)
  if (isEndOfDay()) {
    sendDailySummary();
  }
}

// Function to send compliance notification
function sendComplianceNotification(data) {
  const subject = `Windows Update Compliance Alert - ${data.ComputerName}`;
  const body = `
    <h3>Windows Update Compliance Alert</h3>
    <p><strong>Computer:</strong> ${data.ComputerName}</p>
    <p><strong>User:</strong> ${data.UserName}</p>
    <p><strong>Email:</strong> ${data.UserEmail}</p>
    <p><strong>Status:</strong> ${data.ComplianceStatus}</p>
    <p><strong>Action:</strong> ${data.RebootAction}</p>
    <p><strong>Reason:</strong> ${data.PostponeReason}</p>
    <p><strong>Timestamp:</strong> ${new Date(data.Timestamp).toLocaleString()}</p>
    
    <p>Please follow up with the user to ensure compliance.</p>
  `;
  
  MailApp.sendEmail({
    to: COMPLIANCE_EMAIL,
    subject: subject,
    htmlBody: body
  });
}

// Function to send postpone notification
function sendPostponeNotification(data) {
  const subject = `Reboot Postponed - ${data.ComputerName}`;
  const body = `
    <h3>Windows Update Reboot Postponed</h3>
    <p><strong>Computer:</strong> ${data.ComputerName}</p>
    <p><strong>User:</strong> ${data.UserName} (${data.UserEmail})</p>
    <p><strong>Reason:</strong> ${data.PostponeReason}</p>
    <p><strong>Postponed Until:</strong> ${new Date(data.PostponeDetails?.PostponeTime).toLocaleString()}</p>
    <p><strong>Updates Installed:</strong> ${data.UpdatesInstalled}</p>
    
    <p>The system will automatically reboot at the scheduled time.</p>
  `;
  
  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    htmlBody: body
  });
}

// Function to check if it's end of day
function isEndOfDay() {
  const now = new Date();
  const hour = now.getHours();
  return hour >= 17 && hour <= 18; // Between 5 PM and 6 PM
}

// Function to send daily summary
function sendDailySummary() {
  const sheet = getOrCreateSheet();
  const today = new Date();
  const startOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  
  // Get today's data
  const data = sheet.getDataRange().getValues();
  const todayData = data.filter(row => {
    const timestamp = new Date(row[0]);
    return timestamp >= startOfDay;
  });
  
  if (todayData.length === 0) return;
  
  const compliantCount = todayData.filter(row => row[14].includes('Compliant')).length;
  const nonCompliantCount = todayData.filter(row => row[14].includes('Non-Compliant')).length;
  const totalUpdates = todayData.reduce((sum, row) => sum + (row[9] || 0), 0);
  
  const subject = `Daily Windows Update Compliance Summary - ${today.toDateString()}`;
  const body = `
    <h3>Daily Windows Update Compliance Summary</h3>
    <p><strong>Date:</strong> ${today.toDateString()}</p>
    <p><strong>Total Systems Processed:</strong> ${todayData.length}</p>
    <p><strong>Compliant Systems:</strong> ${compliantCount}</p>
    <p><strong>Non-Compliant Systems:</strong> ${nonCompliantCount}</p>
    <p><strong>Total Updates Installed:</strong> ${totalUpdates}</p>
    <p><strong>Compliance Rate:</strong> ${((compliantCount / todayData.length) * 100).toFixed(1)}%</p>
    
    <p><a href="${sheet.getParent().getUrl()}">View Full Report</a></p>
  `;
  
  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    htmlBody: body
  });
}

// Function to create charts and dashboard
function createDashboard() {
  const sheet = getOrCreateSheet();
  const spreadsheet = sheet.getParent();
  
  // Create dashboard sheet
  let dashboardSheet = spreadsheet.getSheetByName('Dashboard');
  if (!dashboardSheet) {
    dashboardSheet = spreadsheet.insertSheet('Dashboard');
  }
  
  // Clear existing content
  dashboardSheet.clear();
  
  // Add dashboard title
  dashboardSheet.getRange('A1').setValue('Windows Update Compliance Dashboard');
  dashboardSheet.getRange('A1').setFontSize(18).setFontWeight('bold');
  dashboardSheet.getRange('A1:H1').merge();
  
  // Get data for analysis
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const records = data.slice(1);
  
  if (records.length === 0) {
    dashboardSheet.getRange('A3').setValue('No data available yet.');
    return;
  }
  
  // Calculate statistics
  const totalSystems = records.length;
  const compliantSystems = records.filter(row => row[14] && row[14].toString().includes('Compliant')).length;
  const nonCompliantSystems = records.filter(row => row[14] && row[14].toString().includes('Non-Compliant')).length;
  const pendingSystems = totalSystems - compliantSystems - nonCompliantSystems;
  const complianceRate = totalSystems > 0 ? (compliantSystems / totalSystems * 100).toFixed(1) : 0;
  const totalUpdatesInstalled = records.reduce((sum, row) => sum + (parseInt(row[9]) || 0), 0);
  const systemsRequiringReboot = records.filter(row => row[10] === true || row[10] === 'TRUE').length;
  
  // Add summary statistics
  dashboardSheet.getRange('A3').setValue('Summary Statistics');
  dashboardSheet.getRange('A3').setFontWeight('bold').setFontSize(14);
  
  const summaryData = [
    ['Total Systems:', totalSystems],
    ['Compliant Systems:', compliantSystems],
    ['Non-Compliant Systems:', nonCompliantSystems],
    ['Pending Systems:', pendingSystems],
    ['Compliance Rate:', complianceRate + '%'],
    ['Total Updates Installed:', totalUpdatesInstalled],
    ['Systems Requiring Reboot:', systemsRequiringReboot]
  ];
  
  dashboardSheet.getRange('A4:B10').setValues(summaryData);
  dashboardSheet.getRange('A4:A10').setFontWeight('bold');
  
  // Color code compliance rate
  const complianceRateCell = dashboardSheet.getRange('B8');
  if (parseFloat(complianceRate) >= 90) {
    complianceRateCell.setBackground('#d4edda').setFontColor('#155724');
  } else if (parseFloat(complianceRate) >= 70) {
    complianceRateCell.setBackground('#fff3cd').setFontColor('#856404');
  } else {
    complianceRateCell.setBackground('#f8d7da').setFontColor('#721c24');
  }
  
  // Add recent activity section
  dashboardSheet.getRange('A12').setValue('Recent Activity (Last 24 Hours)');
  dashboardSheet.getRange('A12').setFontWeight('bold').setFontSize(14);
  
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  
  const recentRecords = records.filter(row => {
    const timestamp = new Date(row[0]);
    return timestamp >= yesterday;
  });
  
  dashboardSheet.getRange('A13').setValue('Recent Systems Updated:');
  dashboardSheet.getRange('B13').setValue(recentRecords.length);
  dashboardSheet.getRange('A13').setFontWeight('bold');
  
  // Add top non-compliant systems
  dashboardSheet.getRange('A15').setValue('Non-Compliant Systems Requiring Attention');
  dashboardSheet.getRange('A15').setFontWeight('bold').setFontSize(14);
  
  const nonCompliantRecords = records.filter(row => 
    row[14] && row[14].toString().includes('Non-Compliant')
  ).slice(0, 10); // Top 10
  
  if (nonCompliantRecords.length > 0) {
    const nonCompliantHeaders = ['Computer Name', 'User', 'Status', 'Last Update'];
    dashboardSheet.getRange('A16:D16').setValues([nonCompliantHeaders]);
    dashboardSheet.getRange('A16:D16').setFontWeight('bold').setBackground('#f8f9fa');
    
    const nonCompliantData = nonCompliantRecords.map(row => [
      row[1], // Computer Name
      row[2], // User Name
      row[14], // Compliance Status
      row[0] // Timestamp
    ]);
    
    dashboardSheet.getRange(17, 1, nonCompliantData.length, 4).setValues(nonCompliantData);
  } else {
    dashboardSheet.getRange('A16').setValue('No non-compliant systems found!');
    dashboardSheet.getRange('A16').setFontColor('#155724');
  }
  
  // Add update trends (last 7 days)
  dashboardSheet.getRange('F3').setValue('Update Trends (Last 7 Days)');
  dashboardSheet.getRange('F3').setFontWeight('bold').setFontSize(14);
  
  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  
  const trendData = [];
  for (let i = 6; i >= 0; i--) {
    const date = new Date();
    date.setDate(date.getDate() - i);
    const dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const dayEnd = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1);
    
    const dayRecords = records.filter(row => {
      const timestamp = new Date(row[0]);
      return timestamp >= dayStart && timestamp < dayEnd;
    });
    
    trendData.push([
      date.toLocaleDateString(),
      dayRecords.length,
      dayRecords.filter(row => row[14] && row[14].toString().includes('Compliant')).length
    ]);
  }
  
  if (trendData.length > 0) {
    dashboardSheet.getRange('F4:H4').setValues([['Date', 'Total Updates', 'Compliant']]);
    dashboardSheet.getRange('F4:H4').setFontWeight('bold').setBackground('#f8f9fa');
    dashboardSheet.getRange(5, 6, trendData.length, 3).setValues(trendData);
  }
  
  // Auto-resize columns
  dashboardSheet.autoResizeColumns(1, 8);
  
  // Add last updated timestamp
  dashboardSheet.getRange('A1').setNote(`Last updated: ${new Date().toLocaleString()}`);
}

// Function to get compliance statistics for API calls
function getComplianceStats() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const records = data.slice(1); // Skip headers
  
  if (records.length === 0) {
    return {
      totalSystems: 0,
      compliantSystems: 0,
      nonCompliantSystems: 0,
      complianceRate: 0,
      lastUpdated: new Date().toISOString()
    };
  }
  
  const totalSystems = records.length;
  const compliantSystems = records.filter(row => row[14] && row[14].toString().includes('Compliant')).length;
  const nonCompliantSystems = records.filter(row => row[14] && row[14].toString().includes('Non-Compliant')).length;
  const complianceRate = (compliantSystems / totalSystems * 100).toFixed(1);
  
  return {
    totalSystems: totalSystems,
    compliantSystems: compliantSystems,
    nonCompliantSystems: nonCompliantSystems,
    complianceRate: parseFloat(complianceRate),
    lastUpdated: new Date().toISOString()
  };
}

// Function to manually trigger dashboard creation
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Compliance Tools')
    .addItem('Refresh Dashboard', 'createDashboard')
    .addItem('Send Summary Email', 'sendDailySummary')
    .addItem('Export Non-Compliant', 'exportNonCompliant')
    .addToUi();
}

// Function to export non-compliant systems
function exportNonCompliant() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const records = data.slice(1);
  
  const nonCompliantRecords = records.filter(row => 
    row[14] && row[14].toString().includes('Non-Compliant')
  );
  
  if (nonCompliantRecords.length === 0) {
    SpreadsheetApp.getUi().alert('No non-compliant systems found!');
    return;
  }
  
  // Create new sheet for export
  const spreadsheet = sheet.getParent();
  const exportSheet = spreadsheet.insertSheet('Non-Compliant Export');
  
  // Add headers and data
  exportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  exportSheet.getRange(2, 1, nonCompliantRecords.length, headers.length).setValues(nonCompliantRecords);
  
  // Format headers
  const headerRange = exportSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f8d7da');
  headerRange.setFontColor('#721c24');
  
  exportSheet.autoResizeColumns(1, headers.length);
  
  SpreadsheetApp.getUi().alert(`Export complete! ${nonCompliantRecords.length} non-compliant systems exported to new sheet.`);
}

// Function to clean up old data (optional - run monthly)
function cleanupOldData() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  
  // Filter out records older than 30 days
  const headers = data[0];
  const recentRecords = data.slice(1).filter(row => {
    const timestamp = new Date(row[0]);
    return timestamp >= thirtyDaysAgo;
  });
  
  // Clear sheet and add back recent data
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (recentRecords.length > 0) {
    sheet.getRange(2, 1, recentRecords.length, headers.length).setValues(recentRecords);
  }
  
  // Reapply formatting
  setupSheetHeaders(sheet);
  
  console.log(`Cleanup complete. Removed ${data.length - 1 - recentRecords.length} old records.`);
}
