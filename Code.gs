function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Tableau Dashboard QA Tool')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function uploadImage(imageData, fileName) {
  try {
    // Create a folder for QA images if it doesn't exist
    const folders = DriveApp.getFoldersByName('Tableau QA Images');
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder('Tableau QA Images');
    }
    
    // Convert base64 to blob
    const base64Data = imageData.split(',')[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', fileName);
    
    // Save to Drive
    const file = folder.createFile(blob);
    
    return {
      success: true,
      fileId: file.getId(),
      fileName: fileName,
      url: file.getUrl()
    };
  } catch (error) {
    console.error('Error uploading image:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

function saveQAResults(qaData) {
  try {
    // Create or get the QA Results spreadsheet
    const spreadsheetName = 'Tableau QA Results';
    let spreadsheet;
    
    const files = DriveApp.getFilesByName(spreadsheetName);
    if (files.hasNext()) {
      spreadsheet = SpreadsheetApp.openById(files.next().getId());
    } else {
      spreadsheet = SpreadsheetApp.create(spreadsheetName);
      const sheet = spreadsheet.getActiveSheet();
      
      // Create comprehensive headers for all QA checklist items
      const headers = [
        'Timestamp', 'Dashboard Name', 'Image File ID', 'Notes', 'QA Checkpoints',
        // Functionality & Interactivity
        'Filters Work', 'Dropdowns Update', 'Tooltips Accurate', 'Navigation Works', 'Date Filters Default',
        // Chart Accuracy  
        'Correct Data Points', 'No Unexpected Values', 'Measures Match', 'Totals Correct',
        // Visual Consistency
        'Legends Clean', 'Brand Colors', 'Consistent Fonts', 'No Orphan Colors', 'Null Values Handled',
        // Data Integrity
        'Numbers Match', 'Formatting Consistent', 'Data Refresh Verified'
      ];
      
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format the header row
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#3498db');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    const sheet = spreadsheet.getActiveSheet();
    const timestamp = new Date();
    
    // Format QA checkpoints data
    let checkpointData = '';
    if (qaData.qaCheckboxes && qaData.qaCheckboxes.length > 0) {
      checkpointData = qaData.qaCheckboxes.map(checkpoint => 
        `Checkpoint ${checkpoint.id}: ${checkpoint.checked ? 'PASSED' : 'FAILED'}${checkpoint.notes ? ' - ' + checkpoint.notes : ''}`
      ).join(' | ');
    }
    
    // Prepare the row data with all QA checklist results
    const rowData = [
      timestamp,
      qaData.dashboardName,
      qaData.imageFileId,
      qaData.notes,
      checkpointData,
      // Functionality & Interactivity
      qaData.filtersWork || false,
      qaData.dropdownsUpdate || false,
      qaData.tooltipsAccurate || false,
      qaData.navigationWorks || false,
      qaData.dateFiltersDefault || false,
      // Chart Accuracy
      qaData.correctDataPoints || false,
      qaData.noUnexpectedValues || false,
      qaData.measuresMatch || false,
      qaData.totalsCorrect || false,
      // Visual Consistency
      qaData.legendsClean || false,
      qaData.brandColors || false,
      qaData.consistentFonts || false,
      qaData.noOrphanColors || false,
      qaData.nullValuesHandled || false,
      // Data Integrity
      qaData.numbersMatch || false,
      qaData.formattingConsistent || false,
      qaData.dataRefreshVerified || false
    ];
    
    // Add the QA results
    sheet.appendRow(rowData);
    
    // Auto-resize columns for better readability
    sheet.autoResizeColumns(1, rowData.length);
    
    return {
      success: true,
      spreadsheetUrl: spreadsheet.getUrl()
    };
  } catch (error) {
    console.error('Error saving QA results:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}
