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
      sheet.getRange('A1:I1').setValues([['Timestamp', 'Dashboard Name', 'Image File ID', 'Data Accuracy', 'Visual Consistency', 'Interactivity', 'Performance', 'Notes', 'QA Checkpoints']]);
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
    
    // Add the QA results
    sheet.appendRow([
      timestamp,
      qaData.dashboardName,
      qaData.imageFileId,
      qaData.dataAccuracy,
      qaData.visualConsistency,
      qaData.interactivity,
      qaData.performance,
      qaData.notes,
      checkpointData
    ]);
    
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
