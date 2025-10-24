// Google Apps Script Code for House History Form
// This code should be deployed as a web app in Google Apps Script

function doPost(e) {
  try {
    // Parse the incoming data
    const data = JSON.parse(e.postData.contents);
    
    // Get or create the spreadsheet
    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = getOrCreateSheet(spreadsheet, 'House History Submissions');
    
    // Set up headers if this is the first submission
    setupHeaders(sheet);
    
    // Add the data to the sheet
    addDataToSheet(sheet, data);
    
    // Handle file attachments if present
    if (data.photos && data.photos.length > 0) {
      savePhotosToFolder(data, 'House History Photos');
    }
    
    if (data.documents && data.documents.length > 0) {
      saveDocumentsToFolder(data, 'House History Documents');
    }
    
    // Send confirmation email (optional)
    sendConfirmationEmail(data);
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: 'Data saved successfully' }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader("Access-Control-Allow-Origin", "*")
      .setHeader("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
      .setHeader("Access-Control-Allow-Headers", "Content-Type");
      
  } catch (error) {
    console.error('Error processing submission:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader("Access-Control-Allow-Origin", "*")
      .setHeader("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
      .setHeader("Access-Control-Allow-Headers", "Content-Type");
  }
}

// Add OPTIONS handler for CORS preflight requests
function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}

function getOrCreateSpreadsheet() {
  const SPREADSHEET_NAME = 'House History Form Submissions';
  
  // Try to find existing spreadsheet
  const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) {
    const file = files.next();
    return SpreadsheetApp.openById(file.getId());
  }
  
  // Create new spreadsheet if it doesn't exist
  const spreadsheet = SpreadsheetApp.create(SPREADSHEET_NAME);
  
  // Share with yourself (optional - adjust permissions as needed)
  const email = Session.getActiveUser().getEmail();
  if (email) {
    spreadsheet.addEditor(email);
  }
  
  return spreadsheet;
}

function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

function setupHeaders(sheet) {
  // Check if headers already exist
  if (sheet.getLastRow() === 0) {
    const headers = [
      'Timestamp',
      'Name',
      'Email',
      'Phone',
      'Address',
      'Postcode',
      'Position in Town',
      'Property Description',
      'Special Features',
      'Listed Building',
      'Why Interested',
      'Help Needed',
      'Existing History',
      'Photo Count',
      'Document Count',
      'Submission ID'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
  }
}

function addDataToSheet(sheet, data) {
  // Generate unique submission ID
  const submissionId = Utilities.getUuid().substring(0, 8);
  
  const rowData = [
    new Date(data.timestamp),
    data.name || '',
    data.email || '',
    data.phone || '',
    data.address || '',
    data.postcode || '',
    data.position || '',
    data.description || '',
    data.specialFeatures || '',
    data.listedBuilding || '',
    data.whyInterested || '',
    data.helpNeeded || '',
    data.existingHistory || '',
    data.photoCount || 0,
    data.documentCount || 0,
    submissionId
  ];
  
  // Add row to sheet
  sheet.appendRow(rowData);
  
  // Format the new row
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(lastRow, 1, 1, rowData.length);
  
  // Alternate row colors
  if (lastRow % 2 === 0) {
    range.setBackground('#f8f9fa');
  }
  
  return submissionId;
}

function savePhotosToFolder(data, folderName) {
  try {
    // Get or create photos folder
    const folder = getOrCreateFolder(folderName);
    
    // Create subfolder for this submission
    const submissionFolder = folder.createFolder(`${data.name}_${new Date(data.timestamp).toISOString().split('T')[0]}`);
    
    // Save each photo
    data.photos.forEach((photo, index) => {
      try {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(photo.data),
          photo.type,
          photo.name || `photo_${index + 1}.jpg`
        );
        submissionFolder.createFile(blob);
      } catch (error) {
        console.error(`Error saving photo ${index}:`, error);
      }
    });
    
    return submissionFolder.getUrl();
  } catch (error) {
    console.error('Error saving photos:', error);
    return null;
  }
}

function saveDocumentsToFolder(data, folderName) {
  try {
    // Get or create documents folder
    const folder = getOrCreateFolder(folderName);
    
    // Create subfolder for this submission
    const submissionFolder = folder.createFolder(`${data.name}_${new Date(data.timestamp).toISOString().split('T')[0]}_docs`);
    
    // Save each document
    data.documents.forEach((doc, index) => {
      try {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(doc.data),
          doc.type,
          doc.name || `document_${index + 1}`
        );
        submissionFolder.createFile(blob);
      } catch (error) {
        console.error(`Error saving document ${index}:`, error);
      }
    });
    
    return submissionFolder.getUrl();
  } catch (error) {
    console.error('Error saving documents:', error);
    return null;
  }
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function sendConfirmationEmail(data) {
  try {
    if (!data.email) return;
    
    const subject = 'House History Form Submission Received - Bishops Waltham Heritage';
    const body = `
Dear ${data.name},

Thank you for submitting your house history form for the Bishops Waltham Community Heritage Project.

Property Details:
- Address: ${data.address}
- Postcode: ${data.postcode}

We have received your submission and will review it for inclusion on our community heritage website. 

If you have any questions, please don't hesitate to contact us.

Best regards,
Bishops Waltham Heritage Team

---
This is an automated confirmation email.
    `;
    
    MailApp.sendEmail({
      to: data.email,
      subject: subject,
      body: body
    });
    
  } catch (error) {
    console.error('Error sending confirmation email:', error);
  }
}

// Optional: Function to get all submissions (for admin use)
function getAllSubmissions() {
  const spreadsheet = getOrCreateSpreadsheet();
  const sheet = spreadsheet.getSheetByName('House History Submissions');
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

// Optional: Function to export data as CSV
function exportToCSV() {
  const spreadsheet = getOrCreateSpreadsheet();
  const sheet = spreadsheet.getSheetByName('House History Submissions');
  
  const data = sheet.getDataRange().getValues();
  let csv = '';
  
  data.forEach(row => {
    csv += row.map(cell => `"${cell}"`).join(',') + '\n';
  });
  
  const blob = Utilities.newBlob(csv, 'text/csv', 'house_history_submissions.csv');
  DriveApp.createFile(blob);
}