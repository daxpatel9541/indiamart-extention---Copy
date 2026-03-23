// =====================================================================
// Google Apps Script — Paste this code in your Google Sheet's Apps Script
// =====================================================================
//
// SETUP INSTRUCTIONS:
// 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1_W38dXsP6TQVkpPdspO9qK-_1URfO_qMDjFcUcYVUd0/edit
// 2. Go to Extensions → Apps Script
// 3. Delete any existing code and paste this entire file's contents
// 4. Click "Deploy" → "New deployment"
// 5. Select Type: "Web app"
// 6. Set "Execute as": Me
// 7. Set "Who has access": Anyone
// 8. Click "Deploy"
// 9. Copy the Web App URL
// 10. Paste that URL into the extension popup's "Apps Script URL" field
//
// =====================================================================

function doPost(e) {
  Logger.log('Received a request');
  try {
    if (!e || !e.postData || !e.postData.contents) {
      Logger.log('Error: Empty POST body');
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Empty body' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var data = JSON.parse(e.postData.contents);
    var leads = data.leads;
    Logger.log('Parsed ' + (leads ? leads.length : 0) + ' leads');

    if (!leads || leads.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'No leads provided' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Leads') || ss.insertSheet('Leads');

    // Add headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      var headers = [
        'S.No',
        'Company Name',
        'Mobile Number',
        'Contact Person',
        'Address',
        'Rating',
        'Review Count',
        'Review 1',
        'Review 2',
        'Response Rate',
        'Quality Rate',
        'Delivery Rate',
        'Business Type',
        'Company Owner',
        'Total Employees',
        'Year of Establishment',
        'IndiaMART Member Since',
        'Annual Turnover',
        'GST No.',
        'PAN',
        'Sells / Products',
        'Unique ID',
        'Extracted Date'
      ];
      sheet.appendRow(headers);

      // Style headers
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#1a73e8');
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    // Get existing unique IDs to avoid duplicates
    var lastRow = sheet.getLastRow();
    var existingIds = [];
    if (lastRow > 1) {
      var idColumn = 22; // Column V = Unique ID (was U)
      var idRange = sheet.getRange(2, idColumn, lastRow - 1, 1);
      existingIds = idRange.getValues().flat().map(String);
    }

    var addedCount = 0;
    var skippedCount = 0;
    var nextSNo = lastRow; // Since row 1 is header

    for (var i = 0; i < leads.length; i++) {
      var lead = leads[i];
      var uniqueId = lead.uniqueId || '';

      // Skip duplicates
      if (uniqueId && existingIds.indexOf(uniqueId) !== -1) {
        skippedCount++;
        continue;
      }

      nextSNo++;
      var row = [
        nextSNo - 1,
        lead.companyName || '',
        lead.mobileNumber || '',
        lead.contactPerson || '',
        lead.address || '',
        lead.rating || '',
        lead.reviewCount || '',
        lead.review1 || '',
        lead.review2 || '',
        lead.responseRate || '',
        lead.qualityRate || '',
        lead.deliveryRate || '',
        lead.businessType || '',
        lead.ownerName || '',
        lead.totalEmployees || '',
        lead.yearEstablished || '',
        lead.memberSince || '',
        lead.annualTurnover || '',
        lead.gstNumber || '',
        lead.panNumber || '',
        lead.products || '',
        uniqueId,
        new Date().toLocaleString('en-IN')
      ];

      sheet.appendRow(row);
      existingIds.push(uniqueId);
      addedCount++;
    }

    // Auto-resize columns
    try {
      for (var col = 1; col <= 23; col++) {
        sheet.autoResizeColumn(col);
      }
    } catch (e) { /* ignore resize errors */ }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      added: addedCount,
      skipped: skippedCount,
      total: sheet.getLastRow() - 1
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Test function - run this to verify the script works
function testDoPost() {
  var testData = {
    postData: {
      contents: JSON.stringify({
        leads: [{
          companyName: 'Test Company',
          mobileNumber: '9988776655',
          contactPerson: 'Test Person',
          address: 'Test Address, City, State',
          rating: '4.5',
          reviewCount: '10',
          review1: 'Great service',
          review2: 'Good quality',
          responseRate: '100%',
          qualityRate: '95%',
          deliveryRate: '98%',
          businessType: 'Manufacturer',
          ownerName: 'Test Owner',
          totalEmployees: '50',
          yearEstablished: '2020',
          memberSince: '3 Years',
          annualTurnover: 'Rs. 1 Crore',
          gstNumber: 'TEST12345',
          panNumber: 'ABCDE1234F',
          products: 'Test Products',
          uniqueId: 'test_12345'
        }]
      })
    }
  };

  var result = doPost(testData);
  Logger.log(result.getContent());
}
