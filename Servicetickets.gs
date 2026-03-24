/**
 * SERVICE TICKETS MANAGEMENT
 * For field service, maintenance, and repair work
 * 
 * Flow: Create Ticket → Admin Approval → Convert to Quote → Invoice
 * Inventory is NOT deducted until invoice is created
 */

// ========================================
// HELPER FUNCTION
// ========================================

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ========================================
// GENERATE TICKET NUMBER
// ========================================

function getNextTicketNumber() {
  try {
    var ss = getSpreadsheet();
    var ticketsSheet = ss.getSheetByName('Service Tickets');
    
    if (!ticketsSheet) {
      return 'ST-100001';
    }
    
    var lastRow = ticketsSheet.getLastRow();
    if (lastRow <= 1) {
      return 'ST-100001';
    }
    
    var lastTicketNumber = ticketsSheet.getRange(lastRow, 1).getValue();
    if (!lastTicketNumber || lastTicketNumber === '') {
      return 'ST-100001';
    }
    
    // Extract number from ST-XXXXXX format
    var numberPart = lastTicketNumber.toString().replace('ST-', '');
    var nextNumber = parseInt(numberPart) + 1;
    
    // Pad with zeros
    var paddedNumber = String(nextNumber).padStart(6, '0');
    return 'ST-' + paddedNumber;
    
  } catch (error) {
    Logger.log('Error generating ticket number: ' + error);
    return 'ST-100001';
  }
}

// ========================================
// GET SERVICE TICKETS
// ========================================

function getServiceTickets(filterStatus) {
  Logger.log("=== getServiceTickets called ===");
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ticketsSheet = ss.getSheetByName('Service Tickets');
    
    if (!ticketsSheet) {
      return {success: false, message: 'Service Tickets sheet not found'};
    }
    
    var lastRow = ticketsSheet.getLastRow();
    
    if (lastRow <= 1) {
      return {success: true, tickets: []};
    }
    
    var data = ticketsSheet.getRange(1, 1, lastRow, 19).getValues();
    var tickets = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      if (!row[0] || row[0] === '') {
        continue;
      }
      
      // ✅ FIXED: Parse items from correct column (L = index 11)
      var items = [];
      if (row[11]) {  // Changed from row[7] to row[11]
        try {
          items = JSON.parse(String(row[11]));
        } catch (e) {
          Logger.log('Error parsing items: ' + e);
          items = [];
        }
      }
      
      // ✅ FIXED: Correct column mapping
      var ticket = {
        ticketNumber: String(row[0]),           // A
        date: row[1] ? String(row[1]) : '',     // B
        technicianName: String(row[2] || ''),   // C
        customerId: String(row[3] || ''),       // D
        customerName: String(row[4] || ''),     // E
        customerCompany: String(row[5] || ''),  // F
        customerAddress: String(row[6] || ''),  // G
        customerCity: String(row[7] || ''),     // H
        customerPhone: String(row[8] || ''),    // I
        customerEmail: String(row[9] || ''),    // J
        problemType: String(row[10] || ''),     // K
        items: items,                            // L (row[11])
        solution: String(row[12] || ''),        // M
        status: String(row[13] || 'Pending'),   // N
        createdAt: row[14] ? String(row[14]) : '', // O
        approvedBy: String(row[15] || ''),      // P
        approvedAt: row[16] ? String(row[16]) : '', // Q
        quoteNumber: String(row[17] || ''),     // R
        rejectionReason: String(row[18] || '')  // S
      };
      
      tickets.push(ticket);
    }
    
    Logger.log('Returning ' + tickets.length + ' tickets');
    return {success: true, tickets: tickets};
    
  } catch (error) {
    Logger.log('Error: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// CREATE SERVICE TICKET
// ========================================

function createServiceTicket(ticketData) {
  try {
    Logger.log('Creating service ticket...');
    Logger.log('Received data: ' + JSON.stringify(ticketData));
    
    // ADD THESE SPECIFIC LOGS:
    Logger.log('Customer Address: ' + ticketData.customerAddress);
    Logger.log('Customer City: ' + ticketData.customerCity);
    Logger.log('Customer Phone: ' + ticketData.customerPhone);
    Logger.log('Customer Email: ' + ticketData.customerEmail);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ticketsSheet = ss.getSheetByName('Service Tickets');
    
    if (!ticketsSheet) {
      return {success: false, message: 'Service Tickets sheet not found. Please run setupServiceTicketsSheet first.'};
    }
    
    if (!ticketData || !ticketData.customerId) {
      return {success: false, message: 'Customer is required'};
    }
    
    if (!ticketData.problemType) {
      return {success: false, message: 'Problem type is required'};
    }
    
    var ticketNumber = getNextTicketNumber();
    var createdAt = new Date().toISOString();
    var itemsJSON = JSON.stringify(ticketData.items || []);
    
    // Log the rowData before saving
    Logger.log('About to save - Address: ' + (ticketData.customerAddress || 'EMPTY'));
    Logger.log('About to save - City: ' + (ticketData.customerCity || 'EMPTY'));
    Logger.log('About to save - Phone: ' + (ticketData.customerPhone || 'EMPTY'));
    Logger.log('About to save - Email: ' + (ticketData.customerEmail || 'EMPTY'));
    
    var rowData = [
      ticketNumber,
      ticketData.date || createdAt,
      ticketData.technicianName || '',
      ticketData.customerId || '',
      ticketData.customerName || '',
      ticketData.customerCompany || '',
      ticketData.customerAddress || '',      // G
      ticketData.customerCity || '',         // H
      ticketData.customerPhone || '',        // I
      ticketData.customerEmail || '',        // J
      ticketData.problemType || '',
      itemsJSON,
      ticketData.solution || '',
      'Pending',
      createdAt,
      '',
      '',
      '',
      ''
    ];
    
    Logger.log('RowData length: ' + rowData.length);
    Logger.log('RowData[6] (Address): ' + rowData[6]);
    Logger.log('RowData[7] (City): ' + rowData[7]);
    Logger.log('RowData[8] (Phone): ' + rowData[8]);
    Logger.log('RowData[9] (Email): ' + rowData[9]);
    
    ticketsSheet.appendRow(rowData);
    
    Logger.log('Service ticket created: ' + ticketNumber);
    
    return {
      success: true,
      message: 'Service ticket created successfully',
      ticketNumber: ticketNumber
    };
    
  } catch (error) {
    Logger.log('Error creating service ticket: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// GET TICKET BY NUMBER
// ========================================

function getServiceTicketByNumber(ticketNumber) {
  try {
    var result = getServiceTickets();
    
    if (!result.success) {
      return result;
    }
    
    for (var i = 0; i < result.tickets.length; i++) {
      if (String(result.tickets[i].ticketNumber) === String(ticketNumber)) {
        return {success: true, ticket: result.tickets[i]};
      }
    }
    
    return {success: false, message: 'Service ticket not found'};
    
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// APPROVE SERVICE TICKET
// ========================================

function approveServiceTicket(ticketNumber, approvedBy) {
  try {
    Logger.log('Approving service ticket: ' + ticketNumber);
    
    // Get ticket details
    var ticketResult = getServiceTicketByNumber(ticketNumber);
    if (!ticketResult.success) {
      return ticketResult;
    }
    
    var ticket = ticketResult.ticket;
    
    // Check if already approved
    if (ticket.status !== 'Pending') {
      return {success: false, message: 'Ticket has already been processed (Status: ' + ticket.status + ')'};
    }
    
    // Create quote from ticket
    var quoteData = {
      customerId: ticket.customerId,
      customerName: ticket.customerName,
      customerCompany: ticket.customerCompany,
      customerAddress: '',
      customerCity: '',
      customerPhone: '',
      customerEmail: '',
      lineItems: ticket.items,
      objective: 'Service Ticket: ' + ticketNumber + ' - ' + ticket.problemType,
      materialCost: 0,
      installationCost: 0,
      salesTax: 0,
      downPayment: 0,
      terms: 'Converted from Service Ticket ' + ticketNumber,
      preparedBy: approvedBy
    };
    
    // Call saveQuote function (from Quotes.gs)
    var quoteResult = saveQuote(quoteData);
    
    if (!quoteResult.success) {
      return {success: false, message: 'Failed to create quote: ' + quoteResult.message};
    }
    
    var quoteNumber = quoteResult.quoteNumber;
    
    // Update ticket status
    var ss = getSpreadsheet();
    var ticketsSheet = ss.getSheetByName('Service Tickets');
    var data = ticketsSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(ticketNumber)) {
        var rowIndex = i + 1;
        ticketsSheet.getRange(rowIndex, 10).setValue('Converted');
        ticketsSheet.getRange(rowIndex, 12).setValue(approvedBy);
        ticketsSheet.getRange(rowIndex, 13).setValue(new Date().toISOString());
        ticketsSheet.getRange(rowIndex, 14).setValue(quoteNumber);
        break;
      }
    }
    
    Logger.log('Service ticket approved and converted to quote: ' + quoteNumber);
    
    return {
      success: true,
      message: 'Service ticket approved and converted to quote',
      quoteNumber: quoteNumber
    };
    
  } catch (error) {
    Logger.log('Error approving service ticket: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// REJECT SERVICE TICKET
// ========================================

function rejectServiceTicket(ticketNumber, rejectionReason) {
  try {
    Logger.log('Rejecting service ticket: ' + ticketNumber);
    
    var ss = getSpreadsheet();
    var ticketsSheet = ss.getSheetByName('Service Tickets');
    
    if (!ticketsSheet) {
      return {success: false, message: 'Service Tickets sheet not found'};
    }
    
    var data = ticketsSheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(ticketNumber)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return {success: false, message: 'Service ticket not found'};
    }
    
    ticketsSheet.getRange(rowIndex, 10).setValue('Rejected');
    ticketsSheet.getRange(rowIndex, 15).setValue(rejectionReason);
    
    Logger.log('Service ticket rejected');
    
    return {success: true, message: 'Service ticket rejected'};
    
  } catch (error) {
    Logger.log('Error rejecting service ticket: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// DELETE SERVICE TICKET
// ========================================

function deleteServiceTicket(ticketNumber) {
  try {
    Logger.log('Deleting service ticket: ' + ticketNumber);
    
    var ss = getSpreadsheet();
    var ticketsSheet = ss.getSheetByName('Service Tickets');
    
    if (!ticketsSheet) {
      return {success: false, message: 'Service Tickets sheet not found'};
    }
    
    var data = ticketsSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(ticketNumber)) {
        ticketsSheet.deleteRow(i + 1);
        Logger.log('Service ticket deleted');
        return {success: true, message: 'Service ticket deleted successfully'};
      }
    }
    
    return {success: false, message: 'Service ticket not found'};
    
  } catch (error) {
    Logger.log('Error deleting service ticket: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// SETUP SHEET
// ========================================

// UPDATED setupServiceTicketsSheet function
// Run this to update your sheet with new columns

function setupServiceTicketsSheet() {
  try {
    Logger.log('Setting up Service Tickets sheet...');
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Service Tickets');
    
    if (!sheet) {
      Logger.log('Creating new Service Tickets sheet...');
      sheet = ss.insertSheet('Service Tickets');
    }
    
    // Headers - 19 columns (A-S) - UPDATED WITH NEW FIELDS
    var headers = [
      'Ticket Number',        // A
      'Date',                 // B
      'Technician Name',      // C
      'Customer ID',          // D
      'Customer Name',        // E
      'Customer Company',     // F
      'Customer Address',     // G (NEW)
      'Customer City',        // H (NEW)
      'Customer Phone',       // I (NEW)
      'Customer Email',       // J (NEW)
      'Problem Type',         // K
      'Items (JSON)',         // L
      'Solution/Work Done',   // M
      'Status',               // N
      'Created At',           // O
      'Approved By',          // P
      'Approved At',          // Q
      'Quote Number',         // R
      'Rejection Reason'      // S
    ];
    
    // Set headers
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    
    // Format header
    headerRange.setBackground('#001f3f');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    
    // Set column widths
    sheet.setColumnWidth(1, 120);  // Ticket Number
    sheet.setColumnWidth(2, 100);  // Date
    sheet.setColumnWidth(3, 150);  // Technician Name
    sheet.setColumnWidth(4, 100);  // Customer ID
    sheet.setColumnWidth(5, 150);  // Customer Name
    sheet.setColumnWidth(6, 150);  // Customer Company
    sheet.setColumnWidth(7, 200);  // Customer Address (NEW)
    sheet.setColumnWidth(8, 120);  // Customer City (NEW)
    sheet.setColumnWidth(9, 120);  // Customer Phone (NEW)
    sheet.setColumnWidth(10, 180); // Customer Email (NEW)
    sheet.setColumnWidth(11, 150); // Problem Type
    sheet.setColumnWidth(12, 300); // Items (JSON)
    sheet.setColumnWidth(13, 300); // Solution/Work Done
    sheet.setColumnWidth(14, 100); // Status
    sheet.setColumnWidth(15, 150); // Created At
    sheet.setColumnWidth(16, 120); // Approved By
    sheet.setColumnWidth(17, 150); // Approved At
    sheet.setColumnWidth(18, 120); // Quote Number
    sheet.setColumnWidth(19, 200); // Rejection Reason
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Add data validation for Status column (column 14 = N)
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Approved', 'Rejected', 'Converted'], true)
      .setAllowInvalid(false)
      .build();
    
    sheet.getRange(2, 14, sheet.getMaxRows() - 1, 1).setDataValidation(statusRule);
    
    Logger.log('Service Tickets sheet setup complete!');
    
    return {
      success: true,
      message: 'Service Tickets sheet setup complete with ' + headers.length + ' columns'
    };
    
  } catch (error) {
    Logger.log('Error setting up Service Tickets sheet: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// TEST FUNCTION
// ========================================

function testCreateServiceTicket() {
  var testData = {
    date: '2026-02-24',
    technicianName: 'Test Technician',
    customerId: '123',
    customerName: 'Test Customer',
    customerCompany: 'Test Company',
    problemType: 'Camera Not Working',
    solution: 'Replaced camera module',
    items: [
      {
        itemNumber: 1,
        modelNumber: 'CAM-001',
        itemName: 'Security Camera',
        category: 'Cameras',
        qty: 2
      }
    ]
  };
  
  var result = createServiceTicket(testData);
  Logger.log('Test Result: ' + JSON.stringify(result));
  
  if (result.success) {
    Logger.log('SUCCESS! Ticket Number: ' + result.ticketNumber);
  } else {
    Logger.log('FAILED: ' + result.message);
  }
  
  return result;
}