/**
 * Quote Management Functions - COMPLETE & FIXED
 * Supports new structure with: Objective, Items (Model/Name/Category/QTY), 
 * Material Cost, Installation, Sales Tax, Payment Schedule
 */

function getNextQuoteNumber() {
  try {
    var ss = getSpreadsheet();
    var quotesSheet = ss.getSheetByName('Quotes');

    if (!quotesSheet) {
      return '100001';
    }

    var lastRow = quotesSheet.getLastRow();
    if (lastRow <= 1) {
      return '100001';
    }

    var lastQuoteNumber = quotesSheet.getRange(lastRow, 1).getValue();
    if (!lastQuoteNumber) {
      return '100001';
    }

    var nextNumber = parseInt(lastQuoteNumber) + 1;
    return nextNumber.toString();
  } catch (error) {
    return '100001';
  }
}

// Save quote - UPDATED WITH TAX SYSTEM
function saveQuote(quoteData) {
  try {
    var ss = getSpreadsheet();
    var quotesSheet = ss.getSheetByName('Quotes');

    if (!quotesSheet) {
      return {success: false, message: 'Quotes sheet not found'};
    }

    var quoteNumber = quoteData.quoteNumber || getNextQuoteNumber();
    var createdAt = new Date().toISOString();

    // Convert items array to JSON
    var itemsJSON = JSON.stringify(quoteData.items || quoteData.lineItems);
    
    // Convert services array to JSON (NEW)
    var servicesJSON = JSON.stringify(quoteData.services || []);

    // Get tax values
    var taxType = quoteData.taxType || 'Percentage';
    var taxRate = parseFloat(quoteData.taxRate) || 0;
    
    // Calculate totals
    var materialCost = parseFloat(quoteData.materialCost) || 0;
    var installationCost = parseFloat(quoteData.installationCost) || 0;
    var downPayment = parseFloat(quoteData.downPayment) || 0;
    
    var subTotal = materialCost + installationCost;
    
    // Calculate sales tax based on type
    var salesTax = 0;
    if (taxType === 'Exempt') {
      salesTax = 0;
    } else {
      salesTax = subTotal * (taxRate / 100);
    }
    
    var grandTotal = subTotal + salesTax;
    var finalPayment = grandTotal - downPayment;

    var rowData = [
      quoteNumber,                           // A: Quote Number
      quoteData.date,                        // B: Date
      quoteData.validUntil,                  // C: Valid Until
      quoteData.customerId || '',            // D: Customer ID
      quoteData.customerName,                // E: Customer Name
      quoteData.customerCompany || '',       // F: Company
      quoteData.customerAddress || '',       // G: Address
      quoteData.customerCity || '',          // H: City
      quoteData.customerPhone || '',         // I: Phone
      quoteData.customerEmail || '',         // J: Email
      quoteData.objective || '',             // K: Objective (NOW SUPPORTS HTML)
      itemsJSON,                             // L: Items JSON
      materialCost,                          // M: Material Cost
      installationCost,                      // N: Installation Cost
      subTotal,                              // O: Sub Total
      salesTax,                              // P: Sales Tax (calculated amount)
      grandTotal,                            // Q: Grand Total
      downPayment,                           // R: Down Payment
      finalPayment,                          // S: Final Payment
      quoteData.preparedBy,                  // T: Prepared By
      quoteData.terms || 'Payment due within 30 days', // U: Terms
      createdAt,                             // V: Created At
      quoteData.status || 'Pending',         // W: Status
      taxType,                               // X: Tax Type
      taxRate,                               // Y: Tax Rate
      servicesJSON                           // Z: Services JSON (NEW)
    ];

    quotesSheet.appendRow(rowData);

    return {
      success: true,
      message: 'Quote saved successfully',
      quoteNumber: quoteNumber
    };
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Update existing quote - UPDATED WITH TAX SYSTEM
function updateQuote(quoteData) {
  try {
    Logger.log('=== UPDATE QUOTE START ===');
    Logger.log('Quote Number: ' + quoteData.quoteNumber);
    
    var ss = getSpreadsheet();
    var quotesSheet = ss.getSheetByName('Quotes');

    if (!quotesSheet) {
      return {success: false, message: 'Quotes sheet not found'};
    }

    var data = quotesSheet.getDataRange().getValues();
    var rowIndex = -1;

    // Find the quote row
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(quoteData.quoteNumber)) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return {success: false, message: 'Quote not found'};
    }

    Logger.log('Found quote at row: ' + rowIndex);
    
    // Get original data to preserve customer info
    var originalRow = data[rowIndex - 1];
    
    // Convert items array to JSON
    var itemsJSON = JSON.stringify(quoteData.items || quoteData.lineItems || []);
    Logger.log('Items JSON: ' + itemsJSON);
    
    // Convert services array to JSON (NEW)
    var servicesJSON = JSON.stringify(quoteData.services || []);
    Logger.log('Services JSON: ' + servicesJSON);

    // Get tax values
    var taxType = quoteData.taxType || 'Percentage';
    var taxRate = parseFloat(quoteData.taxRate) || 0;
    
    // Calculate totals
    var materialCost = parseFloat(quoteData.materialCost) || 0;
    var installationCost = parseFloat(quoteData.installationCost) || 0;
    var downPayment = parseFloat(quoteData.downPayment) || 0;
    
    var subTotal = materialCost + installationCost;
    
    // Calculate sales tax based on type
    var salesTax = 0;
    if (taxType === 'Exempt') {
      salesTax = 0;
    } else {
      salesTax = subTotal * (taxRate / 100);
    }
    
    var grandTotal = subTotal + salesTax;
    var finalPayment = grandTotal - downPayment;
    
    Logger.log('Calculated: SubTotal=' + subTotal + ', TaxType=' + taxType + ', TaxRate=' + taxRate + ', SalesTax=' + salesTax + ', GrandTotal=' + grandTotal);

    // Update the row - PRESERVE CUSTOMER DATA from original
    var rowData = [
      quoteData.quoteNumber,                 // A: Quote Number
      quoteData.date || originalRow[1],      // B: Date
      quoteData.validUntil || originalRow[2], // C: Valid Until
      // CUSTOMER DATA - Use original values (READ-ONLY)
      originalRow[3],                        // D: Customer ID (from original)
      originalRow[4],                        // E: Customer Name (from original)
      originalRow[5],                        // F: Company (from original)
      originalRow[6],                        // G: Address (from original)
      originalRow[7],                        // H: City (from original)
      originalRow[8],                        // I: Phone (from original)
      originalRow[9],                        // J: Email (from original)
      // EDITABLE FIELDS
      quoteData.objective || '',             // K: Objective
      itemsJSON,                             // L: Items JSON
      materialCost,                          // M: Material Cost
      installationCost,                      // N: Installation Cost
      subTotal,                              // O: Sub Total
      salesTax,                              // P: Sales Tax (calculated)
      grandTotal,                            // Q: Grand Total
      downPayment,                           // R: Down Payment
      finalPayment,                          // S: Final Payment
      quoteData.preparedBy || originalRow[19], // T: Prepared By
      quoteData.terms || 'Payment due within 30 days', // U: Terms
      originalRow[21],                       // V: Keep original Created At
      originalRow[22],                       // W: Keep original Status
      taxType,                               // X: Tax Type
      taxRate,                               // Y: Tax Rate
      servicesJSON                           // Z: Services JSON (NEW)
    ];

    quotesSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    
    Logger.log('=== UPDATE QUOTE SUCCESS ===');

    return {
      success: true,
      message: 'Quote updated successfully',
      quoteNumber: quoteData.quoteNumber
    };
  } catch (error) {
    Logger.log('=== UPDATE QUOTE ERROR ===');
    Logger.log('Error: ' + error.toString());
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get all quotes - UPDATED TO INCLUDE TAX FIELDS
function getQuotes() {
  try {
    var ss = getSpreadsheet();
    var quotesSheet = ss.getSheetByName('Quotes');

    if (!quotesSheet) {
      return {success: false, message: 'Quotes sheet not found'};
    }

    var data = quotesSheet.getDataRange().getValues();
    var quotes = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        // Parse items JSON
        var items = [];
        try {
          if (row[11]) {
            items = JSON.parse(String(row[11]));
          }
        } catch (e) {
          items = [];
        }
        
        // Parse services JSON (NEW)
        var services = [];
        try {
          if (row[25]) {  // Column Z (index 25)
            services = JSON.parse(String(row[25]));
          }
        } catch (e) {
          services = [];
        }

        quotes.push({
          quoteNumber: String(row[0]),
          date: row[1] ? String(row[1]) : '',
          validUntil: row[2] ? String(row[2]) : '',
          customerId: String(row[3] || ''),
          customerName: String(row[4] || ''),
          customerCompany: String(row[5] || ''),
          customerAddress: String(row[6] || ''),
          customerCity: String(row[7] || ''),
          customerPhone: String(row[8] || ''),
          customerEmail: String(row[9] || ''),
          objective: String(row[10] || ''),
          items: items,
          materialCost: parseFloat(row[12]) || 0,
          installationCost: parseFloat(row[13]) || 0,
          subTotal: parseFloat(row[14]) || 0,
          salesTax: parseFloat(row[15]) || 0,
          grandTotal: parseFloat(row[16]) || 0,
          downPayment: parseFloat(row[17]) || 0,
          finalPayment: parseFloat(row[18]) || 0,
          preparedBy: String(row[19] || ''),
          terms: String(row[20] || ''),
          createdAt: row[21] ? String(row[21]) : '',
          status: String(row[22] || 'Pending'),
          taxType: String(row[23] || 'Percentage'),  // NEW - default to Percentage
          taxRate: parseFloat(row[24]) || 0,         // NEW - default to 0
          services: services                         // NEW - Services array
        });
      }
    }

    return {success: true, quotes: quotes};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get single quote by number
function getQuoteByNumber(quoteNumber) {
  try {
    var result = getQuotes();
    if (!result.success) {
      return result;
    }

    for (var i = 0; i < result.quotes.length; i++) {
      if (String(result.quotes[i].quoteNumber) === String(quoteNumber)) {
        return {success: true, quote: result.quotes[i]};
      }
    }

    return {success: false, message: 'Quote not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Update quote status
function updateQuoteStatus(quoteNumber, status) {
  try {
    var ss = getSpreadsheet();
    var quotesSheet = ss.getSheetByName('Quotes');

    if (!quotesSheet) {
      return {success: false, message: 'Quotes sheet not found'};
    }

    var data = quotesSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(quoteNumber)) {
        quotesSheet.getRange(i + 1, 23).setValue(status); // Column W (Status)
        return {success: true, message: 'Quote status updated successfully'};
      }
    }

    return {success: false, message: 'Quote not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Delete quote
function deleteQuote(quoteNumber) {
  try {
    var ss = getSpreadsheet();
    var quotesSheet = ss.getSheetByName('Quotes');

    if (!quotesSheet) {
      return {success: false, message: 'Quotes sheet not found'};
    }

    var data = quotesSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(quoteNumber)) {
        quotesSheet.deleteRow(i + 1);
        return {success: true, message: 'Quote deleted successfully'};
      }
    }

    return {success: false, message: 'Quote not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Setup Quotes sheet - UPDATED WITH NEW COLUMNS
function setupQuotesSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Quotes');
    
    if (!sheet) {
      sheet = ss.insertSheet('Quotes');
    }
    
    // Clear existing content
    sheet.clear();
    
    // Set headers INCLUDING NEW TAX COLUMNS
    var headers = [
      'Quote Number', 'Date', 'Valid Until', 'Customer ID', 'Customer Name',
      'Company', 'Address', 'City', 'Phone', 'Email', 'Objective',
      'Items JSON', 'Material Cost', 'Installation Cost', 'Sub Total',
      'Sales Tax', 'Grand Total', 'Down Payment', 'Final Payment',
      'Prepared By', 'Terms', 'Created At', 'Status',
      'Tax Type', 'Tax Rate'  // NEW COLUMNS
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    
    return {success: true, message: 'Quotes sheet setup complete'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}