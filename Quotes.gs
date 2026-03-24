/**
 * Quote Management Functions - COMPLETE & FIXED
 * Supports new structure with: Objective, Items (Model/Name/Category/QTY), 
 * Material Cost, Installation, Sales Tax, Payment Schedule
 */

function getNextQuoteNumber() {
  // Use lock to prevent race conditions when multiple users create quotes
  var lock = LockService.getScriptLock();
  try {
    // Wait up to 30 seconds for the lock
    lock.waitLock(30000);
    
    var ss = getSpreadsheet();
    var quotesSheet = ss.getSheetByName('Quotes');

    if (!quotesSheet) {
      return '100001';
    }

    var lastRow = quotesSheet.getLastRow();
    if (lastRow <= 1) {
      return '100001';
    }

    // Get ALL quote numbers and find the maximum
    var quoteNumbers = quotesSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var maxNumber = 100000; // Start from minimum
    
    for (var i = 0; i < quoteNumbers.length; i++) {
      var num = parseInt(quoteNumbers[i][0]);
      if (!isNaN(num) && num > maxNumber) {
        maxNumber = num;
      }
    }
    
    var nextNumber = maxNumber + 1;
    return nextNumber.toString();
    
  } catch (error) {
    Logger.log('Error getting next quote number: ' + error);
    return '100001';
  } finally {
    // Always release the lock
    lock.releaseLock();
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
      quoteData.objective || '',             // K: Objective
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
      taxType,                               // X: Tax Type (NEW)
      taxRate                                // Y: Tax Rate (NEW)
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
      taxType,                               // X: Tax Type (NEW)
      taxRate                                // Y: Tax Rate (NEW)
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
          taxRate: parseFloat(row[24]) || 0          // NEW - default to 0
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

// Generate PDF for quote
function generateQuotePDF(quoteNumber) {
  try {
    var result = getQuoteByNumber(quoteNumber);
    if (!result.success) {
      return result;
    }

    var quote = result.quote;
    
    // Create HTML content for PDF
    var html = '<html><body style="font-family: Arial, sans-serif;">';
    html += '<h1>QUOTATION</h1>';
    html += '<p><strong>Quote #:</strong> ' + quote.quoteNumber + '</p>';
    html += '<p><strong>Date:</strong> ' + new Date(quote.date).toLocaleDateString() + '</p>';
    html += '<p><strong>Customer:</strong> ' + quote.customerName + '</p>';
    html += '<p><strong>Company:</strong> ' + quote.customerCompany + '</p>';
    
    if (quote.objective) {
      html += '<h3>Objective</h3>';
      html += '<p>' + quote.objective + '</p>';
    }
    
    html += '<h3>Items</h3>';
    html += '<table border="1" cellpadding="5" style="border-collapse: collapse; width: 100%;">';
    html += '<tr><th>Item</th><th>Model Number</th><th>Item Name</th><th>Category</th><th>QTY</th></tr>';
    
    for (var i = 0; i < quote.items.length; i++) {
      var item = quote.items[i];
      html += '<tr>';
      html += '<td>' + (i + 1) + '</td>';
      html += '<td>' + (item.modelNumber || '') + '</td>';
      html += '<td>' + (item.itemName || '') + '</td>';
      html += '<td>' + (item.category || '') + '</td>';
      html += '<td>' + (item.qty || '') + '</td>';
      html += '</tr>';
    }
    
    html += '</table>';
    
    html += '<h3>Totals</h3>';
    html += '<p><strong>Material Cost:</strong> $' + quote.materialCost.toFixed(2) + '</p>';
    html += '<p><strong>Installation & Training:</strong> $' + quote.installationCost.toFixed(2) + '</p>';
    html += '<p><strong>Sub Total:</strong> $' + quote.subTotal.toFixed(2) + '</p>';
    html += '<p><strong>Sales Tax:</strong> $' + quote.salesTax.toFixed(2) + '</p>';
    html += '<p><strong>Grand Total:</strong> $' + quote.grandTotal.toFixed(2) + '</p>';
    html += '<h4>Payment Schedule</h4>';
    html += '<p><strong>Down Payment:</strong> $' + quote.downPayment.toFixed(2) + '</p>';
    html += '<p><strong>Final Payment:</strong> $' + quote.finalPayment.toFixed(2) + '</p>';
    
    html += '</body></html>';
    
    // Create PDF blob
    var blob = Utilities.newBlob(html, 'text/html', 'quote.html');
    var pdf = blob.getAs('application/pdf');
    
    return {
      success: true,
      pdf: Utilities.base64Encode(pdf.getBytes()),
      filename: 'Quote_' + quote.quoteNumber + '.pdf'
    };
  } catch (error) {
    return {success: false, message: 'Error generating PDF: ' + error.toString()};
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

/**
 * Create Invoice from Quote
 * Copies quote data to Invoices sheet with new invoice number
 */
function createInvoiceFromQuote(quoteNumber) {
  try {
    var ss = getSpreadsheet();
    var quotesSheet = ss.getSheetByName('Quotes');
    var invoicesSheet = ss.getSheetByName('Invoices');
    
    if (!quotesSheet) {
      return {success: false, message: 'Quotes sheet not found'};
    }
    
    if (!invoicesSheet) {
      return {success: false, message: 'Invoices sheet not found'};
    }
    
    // Find the quote
    var quotesData = quotesSheet.getDataRange().getValues();
    var quoteRow = -1;
    
    for (var i = 1; i < quotesData.length; i++) {
      if (quotesData[i][0].toString() === quoteNumber.toString()) {
        quoteRow = i;
        break;
      }
    }
    
    if (quoteRow === -1) {
      return {success: false, message: 'Quote not found'};
    }
    
    var quote = quotesData[quoteRow];
    
    // Check if quote is completed
    if (quote[20] !== 'Completed') {  // Column 20 is Status (U)
      return {success: false, message: 'Quote must be marked as Completed before creating invoice'};
    }
    
    // Get next invoice number
    var invoiceNumber = getNextInvoiceNumber();
    
    // Prepare invoice data (map quote columns to invoice columns)
    var invoiceData = [
      invoiceNumber,              // Invoice Number
      new Date(),                 // Invoice Date
      quote[3],                   // Customer ID
      quote[4],                   // Customer Name
      quote[5],                   // Company
      quote[6],                   // Address
      quote[7],                   // City
      quote[8],                   // Phone
      quote[9],                   // Email
      quote[10],                  // Objective
      quote[11],                  // Items JSON
      quote[12],                  // Material Cost
      quote[13],                  // Installation Cost
      quote[14],                  // Sub Total
      quote[15],                  // Sales Tax
      quote[16],                  // Other
      quote[17],                  // Down Payment
      quote[18],                  // Final Payment
      quote[19],                  // Grand Total
      'Unpaid',                   // Payment Status
      '',                         // Payment Date
      quote[3],                   // Prepared By (using Customer ID temporarily)
      new Date().toISOString(),   // Created At
      quote[24],                  // Tax Type (X)
      quote[25]                   // Tax Rate (Y)
    ];
    
    // Append to Invoices sheet
    invoicesSheet.appendRow(invoiceData);
    
    return {
      success: true, 
      message: 'Invoice created successfully',
      invoiceNumber: invoiceNumber
    };
    
  } catch (error) {
    Logger.log('Error creating invoice: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

/**
 * Get next invoice number
 */
function getNextInvoiceNumber() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    var ss = getSpreadsheet();
    var invoicesSheet = ss.getSheetByName('Invoices');
    
    if (!invoicesSheet) {
      return '200001';
    }
    
    var lastRow = invoicesSheet.getLastRow();
    if (lastRow <= 1) {
      return '200001';
    }
    
    // Get ALL invoice numbers and find the maximum
    var invoiceNumbers = invoicesSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var maxNumber = 200000;
    
    for (var i = 0; i < invoiceNumbers.length; i++) {
      var num = parseInt(invoiceNumbers[i][0]);
      if (!isNaN(num) && num > maxNumber) {
        maxNumber = num;
      }
    }
    
    var nextNumber = maxNumber + 1;
    return nextNumber.toString();
    
  } catch (error) {
    Logger.log('Error getting next invoice number: ' + error);
    return '200001';
  } finally {
    lock.releaseLock();
  }
}