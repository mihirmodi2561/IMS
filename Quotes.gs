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

// Save quote
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

    // Calculate totals
    var materialCost = parseFloat(quoteData.materialCost) || 0;
    var installationCost = parseFloat(quoteData.installationCost) || 0;
    var salesTax = parseFloat(quoteData.salesTax) || 0;
    var downPayment = parseFloat(quoteData.downPayment) || 0;
    
    var subTotal = materialCost + installationCost;
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
      salesTax,                              // P: Sales Tax
      grandTotal,                            // Q: Grand Total
      downPayment,                           // R: Down Payment
      finalPayment,                          // S: Final Payment
      quoteData.preparedBy,                  // T: Prepared By
      quoteData.terms || 'Payment due within 30 days', // U: Terms
      createdAt,                             // V: Created At
      quoteData.status || 'Pending'          // W: Status
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

// Update existing quote
function updateQuote(quoteData) {
  try {
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

    // Convert items array to JSON
    var itemsJSON = JSON.stringify(quoteData.items || quoteData.lineItems);

    // Calculate totals
    var materialCost = parseFloat(quoteData.materialCost) || 0;
    var installationCost = parseFloat(quoteData.installationCost) || 0;
    var salesTax = parseFloat(quoteData.salesTax) || 0;
    var downPayment = parseFloat(quoteData.downPayment) || 0;
    
    var subTotal = materialCost + installationCost;
    var grandTotal = subTotal + salesTax;
    var finalPayment = grandTotal - downPayment;

    // Update the row (keep quote number, created at, and status)
    var rowData = [
      quoteData.quoteNumber,                 // A: Quote Number
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
      salesTax,                              // P: Sales Tax
      grandTotal,                            // Q: Grand Total
      downPayment,                           // R: Down Payment
      finalPayment,                          // S: Final Payment
      quoteData.preparedBy,                  // T: Prepared By
      quoteData.terms || 'Payment due within 30 days', // U: Terms
      data[rowIndex - 1][21],                // V: Keep original Created At
      data[rowIndex - 1][22]                 // W: Keep original Status
    ];

    quotesSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);

    return {
      success: true,
      message: 'Quote updated successfully',
      quoteNumber: quoteData.quoteNumber
    };
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get all quotes
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
            items = JSON.parse(row[11]);
          }
        } catch (e) {
          items = [];
        }

        quotes.push({
          quoteNumber: row[0],
          date: row[1],
          validUntil: row[2],
          customerId: row[3],
          customerName: row[4],
          customerCompany: row[5],
          customerAddress: row[6],
          customerCity: row[7],
          customerPhone: row[8],
          customerEmail: row[9],
          objective: row[10] || '',
          items: items,
          materialCost: parseFloat(row[12]) || 0,
          installationCost: parseFloat(row[13]) || 0,
          subTotal: parseFloat(row[14]) || 0,
          salesTax: parseFloat(row[15]) || 0,
          grandTotal: parseFloat(row[16]) || 0,
          downPayment: parseFloat(row[17]) || 0,
          finalPayment: parseFloat(row[18]) || 0,
          preparedBy: row[19],
          terms: row[20],
          createdAt: row[21],
          status: row[22]
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

// Setup Quotes sheet with new structure
function setupQuotesSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Quotes');
    
    if (!sheet) {
      sheet = ss.insertSheet('Quotes');
    }
    
    // Clear existing content
    sheet.clear();
    
    // Set headers
    var headers = [
      'Quote Number', 'Date', 'Valid Until', 'Customer ID', 'Customer Name',
      'Company', 'Address', 'City', 'Phone', 'Email', 'Objective',
      'Items JSON', 'Material Cost', 'Installation Cost', 'Sub Total',
      'Sales Tax', 'Grand Total', 'Down Payment', 'Final Payment',
      'Prepared By', 'Terms', 'Created At', 'Status'
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