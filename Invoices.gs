/**
 * Invoice Management Functions
 * Handles invoice creation from quotes, stock management, and status updates
 */

function getNextInvoiceNumber() {
  try {
    var ss = getSpreadsheet();
    var invoicesSheet = ss.getSheetByName('Invoices');

    if (!invoicesSheet) {
      return 'INV-100001';
    }

    var data = invoicesSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return 'INV-100001';
    }

    var lastInvoiceNumber = data[data.length - 1][0];
    var numberPart = parseInt(lastInvoiceNumber.split('-')[1]) || 100000;
    return 'INV-' + (numberPart + 1);
  } catch (error) {
    return 'INV-100001';
  }
}

// Check if invoice already exists for a quote
function checkInvoiceExistsForQuote(quoteNumber) {
  try {
    Logger.log('🔍 Checking if invoice exists for quote: ' + quoteNumber);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invoicesSheet = ss.getSheetByName('Invoices');

    if (!invoicesSheet) {
      Logger.log('❌ Invoices sheet not found');
      return null;
    }

    var data = invoicesSheet.getDataRange().getValues();
    Logger.log('📄 Checking ' + (data.length - 1) + ' invoices');

    for (var i = 1; i < data.length; i++) {
      // Quote Number is at column X (index 23)
      var invoiceQuoteNum = String(data[i][23] || '');
      var searchQuoteNum = String(quoteNumber);
      
      Logger.log('Comparing: Invoice Quote#=' + invoiceQuoteNum + ' vs Search=' + searchQuoteNum);
      
      if (invoiceQuoteNum === searchQuoteNum) {
        var invoiceNumber = data[i][0];
        Logger.log('✅ Found existing invoice: ' + invoiceNumber + ' for quote: ' + quoteNumber);
        return invoiceNumber; // Return invoice number
      }
    }

    Logger.log('✅ No existing invoice found for quote: ' + quoteNumber);
    return null;
  } catch (error) {
    Logger.log('❌ Error checking invoice for quote: ' + error.toString());
    return null;
  }
}

// Check stock availability for line items
function checkStockAvailability(lineItems) {
  try {
    var ss = getSpreadsheet();
    var stocksSheet = ss.getSheetByName('Stocks');

    if (!stocksSheet) {
      return {success: false, message: 'Stocks sheet not found'};
    }

    var stockData = stocksSheet.getDataRange().getValues();
    var unavailableItems = [];

    // Check each line item
    for (var i = 0; i < lineItems.length; i++) {
      var item = lineItems[i];
      var itemFound = false;
      var itemName = item.description;
      var requiredQty = parseInt(item.qty) || 0;

      // Search for the item in stock
      for (var j = 1; j < stockData.length; j++) {
        if (stockData[j][0] === itemName) {
          itemFound = true;
          var availableQty = parseInt(stockData[j][3]) || 0;

          if (availableQty < requiredQty) {
            unavailableItems.push({
              item: itemName,
              required: requiredQty,
              available: availableQty
            });
          }
          break;
        }
      }

      if (!itemFound) {
        unavailableItems.push({
          item: itemName,
          required: requiredQty,
          available: 0
        });
      }
    }

    if (unavailableItems.length > 0) {
      var message = 'Insufficient stock for:\n';
      for (var k = 0; k < unavailableItems.length; k++) {
        message += '\n• ' + unavailableItems[k].item +
                   ' (Required: ' + unavailableItems[k].required +
                   ', Available: ' + unavailableItems[k].available + ')';
      }
      return {success: false, message: message, unavailableItems: unavailableItems};
    }

    return {success: true};
  } catch (error) {
    return {success: false, message: 'Error checking stock: ' + error.toString()};
  }
}

// Update stock quantity (decrease or increase)
/**
 * Update inventory by writing to DetailPO sheet
 * For invoices: writes NEGATIVE quantities deducted from actual warehouse locations (FIFO)
 * For invoice deletion: writes POSITIVE quantities back to original warehouses
 */
function updateInventoryForInvoice(invoiceNumber, lineItems, operation) {
  try {
    Logger.log('📦 updateInventoryForInvoice: ' + operation + ' for invoice ' + invoiceNumber);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var detailSheet = ss.getSheetByName('DetailPO');
    
    if (!detailSheet) {
      Logger.log('❌ DetailPO sheet not found');
      return {success: false, message: 'DetailPO sheet not found'};
    }
    
    // DetailPO structure: Date(0), PO#(1), DetailID(2), SupID(3), SupName(4), BillNum(5),
    //                     Model(6), Name(7), Cat(8), Location(9), QTY(10), Cost(11)...
    
    var timestamp = new Date();
    var detailData = detailSheet.getDataRange().getValues();
    
    // Process each line item
    for (var i = 0; i < lineItems.length; i++) {
      var item = lineItems[i];
      var modelNumber = String(item.modelNumber || '');
      var itemName = String(item.itemName || '');
      var category = String(item.category || '');
      var qtyNeeded = parseInt(item.qty) || 0;
      
      if (!modelNumber || qtyNeeded === 0) {
        Logger.log('⚠️ Skipping item with no model number or zero qty');
        continue;
      }
      
      if (operation === 'decrease') {
        // DEDUCT FROM WAREHOUSES (FIFO - First In First Out)
        Logger.log('🏭 Finding warehouses with stock for Model: ' + modelNumber);
        
        // Calculate current stock by location for this model
        var stockByLocation = {};
        for (var j = 1; j < detailData.length; j++) {
          var detailModel = String(detailData[j][6] || '');
          var detailLocation = String(detailData[j][9] || '');
          var detailQty = parseFloat(detailData[j][10]) || 0;
          
          if (detailModel === modelNumber && detailLocation && detailLocation !== 'SOLD') {
            if (!stockByLocation[detailLocation]) {
              stockByLocation[detailLocation] = 0;
            }
            stockByLocation[detailLocation] += detailQty;
          }
        }
        
        Logger.log('📊 Stock by location for ' + modelNumber + ': ' + JSON.stringify(stockByLocation));
        
        // Deduct from warehouses in order (FIFO)
        var remainingQty = qtyNeeded;
        var warehouses = Object.keys(stockByLocation).sort(); // Sort alphabetically (WH-A, WH-B, etc.)
        
        for (var k = 0; k < warehouses.length && remainingQty > 0; k++) {
          var warehouse = warehouses[k];
          var availableQty = stockByLocation[warehouse];
          
          if (availableQty > 0) {
            var qtyToDeduct = Math.min(remainingQty, availableQty);
            
            // Write NEGATIVE entry for this warehouse
            var detailRow = [
              timestamp,                          // Date
              'INV-' + invoiceNumber,            // PO# (invoice reference)
              'SALE-' + Date.now() + '-' + i + '-' + k, // DetailID (unique)
              'SALE',                            // SupID
              'Sales Transaction',               // SupName
              invoiceNumber,                     // BillNum
              modelNumber,                       // Model
              itemName,                          // Name
              category,                          // Cat
              warehouse,                         // Location (ACTUAL WAREHOUSE)
              -qtyToDeduct,                      // QTY (NEGATIVE)
              0,                                 // Cost
              0,                                 // Tax Rate
              0,                                 // Tax Amount
              0,                                 // Total Cost
              timestamp                          // Timestamp
            ];
            
            detailSheet.appendRow(detailRow);
            Logger.log('✅ Deducted ' + qtyToDeduct + ' from ' + warehouse + ' for ' + modelNumber);
            
            remainingQty -= qtyToDeduct;
          }
        }
        
        if (remainingQty > 0) {
          Logger.log('⚠️ Warning: Not enough stock to fulfill order. Short by ' + remainingQty);
        }
        
      } else if (operation === 'increase') {
        // RESTORE TO WAREHOUSES
        // Find original sale entries for this invoice and reverse them
        Logger.log('🔄 Restoring inventory for invoice ' + invoiceNumber);
        
        var saleEntries = [];
        for (var j = 1; j < detailData.length; j++) {
          var detailPO = String(detailData[j][1] || '');
          var detailModel = String(detailData[j][6] || '');
          var detailQty = parseFloat(detailData[j][10]) || 0;
          var detailLocation = String(detailData[j][9] || '');
          
          // Find negative entries for this invoice and model
          if (detailPO === 'INV-' + invoiceNumber && detailModel === modelNumber && detailQty < 0) {
            saleEntries.push({
              location: detailLocation,
              qty: Math.abs(detailQty)
            });
          }
        }
        
        // Write positive entries to restore stock
        for (var k = 0; k < saleEntries.length; k++) {
          var detailRow = [
            timestamp,                          // Date
            'INV-RESTORE-' + invoiceNumber,    // PO#
            'RESTORE-' + Date.now() + '-' + i + '-' + k, // DetailID
            'RESTORE',                         // SupID
            'Invoice Deletion - Stock Restore', // SupName
            invoiceNumber,                     // BillNum
            modelNumber,                       // Model
            itemName,                          // Name
            category,                          // Cat
            saleEntries[k].location,           // Location (original warehouse)
            saleEntries[k].qty,                // QTY (POSITIVE)
            0,                                 // Cost
            0,                                 // Tax Rate
            0,                                 // Tax Amount
            0,                                 // Total Cost
            timestamp                          // Timestamp
          ];
          
          detailSheet.appendRow(detailRow);
          Logger.log('✅ Restored ' + saleEntries[k].qty + ' to ' + saleEntries[k].location + ' for ' + modelNumber);
        }
      }
    }
    
    Logger.log('✅ Inventory updated successfully for ' + lineItems.length + ' items');
    return {success: true, message: 'Inventory updated successfully'};
    
  } catch (error) {
    Logger.log('❌ Error in updateInventoryForInvoice: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// DEPRECATED: Old function for backward compatibility
function updateStockQuantity(lineItems, operation) {
  Logger.log('⚠️ updateStockQuantity is deprecated, using new DetailPO-based function');
  return {success: true}; // Do nothing, handled by new function
}

// Create invoice from completed quote
function createInvoiceFromQuote(quoteNumber) {
  try {
    var ss = getSpreadsheet();
    var quotesSheet = ss.getSheetByName('Quotes');
    var invoicesSheet = ss.getSheetByName('Invoices');

    if (!quotesSheet) {
      return {success: false, message: 'Quotes sheet not found'};
    }

    if (!invoicesSheet) {
      return {success: false, message: 'Invoices sheet not found. Please run setupDemoData() first.'};
    }

    // Check if invoice already exists for this quote
    var existingInvoice = checkInvoiceExistsForQuote(quoteNumber);
    if (existingInvoice) {
      return {
        success: false,
        message: 'Invoice already exists for this quote!\nInvoice Number: ' + existingInvoice
      };
    }

    // Find the quote
    var quotesData = quotesSheet.getDataRange().getValues();
    var quoteRow = null;

    for (var i = 1; i < quotesData.length; i++) {
      if (String(quotesData[i][0]) === String(quoteNumber)) {
        quoteRow = quotesData[i];
        break;
      }
    }

    if (!quoteRow) {
      return {success: false, message: 'Quote not found'};
    }

    // Parse items from NEW quote structure (Column L)
    var items = [];
    try {
      items = JSON.parse(quoteRow[11]); // Column L - Items JSON (NEW structure)
    } catch (e) {
      return {success: false, message: 'Error parsing quote line items: ' + e.toString()};
    }

    // Generate invoice number
    var invoiceNumber = getNextInvoiceNumber();
    var createdAt = new Date().toISOString();
    var dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + 30);

    // Map NEW quote structure to invoice (adjust columns based on new structure)
    var rowData = [
      invoiceNumber,              // A: Invoice Number
      createdAt,                  // B: Created At
      dueDate.toISOString(),      // C: Due Date
      quoteRow[3],                // D: Customer ID
      quoteRow[4],                // E: Customer Name
      quoteRow[5],                // F: Company
      quoteRow[6],                // G: Address
      quoteRow[7],                // H: City
      quoteRow[8],                // I: Phone
      quoteRow[9],                // J: Email
      JSON.stringify(items),      // K: Items JSON (NEW)
      quoteRow[12],               // L: Material Cost (M from quotes)
      quoteRow[13],               // M: Installation Cost (N from quotes)
      quoteRow[14],               // N: Sub Total (O from quotes)
      quoteRow[15],               // O: Sales Tax (P from quotes)
      quoteRow[16],               // P: Grand Total (Q from quotes)
      quoteRow[17],               // Q: Down Payment (R from quotes)
      quoteRow[18],               // R: Final Payment (S from quotes)
      quoteRow[19],               // S: Prepared By (T from quotes)
      quoteRow[10],               // T: Objective (K from quotes)
      'Payment due within 30 days', // U: Terms
      createdAt,                  // V: Created timestamp
      'Unpaid',                   // W: Status
      quoteNumber                 // X: Quote Number
    ];

    invoicesSheet.appendRow(rowData);
    
    // DEDUCT INVENTORY - Add negative entries to DetailPO
    Logger.log('📦 Deducting inventory for invoice ' + invoiceNumber);
    var inventoryUpdate = updateInventoryForInvoice(invoiceNumber, items, 'decrease');
    
    if (!inventoryUpdate.success) {
      Logger.log('⚠️ Warning: Inventory deduction failed: ' + inventoryUpdate.message);
      // Continue anyway - invoice is created
    } else {
      Logger.log('✅ Inventory deducted successfully');
    }

    return {
      success: true,
      message: 'Invoice created successfully and inventory updated',
      invoiceNumber: invoiceNumber
    };
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get all invoices
function getInvoices() {
  try {
    var ss = getSpreadsheet();
    var invoicesSheet = ss.getSheetByName('Invoices');

    if (!invoicesSheet) {
      return {success: true, invoices: []};
    }

    var data = invoicesSheet.getDataRange().getValues();
    var invoices = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        var lineItems = [];
        try {
          lineItems = row[10] ? JSON.parse(row[10]) : [];
        } catch (e) {
          Logger.log('Error parsing lineItems for invoice ' + row[0] + ': ' + e.toString());
          lineItems = [];
        }

        invoices.push({
          invoiceNumber: row[0],
          date: row[1],
          dueDate: row[2],
          customerId: row[3],
          customerName: row[4],
          customerCompany: row[5],
          customerAddress: row[6],
          customerCity: row[7],
          customerPhone: row[8],
          customerEmail: row[9],
          items: lineItems,
          materialCost: parseFloat(row[11]) || 0,
          installationCost: parseFloat(row[12]) || 0,
          subTotal: parseFloat(row[13]) || 0,
          salesTax: parseFloat(row[14]) || 0,
          grandTotal: parseFloat(row[15]) || 0,
          downPayment: parseFloat(row[16]) || 0,
          finalPayment: parseFloat(row[17]) || 0,
          preparedBy: row[18],
          objective: row[19] || '',
          terms: row[20],
          createdAt: row[21],
          status: row[22],
          quoteNumber: row[23]
        });
      }
    }

    return {success: true, invoices: invoices};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Update invoice status
function updateInvoiceStatus(invoiceNumber, status) {
  try {
    var ss = getSpreadsheet();
    var invoicesSheet = ss.getSheetByName('Invoices');

    if (!invoicesSheet) {
      return {success: false, message: 'Invoices sheet not found'};
    }

    var data = invoicesSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(invoiceNumber)) {
        var rowIndex = i + 1;
        invoicesSheet.getRange(rowIndex, 23).setValue(status); // Column T (Status)
        return {success: true, message: 'Invoice status updated successfully'};
      }
    }

    return {success: false, message: 'Invoice not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Delete invoice
function deleteInvoice(invoiceNumber) {
  try {
    var ss = getSpreadsheet();
    var invoicesSheet = ss.getSheetByName('Invoices');

    if (!invoicesSheet) {
      return {success: false, message: 'Invoices sheet not found'};
    }

    var data = invoicesSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(invoiceNumber)) {
        // Get line items before deleting
        var lineItemsJSON = data[i][10];
        var lineItems = [];

        try {
          lineItems = JSON.parse(lineItemsJSON);
          
          // RESTORE INVENTORY - Add positive entries to DetailPO
          Logger.log('📦 Restoring inventory for deleted invoice ' + invoiceNumber);
          var stockRestore = updateInventoryForInvoice(invoiceNumber, lineItems, 'increase');
          
          if (!stockRestore.success) {
            Logger.log('⚠️ Warning: Stock restore failed: ' + stockRestore.message);
          } else {
            Logger.log('✅ Inventory restored successfully');
          }
        } catch (e) {
          Logger.log('⚠️ Warning: Could not restore stock for deleted invoice: ' + e.toString());
        }

        // Delete the invoice
        invoicesSheet.deleteRow(i + 1);
        return {success: true, message: 'Invoice deleted and inventory restored successfully'};
      }
    }

    return {success: false, message: 'Invoice not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}


// Get invoice by number
function getInvoiceByNumber(invoiceNumber) {
  try {
    var result = getInvoices();
    if (!result.success) {
      return result;
    }

    for (var i = 0; i < result.invoices.length; i++) {
      if (String(result.invoices[i].invoiceNumber) === String(invoiceNumber)) {
        return {success: true, invoice: result.invoices[i]};
      }
    }

    return {success: false, message: 'Invoice not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Generate PDF for invoice
function generateInvoicePDF(invoiceNumber) {
  try {
    var result = getInvoiceByNumber(invoiceNumber);
    
    if (!result.success) {
      return result;
    }

    var invoice = result.invoice;
    var html = createPDFHTML(invoice, 'invoice');
    var blob = Utilities.newBlob(html, 'text/html', 'invoice.html'); 
    var pdf = blob.getAs('application/pdf').setName('Invoice_' + invoiceNumber + '.pdf');

    return {
      success: true,
      pdf: Utilities.base64Encode(pdf.getBytes()),
      filename: 'Invoice_' + invoiceNumber + '.pdf'
    };
  } catch (error) {
    return {success: false, message: 'Error generating PDF: ' + error.toString()};
  }
}
// ========================================
// CREATE MANUAL INVOICE (Direct Sale - No Quote)
// ========================================

/**
 * Create a manual invoice directly without a quote
 * Used for direct sales/walk-in customers
 */
function createManualInvoice(invoiceData) {
  try {
    Logger.log('📝 Creating manual invoice');
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invoicesSheet = ss.getSheetByName('Invoices');
    
    if (!invoicesSheet) {
      return {success: false, message: 'Invoices sheet not found'};
    }
    
    // Generate invoice number
    var invoiceNumber = getNextInvoiceNumber();
    var createdAt = new Date(invoiceData.date || new Date()).toISOString();
    var dueDate = new Date(createdAt);
    dueDate.setDate(dueDate.getDate() + 30); // 30 days from creation
    
    // Parse line items
    var items = [];
    try {
      items = typeof invoiceData.lineItems === 'string' 
        ? JSON.parse(invoiceData.lineItems) 
        : invoiceData.lineItems;
    } catch (e) {
      return {success: false, message: 'Invalid line items format: ' + e.toString()};
    }
    
    // Invoices sheet structure (24 columns):
    var rowData = [
      invoiceNumber,                              // A: Invoice Number
      createdAt,                                  // B: Created At
      dueDate.toISOString(),                      // C: Due Date
      String(invoiceData.customerId || ''),       // D: Customer ID
      String(invoiceData.customerName || ''),     // E: Customer Name
      String(invoiceData.customerCompany || ''),  // F: Company
      String(invoiceData.customerAddress || ''),  // G: Address
      String(invoiceData.customerCity || ''),     // H: City
      String(invoiceData.customerPhone || ''),    // I: Phone
      String(invoiceData.customerEmail || ''),    // J: Email
      JSON.stringify(items),                      // K: Items JSON
      parseFloat(invoiceData.materialCost) || 0,  // L: Material Cost
      parseFloat(invoiceData.installationCost) || 0, // M: Installation Cost
      parseFloat(invoiceData.subTotal) || 0,      // N: Sub Total
      parseFloat(invoiceData.salesTax) || 0,      // O: Sales Tax
      parseFloat(invoiceData.grandTotal) || 0,    // P: Grand Total
      parseFloat(invoiceData.downPayment) || 0,   // Q: Down Payment
      parseFloat(invoiceData.finalPayment) || 0,  // R: Final Payment
      String(invoiceData.preparedBy || 'admin'),  // S: Prepared By
      String(invoiceData.objective || ''),        // T: Objective
      String(invoiceData.terms || 'Payment due within 30 days'), // U: Terms
      createdAt,                                  // V: Created timestamp
      'Unpaid',                                   // W: Status
      'MANUAL'                                    // X: Quote Number (MANUAL for manual invoices)
    ];
    
    invoicesSheet.appendRow(rowData);
    Logger.log('✅ Invoice row created: ' + invoiceNumber);
    
    // DEDUCT INVENTORY
    Logger.log('📦 Deducting inventory for manual invoice ' + invoiceNumber);
    var inventoryUpdate = updateInventoryForInvoice(invoiceNumber, items, 'decrease');
    
    if (!inventoryUpdate.success) {
      Logger.log('⚠️ Warning: Inventory deduction failed: ' + inventoryUpdate.message);
    } else {
      Logger.log('✅ Inventory deducted successfully');
    }
    
    return {
      success: true,
      message: 'Manual invoice created successfully',
      invoiceNumber: invoiceNumber
    };
    
  } catch (error) {
    Logger.log('❌ Error creating manual invoice: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}
// ========================================
// UPDATE INVOICE (Edit existing invoice)
// ========================================

/**
 * Update an existing invoice
 * Handles inventory adjustments automatically
 */
function updateInvoice(invoiceData) {
  try {
    Logger.log('📝 Updating invoice: ' + invoiceData.invoiceNumber);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invoicesSheet = ss.getSheetByName('Invoices');
    
    if (!invoicesSheet) {
      return {success: false, message: 'Invoices sheet not found'};
    }
    
    var data = invoicesSheet.getDataRange().getValues();
    var rowIndex = -1;
    var oldInvoice = null;
    
    // Find the invoice row
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(invoiceData.invoiceNumber)) {
        rowIndex = i + 1; // +1 for sheet row (1-indexed)
        oldInvoice = {
          invoiceNumber: String(data[i][0]),
          items: JSON.parse(data[i][10] || '[]')
        };
        break;
      }
    }
    
    if (rowIndex === -1) {
      return {success: false, message: 'Invoice not found'};
    }
    
    Logger.log('✅ Found invoice at row ' + rowIndex);
    
    // STEP 1: Restore inventory from old invoice (reverse the deduction)
    Logger.log('📦 Restoring inventory from old invoice');
    var restoreResult = updateInventoryForInvoice(oldInvoice.invoiceNumber, oldInvoice.items, 'increase');
    if (!restoreResult.success) {
      Logger.log('⚠️ Warning: Failed to restore inventory: ' + restoreResult.message);
    }
    
    // STEP 2: Parse new line items
    var newItems = [];
    try {
      newItems = typeof invoiceData.lineItems === 'string' 
        ? JSON.parse(invoiceData.lineItems) 
        : invoiceData.lineItems;
    } catch (e) {
      return {success: false, message: 'Invalid line items format: ' + e.toString()};
    }
    
    // STEP 3: Update the invoice row
    var dueDate = new Date(invoiceData.date);
    dueDate.setDate(dueDate.getDate() + 30);
    
    // Update all editable columns (keep status and quote number unchanged)
    invoicesSheet.getRange(rowIndex, 2).setValue(invoiceData.date); // B: Date
    invoicesSheet.getRange(rowIndex, 3).setValue(dueDate.toISOString()); // C: Due Date
    invoicesSheet.getRange(rowIndex, 11).setValue(JSON.stringify(newItems)); // K: Items
    invoicesSheet.getRange(rowIndex, 12).setValue(parseFloat(invoiceData.materialCost) || 0); // L: Material Cost
    invoicesSheet.getRange(rowIndex, 13).setValue(parseFloat(invoiceData.installationCost) || 0); // M: Installation
    invoicesSheet.getRange(rowIndex, 14).setValue(parseFloat(invoiceData.subTotal) || 0); // N: Sub Total
    invoicesSheet.getRange(rowIndex, 15).setValue(parseFloat(invoiceData.salesTax) || 0); // O: Sales Tax
    invoicesSheet.getRange(rowIndex, 16).setValue(parseFloat(invoiceData.grandTotal) || 0); // P: Grand Total
    invoicesSheet.getRange(rowIndex, 17).setValue(parseFloat(invoiceData.downPayment) || 0); // Q: Down Payment
    invoicesSheet.getRange(rowIndex, 18).setValue(parseFloat(invoiceData.finalPayment) || 0); // R: Final Payment
    invoicesSheet.getRange(rowIndex, 20).setValue(String(invoiceData.objective || '')); // T: Objective
    invoicesSheet.getRange(rowIndex, 21).setValue(String(invoiceData.terms || 'Payment due within 30 days')); // U: Terms
    
    Logger.log('✅ Invoice row updated');
    
    // STEP 4: Deduct new inventory
    Logger.log('📦 Deducting new inventory');
    var deductResult = updateInventoryForInvoice(invoiceData.invoiceNumber, newItems, 'decrease');
    if (!deductResult.success) {
      Logger.log('⚠️ Warning: Failed to deduct new inventory: ' + deductResult.message);
    }
    
    return {
      success: true,
      message: 'Invoice updated successfully',
      invoiceNumber: invoiceData.invoiceNumber
    };
    
  } catch (error) {
    Logger.log('❌ Error updating invoice: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// SETUP INVOICES SHEET (Fix/Create Columns)
// ========================================

/**
 * Setup or fix the Invoices sheet with all necessary columns
 * Run this function once to create/repair the sheet structure
 */
function setupInvoicesSheet() {
  try {
    Logger.log('🔧 Setting up Invoices sheet...');
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Invoices');
    
    // If sheet doesn't exist, create it
    if (!sheet) {
      Logger.log('📄 Creating new Invoices sheet...');
      sheet = ss.insertSheet('Invoices');
    } else {
      Logger.log('📄 Invoices sheet found, will update headers');
    }
    
    // Define all column headers (24 columns: A-X)
    var headers = [
      'Invoice Number',        // A (Column 0)
      'Date',                  // B (Column 1)
      'Due Date',              // C (Column 2)
      'Customer ID',           // D (Column 3)
      'Customer Name',         // E (Column 4)
      'Company',               // F (Column 5)
      'Address',               // G (Column 6)
      'City',                  // H (Column 7)
      'Phone',                 // I (Column 8)
      'Email',                 // J (Column 9)
      'Items (JSON)',          // K (Column 10)
      'Material Cost',         // L (Column 11)
      'Installation Cost',     // M (Column 12)
      'Sub Total',             // N (Column 13)
      'Sales Tax',             // O (Column 14)
      'Grand Total',           // P (Column 15)
      'Down Payment',          // Q (Column 16)
      'Final Payment',         // R (Column 17)
      'Prepared By',           // S (Column 18)
      'Objective',             // T (Column 19)
      'Terms',                 // U (Column 20)
      'Created At',            // V (Column 21)
      'Status',                // W (Column 22)
      'Quote Number'           // X (Column 23)
    ];
    
    // Set headers in row 1
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    
    // Format header row
    headerRange.setBackground('#001f3f');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    
    // Set column widths for better readability
    sheet.setColumnWidth(1, 120);  // Invoice Number
    sheet.setColumnWidth(2, 100);  // Date
    sheet.setColumnWidth(3, 100);  // Due Date
    sheet.setColumnWidth(4, 100);  // Customer ID
    sheet.setColumnWidth(5, 150);  // Customer Name
    sheet.setColumnWidth(6, 150);  // Company
    sheet.setColumnWidth(7, 200);  // Address
    sheet.setColumnWidth(8, 120);  // City
    sheet.setColumnWidth(9, 120);  // Phone
    sheet.setColumnWidth(10, 200); // Email
    sheet.setColumnWidth(11, 250); // Items (JSON)
    sheet.setColumnWidth(12, 100); // Material Cost
    sheet.setColumnWidth(13, 120); // Installation Cost
    sheet.setColumnWidth(14, 100); // Sub Total
    sheet.setColumnWidth(15, 100); // Sales Tax
    sheet.setColumnWidth(16, 100); // Grand Total
    sheet.setColumnWidth(17, 100); // Down Payment
    sheet.setColumnWidth(18, 100); // Final Payment
    sheet.setColumnWidth(19, 100); // Prepared By
    sheet.setColumnWidth(20, 250); // Objective
    sheet.setColumnWidth(21, 200); // Terms
    sheet.setColumnWidth(22, 150); // Created At
    sheet.setColumnWidth(23, 100); // Status
    sheet.setColumnWidth(24, 120); // Quote Number
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Add data validation for Status column (W)
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Paid', 'Unpaid'], true)
      .setAllowInvalid(false)
      .build();
    
    // Apply to entire Status column (starting from row 2)
    sheet.getRange(2, 23, sheet.getMaxRows() - 1, 1).setDataValidation(statusRule);
    
    Logger.log('✅ Invoices sheet setup complete!');
    Logger.log('📊 Headers: ' + headers.length + ' columns');
    
    return {
      success: true,
      message: 'Invoices sheet setup complete with ' + headers.length + ' columns'
    };
    
  } catch (error) {
    Logger.log('❌ Error setting up Invoices sheet: ' + error);
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

/**
 * Quick check function to verify sheet structure
 */
function verifyInvoicesSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Invoices');
    
    if (!sheet) {
      return {
        success: false,
        message: 'Invoices sheet not found. Run setupInvoicesSheet() first.'
      };
    }
    
    var headers = sheet.getRange(1, 1, 1, 24).getValues()[0];
    var expectedHeaders = [
      'Invoice Number', 'Date', 'Due Date', 'Customer ID', 'Customer Name',
      'Company', 'Address', 'City', 'Phone', 'Email', 'Items (JSON)',
      'Material Cost', 'Installation Cost', 'Sub Total', 'Sales Tax',
      'Grand Total', 'Down Payment', 'Final Payment', 'Prepared By',
      'Objective', 'Terms', 'Created At', 'Status', 'Quote Number'
    ];
    
    var missingColumns = [];
    for (var i = 0; i < expectedHeaders.length; i++) {
      if (headers[i] !== expectedHeaders[i]) {
        missingColumns.push('Column ' + String.fromCharCode(65 + i) + ': Expected "' + expectedHeaders[i] + '", Found "' + headers[i] + '"');
      }
    }
    
    if (missingColumns.length > 0) {
      return {
        success: false,
        message: 'Column mismatch detected',
        issues: missingColumns
      };
    }
    
    return {
      success: true,
      message: 'Invoices sheet structure is correct',
      columnCount: headers.length,
      rowCount: sheet.getLastRow()
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}