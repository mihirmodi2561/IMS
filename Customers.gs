/**
 * Customer Management Functions
 * Handles customer records and information
 */

function getCustomers() {
  try {
    var ss = getSpreadsheet();
    var customersSheet = ss.getSheetByName('Customers');

    if (!customersSheet) {
      return {success: false, message: 'Customers sheet not found'};
    }

    var data = customersSheet.getDataRange().getValues();
    var customers = [];

    // Get quotes and invoices to calculate counts per customer
    var quotesSheet = ss.getSheetByName('Quotes');
    var invoicesSheet = ss.getSheetByName('Invoices');
    var quoteCount = {};
    var invoiceCount = {};

    if (quotesSheet) {
      var quotesData = quotesSheet.getDataRange().getValues();
      // Skip header row and count quotes per customer ID
      for (var j = 1; j < quotesData.length; j++) {
        if (quotesData[j][3]) { // Customer ID is in column 4 (index 3)
          var custId = String(quotesData[j][3]);
          quoteCount[custId] = (quoteCount[custId] || 0) + 1;
        }
      }
    }

    if (invoicesSheet) {
      var invoicesData = invoicesSheet.getDataRange().getValues();
      // Skip header row and count invoices per customer ID
      for (var k = 1; k < invoicesData.length; k++) {
        if (invoicesData[k][3]) { // Customer ID is in column 4 (index 3)
          var custId2 = String(invoicesData[k][3]);
          invoiceCount[custId2] = (invoiceCount[custId2] || 0) + 1;
        }
      }
    }

    // Skip header row
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) { // Check if ID exists
        var customerId = String(row[0]);
        customers.push({
          id: row[0],
          name: row[1],
          companyName: row[2],
          streetAddress: row[3],
          city: row[4],
          phone: row[5],
          email: row[6] || '',
          quoteCount: quoteCount[customerId] || 0,
          invoiceCount: invoiceCount[customerId] || 0,
          rowIndex: i + 1 // Store row index for updates/deletes
        });
      }
    }

    return {success: true, customers: customers};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get next customer ID
function getNextCustomerId() {
  try {
    var ss = getSpreadsheet();
    var customersSheet = ss.getSheetByName('Customers');

    if (!customersSheet) {
      return '1001';
    }

    var lastRow = customersSheet.getLastRow();
    if (lastRow <= 1) {
      return '1001';
    }

    var lastId = customersSheet.getRange(lastRow, 1).getValue();
    if (!lastId) {
      return '1001';
    }

    var nextId = parseInt(lastId) + 1;
    return nextId.toString();
  } catch (error) {
    return '1001';
  }
}

// Add new customer
function addCustomer(customerData) {
  try {
    var ss = getSpreadsheet();
    var customersSheet = ss.getSheetByName('Customers');

    if (!customersSheet) {
      return {success: false, message: 'Customers sheet not found'};
    }

    var customerId = customerData.id || getNextCustomerId();

    var rowData = [
      customerId,
      customerData.name,
      customerData.companyName || '',
      customerData.streetAddress || '',
      customerData.city || '',
      customerData.phone || '',
      customerData.email || ''
    ];

    customersSheet.appendRow(rowData);

    return {
      success: true,
      message: 'Customer added successfully',
      customerId: customerId
    };
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Update customer
function updateCustomer(customerData) {
  try {
    var ss = getSpreadsheet();
    var customersSheet = ss.getSheetByName('Customers');

    if (!customersSheet) {
      return {success: false, message: 'Customers sheet not found'};
    }

    var data = customersSheet.getDataRange().getValues();

    // Find the customer by ID (convert both to strings for comparison)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(customerData.id)) {
        var rowIndex = i + 1;
        customersSheet.getRange(rowIndex, 2).setValue(customerData.name);
        customersSheet.getRange(rowIndex, 3).setValue(customerData.companyName || '');
        customersSheet.getRange(rowIndex, 4).setValue(customerData.streetAddress || '');
        customersSheet.getRange(rowIndex, 5).setValue(customerData.city || '');
        customersSheet.getRange(rowIndex, 6).setValue(customerData.phone || '');
        customersSheet.getRange(rowIndex, 7).setValue(customerData.email || '');

        return {success: true, message: 'Customer updated successfully'};
      }
    }

    return {success: false, message: 'Customer not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Delete customer
function deleteCustomer(customerId) {
  try {
    var ss = getSpreadsheet();
    var customersSheet = ss.getSheetByName('Customers');

    if (!customersSheet) {
      return {success: false, message: 'Customers sheet not found'};
    }

    var data = customersSheet.getDataRange().getValues();

    // Find and delete the row (convert both to strings for comparison)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(customerId)) {
        customersSheet.deleteRow(i + 1);
        return {success: true, message: 'Customer deleted successfully'};
      }
    }

    return {success: false, message: 'Customer not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

function searchCustomers(searchTerm) {
  try {
    if (!searchTerm || searchTerm.trim() === '') {
      // If no search term, return all customers
      return getCustomers();
    }

    var ss = getSpreadsheet();
    var customersSheet = ss.getSheetByName('Customers');

    if (!customersSheet) {
      return {success: false, message: 'Customers sheet not found'};
    }

    var data = customersSheet.getDataRange().getValues();
    var customers = [];
    
    // Convert search term to lowercase for case-insensitive search
    var searchLower = searchTerm.toLowerCase().trim();

    // Get quotes and invoices to calculate counts per customer
    var quotesSheet = ss.getSheetByName('Quotes');
    var invoicesSheet = ss.getSheetByName('Invoices');
    var quoteCount = {};
    var invoiceCount = {};

    if (quotesSheet) {
      var quotesData = quotesSheet.getDataRange().getValues();
      for (var j = 1; j < quotesData.length; j++) {
        if (quotesData[j][3]) {
          var custId = String(quotesData[j][3]);
          quoteCount[custId] = (quoteCount[custId] || 0) + 1;
        }
      }
    }

    if (invoicesSheet) {
      var invoicesData = invoicesSheet.getDataRange().getValues();
      for (var k = 1; k < invoicesData.length; k++) {
        if (invoicesData[k][3]) {
          var custId2 = String(invoicesData[k][3]);
          invoiceCount[custId2] = (invoiceCount[custId2] || 0) + 1;
        }
      }
    }

    // Skip header row and filter customers
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) { // Check if ID exists
        var customerId = String(row[0]).toLowerCase();
        var customerName = String(row[1] || '').toLowerCase();
        var companyName = String(row[2] || '').toLowerCase();
        var email = String(row[6] || '').toLowerCase();

        // Check if search term matches any of the fields
        if (customerId.indexOf(searchLower) !== -1 ||
            customerName.indexOf(searchLower) !== -1 ||
            companyName.indexOf(searchLower) !== -1 ||
            email.indexOf(searchLower) !== -1) {
          
          var customerIdStr = String(row[0]);
          customers.push({
            id: row[0],
            name: row[1],
            companyName: row[2],
            streetAddress: row[3],
            city: row[4],
            phone: row[5],
            email: row[6] || '',
            quoteCount: quoteCount[customerIdStr] || 0,
            invoiceCount: invoiceCount[customerIdStr] || 0,
            rowIndex: i + 1
          });
        }
      }
    }

    return {success: true, customers: customers};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}
// Stock Management Functions

// Add new stock item

// ========================================
// GET CUSTOMER COMPLETE HISTORY
// ========================================

/**
 * Get complete customer history including quotes, invoices, and service tickets
 * Used for Customer Profile View modal
 */
function getCustomerHistory(customerId) {
  try {
    Logger.log('Getting history for customer: ' + customerId);
    
    var ss = getSpreadsheet();
    
    // Get customer info
    var customersSheet = ss.getSheetByName('Customers');
    var customerInfo = null;
    
    if (customersSheet) {
      var custData = customersSheet.getDataRange().getValues();
      for (var i = 1; i < custData.length; i++) {
        if (String(custData[i][0]) === String(customerId)) {
          customerInfo = {
            id: String(custData[i][0]),
            name: custData[i][1],
            companyName: custData[i][2],
            streetAddress: custData[i][3],
            city: custData[i][4],
            phone: custData[i][5],
            email: custData[i][6]
          };
          break;
        }
      }
    }
    
    if (!customerInfo) {
      return {success: false, message: 'Customer not found'};
    }
    
    // Get quotes for this customer
    var quotes = [];
    var quotesSheet = ss.getSheetByName('Quotes');
    if (quotesSheet) {
      var quotesData = quotesSheet.getDataRange().getValues();
      for (var q = 1; q < quotesData.length; q++) {
        if (String(quotesData[q][3]) === String(customerId)) {
          quotes.push({
            quoteNumber: String(quotesData[q][0]),
            date: quotesData[q][1],
            total: parseFloat(quotesData[q][15]) || 0, // Grand Total
            status: quotesData[q][21] || 'Pending',
            objective: quotesData[q][19] || ''
          });
        }
      }
    }
    
    // Get invoices for this customer
    var invoices = [];
    var totalRevenue = 0;
    var outstandingBalance = 0;
    var invoicesSheet = ss.getSheetByName('Invoices');
    if (invoicesSheet) {
      var invoicesData = invoicesSheet.getDataRange().getValues();
      for (var inv = 1; inv < invoicesData.length; inv++) {
        if (String(invoicesData[inv][3]) === String(customerId)) {
          var grandTotal = parseFloat(invoicesData[inv][15]) || 0;
          var status = invoicesData[inv][22] || 'Unpaid';
          
          invoices.push({
            invoiceNumber: String(invoicesData[inv][0]),
            date: invoicesData[inv][1],
            dueDate: invoicesData[inv][2],
            total: grandTotal,
            status: status,
            quoteNumber: invoicesData[inv][23] || ''
          });
          
          totalRevenue += grandTotal;
          if (status === 'Unpaid') {
            outstandingBalance += grandTotal;
          }
        }
      }
    }
    
    // Get service tickets for this customer
    var tickets = [];
    var ticketsSheet = ss.getSheetByName('Service Tickets');
    if (ticketsSheet) {
      var ticketsData = ticketsSheet.getDataRange().getValues();
      for (var t = 1; t < ticketsData.length; t++) {
        if (String(ticketsData[t][3]) === String(customerId)) {
          var items = [];
          try {
            items = ticketsData[t][7] ? JSON.parse(ticketsData[t][7]) : [];
          } catch (e) {
            items = [];
          }
          
          tickets.push({
            ticketNumber: String(ticketsData[t][0]),
            date: ticketsData[t][1],
            technicianName: ticketsData[t][2],
            problemType: ticketsData[t][6],
            itemCount: items.length,
            status: ticketsData[t][9],
            quoteNumber: ticketsData[t][13] || ''
          });
        }
      }
    }
    
    // Calculate summary
    var summary = {
      totalQuotes: quotes.length,
      totalInvoices: invoices.length,
      totalTickets: tickets.length,
      totalRevenue: totalRevenue,
      outstandingBalance: outstandingBalance
    };
    
    return {
      success: true,
      customer: customerInfo,
      summary: summary,
      quotes: quotes,
      invoices: invoices,
      tickets: tickets
    };
    
  } catch (error) {
    Logger.log('❌ Error getting customer history: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ============================================
// ADD THESE FUNCTIONS TO Customers.gs
// ============================================

/**
 * Get all quotes for a specific customer
 */
function getCustomerQuotes(customerId) {
  try {
    Logger.log('Getting quotes for customer: ' + customerId);
    
    var ss = getSpreadsheet();
    var quotesSheet = ss.getSheetByName('Quotes');

    if (!quotesSheet) {
      return {success: false, message: 'Quotes sheet not found'};
    }

    var data = quotesSheet.getDataRange().getValues();
    var customerQuotes = [];

    // Skip header row
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Customer ID is in column D (index 3)
      if (String(row[3]) === String(customerId)) {
        // Parse items JSON
        var items = [];
        try {
          if (row[11]) {
            items = JSON.parse(String(row[11]));
          }
        } catch (e) {
          items = [];
        }

        customerQuotes.push({
          quoteNumber: String(row[0]),
          date: row[1] ? String(row[1]) : '',
          validUntil: row[2] ? String(row[2]) : '',
          customerId: String(row[3]),
          customerName: String(row[4]),
          customerCompany: String(row[5]),
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
          preparedBy: String(row[19]),
          terms: String(row[20] || ''),
          createdAt: row[21] ? String(row[21]) : '',
          status: String(row[22] || 'Pending')
        });
      }
    }

    Logger.log('Found ' + customerQuotes.length + ' quotes for customer ' + customerId);

    return {
      success: true,
      quotes: customerQuotes
    };
  } catch (error) {
    Logger.log('Error getting customer quotes: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

/**
 * Get all invoices for a specific customer
 */
function getCustomerInvoices(customerId) {
  try {
    Logger.log('Getting invoices for customer: ' + customerId);
    
    var ss = getSpreadsheet();
    var invoicesSheet = ss.getSheetByName('Invoices');

    if (!invoicesSheet) {
      return {success: false, message: 'Invoices sheet not found'};
    }

    var data = invoicesSheet.getDataRange().getValues();
    var customerInvoices = [];

    // Skip header row
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Customer ID is in column D (index 3)
      if (String(row[3]) === String(customerId)) {
        // Parse items JSON if exists
        var items = [];
        try {
          if (row[11]) {
            items = JSON.parse(String(row[11]));
          }
        } catch (e) {
          items = [];
        }

        customerInvoices.push({
          invoiceNumber: String(row[0]),
          date: row[1] ? String(row[1]) : '',
          dueDate: row[2] ? String(row[2]) : '',
          customerId: String(row[3]),
          customerName: String(row[4]),
          customerCompany: String(row[5]),
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
          terms: String(row[19] || ''),
          createdAt: row[20] ? String(row[20]) : '',
          status: String(row[22] || 'Unpaid'),
          paidAmount: parseFloat(row[22]) || 0,
          quoteNumber: String(row[23] || '')
        });
      }
    }

    Logger.log('Found ' + customerInvoices.length + ' invoices for customer ' + customerId);

    return {
      success: true,
      invoices: customerInvoices
    };
  } catch (error) {
    Logger.log('Error getting customer invoices: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}