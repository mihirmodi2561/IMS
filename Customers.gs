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


