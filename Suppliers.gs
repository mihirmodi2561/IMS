// Setup Suppliers sheet with headers and demo data
function setupSuppliersSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Suppliers');

    if (!sheet) {
      sheet = ss.insertSheet('Suppliers');
    }

    sheet.clear();

    var headers = ['Supplier ID', 'Name', 'Company Name', 'Contact Person', 'Phone', 'Email', 'Address', 'City', 'Country'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');

    // Demo suppliers
    var demoData = [
      ['SUP-1001', 'Tech Distributor Inc', 'Tech Dist', 'John Smith', '555-1111', 'sales@techdist.com', '123 Tech Lane', 'San Francisco', 'USA'],
      ['SUP-1002', 'Global Components', 'Global Comp', 'Jane Doe', '555-2222', 'info@globalcomp.com', '456 Component Ave', 'New York', 'USA'],
      ['SUP-1003', 'Premium Parts Co', 'Premium Parts', 'Bob Wilson', '555-3333', 'bob@premiumparts.com', '789 Parts Street', 'Chicago', 'USA']
    ];
    sheet.getRange(2, 1, demoData.length, demoData[0].length).setValues(demoData);

    return {success: true, message: 'Suppliers sheet created with demo data'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get next supplier ID (SUP-1001, SUP-1002, etc.)
function getNextSupplierId() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Suppliers');

    if (!sheet) return 'SUP-1001';

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 'SUP-1001';

    var lastId = sheet.getRange(lastRow, 1).getValue();
    if (!lastId) return 'SUP-1001';

    var numPart = parseInt(lastId.replace('SUP-', '')) || 1000;
    return 'SUP-' + (numPart + 1);
  } catch (error) {
    return 'SUP-1001';
  }
}

// Get all suppliers with purchase counts and totals
function getSuppliers() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Suppliers');

    if (!sheet) {
      return {success: false, message: 'Suppliers sheet not found. Run setupSuppliersSheet() first.', suppliers: []};
    }

    var data = sheet.getDataRange().getValues();
    var suppliers = [];

    // Read Purchases sheet to count purchases and totals per supplier
    var purchaseCounts = {};
    var purchaseTotals = {};
    var purchasesSheet = ss.getSheetByName('Purchases');
    if (purchasesSheet && purchasesSheet.getLastRow() > 1) {
      var pData = purchasesSheet.getDataRange().getValues();
      for (var p = 1; p < pData.length; p++) {
        var suppId = pData[p][2]; // Column C = Supplier ID
        if (suppId) {
          var key = suppId.toString();
          purchaseCounts[key] = (purchaseCounts[key] || 0) + 1;
          purchaseTotals[key] = (purchaseTotals[key] || 0) + (parseFloat(pData[p][11]) || 0); // Column L = Total
        }
      }
    }

    // Build supplier objects (skip header)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue; // skip empty rows

      var id = row[0].toString();
      suppliers.push({
        id: id,
        name: row[1] || '',
        companyName: row[2] || '',
        contactPerson: row[3] || '',
        phone: row[4] || '',
        email: row[5] || '',
        address: row[6] || '',
        city: row[7] || '',
        country: row[8] || '',
        purchaseCount: purchaseCounts[id] || 0,
        totalPurchaseValue: purchaseTotals[id] || 0
      });
    }

    return {success: true, suppliers: suppliers};
  } catch (error) {
    Logger.log('Error in getSuppliers: ' + error.toString());
    return {success: false, message: 'Error: ' + error.toString(), suppliers: []};
  }
}

// Add new supplier
function addSupplier(data) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Suppliers');

    if (!sheet) {
      return {success: false, message: 'Suppliers sheet not found'};
    }

    var newId = getNextSupplierId();

    var row = [
      newId,
      data.name || '',
      data.companyName || '',
      data.contactPerson || '',
      data.phone || '',
      data.email || '',
      data.address || '',
      data.city || '',
      data.country || ''
    ];

    sheet.appendRow(row);
    return {success: true, message: 'Supplier added successfully', supplierId: newId};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Update existing supplier
function updateSupplier(data) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Suppliers');

    if (!sheet) {
      return {success: false, message: 'Suppliers sheet not found'};
    }

    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.id.toString()) {
        var rowIndex = i + 1;
        sheet.getRange(rowIndex, 2).setValue(data.name || '');
        sheet.getRange(rowIndex, 3).setValue(data.companyName || '');
        sheet.getRange(rowIndex, 4).setValue(data.contactPerson || '');
        sheet.getRange(rowIndex, 5).setValue(data.phone || '');
        sheet.getRange(rowIndex, 6).setValue(data.email || '');
        sheet.getRange(rowIndex, 7).setValue(data.address || '');
        sheet.getRange(rowIndex, 8).setValue(data.city || '');
        sheet.getRange(rowIndex, 9).setValue(data.country || '');
        return {success: true, message: 'Supplier updated successfully'};
      }
    }

    return {success: false, message: 'Supplier not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Delete supplier (blocked if they have existing purchases)
function deleteSupplier(supplierId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Suppliers');

    if (!sheet) {
      return {success: false, message: 'Suppliers sheet not found'};
    }

    // Check if supplier has any purchases
    var purchasesSheet = ss.getSheetByName('Purchases');
    if (purchasesSheet && purchasesSheet.getLastRow() > 1) {
      var pData = purchasesSheet.getDataRange().getValues();
      for (var p = 1; p < pData.length; p++) {
        if (pData[p][2] && pData[p][2].toString() === supplierId.toString()) {
          return {success: false, message: 'Cannot delete supplier — they have existing purchases. Delete purchases first.'};
        }
      }
    }

    // Find and delete
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === supplierId.toString()) {
        sheet.deleteRow(i + 1);
        return {success: true, message: 'Supplier deleted successfully'};
      }
    }

    return {success: false, message: 'Supplier not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}
