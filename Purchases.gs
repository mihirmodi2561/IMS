/**
 * ============================================
 * PURCHASE MANAGEMENT - LOCATION AT ITEM LEVEL
 * ============================================
 * Two-sheet system with Location per ITEM (not per PO):
 * - PurchaseEntry: Summary (NO location)
 * - DetailPO: Details (location for EACH item)
 */

// ========================================
// SHEET SETUP
// ========================================

/**
 * Setup PurchaseEntry sheet (NO LOCATION)
 */
function setupPurchaseEntrySheet() {
  try {
    var ss = getSpreadsheet();
    
    var oldSheet = ss.getSheetByName('Purchases');
    if (oldSheet) {
      ss.deleteSheet(oldSheet);
      Logger.log('✅ Deleted old Purchases sheet');
    }
    
    var sheet = ss.getSheetByName('PurchaseEntry');
    
    if (!sheet) {
      sheet = ss.insertSheet('PurchaseEntry');
    } else {
      sheet.clear();
    }
    
    // Headers: Date | PO Number | Supplier ID | Supplier Name | Bill Num | Item Count | Total Amount
    var headers = ['Date', 'PO Number', 'Supplier ID', 'Supplier Name', 'Bill Num', 'Item Count', 'Total Amount'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    sheet.setColumnWidth(1, 120);  // Date
    sheet.setColumnWidth(2, 120);  // PO Number
    sheet.setColumnWidth(3, 120);  // Supplier ID
    sheet.setColumnWidth(4, 200);  // Supplier Name
    sheet.setColumnWidth(5, 150);  // Bill Num
    sheet.setColumnWidth(6, 100);  // Item Count
    sheet.setColumnWidth(7, 150);  // Total Amount
    
    Logger.log('✅ PurchaseEntry sheet created (no location)');
    return {success: true, message: 'PurchaseEntry sheet created successfully'};
    
  } catch (error) {
    Logger.log('❌ Error in setupPurchaseEntrySheet: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

/**
 * Setup DetailPO sheet (LOCATION PER ITEM)
 */
function setupDetailPOSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('DetailPO');
    
    if (!sheet) {
      sheet = ss.insertSheet('DetailPO');
    } else {
      sheet.clear();
    }
    
    // Headers: Date | PO Number | Detail ID | Supplier ID | Supplier Name | Bill Num | 
    //          Model Number | Item Name | Category | Location | QTY | Unit Cost | Cost Excl Tax | 
    //          Tax Rate (%) | Total Tax | Cost Incl Tax | Total Price
    var headers = [
      'Date', 'PO Number', 'Detail ID', 'Supplier ID', 'Supplier Name', 'Bill Num',
      'Model Number', 'Item Name', 'Category', 'Location', 'QTY', 'Unit Cost', 'Cost Excl Tax',
      'Tax Rate (%)', 'Total Tax', 'Cost Incl Tax', 'Total Price'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    sheet.setColumnWidth(1, 120);   // Date
    sheet.setColumnWidth(2, 120);   // PO Number
    sheet.setColumnWidth(3, 120);   // Detail ID
    sheet.setColumnWidth(4, 120);   // Supplier ID
    sheet.setColumnWidth(5, 200);   // Supplier Name
    sheet.setColumnWidth(6, 150);   // Bill Num
    sheet.setColumnWidth(7, 150);   // Model Number
    sheet.setColumnWidth(8, 250);   // Item Name
    sheet.setColumnWidth(9, 150);   // Category
    sheet.setColumnWidth(10, 120);  // Location ← HERE
    sheet.setColumnWidth(11, 80);   // QTY
    sheet.setColumnWidth(12, 100);  // Unit Cost
    sheet.setColumnWidth(13, 120);  // Cost Excl Tax
    sheet.setColumnWidth(14, 100);  // Tax Rate
    sheet.setColumnWidth(15, 100);  // Total Tax
    sheet.setColumnWidth(16, 120);  // Cost Incl Tax
    sheet.setColumnWidth(17, 120);  // Total Price
    
    Logger.log('✅ DetailPO sheet created with Location per item');
    return {success: true, message: 'DetailPO sheet created successfully'};
    
  } catch (error) {
    Logger.log('❌ Error in setupDetailPOSheet: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

function setupPurchaseSheets() {
  var result1 = setupPurchaseEntrySheet();
  var result2 = setupDetailPOSheet();
  
  if (result1.success && result2.success) {
    return {success: true, message: 'Both purchase sheets created successfully with Location at item level'};
  } else {
    return {success: false, message: 'Error creating sheets'};
  }
}

// ========================================
// PO NUMBER & DETAIL ID GENERATION
// ========================================

function getNextPONumber() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('PurchaseEntry');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return 'PO-0001';
    }
    
    var data = sheet.getDataRange().getValues();
    var maxNum = 0;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1]) {
        var poNum = data[i][1].toString();
        var numPart = parseInt(poNum.replace('PO-', ''));
        if (!isNaN(numPart) && numPart > maxNum) {
          maxNum = numPart;
        }
      }
    }
    
    var nextNum = maxNum + 1;
    return 'PO-' + String(nextNum).padStart(4, '0');
    
  } catch (error) {
    Logger.log('❌ Error in getNextPONumber: ' + error);
    return 'PO-0001';
  }
}

function getNextDetailId() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('DetailPO');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return 'DT-0001';
    }
    
    var data = sheet.getDataRange().getValues();
    var maxNum = 0;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][2]) {
        var detailId = data[i][2].toString();
        var numPart = parseInt(detailId.replace('DT-', ''));
        if (!isNaN(numPart) && numPart > maxNum) {
          maxNum = numPart;
        }
      }
    }
    
    var nextNum = maxNum + 1;
    return 'DT-' + String(nextNum).padStart(4, '0');
    
  } catch (error) {
    Logger.log('❌ Error in getNextDetailId: ' + error);
    return 'DT-0001';
  }
}

// ========================================
// GET ALL PURCHASE ORDERS
// ========================================

function getAllPurchaseOrders() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('PurchaseEntry');
    
    if (!sheet) {
      return {success: false, message: 'PurchaseEntry sheet not found', purchases: []};
    }
    
    var data = sheet.getDataRange().getValues();
    var purchases = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[1]) {
        purchases.push({
          date: toISOString(row[0]),
          poNumber: row[1],
          supplierId: row[2],
          supplierName: row[3],
          billNum: row[4],
          itemCount: row[5] || 0,
          totalAmount: parseFloat(row[6]) || 0,
          rowIndex: i + 1
        });
      }
    }
    
    Logger.log('✅ Retrieved ' + purchases.length + ' purchase orders');
    return {success: true, purchases: purchases};
    
  } catch (error) {
    Logger.log('❌ Error in getAllPurchaseOrders: ' + error);
    return {success: false, message: 'Error: ' + error.toString(), purchases: []};
  }
}

// ========================================
// SEARCH PURCHASE ORDERS
// ========================================

function searchPurchaseOrders(searchTerm) {
  try {
    if (!searchTerm || searchTerm.trim() === '') {
      return getAllPurchaseOrders();
    }
    
    var result = getAllPurchaseOrders();
    if (!result.success) {
      return result;
    }
    
    var searchLower = searchTerm.toLowerCase().trim();
    var filtered = [];
    
    for (var i = 0; i < result.purchases.length; i++) {
      var po = result.purchases[i];
      var dateStr = po.date ? new Date(po.date).toLocaleDateString() : '';
      
      if (po.poNumber.toLowerCase().indexOf(searchLower) !== -1 ||
          dateStr.toLowerCase().indexOf(searchLower) !== -1 ||
          (po.billNum && po.billNum.toLowerCase().indexOf(searchLower) !== -1)) {
        filtered.push(po);
      }
    }
    
    Logger.log('✅ Search found ' + filtered.length + ' results');
    return {success: true, purchases: filtered};
    
  } catch (error) {
    Logger.log('❌ Error in searchPurchaseOrders: ' + error);
    return {success: false, message: 'Error: ' + error.toString(), purchases: []};
  }
}

// ========================================
// SAVE PURCHASE ORDER - LOCATION PER ITEM
// ========================================

function savePurchaseOrder(poData) {
  try {
    var ss = getSpreadsheet();
    var entrySheet = ss.getSheetByName('PurchaseEntry');
    var detailSheet = ss.getSheetByName('DetailPO');
    
    if (!entrySheet || !detailSheet) {
      return {success: false, message: 'Required sheets not found. Run setupPurchaseSheets() first.'};
    }
    
    // Validate PO level
    if (!poData.date || !poData.poNumber || !poData.supplierName || !poData.billNum) {
      return {success: false, message: 'Missing required fields: Date, PO Number, Supplier Name, Bill Number'};
    }
    
    if (!poData.items || poData.items.length === 0) {
      return {success: false, message: 'At least one item is required'};
    }
    
    // Validate each item has location
    for (var i = 0; i < poData.items.length; i++) {
      if (!poData.items[i].location || poData.items[i].location.trim() === '') {
        return {success: false, message: 'Item ' + (i + 1) + ': Please select Location (Warehouse)'};
      }
    }
    
    // Validate no duplicate items
    var modelNumbers = {};
    for (var j = 0; j < poData.items.length; j++) {
      var model = poData.items[j].modelNumber;
      if (modelNumbers[model]) {
        return {success: false, message: 'Duplicate item: "' + model + '" is already added to this PO'};
      }
      modelNumbers[model] = true;
    }
    
    // Calculate totals
    var itemCount = poData.items.length;
    var totalAmount = 0;
    for (var k = 0; k < poData.items.length; k++) {
      totalAmount += parseFloat(poData.items[k].totalPrice) || 0;
    }
    
    // Save to PurchaseEntry (NO LOCATION)
    var entryRow = [
      poData.date,
      poData.poNumber,
      poData.supplierId,
      poData.supplierName,
      poData.billNum,
      itemCount,
      totalAmount
    ];
    entrySheet.appendRow(entryRow);
    
    // Save to DetailPO (WITH LOCATION PER ITEM)
    for (var m = 0; m < poData.items.length; m++) {
      var item = poData.items[m];
      var detailId = getNextDetailId();
      
      var detailRow = [
        poData.date,
        poData.poNumber,
        detailId,
        poData.supplierId,
        poData.supplierName,
        poData.billNum,
        item.modelNumber,
        item.itemName,
        item.itemCategory,
        item.location,  // ← LOCATION HERE (per item)
        parseFloat(item.qty) || 0,
        parseFloat(item.unitCost) || 0,
        parseFloat(item.costExclTax) || 0,
        parseFloat(item.taxRate) || 0,
        parseFloat(item.totalTax) || 0,
        parseFloat(item.costInclTax) || 0,
        parseFloat(item.totalPrice) || 0
      ];
      detailSheet.appendRow(detailRow);
    }
    
    Logger.log('✅ Saved PO: ' + poData.poNumber + ' with ' + itemCount + ' items to various warehouses');
    
    return {
      success: true,
      message: 'Purchase order saved successfully',
      poNumber: poData.poNumber,
      itemCount: itemCount,
      totalAmount: totalAmount
    };
    
  } catch (error) {
    Logger.log('❌ Error in savePurchaseOrder: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// GET PO DETAILS
// ========================================

function getPOById(poNumber) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('PurchaseEntry');
    
    if (!sheet) {
      return {success: false, message: 'PurchaseEntry sheet not found'};
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === poNumber) {
        var po = {
          date: toISOString(data[i][0]),
          poNumber: data[i][1],
          supplierId: data[i][2],
          supplierName: data[i][3],
          billNum: data[i][4],
          itemCount: data[i][5] || 0,
          totalAmount: parseFloat(data[i][6]) || 0
        };
        
        return {success: true, po: po};
      }
    }
    
    return {success: false, message: 'PO not found'};
    
  } catch (error) {
    Logger.log('❌ Error in getPOById: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

function getPOItems(poNumber) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('DetailPO');
    
    if (!sheet) {
      return {success: false, message: 'DetailPO sheet not found', items: []};
    }
    
    var data = sheet.getDataRange().getValues();
    var items = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === poNumber) {
        items.push({
          date: toISOString(data[i][0]),
          poNumber: data[i][1],
          detailId: data[i][2],
          supplierId: data[i][3],
          supplierName: data[i][4],
          billNum: data[i][5],
          modelNumber: data[i][6],
          itemName: data[i][7],
          itemCategory: data[i][8],
          location: data[i][9] || '',  // ← LOCATION HERE
          qty: parseFloat(data[i][10]) || 0,
          unitCost: parseFloat(data[i][11]) || 0,
          costExclTax: parseFloat(data[i][12]) || 0,
          taxRate: parseFloat(data[i][13]) || 0,
          totalTax: parseFloat(data[i][14]) || 0,
          costInclTax: parseFloat(data[i][15]) || 0,
          totalPrice: parseFloat(data[i][16]) || 0
        });
      }
    }
    
    return {success: true, items: items};
    
  } catch (error) {
    Logger.log('❌ Error in getPOItems: ' + error);
    return {success: false, message: 'Error: ' + error.toString(), items: []};
  }
}

// ========================================
// DELETE PURCHASE ORDER
// ========================================

function deletePurchaseOrder(poNumber) {
  try {
    var ss = getSpreadsheet();
    var entrySheet = ss.getSheetByName('PurchaseEntry');
    var detailSheet = ss.getSheetByName('DetailPO');
    
    if (!entrySheet || !detailSheet) {
      return {success: false, message: 'Required sheets not found'};
    }
    
    var entryData = entrySheet.getDataRange().getValues();
    var entryRowIndex = -1;
    
    for (var i = 1; i < entryData.length; i++) {
      if (entryData[i][1] && entryData[i][1].toString() === poNumber) {
        entryRowIndex = i + 1;
        break;
      }
    }
    
    if (entryRowIndex === -1) {
      return {success: false, message: 'PO not found'};
    }
    
    entrySheet.deleteRow(entryRowIndex);
    
    var detailData = detailSheet.getDataRange().getValues();
    var rowsToDelete = [];
    
    for (var j = 1; j < detailData.length; j++) {
      if (detailData[j][1] && detailData[j][1].toString() === poNumber) {
        rowsToDelete.push(j + 1);
      }
    }
    
    for (var k = rowsToDelete.length - 1; k >= 0; k--) {
      detailSheet.deleteRow(rowsToDelete[k]);
    }
    
    Logger.log('✅ Deleted PO: ' + poNumber + ' and ' + rowsToDelete.length + ' items');
    
    return {
      success: true,
      message: 'Purchase order deleted successfully'
    };
    
  } catch (error) {
    Logger.log('❌ Error in deletePurchaseOrder: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// UPDATE PURCHASE ORDER - ADD THIS TO Purchases.gs
// ========================================

/**
 * Update an existing purchase order
 * Deletes old entries and creates new ones with updated data
 */
function updatePurchaseOrder(poData) {
  try {
    var ss = getSpreadsheet();
    var entrySheet = ss.getSheetByName('PurchaseEntry');
    var detailSheet = ss.getSheetByName('DetailPO');
    
    if (!entrySheet || !detailSheet) {
      return {success: false, message: 'Required sheets not found'};
    }
    
    // Validate input
    if (!poData.poNumber) {
      return {success: false, message: 'PO Number is required'};
    }
    
    if (!poData.date || !poData.supplierName || !poData.billNum) {
      return {success: false, message: 'Missing required fields: Date, Supplier Name, Bill Number'};
    }
    
    if (!poData.items || poData.items.length === 0) {
      return {success: false, message: 'At least one item is required'};
    }
    
    // Validate each item has location
    for (var i = 0; i < poData.items.length; i++) {
      if (!poData.items[i].location || poData.items[i].location.trim() === '') {
        return {success: false, message: 'Item ' + (i + 1) + ': Please select Location (Warehouse)'};
      }
    }
    
    // Validate no duplicate items
    var modelNumbers = {};
    for (var j = 0; j < poData.items.length; j++) {
      var model = poData.items[j].modelNumber;
      if (modelNumbers[model]) {
        return {success: false, message: 'Duplicate item: "' + model + '" is already added to this PO'};
      }
      modelNumbers[model] = true;
    }
    
    // STEP 1: Find and delete old entry from PurchaseEntry
    var entryData = entrySheet.getDataRange().getValues();
    var entryRowIndex = -1;
    
    for (var k = 1; k < entryData.length; k++) {
      if (entryData[k][1] && entryData[k][1].toString() === poData.poNumber) {
        entryRowIndex = k + 1;
        break;
      }
    }
    
    if (entryRowIndex === -1) {
      return {success: false, message: 'Purchase order not found'};
    }
    
    // Delete old entry
    entrySheet.deleteRow(entryRowIndex);
    
    // STEP 2: Find and delete old items from DetailPO
    var detailData = detailSheet.getDataRange().getValues();
    var rowsToDelete = [];
    
    for (var m = 1; m < detailData.length; m++) {
      if (detailData[m][1] && detailData[m][1].toString() === poData.poNumber) {
        rowsToDelete.push(m + 1);
      }
    }
    
    // Delete old items (in reverse order to maintain indices)
    for (var n = rowsToDelete.length - 1; n >= 0; n--) {
      detailSheet.deleteRow(rowsToDelete[n]);
    }
    
    Logger.log('✅ Deleted old PO data: ' + poData.poNumber);
    
    // STEP 3: Calculate new totals
    var itemCount = poData.items.length;
    var totalAmount = 0;
    for (var p = 0; p < poData.items.length; p++) {
      totalAmount += parseFloat(poData.items[p].totalPrice) || 0;
    }
    
    // STEP 4: Save updated entry to PurchaseEntry
    var newEntryRow = [
      poData.date,
      poData.poNumber,
      poData.supplierId,
      poData.supplierName,
      poData.billNum,
      itemCount,
      totalAmount
    ];
    entrySheet.appendRow(newEntryRow);
    
    // STEP 5: Save updated items to DetailPO
    for (var q = 0; q < poData.items.length; q++) {
      var item = poData.items[q];
      var detailId = getNextDetailId();
      
      var newDetailRow = [
        poData.date,
        poData.poNumber,
        detailId,
        poData.supplierId,
        poData.supplierName,
        poData.billNum,
        item.modelNumber,
        item.itemName,
        item.itemCategory,
        item.location,
        parseFloat(item.qty) || 0,
        parseFloat(item.unitCost) || 0,
        parseFloat(item.costExclTax) || 0,
        parseFloat(item.taxRate) || 0,
        parseFloat(item.totalTax) || 0,
        parseFloat(item.costInclTax) || 0,
        parseFloat(item.totalPrice) || 0
      ];
      detailSheet.appendRow(newDetailRow);
    }
    
    Logger.log('✅ Updated PO: ' + poData.poNumber + ' with ' + itemCount + ' items');
    
    return {
      success: true,
      message: 'Purchase order updated successfully',
      poNumber: poData.poNumber,
      itemCount: itemCount,
      totalAmount: totalAmount
    };
    
  } catch (error) {
    Logger.log('❌ Error in updatePurchaseOrder: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// HELPER FUNCTIONS
// ========================================

function toISOString(date) {
  if (!date) return '';
  try {
    if (typeof date === 'string') return date;
    if (date instanceof Date) return date.toISOString();
    var d = new Date(date);
    return d.toISOString();
  } catch (e) {
    return '';
  }
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ========================================
// TESTING
// ========================================

function runPurchaseTest() {
  Logger.log('=== PURCHASE TEST - LOCATION PER ITEM ===');
  setupPurchaseSheets();
  Logger.log('Next PO: ' + getNextPONumber());
  Logger.log('=== TEST COMPLETE ===');
}