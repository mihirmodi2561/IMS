/**
 * ============================================
 * ITEMS MANAGEMENT - BACKEND FUNCTIONS
 * ============================================
 * Complete Google Apps Script backend for Items page
 * 
 * Features:
 * - Auto-incrementing Item ID (IT-0001, IT-0002...)
 * - Unique Model Number validation
 * - Delete validation (check if used in purchases)
 * - Stock QTY calculation from Purchases
 * - Reorder Required auto-calculation
 */

// ========================================
// ITEMS SHEET SETUP
// ========================================

/**
 * Setup Items sheet with proper structure
 */
function setupItemsSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Items');
    
    if (!sheet) {
      sheet = ss.insertSheet('Items');
    }
    
    sheet.clear();
    
    // Headers: Item ID, Model Number, Item Name, Category, Description, Reorder Level, Date Added
    var headers = ['Item ID', 'Model Number', 'Item Name', 'Category', 'Description', 'Reorder Level', 'Date Added'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    // Set column widths for better visibility
    sheet.setColumnWidth(1, 100);  // Item ID
    sheet.setColumnWidth(2, 150);  // Model Number
    sheet.setColumnWidth(3, 250);  // Item Name
    sheet.setColumnWidth(4, 150);  // Category
    sheet.setColumnWidth(5, 300);  // Description
    sheet.setColumnWidth(6, 120);  // Reorder Level
    sheet.setColumnWidth(7, 150);  // Date Added
    
    Logger.log('✅ Items sheet created successfully');
    return {success: true, message: 'Items sheet created successfully'};
    
  } catch (error) {
    Logger.log('❌ Error in setupItemsSheet: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// ITEM ID GENERATION
// ========================================

/**
 * Get next Item ID (IT-0001, IT-0002, etc.)
 */
function getNextItemId() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Items');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return 'IT-0001';
    }
    
    var data = sheet.getDataRange().getValues();
    var maxNum = 0;
    
    // Find highest item number
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var itemId = data[i][0].toString();
        var numPart = parseInt(itemId.replace('IT-', ''));
        if (!isNaN(numPart) && numPart > maxNum) {
          maxNum = numPart;
        }
      }
    }
    
    var nextNum = maxNum + 1;
    return 'IT-' + String(nextNum).padStart(4, '0');
    
  } catch (error) {
    Logger.log('❌ Error in getNextItemId: ' + error);
    return 'IT-0001';
  }
}

// ========================================
// GET ALL ITEMS WITH STOCK CALCULATION
// ========================================

/**
 * Get all items with stock qty calculated from purchases
 */
function getAllItems() {
  try {
    Logger.log('🔍 getAllItems() started');
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('❌ Spreadsheet is null');
      return {success: false, message: 'Cannot access spreadsheet', items: []};
    }
    
    var itemsSheet = ss.getSheetByName('Items');
    
    if (!itemsSheet) {
      Logger.log('❌ Items sheet not found');
      return {success: false, message: 'Items sheet not found', items: []};
    }
    
    Logger.log('✅ Items sheet found');
    
    var data = itemsSheet.getDataRange().getValues();
    Logger.log('📄 Retrieved ' + data.length + ' rows from Items sheet');
    
    var items = [];
    
    // Get stock quantities from DetailPO sheet (simplified, no error if fails)
    var stockQtys = {};
    try {
      var detailSheet = ss.getSheetByName('DetailPO');
      if (detailSheet && detailSheet.getLastRow() > 1) {
        var detailData = detailSheet.getDataRange().getValues();
        for (var j = 1; j < detailData.length; j++) {
          var modelNum = detailData[j][6]; // Column 6 = Model
          var qty = parseFloat(detailData[j][10]) || 0; // Column 10 = QTY
          if (modelNum) {
            stockQtys[modelNum] = (stockQtys[modelNum] || 0) + qty;
          }
        }
      }
      Logger.log('✅ Stock quantities loaded: ' + Object.keys(stockQtys).length + ' models');
    } catch (e) {
      Logger.log('⚠️ Could not load stock: ' + e);
    }
    
    // Process items (skip header row)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) { // If Item ID exists
        var modelNumber = String(row[1] || '');
        var stockQty = stockQtys[modelNumber] || 0;
        var reorderLevel = parseFloat(row[5]) || 0;
        
        // Simple date handling - convert to string
        var dateValue = row[6];
        var dateString = '';
        if (dateValue) {
          try {
            if (dateValue instanceof Date) {
              dateString = dateValue.toISOString();
            } else {
              dateString = new Date(dateValue).toISOString();
            }
          } catch (e) {
            dateString = String(dateValue);
          }
        }
        
        items.push({
          itemId: String(row[0]),
          modelNumber: modelNumber,
          itemName: String(row[2] || ''),
          category: String(row[3] || ''),
          description: String(row[4] || ''),
          reorderLevel: reorderLevel,
          dateAdded: dateString,
          stockQty: stockQty,
          reorderRequired: stockQty < reorderLevel ? 'Yes' : 'No',
          rowIndex: i + 1
        });
      }
    }
    
    Logger.log('✅ Successfully retrieved ' + items.length + ' items');
    
    // Return simple object
    var result = {
      success: true,
      items: items
    };
    
    return result;
    
  } catch (error) {
    Logger.log('❌ CRITICAL ERROR in getAllItems: ' + error);
    Logger.log('❌ Error stack: ' + error.stack);
    
    // Return error object
    return {
      success: false,
      message: String(error),
      items: []
    };
  }
}

/**
 * Calculate stock quantities from Purchases sheet
 * Returns object with modelNumber as key and total qty as value
 */
function getStockQuantities() {
  try {
    var ss = getSpreadsheet();
    var purchasesSheet = ss.getSheetByName('DetailPO'); // FIXED: Changed from 'Purchases' to 'DetailPO'
    
    if (!purchasesSheet || purchasesSheet.getLastRow() <= 1) {
      Logger.log('⚠️ DetailPO sheet is empty or not found');
      return {};
    }
    
    var data = purchasesSheet.getDataRange().getValues();
    var stockQtys = {};
    
    // DetailPO structure: Date(0), PO#(1), DetailID(2), SupID(3), SupName(4), BillNum(5),
    //                     Model(6), Name(7), Cat(8), Location(9), QTY(10), Cost(11)...
    var modelNumCol = 6;  // Model Number is column 6
    var qtyCol = 10;       // QTY is column 10
    
    Logger.log('📊 Processing ' + (data.length - 1) + ' purchase records from DetailPO');
    
    // Sum quantities by model number (skip header row)
    for (var j = 1; j < data.length; j++) {
      var modelNumber = data[j][modelNumCol];
      var qty = parseFloat(data[j][qtyCol]) || 0;
      
      if (modelNumber) {
        if (!stockQtys[modelNumber]) {
          stockQtys[modelNumber] = 0;
        }
        stockQtys[modelNumber] += qty;
      }
    }
    
    Logger.log('✅ Calculated stock for ' + Object.keys(stockQtys).length + ' model numbers');
    return stockQtys;
    
  } catch (error) {
    Logger.log('❌ Error in getStockQuantities: ' + error);
    return {};
  }
}

// ========================================
// ADD NEW ITEM
// ========================================

/**
 * Add new item to Items sheet
 */
function addItem(itemData) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Items');
    
    if (!sheet) {
      return {success: false, message: 'Items sheet not found'};
    }
    
    // Validate required fields
    if (!itemData.modelNumber || itemData.modelNumber.trim() === '') {
      return {success: false, message: 'Model Number is required'};
    }
    if (!itemData.itemName || itemData.itemName.trim() === '') {
      return {success: false, message: 'Item Name is required'};
    }
    if (!itemData.category || itemData.category.trim() === '') {
      return {success: false, message: 'Category is required'};
    }
    if (itemData.reorderLevel === undefined || itemData.reorderLevel === null || itemData.reorderLevel === '') {
      return {success: false, message: 'Reorder Level is required'};
    }
    
    var reorderLevel = parseFloat(itemData.reorderLevel);
    if (isNaN(reorderLevel) || reorderLevel < 0) {
      return {success: false, message: 'Reorder Level must be a number >= 0'};
    }
    
    // Check if Model Number already exists (must be unique)
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().toLowerCase() === itemData.modelNumber.trim().toLowerCase()) {
        return {success: false, message: 'Model Number "' + itemData.modelNumber + '" already exists'};
      }
    }
    
    // Generate next Item ID
    var itemId = getNextItemId();
    var now = new Date();
    
    // Row: Item ID, Model Number, Item Name, Category, Description, Reorder Level, Date Added
    var row = [
      itemId,
      itemData.modelNumber.trim(),
      itemData.itemName.trim(),
      itemData.category.trim(),
      itemData.description ? itemData.description.trim() : '',
      reorderLevel,
      now
    ];
    
    sheet.appendRow(row);
    
    Logger.log('✅ Added item: ' + itemId + ' - ' + itemData.itemName);
    
    return {
      success: true,
      message: 'Item added successfully',
      itemId: itemId
    };
    
  } catch (error) {
    Logger.log('❌ Error in addItem: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// UPDATE ITEM
// ========================================

/**
 * Update existing item
 */
function updateItem(itemId, itemData) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Items');
    
    if (!sheet) {
      return {success: false, message: 'Items sheet not found'};
    }
    
    // Validate required fields
    if (!itemData.modelNumber || itemData.modelNumber.trim() === '') {
      return {success: false, message: 'Model Number is required'};
    }
    if (!itemData.itemName || itemData.itemName.trim() === '') {
      return {success: false, message: 'Item Name is required'};
    }
    if (!itemData.category || itemData.category.trim() === '') {
      return {success: false, message: 'Category is required'};
    }
    if (itemData.reorderLevel === undefined || itemData.reorderLevel === null || itemData.reorderLevel === '') {
      return {success: false, message: 'Reorder Level is required'};
    }
    
    var reorderLevel = parseFloat(itemData.reorderLevel);
    if (isNaN(reorderLevel) || reorderLevel < 0) {
      return {success: false, message: 'Reorder Level must be a number >= 0'};
    }
    
    // Find the item row
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var currentModelNumber = '';
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === itemId) {
        rowIndex = i + 1;  // Sheet rows are 1-indexed
        currentModelNumber = data[i][1];
        break;
      }
    }
    
    if (rowIndex === -1) {
      return {success: false, message: 'Item not found'};
    }
    
    // Check if new Model Number conflicts with existing items (except current item)
    if (itemData.modelNumber.trim().toLowerCase() !== currentModelNumber.toString().toLowerCase()) {
      for (var j = 1; j < data.length; j++) {
        if (data[j][1] && data[j][1].toString().toLowerCase() === itemData.modelNumber.trim().toLowerCase()) {
          return {success: false, message: 'Model Number "' + itemData.modelNumber + '" already exists'};
        }
      }
    }
    
    // Update the row (keep Item ID and Date Added, update rest)
    // Row: Item ID, Model Number, Item Name, Category, Description, Reorder Level, Date Added
    sheet.getRange(rowIndex, 2).setValue(itemData.modelNumber.trim());  // Model Number
    sheet.getRange(rowIndex, 3).setValue(itemData.itemName.trim());     // Item Name
    sheet.getRange(rowIndex, 4).setValue(itemData.category.trim());     // Category
    sheet.getRange(rowIndex, 5).setValue(itemData.description ? itemData.description.trim() : '');  // Description
    sheet.getRange(rowIndex, 6).setValue(reorderLevel);                 // Reorder Level
    
    Logger.log('✅ Updated item: ' + itemId);
    
    return {
      success: true,
      message: 'Item updated successfully'
    };
    
  } catch (error) {
    Logger.log('❌ Error in updateItem: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// DELETE ITEM
// ========================================

/**
 * Delete item (with validation - check if used in purchases)
 */
function deleteItem(itemId) {
  try {
    var ss = getSpreadsheet();
    var itemsSheet = ss.getSheetByName('Items');
    
    if (!itemsSheet) {
      return {success: false, message: 'Items sheet not found'};
    }
    
    // Find the item and get its model number
    var data = itemsSheet.getDataRange().getValues();
    var rowIndex = -1;
    var modelNumber = '';
    var itemName = '';
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === itemId) {
        rowIndex = i + 1;
        modelNumber = data[i][1];
        itemName = data[i][2];
        break;
      }
    }
    
    if (rowIndex === -1) {
      return {success: false, message: 'Item not found'};
    }
    
    // Check if item is used in Purchases
    var purchasesSheet = ss.getSheetByName('Purchases');
    if (purchasesSheet && purchasesSheet.getLastRow() > 1) {
      var purchasesData = purchasesSheet.getDataRange().getValues();
      var headers = purchasesData[0];
      var modelNumCol = -1;
      
      // Find Model Number column in Purchases
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j].toString().toLowerCase();
        if (header.indexOf('model') !== -1 && header.indexOf('number') !== -1) {
          modelNumCol = j;
          break;
        }
      }
      
      if (modelNumCol !== -1) {
        // Check if this model number is used in any purchase
        for (var k = 1; k < purchasesData.length; k++) {
          if (purchasesData[k][modelNumCol] && 
              purchasesData[k][modelNumCol].toString().toLowerCase() === modelNumber.toString().toLowerCase()) {
            return {
              success: false,
              message: 'Cannot delete item "' + itemName + '" - it is used in existing purchases'
            };
          }
        }
      }
    }
    
    // Safe to delete - item not used in purchases
    itemsSheet.deleteRow(rowIndex);
    
    Logger.log('✅ Deleted item: ' + itemId + ' - ' + itemName);
    
    return {
      success: true,
      message: 'Item deleted successfully'
    };
    
  } catch (error) {
    Logger.log('❌ Error in deleteItem: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// SEARCH ITEMS
// ========================================

/**
 * Search items by keyword
 */
function searchItems(searchTerm) {
  try {
    if (!searchTerm || searchTerm.trim() === '') {
      return getAllItems();
    }
    
    var result = getAllItems();
    if (!result.success) {
      return result;
    }
    
    var searchLower = searchTerm.toLowerCase().trim();
    var filtered = [];
    
    for (var i = 0; i < result.items.length; i++) {
      var item = result.items[i];
      if (item.itemId.toLowerCase().indexOf(searchLower) !== -1 ||
          item.modelNumber.toLowerCase().indexOf(searchLower) !== -1 ||
          item.itemName.toLowerCase().indexOf(searchLower) !== -1 ||
          item.category.toLowerCase().indexOf(searchLower) !== -1 ||
          (item.description && item.description.toLowerCase().indexOf(searchLower) !== -1)) {
        filtered.push(item);
      }
    }
    
    return {success: true, items: filtered};
    
  } catch (error) {
    Logger.log('❌ Error in searchItems: ' + error);
    return {success: false, message: 'Error: ' + error.toString(), items: []};
  }
}

// ========================================
// GET ITEM BY ID (for edit)
// ========================================

/**
 * Get single item by ID
 */
function getItemById(itemId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Items');
    
    if (!sheet) {
      return {success: false, message: 'Items sheet not found'};
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === itemId) {
        var item = {
          itemId: data[i][0],
          modelNumber: data[i][1] || '',
          itemName: data[i][2] || '',
          category: data[i][3] || '',
          description: data[i][4] || '',
          reorderLevel: parseFloat(data[i][5]) || 0,
          dateAdded: toISOString(data[i][6])
        };
        
        return {success: true, item: item};
      }
    }
    
    return {success: false, message: 'Item not found'};
    
  } catch (error) {
    Logger.log('❌ Error in getItemById: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// HELPER FUNCTIONS
// ========================================

/**
 * Convert Google Sheets date to ISO string
 */
function toISOString(date) {
  if (!date) return '';
  try {
    if (typeof date === 'string') return date;
    if (date instanceof Date) return date.toISOString();
    // Handle serial date numbers
    var d = new Date(date);
    return d.toISOString();
  } catch (e) {
    return '';
  }
}

/**
 * Get spreadsheet (this function should already exist in your code)
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ========================================
// TESTING FUNCTIONS
// ========================================

/**
 * Test setup - creates Items sheet
 */
function testSetupItems() {
  var result = setupItemsSheet();
  Logger.log(result.message);
}

/**
 * Test add item
 */
function testAddItem() {
  var itemData = {
    modelNumber: 'CPU-I5-12400',
    itemName: 'Intel Core i5-12400 Processor',
    category: 'Processors',
    description: '6-Core, 12-Thread CPU, 2.5GHz Base, 4.4GHz Turbo',
    reorderLevel: 10
  };
  
  var result = addItem(itemData);
  Logger.log(result.message);
  if (result.success) {
    Logger.log('Item ID: ' + result.itemId);
  }
}

/**
 * Test get all items
 */
function testGetAllItems() {
  var result = getAllItems();
  Logger.log('Success: ' + result.success);
  Logger.log('Items count: ' + result.items.length);
  
  if (result.items.length > 0) {
    Logger.log('First item:');
    Logger.log(JSON.stringify(result.items[0], null, 2));
  }
}

/**
 * Full test sequence
 */
function runItemsTest() {
  Logger.log('=== ITEMS BACKEND TEST ===');
  
  Logger.log('\n1. Setup Items sheet...');
  setupItemsSheet();
  
  Logger.log('\n2. Get next Item ID...');
  var itemId = getNextItemId();
  Logger.log('Next ID: ' + itemId);
  
  Logger.log('\n3. Add test items...');
  var items = [
    {
      modelNumber: 'CPU-I5-12400',
      itemName: 'Intel Core i5-12400',
      category: 'Processors',
      description: '6-Core Processor',
      reorderLevel: 10
    },
    {
      modelNumber: 'MEM-8GB-DDR4',
      itemName: '8GB DDR4 RAM Module',
      category: 'Memory',
      description: 'DDR4 3200MHz',
      reorderLevel: 20
    },
    {
      modelNumber: 'SSD-1TB-NVME',
      itemName: '1TB NVMe SSD',
      category: 'Storage',
      description: 'PCIe 4.0 NVMe',
      reorderLevel: 5
    }
  ];
  
  items.forEach(function(item) {
    var result = addItem(item);
    Logger.log(result.message);
  });
  
  Logger.log('\n4. Get all items...');
  var allItems = getAllItems();
  Logger.log('Total items: ' + allItems.items.length);
  
  Logger.log('\n5. Search test...');
  var searchResult = searchItems('processor');
  Logger.log('Found: ' + searchResult.items.length + ' items');
  
  Logger.log('\n=== TEST COMPLETE ===');
}