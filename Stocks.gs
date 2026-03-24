/**
 * ========================================
 * STOCKS MANAGEMENT (NEW 10-COLUMN STRUCTURE)
 * Item ID | Category | Name | Purchase | Sale | Qty | Reorder Level | Reorder Required
 * ========================================
 */

// Setup Stocks Sheet with new structure
function setupStocksSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Stocks');
    
    if (!sheet) {
      sheet = ss.insertSheet('Stocks');
    }
    
    sheet.clear();
    
    var headers = [
      'Item ID',
      'Item Category', 
      'Item Name',
      'Purchase Price',
      'Sale Price',
      'Stock QTY',
      'Reorder Level',
      'Reorder Required'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    // Set column widths
    sheet.setColumnWidth(1, 100);  // Item ID
    sheet.setColumnWidth(2, 150);  // Category
    sheet.setColumnWidth(3, 250);  // Item Name
    sheet.setColumnWidth(4, 120);  // Purchase Price
    sheet.setColumnWidth(5, 120);  // Sale Price
    sheet.setColumnWidth(6, 100);  // Stock QTY
    sheet.setColumnWidth(7, 120);  // Reorder Level
    sheet.setColumnWidth(8, 120);  // Reorder Required
    
    return {
      success: true,
      message: 'Stocks sheet created successfully!'
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

// Get next Item ID (IT001, IT002, etc.)
function getNextItemId() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Stocks');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return 'IT001';
    }
    
    var lastRow = sheet.getLastRow();
    var lastId = sheet.getRange(lastRow, 1).getValue();
    
    if (!lastId) {
      return 'IT001';
    }
    
    // Extract number from IT001, IT002 format
    var numPart = parseInt(lastId.replace('IT', '')) || 0;
    var nextNum = numPart + 1;
    
    // Format: IT001, IT002... IT999, IT1000 (no padding after 999)
    if (nextNum <= 999) {
      return 'IT' + String(nextNum).padStart(3, '0');
    } else {
      return 'IT' + nextNum;
    }
  } catch (error) {
    return 'IT001';
  }
}

// Add new stock item
function addStock(stockData) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Stocks');
    
    if (!sheet) {
      var setupResult = setupStocksSheet();
      if (!setupResult.success) {
        return setupResult;
      }
      sheet = ss.getSheetByName('Stocks');
    }
    
    var itemId = getNextItemId();
    var stockQty = parseInt(stockData.quantity) || 0;
    var reorderLevel = parseInt(stockData.reorderLevel) || 0;
    var reorderRequired = stockQty < reorderLevel ? 'Yes' : 'No';
    
    var row = [
      itemId,
      stockData.category || 'Uncategorized',
      stockData.itemName,
      parseFloat(stockData.purchasePrice) || 0,
      parseFloat(stockData.salePrice) || 0,
      stockQty,
      reorderLevel,
      reorderRequired
    ];
    
    sheet.appendRow(row);
    
    return {
      success: true,
      message: 'Stock item added successfully',
      itemId: itemId
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

// Get all stocks
function getStocks() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Stocks');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        success: true,
        stocks: []
      };
    }
    
    var data = sheet.getDataRange().getValues();
    var stocks = [];
    
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // Skip if no Item ID
      
      var stockQty = parseInt(data[i][5]) || 0;
      var reorderLevel = parseInt(data[i][6]) || 0;
      var reorderRequired = stockQty < reorderLevel ? 'Yes' : 'No';
      
      stocks.push({
        itemId: data[i][0],
        category: data[i][1] || 'Uncategorized',
        itemName: data[i][2],
        purchasePrice: data[i][3],
        salePrice: data[i][4],
        quantity: stockQty,
        reorderLevel: reorderLevel,
        reorderRequired: reorderRequired
      });
    }
    
    return {
      success: true,
      stocks: stocks
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString(),
      stocks: []
    };
  }
}

// Update stock item
function updateStock(itemId, stockData) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Stocks');
    
    if (!sheet) {
      return {
        success: false,
        message: 'Stocks sheet not found'
      };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === itemId) {
        var rowIndex = i + 1;
        
        var stockQty = parseInt(stockData.quantity) || 0;
        var reorderLevel = parseInt(stockData.reorderLevel) || 0;
        var reorderRequired = stockQty < reorderLevel ? 'Yes' : 'No';
        
        // Item ID cannot be changed (column 1)
        sheet.getRange(rowIndex, 2).setValue(stockData.category || 'Uncategorized');
        sheet.getRange(rowIndex, 3).setValue(stockData.itemName);
        sheet.getRange(rowIndex, 4).setValue(parseFloat(stockData.purchasePrice) || 0);
        sheet.getRange(rowIndex, 5).setValue(parseFloat(stockData.salePrice) || 0);
        sheet.getRange(rowIndex, 6).setValue(stockQty);
        sheet.getRange(rowIndex, 7).setValue(reorderLevel);
        sheet.getRange(rowIndex, 8).setValue(reorderRequired);
        
        return {
          success: true,
          message: 'Stock item updated successfully'
        };
      }
    }
    
    return {
      success: false,
      message: 'Stock item not found'
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

// Delete stock item
function deleteStock(itemId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Stocks');
    
    if (!sheet) {
      return {
        success: false,
        message: 'Stocks sheet not found'
      };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === itemId) {
        sheet.deleteRow(i + 1);
        return {
          success: true,
          message: 'Stock item deleted successfully'
        };
      }
    }
    
    return {
      success: false,
      message: 'Stock item not found'
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

// Get stock by Item ID
function getStockByItemId(itemId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Stocks');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        success: false,
        message: 'No stocks found'
      };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === itemId) {
        var stockQty = parseInt(data[i][5]) || 0;
        var reorderLevel = parseInt(data[i][6]) || 0;
        var reorderRequired = stockQty < reorderLevel ? 'Yes' : 'No';
        
        return {
          success: true,
          stock: {
            itemId: data[i][0],
            category: data[i][1] || 'Uncategorized',
            itemName: data[i][2],
            purchasePrice: data[i][3],
            salePrice: data[i][4],
            quantity: stockQty,
            reorderLevel: reorderLevel,
            reorderRequired: reorderRequired
          }
        };
      }
    }
    
    return {
      success: false,
      message: 'Stock item not found'
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

// BACKWARD COMPATIBILITY FUNCTION for other pages
// Returns old format: itemName, purchasePrice, salePrice, quantity, description
function getStocksOldFormat() {
  try {
    var result = getStocks();
    
    if (!result.success) {
      return result;
    }
    
    var oldFormatStocks = [];
    
    for (var i = 0; i < result.stocks.length; i++) {
      var stock = result.stocks[i];
      oldFormatStocks.push({
        itemName: stock.itemName,
        purchasePrice: stock.purchasePrice,
        salePrice: stock.salePrice,
        quantity: stock.quantity,
        description: stock.category // Use category as description for compatibility
      });
    }
    
    return {
      success: true,
      stocks: oldFormatStocks
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString(),
      stocks: []
    };
  }
}