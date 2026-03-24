/**
 * Barcode Scanning System
 * Hybrid approach: Barcode Mapping + History Tracking
 */

// Setup Barcode Sheets
function setupBarcodeSheets() {
  try {
    var ss = getSpreadsheet();
    var result = {
      mappingSheet: false,
      historySheet: false
    };
    
    // 1. Barcode_Mapping Sheet
    var mappingSheet = ss.getSheetByName('Barcode_Mapping');
    if (!mappingSheet) {
      mappingSheet = ss.insertSheet('Barcode_Mapping');
    }
    mappingSheet.clear();
    var headers1 = ['Barcode', 'Item Name', 'Product Code', 'Category', 'Created Date', 'Created By'];
    mappingSheet.getRange(1, 1, 1, headers1.length).setValues([headers1]);
    mappingSheet.getRange(1, 1, 1, headers1.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    result.mappingSheet = true;
    
    // 2. Barcode_History Sheet
    var historySheet = ss.getSheetByName('Barcode_History');
    if (!historySheet) {
      historySheet = ss.insertSheet('Barcode_History');
    }
    historySheet.clear();
    var headers2 = ['Timestamp', 'Barcode', 'Item Name', 'Action', 'User', 'Reference', 'Location', 'Notes'];
    historySheet.getRange(1, 1, 1, headers2.length).setValues([headers2]);
    historySheet.getRange(1, 1, 1, headers2.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    result.historySheet = true;
    
    return {
      success: true,
      message: 'Barcode sheets created successfully!',
      details: result
    };
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Lookup barcode - returns item info if mapped, null if not
function lookupBarcode(barcode) {
  try {
    var ss = getSpreadsheet();
    var mappingSheet = ss.getSheetByName('Barcode_Mapping');
    
    if (!mappingSheet || mappingSheet.getLastRow() <= 1) {
      return {success: true, found: false, barcode: barcode};
    }
    
    var data = mappingSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === barcode.toString()) {
        return {
          success: true,
          found: true,
          barcode: data[i][0],
          itemName: data[i][1],
          productCode: data[i][2] || '',
          category: data[i][3] || '',
          createdDate: data[i][4],
          createdBy: data[i][5]
        };
      }
    }
    
    return {success: true, found: false, barcode: barcode};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Map new barcode to product
function mapBarcode(data) {
  try {
    var ss = getSpreadsheet();
    var mappingSheet = ss.getSheetByName('Barcode_Mapping');
    
    if (!mappingSheet) {
      return {success: false, message: 'Barcode_Mapping sheet not found. Run setupBarcodeSheets() first.'};
    }
    
    // Check if barcode already exists
    var existing = lookupBarcode(data.barcode);
    if (existing.found) {
      return {success: false, message: 'This barcode is already mapped to: ' + existing.itemName};
    }
    
    // Verify item exists in Stocks
    var stocksSheet = ss.getSheetByName('Stocks');
    if (!stocksSheet) {
      return {success: false, message: 'Stocks sheet not found'};
    }
    
    var stockData = stocksSheet.getDataRange().getValues();
    var itemExists = false;
    
    for (var i = 1; i < stockData.length; i++) {
      if (stockData[i][0] && stockData[i][0].toString().toLowerCase() === data.itemName.toString().toLowerCase()) {
        itemExists = true;
        break;
      }
    }
    
    if (!itemExists) {
      return {success: false, message: 'Item "' + data.itemName + '" not found in Stocks. Please add it first.'};
    }
    
    var now = new Date().toISOString();
    
    var row = [
      data.barcode,
      data.itemName,
      data.productCode || '',
      data.category || '',
      now,
      data.createdBy || 'User'
    ];
    
    mappingSheet.appendRow(row);
    
    // Log to history
    logBarcodeAction({
      barcode: data.barcode,
      itemName: data.itemName,
      action: 'Mapped',
      user: data.createdBy || 'User',
      reference: 'Barcode Mapping',
      location: '',
      notes: 'New barcode mapped to product'
    });
    
    return {
      success: true,
      message: 'Barcode mapped successfully',
      barcode: data.barcode,
      itemName: data.itemName
    };
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Log barcode action to history
function logBarcodeAction(data) {
  try {
    var ss = getSpreadsheet();
    var historySheet = ss.getSheetByName('Barcode_History');
    
    if (!historySheet) {
      return {success: false, message: 'Barcode_History sheet not found'};
    }
    
    var now = new Date().toISOString();
    
    var row = [
      now,
      data.barcode || '',
      data.itemName || '',
      data.action || '',
      data.user || 'User',
      data.reference || '',
      data.location || '',
      data.notes || ''
    ];
    
    historySheet.appendRow(row);
    
    return {success: true};
  } catch (error) {
    Logger.log('Error logging barcode action: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get barcode history
function getBarcodeHistory(barcode) {
  try {
    var ss = getSpreadsheet();
    var historySheet = ss.getSheetByName('Barcode_History');
    
    if (!historySheet || historySheet.getLastRow() <= 1) {
      return {success: true, history: []};
    }
    
    var data = historySheet.getDataRange().getValues();
    var history = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === barcode.toString()) {
        history.push({
          timestamp: data[i][0],
          barcode: data[i][1],
          itemName: data[i][2],
          action: data[i][3],
          user: data[i][4],
          reference: data[i][5],
          location: data[i][6],
          notes: data[i][7]
        });
      }
    }
    
    return {success: true, history: history};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString(), history: []};
  }
}

// Get all barcode mappings
function getAllBarcodeMappings() {
  try {
    var ss = getSpreadsheet();
    var mappingSheet = ss.getSheetByName('Barcode_Mapping');
    
    if (!mappingSheet || mappingSheet.getLastRow() <= 1) {
      return {success: true, mappings: []};
    }
    
    var data = mappingSheet.getDataRange().getValues();
    var mappings = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        mappings.push({
          barcode: data[i][0],
          itemName: data[i][1],
          productCode: data[i][2] || '',
          category: data[i][3] || '',
          createdDate: data[i][4],
          createdBy: data[i][5]
        });
      }
    }
    
    return {success: true, mappings: mappings};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString(), mappings: []};
  }
}

// Delete barcode mapping
function deleteBarcode(barcode) {
  try {
    var ss = getSpreadsheet();
    var mappingSheet = ss.getSheetByName('Barcode_Mapping');
    
    if (!mappingSheet) {
      return {success: false, message: 'Barcode_Mapping sheet not found'};
    }
    
    var data = mappingSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === barcode.toString()) {
        mappingSheet.deleteRow(i + 1);
        
        // Log deletion
        logBarcodeAction({
          barcode: barcode,
          itemName: data[i][1],
          action: 'Deleted',
          user: 'Admin',
          reference: 'Barcode Mapping',
          notes: 'Barcode mapping deleted'
        });
        
        return {success: true, message: 'Barcode mapping deleted successfully'};
      }
    }
    
    return {success: false, message: 'Barcode not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Process barcode scan for purchase
function processBarcodeForPurchase(data) {
  try {
    // 1. Lookup barcode
    var lookup = lookupBarcode(data.barcode);
    
    if (!lookup.success) {
      return lookup;
    }
    
    if (!lookup.found) {
      // Return barcode not mapped - frontend will handle
      return {
        success: true,
        needsMapping: true,
        barcode: data.barcode,
        message: 'Barcode not mapped. Please select the product.'
      };
    }
    
    // 2. Get item details from Stocks
    var ss = getSpreadsheet();
    var stocksSheet = ss.getSheetByName('Stocks');
    
    if (!stocksSheet) {
      return {success: false, message: 'Stocks sheet not found'};
    }
    
    var stockData = stocksSheet.getDataRange().getValues();
    var itemInfo = null;
    
    for (var i = 1; i < stockData.length; i++) {
      if (stockData[i][0] && stockData[i][0].toString().toLowerCase() === lookup.itemName.toLowerCase()) {
        itemInfo = {
          itemName: stockData[i][0],
          purchasePrice: stockData[i][1] || 0,
          salePrice: stockData[i][2] || 0,
          quantity: stockData[i][3] || 0,
          description: stockData[i][4] || ''
        };
        break;
      }
    }
    
    if (!itemInfo) {
      return {success: false, message: 'Item not found in Stocks: ' + lookup.itemName};
    }
    
    // 3. Log scan action
    logBarcodeAction({
      barcode: data.barcode,
      itemName: lookup.itemName,
      action: 'Scanned for Purchase',
      user: data.user || 'User',
      reference: data.purchaseNumber || 'Purchase',
      location: data.location || '',
      notes: 'Barcode scanned during purchase entry'
    });
    
    // 4. Return item info
    return {
      success: true,
      found: true,
      needsMapping: false,
      barcode: data.barcode,
      itemInfo: itemInfo
    };
    
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Process barcode scan for sale/quote
function processBarcodeForSale(data) {
  try {
    // Same as purchase but different action logging
    var lookup = lookupBarcode(data.barcode);
    
    if (!lookup.success) {
      return lookup;
    }
    
    if (!lookup.found) {
      return {
        success: true,
        needsMapping: true,
        barcode: data.barcode,
        message: 'Barcode not mapped. Please select the product.'
      };
    }
    
    // Get item details from Stocks
    var ss = getSpreadsheet();
    var stocksSheet = ss.getSheetByName('Stocks');
    
    if (!stocksSheet) {
      return {success: false, message: 'Stocks sheet not found'};
    }
    
    var stockData = stocksSheet.getDataRange().getValues();
    var itemInfo = null;
    
    for (var i = 1; i < stockData.length; i++) {
      if (stockData[i][0] && stockData[i][0].toString().toLowerCase() === lookup.itemName.toLowerCase()) {
        itemInfo = {
          itemName: stockData[i][0],
          purchasePrice: stockData[i][1] || 0,
          salePrice: stockData[i][2] || 0,
          quantity: stockData[i][3] || 0,
          description: stockData[i][4] || ''
        };
        break;
      }
    }
    
    if (!itemInfo) {
      return {success: false, message: 'Item not found in Stocks: ' + lookup.itemName};
    }
    
    // Log scan action
    logBarcodeAction({
      barcode: data.barcode,
      itemName: lookup.itemName,
      action: 'Scanned for Sale',
      user: data.user || 'User',
      reference: data.quoteNumber || 'Quote',
      location: data.location || '',
      notes: 'Barcode scanned during quote/sale entry'
    });
    
    return {
      success: true,
      found: true,
      needsMapping: false,
      barcode: data.barcode,
      itemInfo: itemInfo
    };
    
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get barcode statistics
function getBarcodeStats() {
  try {
    var ss = getSpreadsheet();
    var mappingSheet = ss.getSheetByName('Barcode_Mapping');
    var historySheet = ss.getSheetByName('Barcode_History');
    
    var stats = {
      totalMappings: 0,
      totalScans: 0,
      recentScans: []
    };
    
    if (mappingSheet && mappingSheet.getLastRow() > 1) {
      stats.totalMappings = mappingSheet.getLastRow() - 1;
    }
    
    if (historySheet && historySheet.getLastRow() > 1) {
      stats.totalScans = historySheet.getLastRow() - 1;
      
      // Get last 10 scans
      var data = historySheet.getDataRange().getValues();
      for (var i = Math.max(1, data.length - 10); i < data.length; i++) {
        stats.recentScans.push({
          timestamp: data[i][0],
          barcode: data[i][1],
          itemName: data[i][2],
          action: data[i][3],
          user: data[i][4]
        });
      }
      stats.recentScans.reverse();
    }
    
    return {success: true, stats: stats};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}