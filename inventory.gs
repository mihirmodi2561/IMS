/**
 * ============================================
 * INVENTORY MANAGEMENT - BACKEND
 * ============================================
 * Read-only inventory system that calculates stock from:
 * - Items sheet (master product data)
 * - DetailPO sheet (all purchases with locations)
 * 
 * Features:
 * - Real-time stock calculation
 * - Stock by location (10 warehouses)
 * - Purchase history per item
 * - FIFO stock valuation
 * - Reorder alerts
 */

// ========================================
// GET INVENTORY SUMMARY
// ========================================

/**
 * Get complete inventory with stock calculated from purchases
 */
function getInventorySummary() {
  try {
    Logger.log('🔍 getInventorySummary() started');
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('❌ Spreadsheet is null');
      return {success: false, message: 'Cannot access spreadsheet', inventory: []};
    }
    
    var itemsSheet = ss.getSheetByName('Items');
    
    if (!itemsSheet) {
      Logger.log('❌ Items sheet not found');
      return {success: false, message: 'Items sheet not found', inventory: []};
    }
    
    Logger.log('✅ Items sheet found');
    
    // Get all items
    var itemsData = itemsSheet.getDataRange().getValues();
    Logger.log('📄 Retrieved ' + itemsData.length + ' items');
    
    var inventory = [];
    
    // Calculate stock from DetailPO (inline - no external function)
    var stockData = {};
    try {
      var detailSheet = ss.getSheetByName('DetailPO');
      
      if (detailSheet && detailSheet.getLastRow() > 1) {
        Logger.log('📦 Processing DetailPO for stock calculation');
        var detailData = detailSheet.getDataRange().getValues();
        
        // DetailPO columns: Date(0), PO#(1), DetailID(2), SupID(3), SupName(4), BillNum(5),
        //                   Model(6), Name(7), Cat(8), Location(9), QTY(10), Cost(11)...
        
        for (var j = 1; j < detailData.length; j++) {
          var modelNumber = String(detailData[j][6] || '');
          var location = String(detailData[j][9] || 'Unknown');
          var qty = parseFloat(detailData[j][10]) || 0;
          
          if (modelNumber) {
            // Initialize if doesn't exist
            if (!stockData[modelNumber]) {
              stockData[modelNumber] = {
                totalStock: 0,
                locations: {},
                purchaseHistory: []
              };
            }
            
            // Add to total
            stockData[modelNumber].totalStock += qty;
            
            // Add to location
            if (!stockData[modelNumber].locations[location]) {
              stockData[modelNumber].locations[location] = 0;
            }
            stockData[modelNumber].locations[location] += qty;
            
            // Add to purchase history (simplified)
            stockData[modelNumber].purchaseHistory.push({
              date: String(detailData[j][0] || ''),
              qty: qty,
              location: location
            });
          }
        }
        Logger.log('✅ Stock calculated for ' + Object.keys(stockData).length + ' models');
      } else {
        Logger.log('⚠️ DetailPO sheet is empty or not found');
      }
    } catch (stockError) {
      Logger.log('⚠️ Error calculating stock: ' + stockError);
    }
    
    // Process each item
    for (var i = 1; i < itemsData.length; i++) {
      if (!itemsData[i][0]) continue;
      
      var itemId = String(itemsData[i][0]);
      var modelNumber = String(itemsData[i][1] || '');
      var itemName = String(itemsData[i][2] || '');
      var category = String(itemsData[i][3] || '');
      var reorderLevel = parseFloat(itemsData[i][5]) || 0;
      
      // Get stock for this item
      var itemStock = stockData[modelNumber] || {
        totalStock: 0,
        locations: {},
        purchaseHistory: []
      };
      
      // Calculate reorder status
      var reorderRequired = itemStock.totalStock < reorderLevel ? 'Yes' : 'No';
      
      inventory.push({
        itemId: itemId,
        modelNumber: modelNumber,
        itemName: itemName,
        category: category,
        reorderLevel: reorderLevel,
        totalStock: itemStock.totalStock,
        stockByLocation: itemStock.locations,
        reorderRequired: reorderRequired,
        purchaseHistory: itemStock.purchaseHistory
      });
    }
    
    Logger.log('✅ Generated inventory for ' + inventory.length + ' items');
    
    var result = {
      success: true,
      inventory: inventory
    };
    
    return result;
    
  } catch (error) {
    Logger.log('❌ CRITICAL ERROR in getInventorySummary: ' + error);
    Logger.log('❌ Error stack: ' + error.stack);
    return {success: false, message: String(error), inventory: []};
  }
}

// ========================================
// CALCULATE STOCK FROM PURCHASES
// ========================================

/**
 * Calculate stock quantities from DetailPO sheet
 * Returns: {modelNumber: {totalStock, locations: {WH-A: qty}, purchaseHistory: [...]}}
 */
function calculateStockFromPurchases() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('DetailPO');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {};
    }
    
    var data = sheet.getDataRange().getValues();
    var stockData = {};
    
    // Process each purchase record
    // Columns: Date(0), PO#(1), DetailID(2), SupID(3), SupName(4), BillNum(5),
    //          Model(6), Name(7), Cat(8), Location(9), QTY(10), Cost(11)...
    
    for (var i = 1; i < data.length; i++) {
      var modelNumber = data[i][6];
      var location = data[i][9] || 'Unknown';
      var qty = parseFloat(data[i][10]) || 0;
      var unitCost = parseFloat(data[i][11]) || 0;
      
      if (!modelNumber) continue;
      
      // Initialize if not exists
      if (!stockData[modelNumber]) {
        stockData[modelNumber] = {
          totalStock: 0,
          locations: {},
          purchaseHistory: []
        };
      }
      
      // Add to total stock
      stockData[modelNumber].totalStock += qty;
      
      // Add to location stock
      if (!stockData[modelNumber].locations[location]) {
        stockData[modelNumber].locations[location] = 0;
      }
      stockData[modelNumber].locations[location] += qty;
      
      // Add to purchase history (for FIFO and details)
      stockData[modelNumber].purchaseHistory.push({
        date: toISOString(data[i][0]),
        poNumber: data[i][1],
        supplier: data[i][4],
        location: location,
        qty: qty,
        unitCost: unitCost,
        total: qty * unitCost
      });
    }
    
    return stockData;
    
  } catch (error) {
    Logger.log('❌ Error in calculateStockFromPurchases: ' + error);
    return {};
  }
}

// ========================================
// GET ITEM DETAILS
// ========================================

/**
 * Get detailed information for a specific item
 */
function getItemInventoryDetails(modelNumber) {
  try {
    var ss = getSpreadsheet();
    var itemsSheet = ss.getSheetByName('Items');
    
    if (!itemsSheet) {
      return {success: false, message: 'Items sheet not found'};
    }
    
    // Find item in Items sheet
    var itemsData = itemsSheet.getDataRange().getValues();
    var item = null;
    
    for (var i = 1; i < itemsData.length; i++) {
      if (itemsData[i][1] === modelNumber) {
        item = {
          itemId: itemsData[i][0],
          modelNumber: itemsData[i][1],
          itemName: itemsData[i][2],
          category: itemsData[i][3],
          description: itemsData[i][4] || '',
          reorderLevel: parseFloat(itemsData[i][5]) || 0
        };
        break;
      }
    }
    
    if (!item) {
      return {success: false, message: 'Item not found'};
    }
    
    // Get stock data
    var stockData = calculateStockFromPurchases();
    var itemStock = stockData[modelNumber] || {
      totalStock: 0,
      locations: {},
      purchaseHistory: []
    };
    
    // Calculate stock value (FIFO)
    var stockValue = 0;
    var remainingQty = itemStock.totalStock;
    var fifoDetails = [];
    
    // Sort purchase history by date (oldest first for FIFO)
    var sortedHistory = itemStock.purchaseHistory.sort(function(a, b) {
      return new Date(a.date) - new Date(b.date);
    });
    
    for (var j = 0; j < sortedHistory.length && remainingQty > 0; j++) {
      var purchase = sortedHistory[j];
      var qtyToUse = Math.min(remainingQty, purchase.qty);
      stockValue += qtyToUse * purchase.unitCost;
      
      fifoDetails.push({
        qty: qtyToUse,
        unitCost: purchase.unitCost,
        subtotal: qtyToUse * purchase.unitCost
      });
      
      remainingQty -= qtyToUse;
    }
    
    // Prepare result
    var result = {
      item: item,
      totalStock: itemStock.totalStock,
      stockByLocation: itemStock.locations,
      reorderRequired: itemStock.totalStock < item.reorderLevel ? 'Yes' : 'No',
      stockValue: stockValue,
      averageCost: itemStock.totalStock > 0 ? stockValue / itemStock.totalStock : 0,
      purchaseHistory: sortedHistory,
      fifoBreakdown: fifoDetails
    };
    
    return {success: true, details: result};
    
  } catch (error) {
    Logger.log('❌ Error in getItemInventoryDetails: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// SEARCH INVENTORY
// ========================================

/**
 * Search inventory by model, name, or category
 */
function searchInventory(searchTerm) {
  try {
    if (!searchTerm || searchTerm.trim() === '') {
      return getInventorySummary();
    }
    
    var result = getInventorySummary();
    if (!result.success) {
      return result;
    }
    
    var searchLower = searchTerm.toLowerCase().trim();
    var filtered = [];
    
    for (var i = 0; i < result.inventory.length; i++) {
      var item = result.inventory[i];
      if (item.modelNumber.toLowerCase().indexOf(searchLower) !== -1 ||
          item.itemName.toLowerCase().indexOf(searchLower) !== -1 ||
          item.category.toLowerCase().indexOf(searchLower) !== -1) {
        filtered.push(item);
      }
    }
    
    Logger.log('✅ Search found ' + filtered.length + ' items');
    return {success: true, inventory: filtered};
    
  } catch (error) {
    Logger.log('❌ Error in searchInventory: ' + error);
    return {success: false, message: 'Error: ' + error.toString(), inventory: []};
  }
}

// ========================================
// GET LOW STOCK ITEMS
// ========================================

/**
 * Get all items that need reordering
 */
function getLowStockItems() {
  try {
    var result = getInventorySummary();
    if (!result.success) {
      return result;
    }
    
    var lowStock = [];
    
    for (var i = 0; i < result.inventory.length; i++) {
      if (result.inventory[i].reorderRequired === 'Yes') {
        lowStock.push(result.inventory[i]);
      }
    }
    
    Logger.log('✅ Found ' + lowStock.length + ' low stock items');
    return {success: true, items: lowStock};
    
  } catch (error) {
    Logger.log('❌ Error in getLowStockItems: ' + error);
    return {success: false, message: 'Error: ' + error.toString(), items: []};
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
// TESTING FUNCTIONS
// ========================================

/**
 * Test inventory functions
 */
function testInventory() {
  Logger.log('=== INVENTORY TEST ===');
  
  Logger.log('\n1. Get inventory summary...');
  var summary = getInventorySummary();
  Logger.log('Total items: ' + summary.inventory.length);
  
  if (summary.inventory.length > 0) {
    Logger.log('\nFirst item:');
    Logger.log(JSON.stringify(summary.inventory[0], null, 2));
    
    Logger.log('\n2. Get item details...');
    var details = getItemInventoryDetails(summary.inventory[0].modelNumber);
    if (details.success) {
      Logger.log('Stock value: $' + details.details.stockValue.toFixed(2));
      Logger.log('Average cost: $' + details.details.averageCost.toFixed(2));
    }
  }
  
  Logger.log('\n3. Get low stock items...');
  var lowStock = getLowStockItems();
  Logger.log('Low stock items: ' + lowStock.items.length);
  
  Logger.log('\n=== TEST COMPLETE ===');
}