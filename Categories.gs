/**
 * ========================================
 * CATEGORIES MANAGEMENT (SIMPLE VERSION)
 * Just manages category names for inventory
 * ========================================
 */

// Setup Categories Sheet
function setupCategoriesSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Categories');
    
    if (!sheet) {
      sheet = ss.insertSheet('Categories');
    }
    
    sheet.clear();
    
    var headers = ['Category Name', 'Created Date'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    sheet.setColumnWidth(1, 250); // Category Name
    sheet.setColumnWidth(2, 150); // Created Date
    
    return {
      success: true,
      message: 'Categories sheet created successfully!'
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

// Add new category
function addCategory(categoryName) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Categories');
    
    if (!sheet) {
      var setupResult = setupCategoriesSheet();
      if (!setupResult.success) {
        return setupResult;
      }
      sheet = ss.getSheetByName('Categories');
    }
    
    // Check if category already exists
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === categoryName.toLowerCase()) {
        return {
          success: false,
          message: 'Category "' + categoryName + '" already exists'
        };
      }
    }
    
    var now = new Date().toISOString();
    var row = [categoryName, now];
    
    sheet.appendRow(row);
    
    return {
      success: true,
      message: 'Category added successfully'
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

// Get all categories
function getCategories() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Categories');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        success: true,
        categories: []
      };
    }
    
    var data = sheet.getDataRange().getValues();
    var categories = [];
    
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      
      categories.push({
        categoryId: 'CAT-' + String(i).padStart(4, '0'),
        categoryName: data[i][0],
        createdDate: data[i][1] || ''
      });
    }
    
    return {
      success: true,
      categories: categories
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString(),
      categories: []
    };
  }
}

// Delete category
function deleteCategory(categoryName) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Categories');
    
    if (!sheet) {
      return {
        success: false,
        message: 'Categories sheet not found'
      };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === categoryName) {
        sheet.deleteRow(i + 1);
        return {
          success: true,
          message: 'Category deleted successfully'
        };
      }
    }
    
    return {
      success: false,
      message: 'Category not found'
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}