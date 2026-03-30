/**
 * ========================================
 * SERVICES MANAGEMENT - BACKEND
 * ========================================
 * Manages service codes and names for the system
 * 
 * Features:
 * - Auto-incrementing Service Code (SRV-0001, SRV-0002...)
 * - Service Name storage
 * - CRUD operations
 * - Search functionality
 */

// ========================================
// SETUP SERVICES SHEET
// ========================================

/**
 * Setup Services sheet with proper structure
 */
function setupServicesSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Services');
    
    if (!sheet) {
      sheet = ss.insertSheet('Services');
    } else {
      // Don't clear if sheet exists - preserve data
      // Only update headers if needed
    }
    
    // Headers: Service Code, Service Name, Created Date
    var headers = ['Service Code', 'Service Name', 'Created Date'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // Set column widths
    sheet.setColumnWidth(1, 150);  // Service Code
    sheet.setColumnWidth(2, 400);  // Service Name
    sheet.setColumnWidth(3, 180);  // Created Date
    
    Logger.log('✅ Services sheet setup complete');
    return {success: true, message: 'Services sheet created successfully'};
    
  } catch (error) {
    Logger.log('❌ Error in setupServicesSheet: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// SERVICE CODE GENERATION
// ========================================

/**
 * Get next Service Code (SRV-0001, SRV-0002, etc.)
 */
function getNextServiceCode() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Services');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return 'SRV-0001';
    }
    
    var data = sheet.getDataRange().getValues();
    var maxNum = 0;
    
    // Find highest service number
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var code = String(data[i][0]);
        if (code.startsWith('SRV-')) {
          var num = parseInt(code.replace('SRV-', '')) || 0;
          if (num > maxNum) {
            maxNum = num;
          }
        }
      }
    }
    
    var nextNum = maxNum + 1;
    return 'SRV-' + String(nextNum).padStart(4, '0');
    
  } catch (error) {
    Logger.log('❌ Error in getNextServiceCode: ' + error);
    return 'SRV-0001';
  }
}

// ========================================
// ADD SERVICE
// ========================================

/**
 * Add a new service
 * Returns the auto-generated service code
 */
function addService(serviceName) {
  try {
    Logger.log('🔧 Adding service: ' + serviceName);
    
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Services');
    
    if (!sheet) {
      var setupResult = setupServicesSheet();
      if (!setupResult.success) {
        return setupResult;
      }
      sheet = ss.getSheetByName('Services');
    }
    
    // Validate service name
    if (!serviceName || serviceName.trim() === '') {
      return {success: false, message: 'Service name is required'};
    }
    
    serviceName = serviceName.trim();
    
    // Check for duplicates
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().toLowerCase() === serviceName.toLowerCase()) {
        return {
          success: false,
          message: 'Service "' + serviceName + '" already exists'
        };
      }
    }
    
    // Generate service code
    var serviceCode = getNextServiceCode();
    var createdDate = new Date().toISOString();
    
    // Add row
    var rowData = [serviceCode, serviceName, createdDate];
    sheet.appendRow(rowData);
    
    Logger.log('✅ Service added: ' + serviceCode + ' - ' + serviceName);
    
    return {
      success: true,
      message: 'Service added successfully',
      serviceCode: serviceCode,
      serviceName: serviceName,
      createdDate: createdDate
    };
    
  } catch (error) {
    Logger.log('❌ Error in addService: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// GET ALL SERVICES
// ========================================

/**
 * Get all services from the sheet
 */
function getAllServices() {
  try {
    Logger.log('🔍 Getting all services');
    
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Services');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      Logger.log('⚠️ Services sheet empty or not found');
      return {success: true, services: []};
    }
    
    var data = sheet.getDataRange().getValues();
    var services = [];
    
    // Skip header row
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {  // Check if Service Code exists
        services.push({
          serviceCode: data[i][0],
          serviceName: data[i][1] || '',
          createdDate: data[i][2] || '',
          rowIndex: i + 1  // Store for edit/delete
        });
      }
    }
    
    Logger.log('✅ Retrieved ' + services.length + ' services');
    return {success: true, services: services};
    
  } catch (error) {
    Logger.log('❌ Error in getAllServices: ' + error);
    return {success: false, message: 'Error: ' + error.toString(), services: []};
  }
}

// ========================================
// SEARCH SERVICES
// ========================================

/**
 * Search services by code or name
 */
function searchServices(searchTerm) {
  try {
    Logger.log('🔍 Searching services: ' + searchTerm);
    
    if (!searchTerm || searchTerm.trim() === '') {
      return getAllServices();
    }
    
    var result = getAllServices();
    if (!result.success) {
      return result;
    }
    
    var searchLower = searchTerm.toLowerCase().trim();
    var filtered = [];
    
    for (var i = 0; i < result.services.length; i++) {
      var service = result.services[i];
      
      if (service.serviceCode.toLowerCase().indexOf(searchLower) !== -1 ||
          service.serviceName.toLowerCase().indexOf(searchLower) !== -1) {
        filtered.push(service);
      }
    }
    
    Logger.log('✅ Search found ' + filtered.length + ' results');
    return {success: true, services: filtered};
    
  } catch (error) {
    Logger.log('❌ Error in searchServices: ' + error);
    return {success: false, message: 'Error: ' + error.toString(), services: []};
  }
}

// ========================================
// UPDATE SERVICE
// ========================================

/**
 * Update an existing service
 */
function updateService(serviceCode, newServiceName) {
  try {
    Logger.log('🔄 Updating service: ' + serviceCode);
    
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Services');
    
    if (!sheet) {
      return {success: false, message: 'Services sheet not found'};
    }
    
    if (!newServiceName || newServiceName.trim() === '') {
      return {success: false, message: 'Service name is required'};
    }
    
    newServiceName = newServiceName.trim();
    
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    // Find the service
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === serviceCode) {
        rowIndex = i + 1;  // 1-indexed for sheet
        break;
      }
    }
    
    if (rowIndex === -1) {
      return {success: false, message: 'Service not found'};
    }
    
    // Check for duplicate name (excluding current service)
    for (var j = 1; j < data.length; j++) {
      if ((j + 1) !== rowIndex && 
          data[j][1] && 
          data[j][1].toString().toLowerCase() === newServiceName.toLowerCase()) {
        return {
          success: false,
          message: 'Service name "' + newServiceName + '" already exists'
        };
      }
    }
    
    // Update service name
    sheet.getRange(rowIndex, 2).setValue(newServiceName);
    
    Logger.log('✅ Service updated: ' + serviceCode);
    
    return {
      success: true,
      message: 'Service updated successfully'
    };
    
  } catch (error) {
    Logger.log('❌ Error in updateService: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// DELETE SERVICE
// ========================================

/**
 * Delete a service
 */
function deleteService(serviceCode) {
  try {
    Logger.log('🗑️ Deleting service: ' + serviceCode);
    
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Services');
    
    if (!sheet) {
      return {success: false, message: 'Services sheet not found'};
    }
    
    var data = sheet.getDataRange().getValues();
    
    // Find and delete the row
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === serviceCode) {
        sheet.deleteRow(i + 1);
        Logger.log('✅ Service deleted: ' + serviceCode);
        return {
          success: true,
          message: 'Service deleted successfully'
        };
      }
    }
    
    return {success: false, message: 'Service not found'};
    
  } catch (error) {
    Logger.log('❌ Error in deleteService: ' + error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ========================================
// HELPER FUNCTION
// ========================================

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}