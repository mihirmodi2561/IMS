/**
 * Technical Library Management Functions
 * Handles reference documentation, system diagrams, equipment lists, and installation procedures
 */

// Setup Technical Library Sheets
function setupTechnicalLibrarySheets() {
  try {
    var ss = getSpreadsheet();
    var result = {
      librarySheet: false,
      filesSheet: false,
      equipmentSheet: false,
      stepsSheet: false,
      scopeSheet: false
    };
    
    // 1. Technical_Library sheet
    var librarySheet = ss.getSheetByName('Technical_Library');
    if (!librarySheet) {
      librarySheet = ss.insertSheet('Technical_Library');
    }
    librarySheet.clear();
    var headers1 = ['Library ID', 'Name', 'Description', 'Created Date', 'Created By'];
    librarySheet.getRange(1, 1, 1, headers1.length).setValues([headers1]);
    librarySheet.getRange(1, 1, 1, headers1.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    result.librarySheet = true;
    
    // 2. Library_Files sheet (diagrams/images)
    var filesSheet = ss.getSheetByName('Library_Files');
    if (!filesSheet) {
      filesSheet = ss.insertSheet('Library_Files');
    }
    filesSheet.clear();
    var headers2 = ['File ID', 'Library ID', 'File Name', 'File Type', 'Base64 Data', 'Upload Date'];
    filesSheet.getRange(1, 1, 1, headers2.length).setValues([headers2]);
    filesSheet.getRange(1, 1, 1, headers2.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    result.filesSheet = true;
    
    // 3. Library_Equipment sheet (references Stocks)
    var equipmentSheet = ss.getSheetByName('Library_Equipment');
    if (!equipmentSheet) {
      equipmentSheet = ss.insertSheet('Library_Equipment');
    }
    equipmentSheet.clear();
    var headers3 = ['Entry ID', 'Library ID', 'Item Name', 'Quantity Needed'];
    equipmentSheet.getRange(1, 1, 1, headers3.length).setValues([headers3]);
    equipmentSheet.getRange(1, 1, 1, headers3.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    result.equipmentSheet = true;
    
    // 4. Library_Steps sheet (installation checklist)
    var stepsSheet = ss.getSheetByName('Library_Steps');
    if (!stepsSheet) {
      stepsSheet = ss.insertSheet('Library_Steps');
    }
    stepsSheet.clear();
    var headers4 = ['Step ID', 'Library ID', 'Step Number', 'Step Text'];
    stepsSheet.getRange(1, 1, 1, headers4.length).setValues([headers4]);
    stepsSheet.getRange(1, 1, 1, headers4.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    result.stepsSheet = true;
    
    // 5. Library_Scope sheet (scope of work text)
    var scopeSheet = ss.getSheetByName('Library_Scope');
    if (!scopeSheet) {
      scopeSheet = ss.insertSheet('Library_Scope');
    }
    scopeSheet.clear();
    var headers5 = ['Library ID', 'Scope Text'];
    scopeSheet.getRange(1, 1, 1, headers5.length).setValues([headers5]);
    scopeSheet.getRange(1, 1, 1, headers5.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    result.scopeSheet = true;
    
    return {
      success: true,
      message: 'All 5 Technical Library sheets created successfully!',
      details: result
    };
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get next Library ID
function getNextLibraryId() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Technical_Library');
    if (!sheet) return 'LIB-1001';
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 'LIB-1001';
    
    var lastId = sheet.getRange(lastRow, 1).getValue();
    if (!lastId) return 'LIB-1001';
    
    var numPart = parseInt(lastId.replace('LIB-', '')) || 1000;
    return 'LIB-' + (numPart + 1);
  } catch (error) {
    return 'LIB-1001';
  }
}

// Create new library (basic info only)
function createLibrary(data) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Technical_Library');
    
    if (!sheet) {
      return {success: false, message: 'Technical_Library sheet not found. Run setupTechnicalLibrarySheets() first.'};
    }
    
    var libraryId = getNextLibraryId();
    var now = new Date().toISOString();
    
    var row = [
      libraryId,
      data.name || '',
      data.description || '',
      now,
      data.createdBy || 'User'
    ];
    
    sheet.appendRow(row);
    
    return {
      success: true,
      message: 'Library created successfully',
      libraryId: libraryId
    };
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get all libraries (with counts)
function getLibraries() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Technical_Library');
    
    if (!sheet) {
      return {success: false, message: 'Technical_Library sheet not found', libraries: []};
    }
    
    var data = sheet.getDataRange().getValues();
    var libraries = [];
    
    // Get counts from related sheets
    var filesSheet = ss.getSheetByName('Library_Files');
    var equipmentSheet = ss.getSheetByName('Library_Equipment');
    
    var fileCounts = {};
    var equipmentCounts = {};
    
    if (filesSheet && filesSheet.getLastRow() > 1) {
      var filesData = filesSheet.getDataRange().getValues();
      for (var f = 1; f < filesData.length; f++) {
        var libId = filesData[f][1];
        if (libId) {
          fileCounts[libId] = (fileCounts[libId] || 0) + 1;
        }
      }
    }
    
    if (equipmentSheet && equipmentSheet.getLastRow() > 1) {
      var eqData = equipmentSheet.getDataRange().getValues();
      for (var e = 1; e < eqData.length; e++) {
        var libId2 = eqData[e][1];
        if (libId2) {
          equipmentCounts[libId2] = (equipmentCounts[libId2] || 0) + 1;
        }
      }
    }
    
    // Build library objects
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      
      var libId = row[0].toString();
      libraries.push({
        libraryId: libId,
        name: row[1] || '',
        description: row[2] || '',
        createdDate: row[3] || '',
        createdBy: row[4] || '',
        fileCount: fileCounts[libId] || 0,
        equipmentCount: equipmentCounts[libId] || 0
      });
    }
    
    return {success: true, libraries: libraries};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString(), libraries: []};
  }
}

// Get single library details (with all related data)
function getLibraryDetails(libraryId) {
  try {
    var ss = getSpreadsheet();
    
    // Get library basic info
    var librarySheet = ss.getSheetByName('Technical_Library');
    if (!librarySheet) {
      return {success: false, message: 'Technical_Library sheet not found'};
    }
    
    var data = librarySheet.getDataRange().getValues();
    var library = null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === libraryId.toString()) {
        library = {
          libraryId: data[i][0],
          name: data[i][1],
          description: data[i][2],
          createdDate: data[i][3],
          createdBy: data[i][4]
        };
        break;
      }
    }
    
    if (!library) {
      return {success: false, message: 'Library not found'};
    }
    
    // Get files
    library.files = getLibraryFiles(libraryId).files || [];
    
    // Get equipment (with current stock data)
    library.equipment = getLibraryEquipment(libraryId).equipment || [];
    
    // Get installation steps
    library.steps = getLibrarySteps(libraryId).steps || [];
    
    // Get scope of work
    library.scope = getLibraryScope(libraryId).scope || '';
    
    return {success: true, library: library};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Update library basic info
function updateLibrary(data) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Technical_Library');
    
    if (!sheet) {
      return {success: false, message: 'Technical_Library sheet not found'};
    }
    
    var rows = sheet.getDataRange().getValues();
    
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.libraryId.toString()) {
        var rowIndex = i + 1;
        sheet.getRange(rowIndex, 2).setValue(data.name || '');
        sheet.getRange(rowIndex, 3).setValue(data.description || '');
        return {success: true, message: 'Library updated successfully'};
      }
    }
    
    return {success: false, message: 'Library not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Delete library (and all related data)
function deleteLibrary(libraryId) {
  try {
    var ss = getSpreadsheet();
    
    // Delete from all 5 sheets
    var sheets = [
      'Technical_Library',
      'Library_Files',
      'Library_Equipment',
      'Library_Steps',
      'Library_Scope'
    ];
    
    var deleted = false;
    
    for (var s = 0; s < sheets.length; s++) {
      var sheet = ss.getSheetByName(sheets[s]);
      if (!sheet) continue;
      
      var data = sheet.getDataRange().getValues();
      var colIndex = (sheets[s] === 'Technical_Library') ? 0 : 1; // Library ID is col A or B
      
      for (var i = data.length - 1; i >= 1; i--) {
        if (data[i][colIndex] && data[i][colIndex].toString() === libraryId.toString()) {
          sheet.deleteRow(i + 1);
          deleted = true;
        }
      }
    }
    
    if (deleted) {
      return {success: true, message: 'Library and all related data deleted successfully'};
    } else {
      return {success: false, message: 'Library not found'};
    }
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ===== FILE MANAGEMENT =====

// Add file to library
function addLibraryFile(data) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Library_Files');
    
    if (!sheet) {
      return {success: false, message: 'Library_Files sheet not found'};
    }
    
    var fileId = 'FILE-' + new Date().getTime();
    var now = new Date().toISOString();
    
    var row = [
      fileId,
      data.libraryId,
      data.fileName || '',
      data.fileType || '',
      data.base64Data || '',
      now
    ];
    
    sheet.appendRow(row);
    
    return {success: true, message: 'File added successfully', fileId: fileId};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get files for a library
function getLibraryFiles(libraryId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Library_Files');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {success: true, files: []};
    }
    
    var data = sheet.getDataRange().getValues();
    var files = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === libraryId.toString()) {
        files.push({
          fileId: data[i][0],
          libraryId: data[i][1],
          fileName: data[i][2],
          fileType: data[i][3],
          base64Data: data[i][4],
          uploadDate: data[i][5]
        });
      }
    }
    
    return {success: true, files: files};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString(), files: []};
  }
}

// Delete file
function deleteLibraryFile(fileId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Library_Files');
    
    if (!sheet) {
      return {success: false, message: 'Library_Files sheet not found'};
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === fileId.toString()) {
        sheet.deleteRow(i + 1);
        return {success: true, message: 'File deleted successfully'};
      }
    }
    
    return {success: false, message: 'File not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ===== EQUIPMENT MANAGEMENT =====

// Add equipment to library
function addLibraryEquipment(data) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Library_Equipment');
    
    if (!sheet) {
      return {success: false, message: 'Library_Equipment sheet not found'};
    }
    
    var entryId = 'EQ-' + new Date().getTime();
    
    var row = [
      entryId,
      data.libraryId,
      data.itemName,
      data.quantityNeeded || 1
    ];
    
    sheet.appendRow(row);
    
    return {success: true, message: 'Equipment added successfully', entryId: entryId};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get equipment for a library (with current stock data)
function getLibraryEquipment(libraryId) {
  try {
    var ss = getSpreadsheet();
    var equipmentSheet = ss.getSheetByName('Library_Equipment');
    var stocksSheet = ss.getSheetByName('Stocks');
    
    if (!equipmentSheet || equipmentSheet.getLastRow() <= 1) {
      return {success: true, equipment: []};
    }
    
    var eqData = equipmentSheet.getDataRange().getValues();
    var equipment = [];
    
    // Build stock lookup map
    var stockMap = {};
    if (stocksSheet && stocksSheet.getLastRow() > 1) {
      var stockData = stocksSheet.getDataRange().getValues();
      for (var s = 1; s < stockData.length; s++) {
        if (stockData[s][0]) {
          stockMap[stockData[s][0].toString().toLowerCase()] = {
            itemName: stockData[s][0],
            purchasePrice: stockData[s][1],
            salePrice: stockData[s][2],
            quantityAvailable: stockData[s][3],
            description: stockData[s][4]
          };
        }
      }
    }
    
    // Get equipment items and merge with current stock data
    for (var i = 1; i < eqData.length; i++) {
      if (eqData[i][1] && eqData[i][1].toString() === libraryId.toString()) {
        var itemName = eqData[i][2];
        var stockInfo = stockMap[itemName.toLowerCase()] || {};
        
        equipment.push({
          entryId: eqData[i][0],
          libraryId: eqData[i][1],
          itemName: itemName,
          quantityNeeded: eqData[i][3],
          // Current stock data (live)
          purchasePrice: stockInfo.purchasePrice || 0,
          salePrice: stockInfo.salePrice || 0,
          quantityAvailable: stockInfo.quantityAvailable || 0,
          description: stockInfo.description || ''
        });
      }
    }
    
    return {success: true, equipment: equipment};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString(), equipment: []};
  }
}

// Delete equipment item
function deleteLibraryEquipment(entryId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Library_Equipment');
    
    if (!sheet) {
      return {success: false, message: 'Library_Equipment sheet not found'};
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === entryId.toString()) {
        sheet.deleteRow(i + 1);
        return {success: true, message: 'Equipment item removed successfully'};
      }
    }
    
    return {success: false, message: 'Equipment item not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ===== INSTALLATION STEPS =====

// Add installation step
function addLibraryStep(data) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Library_Steps');
    
    if (!sheet) {
      return {success: false, message: 'Library_Steps sheet not found'};
    }
    
    var stepId = 'STEP-' + new Date().getTime();
    
    var row = [
      stepId,
      data.libraryId,
      data.stepNumber || 1,
      data.stepText || ''
    ];
    
    sheet.appendRow(row);
    
    return {success: true, message: 'Step added successfully', stepId: stepId};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get installation steps for a library
function getLibrarySteps(libraryId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Library_Steps');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {success: true, steps: []};
    }
    
    var data = sheet.getDataRange().getValues();
    var steps = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === libraryId.toString()) {
        steps.push({
          stepId: data[i][0],
          libraryId: data[i][1],
          stepNumber: data[i][2],
          stepText: data[i][3]
        });
      }
    }
    
    // Sort by step number
    steps.sort(function(a, b) { return a.stepNumber - b.stepNumber; });
    
    return {success: true, steps: steps};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString(), steps: []};
  }
}

// Delete installation step
function deleteLibraryStep(stepId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Library_Steps');
    
    if (!sheet) {
      return {success: false, message: 'Library_Steps sheet not found'};
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === stepId.toString()) {
        sheet.deleteRow(i + 1);
        return {success: true, message: 'Step deleted successfully'};
      }
    }
    
    return {success: false, message: 'Step not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ===== SCOPE OF WORK =====

// Save/Update scope of work
function saveLibraryScope(data) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Library_Scope');
    
    if (!sheet) {
      return {success: false, message: 'Library_Scope sheet not found'};
    }
    
    var rows = sheet.getDataRange().getValues();
    var found = false;
    
    // Check if scope already exists for this library
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] && rows[i][0].toString() === data.libraryId.toString()) {
        // Update existing
        sheet.getRange(i + 1, 2).setValue(data.scopeText || '');
        found = true;
        break;
      }
    }
    
    if (!found) {
      // Create new
      sheet.appendRow([data.libraryId, data.scopeText || '']);
    }
    
    return {success: true, message: 'Scope of work saved successfully'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get scope of work for a library
function getLibraryScope(libraryId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Library_Scope');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {success: true, scope: ''};
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === libraryId.toString()) {
        return {success: true, scope: data[i][1] || ''};
      }
    }
    
    return {success: true, scope: ''};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString(), scope: ''};
  }
}