// ============================================
// SETUP USERS SHEET (Updated structure)
// ============================================

function setupUsersSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Users');
    
    if (!sheet) {
      sheet = ss.insertSheet('Users');
    } else {
      // Clear existing data
      sheet.clear();
    }
    
    // Headers
    var headers = [
      'UserID',        // A: U001, U002, U003...
      'Username',      // B: admin, john, sarah...
      'Email',         // C: admin@company.com
      'FullName',      // D: John Smith
      'Status',        // E: Active/Blocked/Suspended
      'RoleID',        // F: R001 (optional - for role templates)
      'CreatedDate',   // G: 2024-03-19
      'LastLogin',     // H: 2024-03-19 10:30:00
      'BlockedDate',   // I: Date when blocked
      'BlockedBy',     // J: Admin who blocked
      'BlockReason'    // K: Reason for blocking
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Set column widths
    sheet.setColumnWidth(1, 80);   // UserID
    sheet.setColumnWidth(2, 120);  // Username
    sheet.setColumnWidth(3, 200);  // Email
    sheet.setColumnWidth(4, 150);  // FullName
    sheet.setColumnWidth(5, 100);  // Status
    sheet.setColumnWidth(6, 80);   // RoleID
    sheet.setColumnWidth(7, 120);  // CreatedDate
    sheet.setColumnWidth(8, 150);  // LastLogin
    sheet.setColumnWidth(9, 120);  // BlockedDate
    sheet.setColumnWidth(10, 120); // BlockedBy
    sheet.setColumnWidth(11, 200); // BlockReason
    
    // Add default admin user
    var adminData = [
      'U001',                    // UserID
      'admin',                   // Username
      'admin@company.com',       // Email
      'System Administrator',    // FullName
      'Active',                  // Status
      'R001',                    // RoleID (Super Admin)
      new Date(),                // CreatedDate
      '',                        // LastLogin
      '',                        // BlockedDate
      '',                        // BlockedBy
      ''                         // BlockReason
    ];
    
    sheet.getRange(2, 1, 1, adminData.length).setValues([adminData]);
    
    return {success: true, message: 'Users sheet created successfully'};
    
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ============================================
// SETUP PAGE ACCESS SHEET
// ============================================
function setupPageAccessSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('PageAccess');
    
    if (!sheet) {
      sheet = ss.insertSheet('PageAccess');
    } else {
      sheet.clear();
    }
    
    // Headers
    var headers = [
      'UserID',      // A: U001, U002...
      'Page',        // B: dashboard, new-quote, all-quotes...
      'Allowed'      // C: TRUE/FALSE
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Set column widths
    sheet.setColumnWidth(1, 100);  // UserID
    sheet.setColumnWidth(2, 200);  // Page
    sheet.setColumnWidth(3, 100);  // Allowed
    
    // Add default admin access (all pages)
    var pages = [
      'dashboard',
      'new-quote',
      'all-quotes',
      'invoices',
      'customers',
      'suppliers',
      'items',
      'inventory',
      'create-purchase',
      'all-purchases',
      'technical-library',
      'user-management'
    ];
    
    var adminAccess = [];
    for (var i = 0; i < pages.length; i++) {
      adminAccess.push(['U001', pages[i], true]);
    }
    
    if (adminAccess.length > 0) {
      sheet.getRange(2, 1, adminAccess.length, 3).setValues(adminAccess);
    }
    
    return {success: true, message: 'PageAccess sheet created successfully'};
    
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ============================================
// SETUP ACTION ACCESS SHEET
// ============================================
function setupActionAccessSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('ActionAccess');
    
    if (!sheet) {
      sheet = ss.insertSheet('ActionAccess');
    } else {
      sheet.clear();
    }
    
    // Headers
    var headers = [
      'UserID',      // A: U001, U002...
      'Page',        // B: dashboard, all-quotes...
      'Action',      // C: view-quote, edit-quote, delete-quote...
      'Allowed'      // D: TRUE/FALSE/OWN/ALL
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Set column widths
    sheet.setColumnWidth(1, 100);  // UserID
    sheet.setColumnWidth(2, 200);  // Page
    sheet.setColumnWidth(3, 250);  // Action
    sheet.setColumnWidth(4, 100);  // Allowed
    
    // Add default admin access (all actions on all pages)
    var adminActions = [
      // Dashboard
      ['U001', 'dashboard', 'view-dashboard', true],
      ['U001', 'dashboard', 'view-statistics', true],
      ['U001', 'dashboard', 'view-charts', true],
      ['U001', 'dashboard', 'view-recent-quotes', true],
      ['U001', 'dashboard', 'view-recent-invoices', true],
      
      // New Quote
      ['U001', 'new-quote', 'save-quote', true],
      ['U001', 'new-quote', 'download-preview', true],
      
      // All Quotes
      ['U001', 'all-quotes', 'view-quote', true],
      ['U001', 'all-quotes', 'edit-quote', 'ALL'],
      ['U001', 'all-quotes', 'download-pdf', true],
      ['U001', 'all-quotes', 'send-email', true],
      ['U001', 'all-quotes', 'delete-quote', true],
      ['U001', 'all-quotes', 'mark-completed', true],
      ['U001', 'all-quotes', 'mark-pending', true],
      ['U001', 'all-quotes', 'create-invoice', true],
      
      // Invoices
      ['U001', 'invoices', 'view-invoice', true],
      ['U001', 'invoices', 'download-pdf', true],
      ['U001', 'invoices', 'send-email', true],
      ['U001', 'invoices', 'edit-invoice', true],
      ['U001', 'invoices', 'delete-invoice', true],
      ['U001', 'invoices', 'mark-paid-unpaid', true],
      ['U001', 'invoices', 'create-manual-invoice', true],
      
      // Customers
      ['U001', 'customers', 'view-customer', true],
      ['U001', 'customers', 'add-customer', true],
      ['U001', 'customers', 'edit-customer', true],
      ['U001', 'customers', 'delete-customer', true],
      ['U001', 'customers', 'view-customer-invoices', true],
      ['U001', 'customers', 'view-customer-quotes', true],
      
      // Suppliers
      ['U001', 'suppliers', 'view-supplier', true],
      ['U001', 'suppliers', 'add-supplier', true],
      ['U001', 'suppliers', 'edit-supplier', true],
      ['U001', 'suppliers', 'delete-supplier', true],
      ['U001', 'suppliers', 'view-purchase-history', true],
      
      // Items
      ['U001', 'items', 'view-items', true],
      ['U001', 'items', 'add-item', true],
      ['U001', 'items', 'edit-item', true],
      ['U001', 'items', 'delete-item', true],
      ['U001', 'items', 'add-category', true],
      ['U001', 'items', 'edit-category', true],
      ['U001', 'items', 'delete-category', true],
      
      // Inventory
      ['U001', 'inventory', 'view-inventory', true],
      ['U001', 'inventory', 'add-stock', true],
      ['U001', 'inventory', 'edit-stock', true],
      ['U001', 'inventory', 'delete-stock', true],
      ['U001', 'inventory', 'update-quantities', true],
      
      // Create Purchase
      ['U001', 'create-purchase', 'create-purchase', true],
      
      // All Purchases
      ['U001', 'all-purchases', 'view-purchase', true],
      ['U001', 'all-purchases', 'edit-purchase', true],
      ['U001', 'all-purchases', 'delete-purchase', true],
      ['U001', 'all-purchases', 'download-pdf', true],
      
      // Technical Library
      ['U001', 'technical-library', 'view-library', true],
      ['U001', 'technical-library', 'add-file', true],
      ['U001', 'technical-library', 'upload-document', true],
      ['U001', 'technical-library', 'download-document', true],
      ['U001', 'technical-library', 'edit-file', true],
      ['U001', 'technical-library', 'delete-file', true],
      
      // User Management
      ['U001', 'user-management', 'view-users', true],
      ['U001', 'user-management', 'add-user', true],
      ['U001', 'user-management', 'edit-user', true],
      ['U001', 'user-management', 'delete-user', true],
      ['U001', 'user-management', 'block-user', true],
      ['U001', 'user-management', 'unblock-user', true],
      ['U001', 'user-management', 'change-password', true],
      ['U001', 'user-management', 'assign-roles', true],
      ['U001', 'user-management', 'assign-permissions', true],
      ['U001', 'user-management', 'view-activity-log', true]
    ];
    
    if (adminActions.length > 0) {
      sheet.getRange(2, 1, adminActions.length, 4).setValues(adminActions);
    }
    
    return {success: true, message: 'ActionAccess sheet created successfully with ' + adminActions.length + ' actions'};
    
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ============================================
// SETUP PASSWORD HASHES SHEET
// ============================================
function setupPasswordHashesSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('PasswordHashes');
    
    if (!sheet) {
      sheet = ss.insertSheet('PasswordHashes');
    } else {
      sheet.clear();
    }
    
    // Headers
    var headers = [
      'UserID',         // A: U001, U002...
      'PasswordHash',   // B: SHA-256 hash
      'LastChanged',    // C: Date
      'MustChange'      // D: TRUE/FALSE (force password reset)
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Set column widths
    sheet.setColumnWidth(1, 100);  // UserID
    sheet.setColumnWidth(2, 500);  // PasswordHash
    sheet.setColumnWidth(3, 150);  // LastChanged
    sheet.setColumnWidth(4, 100);  // MustChange
    
    // Add default admin password hash (password: "admin123" - CHANGE THIS!)
    // For now, using plain text - will implement hashing later
    var adminPassword = [
      'U001',          // UserID
      'admin123',      // Temporary plain text - will be hashed
      new Date(),      // LastChanged
      false            // MustChange
    ];
    
    sheet.getRange(2, 1, 1, adminPassword.length).setValues([adminPassword]);
    
    return {success: true, message: 'PasswordHashes sheet created successfully'};
    
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ============================================
// SETUP AUDIT LOG SHEET
// ============================================
function setupAuditLogSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('AuditLog');
    
    if (!sheet) {
      sheet = ss.insertSheet('AuditLog');
    } else {
      sheet.clear();
    }
    
    // Headers
    var headers = [
      'Timestamp',     // A: 2024-03-19 10:30:15
      'UserID',        // B: U001
      'Username',      // C: admin
      'Action',        // D: LOGIN/LOGOUT/CREATE/EDIT/DELETE/VIEW
      'Module',        // E: Quotes/Invoices/Users/etc.
      'RecordID',      // F: Q-100001, INV-001, U002, etc.
      'OldValue',      // G: Previous value (for edits)
      'NewValue',      // H: New value (for edits)
      'Success',       // I: TRUE/FALSE
      'ErrorMessage'   // J: Error details if failed
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#001f3f')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Set column widths
    sheet.setColumnWidth(1, 150);  // Timestamp
    sheet.setColumnWidth(2, 80);   // UserID
    sheet.setColumnWidth(3, 120);  // Username
    sheet.setColumnWidth(4, 120);  // Action
    sheet.setColumnWidth(5, 120);  // Module
    sheet.setColumnWidth(6, 120);  // RecordID
    sheet.setColumnWidth(7, 200);  // OldValue
    sheet.setColumnWidth(8, 200);  // NewValue
    sheet.setColumnWidth(9, 80);   // Success
    sheet.setColumnWidth(10, 300); // ErrorMessage
    
    // Add initial log entry
    var initialLog = [
      new Date(),      // Timestamp
      'SYSTEM',        // UserID
      'SYSTEM',        // Username
      'SETUP',         // Action
      'AuditLog',      // Module
      '',              // RecordID
      '',              // OldValue
      '',              // NewValue
      true,            // Success
      'Audit log initialized'  // ErrorMessage
    ];
    
    sheet.getRange(2, 1, 1, initialLog.length).setValues([initialLog]);
    
    return {success: true, message: 'AuditLog sheet created successfully'};
    
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ============================================
// SETUP ALL SHEETS AT ONCE
// ============================================
function setupAllPermissionSheets() {
  try {
    var results = [];
    
    // Setup all sheets
    results.push(setupUsersSheet());
    results.push(setupPageAccessSheet());
    results.push(setupActionAccessSheet());
    results.push(setupPasswordHashesSheet());
    results.push(setupAuditLogSheet());
    
    // Check if all succeeded
    var allSuccess = true;
    var messages = [];
    
    for (var i = 0; i < results.length; i++) {
      messages.push(results[i].message);
      if (!results[i].success) {
        allSuccess = false;
      }
    }
    
    if (allSuccess) {
      return {
        success: true,
        message: 'All permission sheets created successfully!\n\n' + messages.join('\n')
      };
    } else {
      return {
        success: false,
        message: 'Some sheets failed:\n\n' + messages.join('\n')
      };
    }
    
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}