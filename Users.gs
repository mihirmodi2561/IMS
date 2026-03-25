// Authenticate user
function authenticateUser(username, password) {
  try {
    var ss = getSpreadsheet();
    var loginSheet = ss.getSheetByName('Login');

    if (!loginSheet) {
      return {success: false, message: 'Login sheet not found. Please run setupDemoData first.'};
    }

    var data = loginSheet.getDataRange().getValues();

    // Skip header row
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0] === username && row[1] === password) {
        if (row[2] === 'Allow') {
          return {
            success: true,
            message: 'Login successful',
            username: username,
            role: row[3] || 'user' // Default to 'user' if role not specified
          };
        } else {
          return {success: false, message: 'Access blocked for this user'};
        }
      }
    }

    return {success: false, message: 'Invalid username or password'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Get all users
function getUsers() {
  try {
    var ss = getSpreadsheet();
    var loginSheet = ss.getSheetByName('Login');

    if (!loginSheet) {
      return {success: false, message: 'Login sheet not found'};
    }

    var data = loginSheet.getDataRange().getValues();
    var users = [];

    // Skip header row
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) { // Check if username exists
        users.push({
          username: row[0],
          password: row[1],
          access: row[2],
          role: row[3] || 'user'
        });
      }
    }

    return {success: true, users: users};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Add new user
function addUser(userData) {
  try {
    var ss = getSpreadsheet();
    var loginSheet = ss.getSheetByName('Login');

    if (!loginSheet) {
      return {success: false, message: 'Login sheet not found'};
    }

    // Check if username already exists
    var data = loginSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userData.username) {
        return {success: false, message: 'Username already exists'};
      }
    }

    var rowData = [
      userData.username,
      userData.password,
      userData.access || 'Allow',
      userData.role || 'user'
    ];

    loginSheet.appendRow(rowData);

    return {success: true, message: 'User added successfully'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Update existing user
function updateUser(oldUsername, userData) {
  try {
    var ss = getSpreadsheet();
    var loginSheet = ss.getSheetByName('Login');

    if (!loginSheet) {
      return {success: false, message: 'Login sheet not found'};
    }

    var data = loginSheet.getDataRange().getValues();

    // Find and update the user
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === oldUsername) {
        var rowIndex = i + 1;
        
        // Check if new username already exists (if username is being changed)
        if (userData.username !== oldUsername) {
          for (var j = 1; j < data.length; j++) {
            if (j !== i && data[j][0] === userData.username) {
              return {success: false, message: 'Username already exists'};
            }
          }
        }
        
        loginSheet.getRange(rowIndex, 1).setValue(userData.username);
        loginSheet.getRange(rowIndex, 2).setValue(userData.password);
        loginSheet.getRange(rowIndex, 3).setValue(userData.access || 'Allow');
        loginSheet.getRange(rowIndex, 4).setValue(userData.role || 'user');
        
        return {success: true, message: 'User updated successfully'};
      }
    }

    return {success: false, message: 'User not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// Delete user
function deleteUser(username) {
  try {
    var ss = getSpreadsheet();
    var loginSheet = ss.getSheetByName('Login');

    if (!loginSheet) {
      return {success: false, message: 'Login sheet not found'};
    }

    // Prevent deleting admin user
    if (username === 'admin') {
      return {success: false, message: 'Cannot delete admin user'};
    }

    var data = loginSheet.getDataRange().getValues();

    // Find and delete the row
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        loginSheet.deleteRow(i + 1);
        return {success: true, message: 'User deleted successfully'};
      }
    }

    return {success: false, message: 'User not found'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}
