// ============================================================
//  MEEG Operations Dashboard — AdminAPI.gs
//  كل عمليات الـ Admin على الـ USERS sheet
//  (Add / Edit / Delete / Toggle Status / Sync)
// ============================================================

// ════════════════════════════════════════════════════════════
//  ADD USER — بيضيف user جديد في الـ USERS sheet
// ════════════════════════════════════════════════════════════
function addUserToSheet(userObj) {
  try {
    var ss       = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    var usersTab = _getOrCreateUsersTab(ss);

    // تحقق إن الإيميل مش موجود قبل كده
    var existing = _findUserRow(usersTab, userObj.email);
    if (existing > 0) {
      return { success: false, error: 'Email already exists: ' + userObj.email };
    }

    // بناء الباسورد
    var plainPass = userObj.password
      ? userObj.password                          // لو Admin حدد باسورد مخصص
      : _autoPassword(userObj.email, userObj.role); // أوتوماتيك

    var created = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

    var newRow = [
      userObj.email,
      userObj.name,
      userObj.role,
      userObj.shift   || 'All',
      userObj.room    || 'All',
      userObj.qtc     || '—',
      _hashPassword(plainPass),
      plainPass,
      created,
      userObj.status  || 'Active',
      userObj.hidden  || '',
      (userObj.perms  || []).join(','),
    ];

    usersTab.appendRow(newRow);
    _styleLastRow(usersTab, userObj.role);

    Logger.log('✅ Added user: ' + userObj.email);
    return {
      success   : true,
      message   : 'User added successfully',
      email     : userObj.email,
      password  : plainPass,
    };

  } catch(e) {
    Logger.log('❌ addUserToSheet error: ' + e.message);
    return { success: false, error: e.message };
  }
}


// ════════════════════════════════════════════════════════════
//  EDIT USER — بيعدل بيانات user موجود
// ════════════════════════════════════════════════════════════
function editUserInSheet(email, updatedObj) {
  try {
    var ss       = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    var usersTab = ss.getSheetByName(CONFIG.USERS_TAB);
    if (!usersTab) return { success: false, error: 'USERS tab not found' };

    var rowIdx = _findUserRow(usersTab, email);
    if (rowIdx < 0) return { success: false, error: 'User not found: ' + email };

    var data    = usersTab.getDataRange().getValues();
    var headers = data[0];
    var row     = data[rowIdx];

    // عدّل الـ fields المطلوبة
    var colMap = _buildColMap(headers);

    function setCol(colName, val) {
      if (colMap[colName] !== undefined && val !== undefined) {
        usersTab.getRange(rowIdx + 1, colMap[colName] + 1).setValue(val);
      }
    }

    setCol('Name',    updatedObj.name);
    setCol('Role',    updatedObj.role);
    setCol('Shift',   updatedObj.shift);
    setCol('Room',    updatedObj.room);
    setCol('QTC',     updatedObj.qtc);
    setCol('Status',  updatedObj.status);
    setCol('Hidden Columns', updatedObj.hidden);
    setCol('Permissions',    (updatedObj.perms || []).join(','));

    // لو عايز يغير الباسورد
    if (updatedObj.newPassword) {
      var hash = _hashPassword(updatedObj.newPassword);
      setCol('Password Hash (SHA-256)', hash);
      setCol('Password (Plain)', updatedObj.newPassword);
    }

    Logger.log('✅ Edited user: ' + email);
    return { success: true, message: 'User updated successfully' };

  } catch(e) {
    Logger.log('❌ editUserInSheet error: ' + e.message);
    return { success: false, error: e.message };
  }
}


// ════════════════════════════════════════════════════════════
//  DELETE USER — بيمسح user من الـ USERS sheet
// ════════════════════════════════════════════════════════════
function deleteUserFromSheet(email) {
  try {
    // حماية — مينفعش تمسح الـ admin الرئيسي
    if (email.toLowerCase() === CONFIG.ADMIN_EMAIL.toLowerCase()) {
      return { success: false, error: 'Cannot delete the main admin account' };
    }

    var ss       = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    var usersTab = ss.getSheetByName(CONFIG.USERS_TAB);
    if (!usersTab) return { success: false, error: 'USERS tab not found' };

    var rowIdx = _findUserRow(usersTab, email);
    if (rowIdx < 0) return { success: false, error: 'User not found: ' + email };

    usersTab.deleteRow(rowIdx + 1);  // +1 عشان Sheets 1-based
    Logger.log('✅ Deleted user: ' + email);
    return { success: true, message: 'User deleted successfully' };

  } catch(e) {
    Logger.log('❌ deleteUserFromSheet error: ' + e.message);
    return { success: false, error: e.message };
  }
}


// ════════════════════════════════════════════════════════════
//  TOGGLE STATUS — Active ↔ Inactive
// ════════════════════════════════════════════════════════════
function toggleUserStatus(email) {
  try {
    var ss       = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    var usersTab = ss.getSheetByName(CONFIG.USERS_TAB);
    if (!usersTab) return { success: false, error: 'USERS tab not found' };

    var data    = usersTab.getDataRange().getValues();
    var headers = data[0];
    var colMap  = _buildColMap(headers);
    var statusCol = colMap['Status'];
    if (statusCol === undefined) return { success: false, error: 'Status column not found' };

    var rowIdx = _findUserRow(usersTab, email);
    if (rowIdx < 0) return { success: false, error: 'User not found' };

    var current    = String(data[rowIdx][statusCol]).trim();
    var newStatus  = current === 'Active' ? 'Inactive' : 'Active';
    usersTab.getRange(rowIdx + 1, statusCol + 1).setValue(newStatus);

    Logger.log('✅ Toggled ' + email + ' → ' + newStatus);
    return { success: true, newStatus: newStatus };

  } catch(e) {
    Logger.log('❌ toggleUserStatus error: ' + e.message);
    return { success: false, error: e.message };
  }
}


// ════════════════════════════════════════════════════════════
//  GET ALL USERS — بيجيب كل الـ users للـ Admin dashboard
// ════════════════════════════════════════════════════════════
function getAllUsers() {
  try {
    var ss       = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    var usersTab = ss.getSheetByName(CONFIG.USERS_TAB);
    if (!usersTab) return { success: false, error: 'USERS tab not found' };

    var data    = usersTab.getDataRange().getValues();
    if (data.length < 2) return { success: true, users: [] };

    var headers = data[0].map(function(h){ return String(h).trim(); });
    var users   = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {};
      headers.forEach(function(h, j){ obj[h] = row[j]; });

      // بنرجع الـ user بدون الـ hash (للأمان)
      users.push({
        email   : obj['Email']   || '',
        name    : obj['Name']    || '',
        role    : obj['Role']    || '',
        shift   : obj['Shift']   || '',
        room    : obj['Room']    || '',
        qtc     : obj['QTC']     || '—',
        status  : obj['Status']  || 'Active',
        hidden  : obj['Hidden Columns'] || '',
        perms   : String(obj['Permissions'] || '').split(',').filter(Boolean),
        created : obj['Created At'] || '',
      });
    }

    return { success: true, users: users };

  } catch(e) {
    Logger.log('❌ getAllUsers error: ' + e.message);
    return { success: false, error: e.message };
  }
}


// ════════════════════════════════════════════════════════════
//  RESET PASSWORD — بيعمل reset للباسورد
// ════════════════════════════════════════════════════════════
function resetUserPassword(email, newPassword) {
  try {
    var autoPass = newPassword || _autoPassword(email, '');
    return editUserInSheet(email, { newPassword: autoPass });
  } catch(e) {
    return { success: false, error: e.message };
  }
}


// ════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════

// بيجيب أو يعمل USERS tab
function _getOrCreateUsersTab(ss) {
  var tab = ss.getSheetByName(CONFIG.USERS_TAB);
  if (!tab) {
    tab = ss.insertSheet(CONFIG.USERS_TAB);
    // كتابة الـ header الموسعة
    tab.getRange(1, 1, 1, 12).setValues([[
      'Email', 'Name', 'Role', 'Shift', 'Room', 'QTC',
      'Password Hash (SHA-256)', 'Password (Plain)',
      'Created At', 'Status', 'Hidden Columns', 'Permissions'
    ]]);
    // Style
    tab.getRange(1, 1, 1, 12)
       .setBackground('#1A2B5F')
       .setFontColor('#2ABFBF')
       .setFontWeight('bold')
       .setFontSize(11)
       .setHorizontalAlignment('center');
    tab.setFrozenRows(1);
  }
  return tab;
}

// بيلاقي row index الـ user (0-based بما فيها الـ header)
function _findUserRow(sheet, email) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email.trim().toLowerCase()) {
      return i;
    }
  }
  return -1;
}

// بيبني map من header name → column index
function _buildColMap(headers) {
  var map = {};
  headers.forEach(function(h, i){ map[String(h).trim()] = i; });
  return map;
}

// باسورد أوتوماتيك حسب الـ email والـ role
function _autoPassword(email, role) {
  if (!role || role === 'Labeler') {
    var m = String(email).match(/^me(\d+)@/i);
    if (m) return CONFIG.PASSWORD_PREFIX + m[1].substring(0, 4);
  }
  // غير Labeler
  return 'MEEG@' + String(email).split('@')[0] + '2025';
}

// Style الـ last row حسب الـ role
function _styleLastRow(sheet, role) {
  var lastRow = sheet.getLastRow();
  var range   = sheet.getRange(lastRow, 1, 1, 12);
  var colors  = {
    'Admin'            : ['#2d1515', '#fca5a5'],
    'Shift Supervisor' : ['#1a1530', '#a89ef0'],
    'Floor Supervisor' : ['#0d2020', '#2ABFBF'],
    'QA & Training'    : ['#1e1a08', '#FFD166'],
    'Workload Manager' : ['#1e1108', '#F0994A'],
    'Labeler'          : ['#0d1b2e', '#e8edf5'],
  };
  var c = colors[role] || colors['Labeler'];
  range.setBackground(c[0]).setFontColor(c[1]);
}


// ════════════════════════════════════════════════════════════
//  doPost router — يستقبل admin actions من الـ frontend
// ════════════════════════════════════════════════════════════
function handleAdminRequest(params, userSession) {
  // تأكد إن الـ user هو Admin
  if (userSession.role !== 'Admin') {
    return { success: false, error: 'Access denied — Admin only' };
  }

  var action = params.action;

  switch (action) {

    case 'getAllUsers':
      return getAllUsers();

    case 'addUser':
      return addUserToSheet(params.user);

    case 'editUser':
      return editUserInSheet(params.email, params.updates);

    case 'deleteUser':
      return deleteUserFromSheet(params.email);

    case 'toggleStatus':
      return toggleUserStatus(params.email);

    case 'resetPassword':
      return resetUserPassword(params.email, params.newPassword);

    default:
      return { success: false, error: 'Unknown admin action: ' + action };
  }
}
