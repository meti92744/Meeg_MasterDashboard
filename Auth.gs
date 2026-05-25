// ============================================================
//  MEEG Operations Dashboard — Auth.gs
//  يقرأ TEAMS tab ويولد USERS tab بكل الأكونتات تلقائياً
// ============================================================

// ── CONFIG ─────────────────────────────────────────────────
var CONFIG = {
  MASTER_SHEET_ID : '1_Tg0SNhFcRYy00U5GyUADnRLjoeHp0hOYxEmC3Z6bxI',
  TEAMS_TAB       : 'TEAMS',
  USERS_TAB       : 'USERS',
  PASSWORD_PREFIX : '$M@@g#ME',
  ADMIN_EMAIL     : 'admin@meeg.ai',
  ADMIN_PASSWORD  : 'MEEGAdmin@2025',
  ADMIN_NAME      : 'Meti Admin',
};

// ── COLUMNS في TEAMS tab (1-based) ─────────────────────────
var TEAMS_COL = {
  EMAIL : 1,
  NAME  : 2,
  ROOM  : 3,
  SHIFT : 4,
  QTC   : 5,
};

// ── COLUMNS في USERS tab ───────────────────────────────────
var USERS_COL = {
  EMAIL       : 1,
  NAME        : 2,
  ROLE        : 3,
  SHIFT       : 4,
  ROOM        : 5,
  QTC         : 6,
  PASS_HASH   : 7,
  PASS_PLAIN  : 8,
  CREATED_AT  : 9,
  STATUS      : 10,
};


// ════════════════════════════════════════════════════════════
//  1. الدالة الرئيسية — شغّلها مرة واحدة من الـ Editor
// ════════════════════════════════════════════════════════════
function generateUsersFromTeams() {
  var ss        = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
  var teamsTab  = ss.getSheetByName(CONFIG.TEAMS_TAB);

  if (!teamsTab) {
    SpreadsheetApp.getUi().alert('❌ مش لاقي تاب اسمها TEAMS — تأكد من الاسم.');
    return;
  }

  var usersTab = ss.getSheetByName(CONFIG.USERS_TAB);
  if (!usersTab) {
    usersTab = ss.insertSheet(CONFIG.USERS_TAB);
    Logger.log('✅ عملت تاب USERS جديدة');
  } else {
    if (usersTab.getLastRow() > 1) {
      usersTab.getRange(2, 1, usersTab.getLastRow() - 1, usersTab.getLastColumn()).clearContent();
    }
    Logger.log('♻️ هعيد بناء الـ USERS tab');
  }

  _writeHeader(usersTab);
  _styleHeader(usersTab);

  var rows       = [];
  var teamsData  = teamsTab.getDataRange().getValues();
  var now        = new Date();
  var created    = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  // Admin row أول حاجة
  rows.push(_buildUserRow(
    CONFIG.ADMIN_EMAIL,
    CONFIG.ADMIN_NAME,
    'Admin',
    'All',
    'All',
    'All',
    CONFIG.ADMIN_PASSWORD,
    created
  ));

  // Labeler rows من TEAMS
  var skipped = 0;
  for (var i = 1; i < teamsData.length; i++) {
    var row   = teamsData[i];
    var email = String(row[TEAMS_COL.EMAIL - 1]).trim();
    var name  = String(row[TEAMS_COL.NAME  - 1]).trim();
    var room  = String(row[TEAMS_COL.ROOM  - 1]).trim();
    var shift = String(row[TEAMS_COL.SHIFT - 1]).trim();
    var qtc   = String(row[TEAMS_COL.QTC   - 1]).trim();

    if (!email || email === '' || email.toLowerCase() === 'email') {
      skipped++;
      continue;
    }

    var password = _buildLabelerPassword(email);
    rows.push(_buildUserRow(email, name, 'Labeler', shift, room, qtc, password, created));
  }

  if (rows.length > 0) {
    usersTab.getRange(2, 1, rows.length, Object.keys(USERS_COL).length).setValues(rows);
  }

  _styleUsersSheet(usersTab, rows.length);

  var msg = '✅ تم بنجاح!\n\n'
          + '👤 Admin: ' + CONFIG.ADMIN_EMAIL + '\n'
          + '👥 Labelers: ' + (rows.length - 1) + '\n'
          + '⏭️ Skipped: ' + skipped + '\n\n'
          + 'شوف تاب USERS.';
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}


// ════════════════════════════════════════════════════════════
//  2. بناء الباسورد من الإيميل
// ════════════════════════════════════════════════════════════
function _buildLabelerPassword(email) {
  var match = email.match(/^me(\d+)@/i);
  if (!match) return CONFIG.PASSWORD_PREFIX + '0000';
  var digits = match[1].substring(0, 4);
  return CONFIG.PASSWORD_PREFIX + digits;
}


// ════════════════════════════════════════════════════════════
//  3. SHA-256 Hash
// ════════════════════════════════════════════════════════════
function _hashPassword(plainText) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    plainText,
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}


// ════════════════════════════════════════════════════════════
//  4. بناء row واحدة
// ════════════════════════════════════════════════════════════
function _buildUserRow(email, name, role, shift, room, qtc, plainPass, createdAt) {
  return [
    email, name, role, shift, room, qtc,
    _hashPassword(plainPass),
    plainPass,
    createdAt,
    'Active'
  ];
}


// ════════════════════════════════════════════════════════════
//  5. Header
// ════════════════════════════════════════════════════════════
function _writeHeader(sheet) {
  sheet.getRange(1, 1, 1, 10).setValues([[
    'Email', 'Name', 'Role', 'Shift', 'Room', 'QTC',
    'Password Hash (SHA-256)', 'Password (Plain)', 'Created At', 'Status'
  ]]);
}

function _styleHeader(sheet) {
  sheet.getRange(1, 1, 1, 10)
    .setBackground('#1A2B5F')
    .setFontColor('#2ABFBF')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

function _styleUsersSheet(sheet, dataRows) {
  if (dataRows < 1) return;
  sheet.getRange(2, 1, 1, 10).setBackground('#2d1515').setFontColor('#fca5a5');
  if (dataRows > 1) {
    sheet.getRange(3, 1, dataRows - 1, 10).setBackground('#0d1b2e').setFontColor('#e8edf5');
  }
  sheet.setColumnWidth(1, 220); sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 140); sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 140); sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 300); sheet.setColumnWidth(8, 160);
  sheet.setColumnWidth(9, 140); sheet.setColumnWidth(10, 80);
}


// ════════════════════════════════════════════════════════════
//  6. LOGIN
// ════════════════════════════════════════════════════════════
function loginUser(email, plainPassword) {
  try {
    var ss       = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    var usersTab = ss.getSheetByName(CONFIG.USERS_TAB);
    if (!usersTab) return { success: false, error: 'USERS tab not found' };

    var data      = usersTab.getDataRange().getValues();
    var inputHash = _hashPassword(plainPassword.trim());

    for (var i = 1; i < data.length; i++) {
      var row       = data[i];
      var rowEmail  = String(row[USERS_COL.EMAIL     - 1]).trim().toLowerCase();
      var rowHash   = String(row[USERS_COL.PASS_HASH - 1]).trim();
      var rowStatus = String(row[USERS_COL.STATUS    - 1]).trim();

      if (rowEmail === email.trim().toLowerCase() && rowHash === inputHash) {
        if (rowStatus !== 'Active') {
          return { success: false, error: 'Account inactive' };
        }
        return {
          success : true,
          email   : rowEmail,
          name    : String(row[USERS_COL.NAME  - 1]).trim(),
          role    : String(row[USERS_COL.ROLE  - 1]).trim(),
          shift   : String(row[USERS_COL.SHIFT - 1]).trim(),
          room    : String(row[USERS_COL.ROOM  - 1]).trim(),
          qtc     : String(row[USERS_COL.QTC   - 1]).trim(),
        };
      }
    }

    return { success: false, error: 'Invalid credentials' };

  } catch(e) {
    return { success: false, error: e.message };
  }
}


// ════════════════════════════════════════════════════════════
//  7. doPost — entry point كامل
// ════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;

    // ── Login — مش محتاج session ──────────────────────────
    if (action === 'login') {
      return _json(loginUser(params.email, params.password));
    }

    // ── كل الـ actions التانية ─────────────────────────────
    var session = params.session || {};

    switch (action) {

      // Data actions
      case 'getLabelerData':
        return _json(getLabelerData(session.email));

      case 'getFloorData':
        return _json(getFloorSupervisorData(
          params.room  || session.room  || null,
          params.shift || session.shift || null,
          params.date  || null
        ));

      case 'getWorkloadData':
        return _json(getWorkloadData(params.date || null));

      case 'getQualityData':
        return _json(getQualityData());

      // Admin actions
      case 'admin_getAllUsers':
        return _json(getAllUsers());

      case 'admin_addUser':
        return _json(addUserToSheet(params.user));

      case 'admin_editUser':
        return _json(editUserInSheet(params.email, params.updates));

      case 'admin_deleteUser':
        return _json(deleteUserFromSheet(params.email));

      case 'admin_toggleStatus':
        return _json(toggleUserStatus(params.email));

      case 'admin_resetPassword':
        return _json(resetUserPassword(params.email, params.newPassword));

      default:
        return _json({ error: 'Unknown action: ' + action });
    }

  } catch (err) {
    return _json({ error: err.message });
  }
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('MEEG Operations Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
