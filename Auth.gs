// ============================================================
//  MEEG Operations Dashboard — Auth.gs
//  يقرأ TEAMS tab ويولد USERS tab بكل الأكونتات تلقائياً
// ============================================================

// ── CONFIG ─────────────────────────────────────────────────
var CONFIG = {
  MASTER_SHEET_ID : 'YOUR_MASTER_SHEET_ID_HERE',  // ← حط هنا الـ ID بتاع الـ Master Sheet
  TEAMS_TAB       : 'TEAMS',
  USERS_TAB       : 'USERS',
  PASSWORD_PREFIX : '$M@@g#ME',
  ADMIN_EMAIL     : 'admin@meeg.ai',
  ADMIN_PASSWORD  : 'MEEGAdmin@2025',
  ADMIN_NAME      : 'Meti Admin',
};

// ── COLUMNS في TEAMS tab (1-based) ─────────────────────────
var TEAMS_COL = {
  EMAIL : 1,   // A
  NAME  : 2,   // B
  ROOM  : 3,   // C
  SHIFT : 4,   // D
  QTC   : 5,   // E
};

// ── COLUMNS في USERS tab ───────────────────────────────────
var USERS_COL = {
  EMAIL       : 1,   // A
  NAME        : 2,   // B
  ROLE        : 3,   // C
  SHIFT       : 4,   // D
  ROOM        : 5,   // E
  QTC         : 6,   // F
  PASS_HASH   : 7,   // G — SHA-256 hash
  PASS_PLAIN  : 8,   // H — للمراجعة فقط (ممكن تشيلها بعدين)
  CREATED_AT  : 9,   // I
  STATUS      : 10,  // J — Active / Inactive
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

  // جيب أو اعمل USERS tab
  var usersTab = ss.getSheetByName(CONFIG.USERS_TAB);
  if (!usersTab) {
    usersTab = ss.insertSheet(CONFIG.USERS_TAB);
    Logger.log('✅ عملت تاب USERS جديدة');
  } else {
    // امسح القديم غير الـ header
    if (usersTab.getLastRow() > 1) {
      usersTab.getRange(2, 1, usersTab.getLastRow() - 1, usersTab.getLastColumn()).clearContent();
    }
    Logger.log('♻️ هعيد بناء الـ USERS tab');
  }

  // ── Header ──────────────────────────────────────────────
  _writeHeader(usersTab);
  _styleHeader(usersTab);

  var rows       = [];
  var teamsData  = teamsTab.getDataRange().getValues();
  var now        = new Date();
  var created    = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  // ── Admin row أول حاجة ──────────────────────────────────
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

  // ── Labeler rows من TEAMS ────────────────────────────────
  var skipped = 0;
  for (var i = 1; i < teamsData.length; i++) {   // i=1 عشان نتخطى الـ header
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

    rows.push(_buildUserRow(
      email, name, 'Labeler', shift, room, qtc, password, created
    ));
  }

  // ── اكتب كل الـ rows دفعة واحدة (أسرع) ─────────────────
  if (rows.length > 0) {
    usersTab.getRange(2, 1, rows.length, Object.keys(USERS_COL).length).setValues(rows);
  }

  // ── Style الـ sheet ──────────────────────────────────────
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
//     me409105@meti.ai  →  $M@@g#ME4091  (أول 4 أرقام من الـ ID)
// ════════════════════════════════════════════════════════════
function _buildLabelerPassword(email) {
  // استخرج الـ ID من الإيميل (الجزء بعد "me" وقبل "@")
  var match = email.match(/^me(\d+)@/i);
  if (!match) return CONFIG.PASSWORD_PREFIX + '0000';   // fallback
  var digits = match[1].substring(0, 4);               // أول 4 أرقام
  return CONFIG.PASSWORD_PREFIX + digits;
}


// ════════════════════════════════════════════════════════════
//  3. SHA-256 Hash للباسورد
// ════════════════════════════════════════════════════════════
function _hashPassword(plainText) {
  var bytes  = Utilities.computeDigest(
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
    email,
    name,
    role,
    shift,
    room,
    qtc,
    _hashPassword(plainPass),   // PASS_HASH
    plainPass,                  // PASS_PLAIN (للمراجعة)
    createdAt,
    'Active'
  ];
}


// ════════════════════════════════════════════════════════════
//  5. كتابة الـ Header
// ════════════════════════════════════════════════════════════
function _writeHeader(sheet) {
  sheet.getRange(1, 1, 1, 10).setValues([[
    'Email', 'Name', 'Role', 'Shift', 'Room', 'QTC',
    'Password Hash (SHA-256)', 'Password (Plain)', 'Created At', 'Status'
  ]]);
}


// ════════════════════════════════════════════════════════════
//  6. تنسيق الـ Header
// ════════════════════════════════════════════════════════════
function _styleHeader(sheet) {
  var hdr = sheet.getRange(1, 1, 1, 10);
  hdr.setBackground('#1A2B5F')
     .setFontColor('#2ABFBF')
     .setFontWeight('bold')
     .setFontSize(11)
     .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}


// ════════════════════════════════════════════════════════════
//  7. تنسيق الـ sheet كاملة
// ════════════════════════════════════════════════════════════
function _styleUsersSheet(sheet, dataRows) {
  if (dataRows < 1) return;

  // Admin row — لون مختلف
  sheet.getRange(2, 1, 1, 10)
       .setBackground('#2d1515')
       .setFontColor('#fca5a5');

  // Labeler rows
  if (dataRows > 1) {
    sheet.getRange(3, 1, dataRows - 1, 10)
         .setBackground('#0d1b2e')
         .setFontColor('#e8edf5');
  }

  // Status column — لون حسب القيمة
  var statusRange = sheet.getRange(2, USERS_COL.STATUS, dataRows, 1);
  statusRange.setFontWeight('bold');

  // عرض الـ columns
  sheet.setColumnWidth(1, 220);   // Email
  sheet.setColumnWidth(2, 160);   // Name
  sheet.setColumnWidth(3, 140);   // Role
  sheet.setColumnWidth(4, 150);   // Shift
  sheet.setColumnWidth(5, 140);   // Room
  sheet.setColumnWidth(6, 80);    // QTC
  sheet.setColumnWidth(7, 300);   // Hash
  sheet.setColumnWidth(8, 160);   // Plain
  sheet.setColumnWidth(9, 140);   // Created
  sheet.setColumnWidth(10, 80);   // Status

  // إخفاء عمود الـ Hash (اختياري — الأمان)
  // sheet.hideColumns(USERS_COL.PASS_HASH);
}


// ════════════════════════════════════════════════════════════
//  8. دالة LOGIN — بتتستخدم من الـ Web App
//     بترجع user object أو null
// ════════════════════════════════════════════════════════════
function loginUser(email, plainPassword) {
  try {
    var ss       = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    var usersTab = ss.getSheetByName(CONFIG.USERS_TAB);

    if (!usersTab) return { success: false, error: 'USERS tab not found' };

    var data     = usersTab.getDataRange().getValues();
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
          name    : row[USERS_COL.NAME  - 1],
          role    : row[USERS_COL.ROLE  - 1],
          shift   : row[USERS_COL.SHIFT - 1],
          room    : row[USERS_COL.ROOM  - 1],
          qtc     : row[USERS_COL.QTC   - 1],
        };
      }
    }

    return { success: false, error: 'Invalid credentials' };

  } catch(e) {
    return { success: false, error: e.message };
  }
}


// ════════════════════════════════════════════════════════════
//  9. Web App entry point
// ════════════════════════════════════════════════════════════
function doPost(e) {
  var params = JSON.parse(e.postData.contents);
  var action = params.action;

  if (action === 'login') {
    var result = loginUser(params.email, params.password);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  // الـ frontend HTML
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('MEEG Operations Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
