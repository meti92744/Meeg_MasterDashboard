// ============================================================
//  MEEG Operations Dashboard — Auth.gs
//  يقرأ TEAMS tab ويولد USERS tab بكل الأكونتات تلقائياً
//  ملاحظة: doPost موجود في Router.gs بس
// ============================================================

var CONFIG = {
  MASTER_SHEET_ID : '1_Tg0SNhFcRYy00U5GyUADnRLjoeHp0hOYxEmC3Z6bxI',
  TEAMS_TAB       : 'TEAMS',
  USERS_TAB       : 'USERS',
  PASSWORD_PREFIX : '$M@@g#ME',
  ADMIN_EMAIL     : 'admin@meeg.ai',
  ADMIN_PASSWORD  : 'MEEGAdmin@2025',
  ADMIN_NAME      : 'Meti Admin',
};

var TEAMS_COL = { EMAIL:1, NAME:2, ROOM:3, SHIFT:4, QTC:5 };

var USERS_COL = {
  EMAIL:1, NAME:2, ROLE:3, SHIFT:4, ROOM:5, QTC:6,
  PASS_HASH:7, PASS_PLAIN:8, CREATED_AT:9, STATUS:10,
};

// ════════════════════════════════════════════════════════════
//  generateUsersFromTeams — شغّلها مرة واحدة من الـ Editor
// ════════════════════════════════════════════════════════════
function generateUsersFromTeams() {
  var ss       = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
  var teamsTab = ss.getSheetByName(CONFIG.TEAMS_TAB);
  if (!teamsTab) { SpreadsheetApp.getUi().alert('❌ مش لاقي TEAMS tab'); return; }

  var usersTab = ss.getSheetByName(CONFIG.USERS_TAB);
  if (!usersTab) {
    usersTab = ss.insertSheet(CONFIG.USERS_TAB);
  } else {
    if (usersTab.getLastRow() > 1)
      usersTab.getRange(2,1,usersTab.getLastRow()-1,usersTab.getLastColumn()).clearContent();
  }

  _writeHeader(usersTab);
  _styleHeader(usersTab);

  var teamsData = teamsTab.getDataRange().getValues();
  var created   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var rows      = [];

  rows.push(_buildUserRow(
    CONFIG.ADMIN_EMAIL, CONFIG.ADMIN_NAME,
    'Admin', 'All', 'All', 'All', CONFIG.ADMIN_PASSWORD, created
  ));

  var skipped = 0;
  for (var i = 1; i < teamsData.length; i++) {
    var row   = teamsData[i];
    var email = String(row[0]).trim();
    var name  = String(row[1]).trim();
    var room  = String(row[2]).trim();
    var shift = String(row[3]).trim();
    var qtc   = String(row[4]).trim();

    if (!email || email.toLowerCase() === 'email') { skipped++; continue; }

    rows.push(_buildUserRow(email, name, 'Labeler', shift, room, qtc, _buildLabelerPassword(email), created));
  }

  if (rows.length > 0)
    usersTab.getRange(2, 1, rows.length, 10).setValues(rows);

  _styleUsersSheet(usersTab, rows.length);

  SpreadsheetApp.getUi().alert(
    '✅ تم!\n👤 Admin: 1\n👥 Labelers: ' + (rows.length-1) + '\n⏭️ Skipped: ' + skipped
  );
}

// ════════════════════════════════════════════════════════════
//  LOGIN — بيتستخدم من Router.gs
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
      var rowEmail  = String(row[USERS_COL.EMAIL    -1]).trim().toLowerCase();
      var rowHash   = String(row[USERS_COL.PASS_HASH-1]).trim();
      var rowStatus = String(row[USERS_COL.STATUS   -1]).trim();

      if (rowEmail === email.trim().toLowerCase() && rowHash === inputHash) {
        if (rowStatus !== 'Active') return { success: false, error: 'Account inactive' };
        return {
          success : true,
          email   : rowEmail,
          name    : String(row[USERS_COL.NAME -1]).trim(),
          role    : String(row[USERS_COL.ROLE -1]).trim(),
          shift   : String(row[USERS_COL.SHIFT-1]).trim(),
          room    : String(row[USERS_COL.ROOM -1]).trim(),
          qtc     : String(row[USERS_COL.QTC  -1]).trim(),
        };
      }
    }
    return { success: false, error: 'Invalid credentials' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════
function _buildLabelerPassword(email) {
  var m = email.match(/^me(\d+)@/i);
  return CONFIG.PASSWORD_PREFIX + (m ? m[1].substring(0,4) : '0000');
}

function _hashPassword(plainText) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, plainText, Utilities.Charset.UTF_8
  );
  return bytes.map(function(b){ return ('0'+(b&0xFF).toString(16)).slice(-2); }).join('');
}

function _buildUserRow(email, name, role, shift, room, qtc, plainPass, createdAt) {
  return [email, name, role, shift, room, qtc, _hashPassword(plainPass), plainPass, createdAt, 'Active'];
}

function _writeHeader(sheet) {
  sheet.getRange(1,1,1,10).setValues([[
    'Email','Name','Role','Shift','Room','QTC',
    'Password Hash (SHA-256)','Password (Plain)','Created At','Status'
  ]]);
}

function _styleHeader(sheet) {
  sheet.getRange(1,1,1,10)
    .setBackground('#1A2B5F').setFontColor('#2ABFBF')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

function _styleUsersSheet(sheet, dataRows) {
  if (dataRows < 1) return;
  sheet.getRange(2,1,1,10).setBackground('#2d1515').setFontColor('#fca5a5');
  if (dataRows > 1)
    sheet.getRange(3,1,dataRows-1,10).setBackground('#0d1b2e').setFontColor('#e8edf5');
  [220,160,140,150,140,80,300,160,140,80].forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
}
