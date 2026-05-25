// ============================================================
//  MEEG Operations Dashboard — DataAPI.gs
//  بيجيب البيانات من الشيتات ويفلترها حسب الـ user
// ============================================================

// ── TAB NAMES في الـ Master Sheet ──────────────────────────
var TABS = {
  SUBMITTED_TASKS : 'Submitted Tasks',
  ATOT_PER_DAY    : 'ATOT Per Day',
  ATOT_SHIFT      : 'ATOT Shift',
  QUALITY         : 'Quality',
  TEAMS           : 'TEAMS',
  USERS           : 'USERS',
};

// ── Column names المفروض تكون موجودة (عدّل حسب شيتاتك) ───
var SUBMITTED_COLS = {
  EMAIL    : 'Email',
  DATE     : 'Date',
  SHIFT    : 'Shift',
  ROOM     : 'Room',
  QUEUE    : 'Queue',
  MODALITY : 'Modality',
  PASS     : 'Pass',
  OBJ_CNT  : 'Object Count',
  DEVICE   : 'Device',
};


// ════════════════════════════════════════════════════════════
//  HELPER — جيب الـ sheet data كـ array of objects
// ════════════════════════════════════════════════════════════
function _getSheetAsObjects(tabName) {
  var ss    = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) return [];

  var data    = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0].map(function(h){ return String(h).trim(); });
  var result  = [];

  for (var i = 1; i < data.length; i++) {
    var obj = {};
    headers.forEach(function(h, j){ obj[h] = data[i][j]; });
    result.push(obj);
  }
  return result;
}


// ════════════════════════════════════════════════════════════
//  1. بيانات الـ Labeler — يرجع بس بيانات الإيميل ده
// ════════════════════════════════════════════════════════════
function getLabelerData(email) {
  var tasks   = _getSheetAsObjects(TABS.SUBMITTED_TASKS);
  var atot    = _getSheetAsObjects(TABS.ATOT_PER_DAY);
  var quality = _getSheetAsObjects(TABS.QUALITY);

  // فلتر على الإيميل
  var myTasks = tasks.filter(function(r){
    return String(r[SUBMITTED_COLS.EMAIL]).trim().toLowerCase() === email.toLowerCase();
  });

  var myAtot = atot.filter(function(r){
    return String(r['Email'] || r['email'] || '').trim().toLowerCase() === email.toLowerCase();
  });

  var myQuality = quality.filter(function(r){
    return String(r['Email'] || r['email'] || '').trim().toLowerCase() === email.toLowerCase();
  });

  // إجمالي اليوم
  var today     = _todayStr();
  var todayTasks = myTasks.filter(function(r){
    return _dateStr(r[SUBMITTED_COLS.DATE]) === today;
  });

  return {
    totalTasksToday : todayTasks.length,
    tasksByQueue    : _groupBy(todayTasks, SUBMITTED_COLS.QUEUE),
    allTasks        : myTasks,
    atot            : myAtot,
    quality         : myQuality,
    latestQuality   : myQuality.length ? myQuality[myQuality.length - 1] : null,
  };
}


// ════════════════════════════════════════════════════════════
//  2. بيانات الـ Floor Supervisor — روم واحدة
// ════════════════════════════════════════════════════════════
function getFloorSupervisorData(room, shift, dateFilter) {
  var tasks   = _getSheetAsObjects(TABS.SUBMITTED_TASKS);
  var atot    = _getSheetAsObjects(TABS.ATOT_PER_DAY);
  var quality = _getSheetAsObjects(TABS.QUALITY);
  var teams   = _getSheetAsObjects(TABS.TEAMS);

  // فلتر الروم والشيفت
  var filtered = tasks.filter(function(r){
    var matchRoom  = !room  || String(r[SUBMITTED_COLS.ROOM ]).trim() === room;
    var matchShift = !shift || String(r[SUBMITTED_COLS.SHIFT]).trim() === shift;
    var matchDate  = !dateFilter || _dateStr(r[SUBMITTED_COLS.DATE]) === dateFilter;
    return matchRoom && matchShift && matchDate;
  });

  var roomTeam = teams.filter(function(r){
    return String(r['Room'] || '').trim() === room;
  });

  // Top labelers
  var byLabeler = _groupBy(filtered, SUBMITTED_COLS.EMAIL);
  var leaderboard = Object.keys(byLabeler)
    .map(function(email){
      var member = roomTeam.find(function(t){
        return String(t['Email']).trim().toLowerCase() === email.toLowerCase();
      }) || {};
      return {
        email  : email,
        name   : member['Name'] || email,
        qtc    : member['QTC']  || '—',
        count  : byLabeler[email].length,
      };
    })
    .sort(function(a,b){ return b.count - a.count });

  // Quality per labeler
  var qualityData = _getSheetAsObjects(TABS.QUALITY);
  var qualMap = {};
  qualityData.forEach(function(r){
    var em = String(r['Email'] || '').trim().toLowerCase();
    qualMap[em] = r['Accuracy'] || r['Score'] || r['Quality'] || 0;
  });

  return {
    totalTasks  : filtered.length,
    leaderboard : leaderboard,
    byQueue     : _groupBy(filtered, SUBMITTED_COLS.QUEUE),
    teamSize    : roomTeam.length,
    qualityMap  : qualMap,
    atot        : atot.filter(function(r){
      return String(r['Room'] || '').trim() === room;
    }),
  };
}


// ════════════════════════════════════════════════════════════
//  3. بيانات الـ Workload Manager — كل حاجة
// ════════════════════════════════════════════════════════════
function getWorkloadData(dateFilter) {
  var tasks = _getSheetAsObjects(TABS.SUBMITTED_TASKS);
  var atot  = _getSheetAsObjects(TABS.ATOT_SHIFT);

  var filtered = dateFilter
    ? tasks.filter(function(r){ return _dateStr(r[SUBMITTED_COLS.DATE]) === dateFilter; })
    : tasks;

  return {
    totalTasks    : filtered.length,
    byRoom        : _groupBy(filtered, SUBMITTED_COLS.ROOM),
    byShift       : _groupBy(filtered, SUBMITTED_COLS.SHIFT),
    byDate        : _groupBy(filtered, function(r){ return _dateStr(r[SUBMITTED_COLS.DATE]); }),
    atotShift     : atot,
  };
}


// ════════════════════════════════════════════════════════════
//  4. بيانات الـ Quality — كل الرومات
// ════════════════════════════════════════════════════════════
function getQualityData() {
  var quality = _getSheetAsObjects(TABS.QUALITY);
  var teams   = _getSheetAsObjects(TABS.TEAMS);

  return {
    raw       : quality,
    byRoom    : _groupByFn(quality, function(r){
      var email  = String(r['Email'] || '').trim().toLowerCase();
      var member = teams.find(function(t){
        return String(t['Email']).trim().toLowerCase() === email;
      });
      return member ? member['Room'] : 'Unknown';
    }),
    summary   : _qualitySummary(quality),
  };
}

function _qualitySummary(qualityData) {
  var scores = qualityData.map(function(r){
    return parseFloat(r['Accuracy'] || r['Score'] || r['Quality'] || 0);
  }).filter(function(s){ return s > 0; });

  if (!scores.length) return { avg: 0, max: 0, min: 0, above90: 0, below80: 0 };

  var sum    = scores.reduce(function(a,b){ return a+b; }, 0);
  return {
    avg     : (sum / scores.length).toFixed(1),
    max     : Math.max.apply(null, scores).toFixed(1),
    min     : Math.min.apply(null, scores).toFixed(1),
    above90 : scores.filter(function(s){ return s >= 90; }).length,
    below80 : scores.filter(function(s){ return s <  80; }).length,
    total   : scores.length,
  };
}


// ════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════
function _todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _dateStr(val) {
  if (!val) return '';
  try {
    var d = val instanceof Date ? val : new Date(val);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch(e) { return String(val).substring(0, 10); }
}

function _groupBy(arr, key) {
  var result = {};
  arr.forEach(function(item){
    var k = typeof key === 'function' ? key(item) : String(item[key] || '').trim();
    if (!result[k]) result[k] = [];
    result[k].push(item);
  });
  return result;
}

function _groupByFn(arr, fn) {
  return _groupBy(arr, fn);
}


// ════════════════════════════════════════════════════════════
//  doPost router — يستقبل calls من الـ frontend
// ════════════════════════════════════════════════════════════
function handleDataRequest(params, userSession) {
  var action = params.action;

  switch(action) {

    case 'getLabelerData':
      // Labeler يشوف بياناته بس
      if (userSession.role !== 'Labeler' && userSession.role !== 'Admin') {
        return { error: 'Access denied' };
      }
      var targetEmail = userSession.role === 'Admin'
        ? (params.email || userSession.email)
        : userSession.email;
      return getLabelerData(targetEmail);

    case 'getFloorData':
      if (!_hasAccess(userSession, ['Floor Supervisor','Shift Supervisor','Workload Manager','Admin'])) {
        return { error: 'Access denied' };
      }
      // Floor Supervisor يشوف روم واحدة بس (روم بتاعته)
      var room = userSession.role === 'Admin' ? params.room : userSession.room;
      return getFloorSupervisorData(room, params.shift, params.date);

    case 'getWorkloadData':
      if (!_hasAccess(userSession, ['Workload Manager','Admin'])) {
        return { error: 'Access denied' };
      }
      return getWorkloadData(params.date);

    case 'getQualityData':
      if (!_hasAccess(userSession, ['QA & Training','Shift Supervisor','Admin'])) {
        return { error: 'Access denied' };
      }
      return getQualityData();

    case 'getUsers':
      if (!_hasAccess(userSession, ['Admin'])) {
        return { error: 'Access denied' };
      }
      return { users: _getSheetAsObjects(TABS.USERS) };

    default:
      return { error: 'Unknown action: ' + action };
  }
}

function _hasAccess(userSession, allowedRoles) {
  return allowedRoles.indexOf(userSession.role) !== -1;
}
