// ============================================================
//  MEEG Operations Dashboard — Router.gs
//  الملف الرئيسي — بيستقبل كل الـ requests ويوزعها
// ============================================================

// ════════════════════════════════════════════════════════════
//  doGet — بيرجع الـ Login HTML
// ════════════════════════════════════════════════════════════
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('MEEG Operations Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// ════════════════════════════════════════════════════════════
//  doPost — الـ main router
//  كل request من الـ frontend بييجي هنا
// ════════════════════════════════════════════════════════════
function doPost(e) {
  var result;
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;

    // ── 1. LOGIN (مش محتاج session) ──────────────────────
    if (action === 'login') {
      result = loginUser(params.email, params.password);
      if (result.success) {
        result.sessionToken = _createSession(result);
      }
      return _json(result);
    }

    // ── 2. LOGOUT ─────────────────────────────────────────
    if (action === 'logout') {
      _destroySession(params.sessionToken);
      return _json({ success: true });
    }

    // ── 3. كل الـ actions التانية — محتاجة session ────────
    var session = _validateSession(params.sessionToken);
    if (!session) {
      return _json({ success: false, error: 'Session expired — please login again', code: 401 });
    }

    // ── 4. توزيع الـ requests ─────────────────────────────
    if (action.startsWith('admin_')) {
      params.action = action.replace('admin_', '');
      result = handleAdminRequest(params, session);

    } else if (action === 'getDashboardHtml') {
      result = { success: true, html: _getDashboardHtml(params.view, session) };

    } else {
      result = handleDataRequest(params, session);
    }

    return _json(result);

  } catch (err) {
    Logger.log('❌ Router error: ' + err.message + '\n' + err.stack);
    return _json({ success: false, error: 'Server error: ' + err.message });
  }
}


// ════════════════════════════════════════════════════════════
//  SESSION MANAGEMENT
//  بيستخدم CacheService — sessions بتنتهي بعد 6 ساعات
// ════════════════════════════════════════════════════════════

var SESSION_TTL = 21600; // 6 ساعات بالثواني

function _createSession(userObj) {
  var token   = _generateToken();
  var cache   = CacheService.getScriptCache();
  var payload = JSON.stringify({
    email : userObj.email,
    name  : userObj.name,
    role  : userObj.role,
    shift : userObj.shift,
    room  : userObj.room,
    qtc   : userObj.qtc,
    created: new Date().getTime(),
  });
  cache.put('session_' + token, payload, SESSION_TTL);
  Logger.log('🔑 Session created for: ' + userObj.email);
  return token;
}

function _validateSession(token) {
  if (!token) return null;
  var cache   = CacheService.getScriptCache();
  var payload = cache.get('session_' + token);
  if (!payload) return null;
  try {
    var session = JSON.parse(payload);
    // تجديد الـ TTL تلقائياً لو لسه شغال
    cache.put('session_' + token, payload, SESSION_TTL);
    return session;
  } catch(e) { return null; }
}

function _destroySession(token) {
  if (!token) return;
  CacheService.getScriptCache().remove('session_' + token);
  Logger.log('🔒 Session destroyed');
}

function _generateToken() {
  var chars  = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var token  = '';
  for (var i = 0; i < 48; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}


// ════════════════════════════════════════════════════════════
//  getDashboardHtml — بيرجع الـ HTML المناسب للـ role
// ════════════════════════════════════════════════════════════
function _getDashboardHtml(view, session) {
  // map من role → HTML file name
  var viewMap = {
    'Labeler'          : 'Labeler',
    'Floor Supervisor' : 'FloorSupervisor',
    'Shift Supervisor' : 'ShiftSupervisor',
    'QA & Training'    : 'QA',
    'Workload Manager' : 'WorkloadManager',
    'Admin'            : 'Admin',
  };

  var fileName = viewMap[session.role] || 'Labeler';

  // لو Admin طلب view مختلف
  if (session.role === 'Admin' && view && viewMap[view]) {
    fileName = viewMap[view];
  }

  try {
    var tmpl = HtmlService.createTemplateFromFile(fileName);
    tmpl.session = JSON.stringify(session);   // بنحقن الـ session في الـ HTML
    return tmpl.evaluate().getContent();
  } catch(e) {
    Logger.log('❌ getDashboardHtml error: ' + e.message);
    return '<p style="color:red">Error loading dashboard: ' + e.message + '</p>';
  }
}


// ════════════════════════════════════════════════════════════
//  HELPER — بيرجع JSON response
// ════════════════════════════════════════════════════════════
function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ════════════════════════════════════════════════════════════
//  include() — بيستخدمها الـ HTML templates
//  مثال: <?= include('styles') ?>
// ════════════════════════════════════════════════════════════
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ════════════════════════════════════════════════════════════
//  getDashboardHtml — callable من الـ frontend مباشرة
//  بيستخدمها Index.html بعد اللوجين
// ════════════════════════════════════════════════════════════
function getDashboardHtml(view, sessionObj) {
  // التحقق من الـ session object (مش token — ده للـ direct calls)
  if (!sessionObj || !sessionObj.role) {
    return '<p style="color:red">Access denied</p>';
  }
  return _getDashboardHtml(view, sessionObj);
}
