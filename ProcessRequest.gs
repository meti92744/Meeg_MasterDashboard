// ============================================================
//  MEEG — ProcessRequest.gs
//  الدالة اللي بيستدعيها الـ frontend مباشرة
//  google.script.run.processRequest(jsonString)
// ============================================================

function processRequest(jsonString) {
  try {
    var params = JSON.parse(jsonString);
    var action = params.action;

    // ── LOGIN ──────────────────────────────────────────────
    if (action === 'login') {
      var result = loginUser(params.email, params.password);
      if (result.success) {
        result.sessionToken = _createSession(result);
      }
      return result;
    }

    // ── LOGOUT ─────────────────────────────────────────────
    if (action === 'logout') {
      _destroySession(params.sessionToken);
      return { success: true };
    }

    // ── كل التاني محتاج session valid ─────────────────────
    var session = _validateSession(params.sessionToken);
    if (!session) {
      return { success: false, error: 'Session expired', code: 401 };
    }

    // ── GET DASHBOARD HTML ─────────────────────────────────
    if (action === 'getDashboardHtml') {
      return { success: true, html: _getDashboardHtml(params.view, session) };
    }

    // ── ADMIN ACTIONS ──────────────────────────────────────
    if (action.startsWith('admin_')) {
      var adminParams = JSON.parse(JSON.stringify(params));
      adminParams.action = action.replace('admin_', '');
      return handleAdminRequest(adminParams, session);
    }

    // ── DATA REQUESTS ──────────────────────────────────────
    return handleDataRequest(params, session);

  } catch(e) {
    Logger.log('processRequest error: ' + e.message);
    return { success: false, error: e.message };
  }
}
