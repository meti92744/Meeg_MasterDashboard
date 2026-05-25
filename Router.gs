// ============================================================
//  MEEG Operations Dashboard — Router.gs
// ============================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('MEEG Operations Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;

    // ── LOGIN ──────────────────────────────────────────────
    if (action === 'login') {
      var result = loginUser(params.email, params.password);
      return _json(result);
    }

    // ── كل الـ actions التانية ─────────────────────────────
    // بنقبل session object مباشرة من الـ frontend
    var session = params.session || {};

    if (!session.email || !session.role) {
      return _json({ success: false, error: 'Invalid session', code: 401 });
    }

    switch (action) {

      // ── Data ──────────────────────────────────────────────
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

      // ── Admin ─────────────────────────────────────────────
      case 'admin_getAllUsers':
        if (session.role !== 'Admin') return _json({ error: 'Access denied' });
        return _json(getAllUsers());

      case 'admin_addUser':
        if (session.role !== 'Admin') return _json({ error: 'Access denied' });
        return _json(addUserToSheet(params.user));

      case 'admin_editUser':
        if (session.role !== 'Admin') return _json({ error: 'Access denied' });
        return _json(editUserInSheet(params.email, params.updates));

      case 'admin_deleteUser':
        if (session.role !== 'Admin') return _json({ error: 'Access denied' });
        return _json(deleteUserFromSheet(params.email));

      case 'admin_toggleStatus':
        if (session.role !== 'Admin') return _json({ error: 'Access denied' });
        return _json(toggleUserStatus(params.email));

      case 'admin_resetPassword':
        if (session.role !== 'Admin') return _json({ error: 'Access denied' });
        return _json(resetUserPassword(params.email, params.newPassword));

      default:
        return _json({ error: 'Unknown action: ' + action });
    }

  } catch (err) {
    Logger.log('❌ Router error: ' + err.message);
    return _json({ success: false, error: 'Server error: ' + err.message });
  }
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
