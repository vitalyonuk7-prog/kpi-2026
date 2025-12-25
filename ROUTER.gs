/** ====== ACCESS GUARD (file-based) ====== */
/** Toggle check quickly if needed */
const REQUIRE_FILE_ACCESS = true;

/** Optional domain restriction, e.g. 'bets.io'
 *  Leave '' to disable
 */
const ALLOW_DOMAIN = ''; // 'bets.io'

/** Check current user against file viewers/editors/owner */
function guardAccess_() {
  if (!REQUIRE_FILE_ACCESS) return null;

  // Current user email; can be "" if deployment is public
  var email = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || '';
  if (!email) return render403_('Your identity could not be verified (empty email).');

  if (ALLOW_DOMAIN) {
    var domain = String(email.split('@')[1] || '').toLowerCase();
    if (domain !== ALLOW_DOMAIN.toLowerCase()) {
      return render403_('Your email domain is not allowed: ' + email);
    }
  }

  var ss   = SpreadsheetApp.getActive();
  var file = DriveApp.getFileById(ss.getId());

  // Collect owner/editors/viewers
  var allowed = new Set();
  try { allowed.add((file.getOwner() && file.getOwner().getEmail()) || ''); } catch(_) {}
  try { file.getEditors().forEach(function(u){ if(u) allowed.add(u.getEmail()); }); } catch(_) {}
  try { file.getViewers().forEach(function(u){ if(u) allowed.add(u.getEmail()); }); } catch(_) {}

  if (!allowed.has(email)) {
    return render403_('Access to this file is required for: ' + email);
  }
  return null; // access granted
}

/** -------- 403 page -------- */
function render403_(reason) {
  var esc = function(s){ return String(s||'').replace(/[&<>"]/g, function(c){
    return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]);
  });};
  var html = [];
  html.push('<div style="font-family:Segoe UI,Arial,sans-serif;padding:24px">');
  html.push('<h1 style="margin:0 0 8px;color:#b91c1c">403 — Access denied</h1>');
  html.push('<p style="opacity:.9;margin:0 0 12px">This web app is restricted to users who have access to the Spreadsheet file.</p>');
  if (reason) html.push('<div style="font-family:monospace;background:#111827;color:#e5e7eb;padding:10px 12px;border-radius:8px">'+esc(reason)+'</div>');
  html.push('</div>');
  return HtmlService.createHtmlOutput(html.join(''))
    .setTitle('Access denied')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** -------- 404 page -------- */
function render404_(viewTried, allowedKeys) {
  var esc = function(s){ return String(s||'').replace(/[&<>"]/g, function(c){
    return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]);
  });};
  var html = [];
  html.push('<div style="font-family:Segoe UI,Arial,sans-serif;padding:24px">');
  html.push('<h1 style="margin:0 0 8px">404 — Page not found</h1>');
  html.push('<div style="opacity:.85;margin-bottom:16px">Requested view: <b>' + esc(viewTried) + '</b></div>');
  html.push('<div style="margin:8px 0 6px;font-weight:700">Allowed views:</div>');
  html.push('<ul style="margin:6px 0 0;padding-left:18px">');
  (allowedKeys||[]).forEach(function(k){ html.push('<li><code>?view=' + esc(k) + '</code></li>'); });
  html.push('</ul>');
  html.push('</div>');
  return HtmlService.createHtmlOutput(html.join(''))
    .setTitle('Not found')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ====== Web App router: exact mapping only (with access guard) ====== */
function doGet(e) {
  // 1) Access guard against file permissions
  var deny = guardAccess_();
  if (deny) return deny;

  // 2) Routing
  var view = (e && e.parameter && typeof e.parameter.view === 'string')
    ? e.parameter.view.trim().toLowerCase()
    : '';

  var routes = {
   
    'kpi':          { file: 'kpi_dash',     title: 'KPI' },

  };

  var route = routes[view];
  if (!route) return render404_(view, Object.keys(routes));

  return HtmlService.createHtmlOutputFromFile(route.file)
    .setTitle(route.title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
