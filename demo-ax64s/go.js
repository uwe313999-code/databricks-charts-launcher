// Dialog start page script (must run on the launcher's own origin); hands over to the app's sign-in target page,
// which the Databricks proxy serves only after the workspace sign-in. Separate file: go.html carries a CSP without inline script.
var app = String((window.DBX_LAUNCHER || {}).app || '').replace(/\/+$/, '');
// The app page itself reports the sign-in when opened with ?signin=1 (Databricks may end the sign-in on the app page instead of a
// deeper path); login-done.html stays for older launchers.
if (app) location.replace(app + '/?signin=1&parent=' + encodeURIComponent(location.origin));
else document.body.textContent = 'config.js: Adresse der App fehlt.';
