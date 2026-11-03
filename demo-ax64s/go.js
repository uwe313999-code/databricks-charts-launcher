// Dialog start page script (must run on the launcher's own origin); hands over to the app's sign-in target page,
// which the Databricks proxy serves only after the workspace sign-in. Separate file: go.html carries a CSP without inline script.
var app = String((window.DBX_LAUNCHER || {}).app || '').replace(/\/+$/, '');
if (app) location.replace(app + '/login-done.html?parent=' + encodeURIComponent(location.origin));
else document.body.textContent = 'config.js: Adresse der App fehlt.';
