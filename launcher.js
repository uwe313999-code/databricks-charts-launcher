// Launcher (APP1): the add-in's start page outside Databricks. Databricks Apps serve nothing without the workspace
// sign-in, and the Microsoft sign-in page refuses to run inside the task pane, so: (1) silent check whether the app
// session already exists (hidden frame on the app's session-check page), (2) otherwise the sign-in in an Office dialog
// (go.html → app's login-done page), (3) then the pane itself navigates to the app. Holds no data and no tokens.
(function () {
  var cfg = window.DBX_LAUNCHER || {};
  var APP = String(cfg.app || '').replace(/\/+$/, '');
  var CHECK_MS = typeof cfg.silentCheckMs === 'number' && cfg.silentCheckMs >= 0 ? cfg.silentCheckMs : 2500;
  // Office on the web: the check runs the whole sign-in redirect chain inside the pane's cookie partition – give it time.
  var WEB_CHECK_MS = typeof cfg.webCheckMs === 'number' && cfg.webCheckMs > 0 ? cfg.webCheckMs : 8000;
  var platform = null;
  var out = document.getElementById('out');
  var statusEl = document.getElementById('status');
  var actions = document.getElementById('actions');
  var done = false;

  function log() { var parts = []; for (var i = 0; i < arguments.length; i++) parts.push(typeof arguments[i] === 'string' ? arguments[i] : JSON.stringify(arguments[i])); out.textContent += '\n' + parts.join(' '); }
  function status(text) { statusEl.textContent = text; }
  function isLoginOk(origin, message) { return message === 'login-ok' && (!origin || origin === APP); }
  function go() { if (done) return; done = true; status('Angemeldet – Databricks Charts wird geladen …'); window.location.href = APP + '/taskpane.html'; }

  // (1) Already signed in? The app's session-check page answers "login-ok" only when the proxy let it through.
  function silentCheck(ms) {
    return new Promise(function (resolve) {
      if (!ms) { resolve(false); return; }
      var frame = document.createElement('iframe');
      frame.hidden = true; frame.setAttribute('aria-hidden', 'true'); frame.setAttribute('title', 'Sitzungsprüfung');
      var timer = setTimeout(function () { finish(false); }, ms);
      function onMessage(event) { if (event.origin === APP && event.data === 'login-ok') finish(true); }
      function finish(ok) { clearTimeout(timer); window.removeEventListener('message', onMessage); frame.remove(); log('Sitzungsprüfung', ok ? 'angemeldet' : 'keine Sitzung'); resolve(ok); }
      window.addEventListener('message', onMessage);
      frame.src = APP + '/session-check.html?parent=' + encodeURIComponent(window.location.origin);
      document.body.appendChild(frame);
    });
  }

  // (2) Sign-in in an Office dialog (a top-level window; the Microsoft sign-in refuses the task pane).
  function openDialog() {
    return new Promise(function (resolve) {
      var url = new URL('go.html', window.location.href).href;
      Office.context.ui.displayDialogAsync(url, { height: 60, width: 40 }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) { log('Dialog-Fehler', res.error.code, res.error.message); resolve(false); return; }
        var dialog = res.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          if (isLoginOk(arg.origin, arg.message)) { dialog.close(); resolve(true); } else log('Nachricht vom Dialog', arg.origin || '', arg.message);
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) { log('Dialog-Ereignis', arg.error); resolve(false); });
      });
    });
  }

  function failed() {
    status('Nicht angemeldet. Mit „Anmelden“ erneut versuchen.' + (platform === 'OfficeOnline' ? ' Zeigt Databricks „Login error“: dort „Back to sign in“ klicken.' : ''));
    actions.hidden = false;
  }

  // Office on the web (APP1-E2E, 17./19.10.2026): the task pane is an iframe of the Office page, so Firefox/Safari keep
  // the app's cookies in a separate partition – a sign-in completed in the dialog never reaches the pane. The sign-in has
  // to run inside the pane's own context: the hidden frame runs the redirect chain there (an existing Entra session
  // completes it without any page), a second round trip finishes what the first only prepared; the dialog is only the
  // way to establish the Entra session when none exists.
  function startWeb(auto) {
    var ms = Math.max(CHECK_MS, WEB_CHECK_MS);
    status('Prüfe Anmeldung …');
    return (auto ? silentCheck(ms) : Promise.resolve(false)).then(function (ok) {
      if (ok) return go();
      status('Anmeldung im Aufgabenbereich …');
      return silentCheck(ms).then(function (ok2) {
        if (ok2) return go();
        status('Anmeldung bei Databricks …');
        return openDialog().then(function () {
          status('Anmeldung im Aufgabenbereich …');
          return silentCheck(ms).then(function (ok3) { if (ok3) return go(); failed(); });
        });
      });
    });
  }

  function start(auto) {
    if (!APP) { status('config.js: Adresse der App fehlt.'); actions.hidden = true; return Promise.resolve(); }
    actions.hidden = true;
    if (platform === 'OfficeOnline') return startWeb(auto);
    var first = auto ? (status('Prüfe Anmeldung …'), silentCheck(CHECK_MS)) : Promise.resolve(false);
    return first.then(function (signedIn) {
      if (signedIn) { go(); return; }
      status('Anmeldung bei Databricks …');
      return openDialog().then(function (ok) {
        if (ok) { go(); return; }
        failed();
      });
    });
  }

  document.getElementById('login').onclick = function () { start(false); };
  document.getElementById('open').onclick = function () { go(); };
  document.getElementById('diag').onclick = function () { done = true; status('Diagnose wird geladen …'); window.location.href = APP + '/diag.html'; };
  Office.onReady(function (info) { platform = info.platform; log('Office bereit', info.host, info.platform); start(true); setTimeout(function () { if (!done) actions.hidden = false; }, 4000); });
})();
