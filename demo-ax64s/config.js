// Konfiguration des Launchers (statische Seite außerhalb von Databricks). Keine Geheimnisse hier.
// app: Adresse der Databricks App (nur Protokoll + Host, ohne Pfad). silentCheckMs: Wartezeit der stillen
// Sitzungsprüfung in Millisekunden (0 = immer direkt den Anmeldedialog öffnen).
window.DBX_LAUNCHER = {
  app: "",
  silentCheckMs: 2500,
  // Office im Web: Wartezeit je Anmeldeversuch im Aufgabenbereich (die Umleitungskette über Entra braucht Sekunden).
  webCheckMs: 8000,
};
