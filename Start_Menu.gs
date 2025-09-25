function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 Menu")
    .addItem("📅 Crea nuovo mese", "avviaCalendario")
    .addItem("📥 Scarica dati","prendiDati")
    .addToUi();

}

function avviaCalendario() {
  var ui = SpreadsheetApp.getUi();
  var input = ui.prompt("Inserisci il nome del mese (es. Ottobre)").getResponseText();
  if (!input) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nuovoFoglio = ss.insertSheet(input);
  nuovoFoglio.activate();

  // genera tutto il mese con settimane + tabelle
  generaCalendario();
}

function prendiDati(){
  importaEventiDaProgetti();
}
