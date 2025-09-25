function generaCalendario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var nomeFoglio = sheet.getName();  // Es. "Ottobre"
  
  var mesi = {
    "Gennaio": 1, "Febbraio": 2, "Marzo": 3, "Aprile": 4,
    "Maggio": 5, "Giugno": 6, "Luglio": 7, "Agosto": 8,
    "Settembre": 9, "Ottobre": 10, "Novembre": 11, "Dicembre": 12
  };
  
  var mese = mesi[nomeFoglio];
  if (!mese) {
    SpreadsheetApp.getUi().alert("Il nome del foglio non corrisponde a un mese valido.");
    return;
  }
  
  var anno = new Date().getFullYear();
  var primoGiornoMese = new Date(anno, mese - 1, 1);
  var ultimoGiornoMese = new Date(anno, mese, 0);

  var oggi = new Date();
  oggi.setHours(0,0,0,0);
  
  var offset = (1 - (primoGiornoMese.getDay() || 7) + 7) % 7;
  var primoLunedi = new Date(primoGiornoMese);
  primoLunedi.setDate(primoGiornoMese.getDate() + offset);
  
  var fineOffset = (7 - (ultimoGiornoMese.getDay() || 7)) % 7;
  var ultimaDomenica = new Date(ultimoGiornoMese);
  ultimaDomenica.setDate(ultimoGiornoMese.getDate() + fineOffset);
  
  var dataCorrente = new Date(primoLunedi);
  var settimana = 0;
  
  while (dataCorrente <= ultimaDomenica) {
    // 📌 calcola la riga per questa settimana
    var baseRow = 1 + settimana * 19;   // 1, 19, 37, …
    
    // scrive i 7 giorni
    for (var giorno = 0; giorno < 7; giorno++) {
      var col = 2 + giorno * 2; // B, D, F, ...
      var cella = sheet.getRange(baseRow, col);
      
      var giornoMese = Utilities.formatDate(dataCorrente, "Europe/Rome", "dd/MM/yyyy");
      cella.setValue(giornoMese);
      
      if (dataCorrente.getTime() === oggi.getTime()) {
        cella.setBackground("#fff176");
      } else {
        cella.setBackground("#e6e6e6");
      }
      
      dataCorrente.setDate(dataCorrente.getDate() + 1);
    }

    // 📌 copia la tabella sotto questa settimana
    copiaTabellaSettimana(baseRow+1);

    settimana++;
  }
}

function copiaTabellaSettimana(startRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var foglioModello = ss.getSheetByName("Modello_Tabella");

  if (!foglioModello) {
    SpreadsheetApp.getUi().alert("Non trovo il foglio 'Modello_Tabella' usato come modello.");
    return;
  }

  // 🔹 calcola automaticamente dimensioni tabella nel modello
  var lastRow = foglioModello.getLastRow();
  var lastCol = foglioModello.getLastColumn();

  var intervallo = foglioModello.getRange(1, 1, lastRow, lastCol);
  var destinazione = sheet.getRange(startRow, 1);

  intervallo.copyTo(destinazione, {contentsOnly: false});
}

