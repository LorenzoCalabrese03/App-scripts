function importaEventiDaProgetti() {
  var ssDest = SpreadsheetApp.getActiveSpreadsheet();

  var progetti = [
    { id: "ID_PROJECTS" },
    { id: "ID_PROJECTS" },
    { id: "ID_PROJECTS"}
  ];

  var fogliSemestri = ["I_Semestre", "II_Semestre"];
  var tz = Session.getScriptTimeZone();

  progetti.forEach(function(progetto) {
    var ssSource = SpreadsheetApp.openById(progetto.id);

    fogliSemestri.forEach(function(nomeFoglio) {
      var sh = ssSource.getSheetByName(nomeFoglio);
      if (!sh) return;

      var lastRow = sh.getLastRow();
      if (lastRow < 2) return;

      var dati = sh.getRange(2, 2, lastRow - 1, 2).getValues();

      dati.forEach(function(riga) {
        var dataEvento = riga[0];
        var nomeEvento = riga[1];

        if (!dataEvento || !nomeEvento) return;
        if (!(dataEvento instanceof Date)) return;

        var mese = dataEvento.getMonth();
        var nomeMese = getNomeMese(mese + 1);
        var shDest = ssDest.getSheetByName(nomeMese);
        if (!shDest) return;

        var evDate = Utilities.formatDate(dataEvento, tz, "dd/MM/yyyy");

        var lastRowDest = shDest.getLastRow();
        var lastColDest = shDest.getLastColumn();

        var trovato = false;
        for (var col = 2; col <= lastColDest; col += 2) {
          var valoriCol = shDest.getRange(1, col, lastRowDest).getValues();
          for (var r = 0; r < valoriCol.length; r++) {
            var cellVal = valoriCol[r][0];
            if (cellVal instanceof Date) {
              var cellDate = Utilities.formatDate(cellVal, tz, "dd/MM/yyyy");
              if (cellDate === evDate) {
                var startRow = r + 2;
                var lastRowCheck = shDest.getLastRow();
                var eventiColonna = shDest.getRange(startRow, col, lastRowCheck - startRow + 1).getValues().flat();

                if (eventiColonna.includes(nomeEvento)) {
                  Logger.log("⏩ Evento già presente: '" + nomeEvento + "' in " + nomeMese + " (" + evDate + ")");
                } else {
                  var rowToWrite = startRow;
                  while (shDest.getRange(rowToWrite, col).getValue()) {
                    rowToWrite++;
                  }
                  shDest.getRange(rowToWrite, col).setValue(nomeEvento);
                  Logger.log("✅ Evento scritto: '" + nomeEvento + "' in " + nomeMese + " (" + evDate + ")");
                }

                trovato = true;
                break;
              }
            }
          }
          if (trovato) break;
        }

        if (!trovato) {
          Logger.log("⚠️ Data " + evDate + " non trovata in foglio " + nomeMese);
        }
      });
    });
  });
}

function getNomeMese(meseNum) {
  var mesi = ["Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
              "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"];
  return mesi[meseNum-1];
}

