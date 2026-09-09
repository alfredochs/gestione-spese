const SPREADSHEET_ID = '1pSot2PRSlUJruaEBsdqAqFf2epJgQU4I9R_hW-WiTrU';
const SHEET_NAME = 'Spese';

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var data = JSON.parse(e.postData.contents);

    var importo  = parseFloat(data.importo)  || 0;
    var categoria = data.categoria            || '';
    var nota      = data.nota                 || '';
    var tipo      = data.tipo                 || 'Personale';   // Personale | Casa | Anticipo
    var metodo    = data.metodo               || 'Contanti';    // Contanti | Carta
    var chi       = data.chi                  || 'Alfredo';     // Alfredo | Alessandra

    var now = new Date();
    var lastRow = sheet.getLastRow() + 1;

    sheet.appendRow([now, importo, categoria, nota, tipo, metodo, chi]);

    // Formato data DD/MM/YYYY nella colonna A della riga appena scritta
    sheet.getRange(lastRow, 1).setNumberFormat('dd/MM/yyyy');

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'OK', riga: lastRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ERRORE', messaggio: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Per test manuali: apri l'URL dello script nel browser
function doGet(e) {
  return ContentService
    .createTextOutput('Gestione Spese API attiva.')
    .setMimeType(ContentService.MimeType.TEXT);
}
