function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSh = ss.getSheetByName(SHEET_MASTER);
  var outSh    = ss.getSheetByName(SHEET_OUT);

  var masterData = masterSh.getDataRange().getValues();
  var inventory = masterData.slice(1)
    .filter(function(row) { return row[2] !== ''; })
    .map(function(row) {
      return {
        cat:    row[0], maker:  row[1], model: row[2],
        note:   row[3], keyword:row[4],
        total:  row[6], stock:  row[7],
        out:    row[8], status: row[9],
      };
    });

  var outData = outSh.getDataRange().getValues();
  var outRows = outData.slice(1)
    .filter(function(row) { return row[10] === '持ち出し中' && row[7] !== ''; })
    .map(function(row) {
      return {
        date:       row[0], fileName:   row[1],
        project:    row[2], staff:      row[3],
        dateOut:    row[4], dateReturn: row[5],
        category:   row[6], model:      row[7],
        qty:        row[8], note:       row[9],
        status:     row[10],
      };
    });

  return ContentService
    .createTextOutput(JSON.stringify({ inventory: inventory, out: outRows }))
    .setMimeType(ContentService.MimeType.JSON);
}
