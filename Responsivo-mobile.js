function onOpen() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var columns = sheets[i].getMaxColumns();
    for (var j = 1; j <= columns; j++) {
      sheets[i].autoResizeColumn(j);
    }
  }
}