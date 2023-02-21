function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Pesquisar')
      .addItem('Mostrar Barra Lateral', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle('Barra lateral de pesquisa');

  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function searchValues(searchTerm) {
  if (!searchTerm || searchTerm.trim() === '') {
    Browser.msgBox('O campo de pesquisa está em branco');
    return;
  }
  
  var values = doSearch(searchTerm);
  var count = values.length;

  if (count > 0) {
    var message = count + ' valores correspondentes encontrados.\n\n';
    message += values.join('\n');
    Browser.msgBox(message);
  } else {
    Browser.msgBox('Nenhum valor correspondente encontrado');
  }
}

function validateSearchTerm(searchTerm) {
  if (!searchTerm || searchTerm.trim() === '') {
    throw new Error('O campo de pesquisa está em branco');
  }
}

function search() {
  var searchTerm = document.getElementById('searchTerm').value;
  try {
    validateSearchTerm(searchTerm);
    google.script.run.searchValues(searchTerm);
  } catch (e) {
    alert(e.message);
  }
}

function doSearch(searchTerm) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheets = spreadsheet.getSheets();
  var values = [];

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var dataRange = sheet.getDataRange();
    var sheetValues = dataRange.getValues();

    for (var row = 0; row < sheetValues.length; row++) {
      for (var col = 0; col < sheetValues[row].length; col++) {
        var cellValue = sheetValues[row][col].toString().toLowerCase();
        var term = searchTerm.toLowerCase();

        if (cellValue.indexOf(term) !== -1) {
          var cell = sheet.getRange(row + 1, col + 1);
          var cellName = cell.getA1Notation();
          var sheetName = sheet.getName();
          var value = 'Valor: ' + cellValue + '\nCélula: ' + cellName + '\nPlanilha: ' + sheetName;
          values.push(value);
        }
      }
    }
  }

  return values;
}