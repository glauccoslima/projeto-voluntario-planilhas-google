function onEdit(event) {
  if (!event || !event.source) {
    return;
  }

  var timezone = "GMT-3";
  var timestamp_format = "dd"; // Timestamp Format. 
  var updateColName = "NOME E SOBRENOME"; // novo nome da coluna
  var publicationColName = "PUBLICAÇÃO";
  var timeStampColName = "DIA ";
  var prepositionsAndArticles = ['e', 'da', 'de', 'do', 'das', 'dos', 'a', 'an', 'and', 'the'];
  var sheet = event.source.getActiveSheet();

  var headers = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName) + 1;
  var publicationCol = headers[0].indexOf(publicationColName) + 1;

  if (dateCol > -1) {
    var numRows = sheet.getLastRow() - 4;
    var range = sheet.getRange(5, updateCol, numRows, 3);
    var values = range.getValues();

    for (var i = 0; i < numRows; i++) {
      var row = i + 5;
      var nameCell = sheet.getRange(row, updateCol);
      var nameValue = values[i][0];
      var publicationCell = sheet.getRange(row, publicationCol);
      var publicationValue = values[i][1];
      var dateCell = sheet.getRange(row, dateCol + 1);
      var dateValue = dateCell.getValue();

      if (!nameCell.isBlank()) {
        if (nameValue != "") {
          var nameArray = nameValue.split(" ");
          for (var j = 0; j < nameArray.length; j++) {
            if (j === 0 || j === nameArray.length - 1) {
              nameArray[j] = nameArray[j].charAt(0).toUpperCase() + nameArray[j].slice(1).toLowerCase();
            } else if (!prepositionsAndArticles.includes(nameArray[j].toLowerCase())) {
              nameArray[j] = nameArray[j].charAt(0).toUpperCase() + nameArray[j].slice(1).toLowerCase();
            }
          }
          var nameValueCapitalized = nameArray.join(" ");
          nameCell.setValue(nameValueCapitalized);
        }
      }

      if (!publicationCell.isBlank()) {
        if (publicationValue != "") {
          publicationValue = publicationValue.charAt(0).toUpperCase() + publicationValue.slice(1);
          publicationCell.setValue(publicationValue);
        }
      }

      if (nameValue == "" && publicationValue == "" && dateValue != "") {
        dateCell.clearContent();
        dateCell.offset(0, 1).removeCheckboxes();
      } else if (nameValue != "" && publicationValue != "" && dateValue == "") {
        var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
        dateCell.setValue(date);
        dateCell.offset(0, 1).insertCheckboxes();
      }
    }
  }

// Protege as células A1:E4
var rangeToProtect = sheet.getRange(1, 1, 4, 5);
var protection = rangeToProtect.protect().setDescription('Protected Range');

// Garante que somente o usuário atual possa editar a planilha protegida
var me = Session.getEffectiveUser();
protection.addEditor(me);
protection.removeEditors(protection.getEditors());

if (protection.canDomainEdit()) {
  protection.setDomainEdit(false);
}
if (protection.canDeleteColumns()) {
protection.setDeleteColumns(false);
}
if (protection.canDeleteRows()) {
protection.setDeleteRows(false);
}
if (protection.canEditRange()) {
protection.setEditRange(null);
}
if (protection.canFormatCells()) {
protection.setFormatCells(false);
}
if (protection.canFormatColumns()) {
protection.setFormatColumns(false);
}
if (protection.canFormatRows()) {
protection.setFormatRows(false);
}
if (protection.canInsertColumns()) {
protection.setInsertColumns(false);
}
if (protection.canInsertRows()) {
protection.setInsertRows(false);
}
if (protection.canSort()) {
protection.setSort(false);
}
if (protection.canUsePivotTables()) {
protection.setUsePivotTables(false);
}

// Adiciona a validação de dados às colunas A e B para todas as novas linhas adicionadas à planilha
var dataRangeA = sheet.getRange("A2:A");
var ruleA = SpreadsheetApp.newDataValidation().requireValueInRange(dataRangeA).build();
sheet.getRange("A2").setDataValidation(ruleA);

var dataRangeB = sheet.getRange("B2:B");
var ruleB = SpreadsheetApp.newDataValidation().requireValueInRange(dataRangeB).build();
sheet.getRange("B2").setDataValidation(ruleB);

// Adiciona um gatilho para executar a função onEdit() sempre que uma nova linha é adicionada à planilha
ScriptApp.newTrigger('onEdit')
.forSpreadsheet(sheet)
.onEdit()
.create();
}