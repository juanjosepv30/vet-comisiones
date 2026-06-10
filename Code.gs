const SHEET_ID = '1YfhXzHGH4j23uilfU1XBlAvDdV0T9BjHK1nwugDgJKI';

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Vet Comisiones')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}
