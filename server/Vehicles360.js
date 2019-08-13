var ControlSheetName = "Аналитика по ТС";
var DataSheetName = "Автопарк ТС";

function FromUnique() {
  var Array = [];
  var Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName);
  var Rows = Sheet.getDataRange();
  var numRows = Rows.getNumRows();
  var values = Rows.getValues();
  var Formula_QueryString = "=UNIQUE({'Журнал заказов'!D2:D})";
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName).getRange("B10:C").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName).getRange('C12').setValue(Formula_QueryString);
}

function FromList() {
  var Array = [];
  var Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName);
  var Rows = Sheet.getDataRange();
  var numRows = Rows.getNumRows();
  var values = Rows.getValues();
  var Formula_QueryString = "={'Автопарк ТС'!C2:D}";
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName).getRange("B10:C").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName).getRange('B12').setValue(Formula_QueryString);
}