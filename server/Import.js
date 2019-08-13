var ControlSheetName = "Управление";
var OrderLogSheetName = "Журнал заказов";

function ImportArray() {
  var Array = [];
  var Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName);
  var Rows = Sheet.getDataRange();
  var numRows = Rows.getNumRows();
  var values = Rows.getValues();
  for (var i = 2; i <= numRows - 1; i++) {
    if (Sheet.getRange('C'+(parseInt(i)+1)).getValues() != "") {
      var value = Sheet.getRange('C'+(parseInt(i)+1)+':E'+(parseInt(i)+1)).getValues() //B4:P6
      //Logger.log(value[0][0])
      Array.push('IMPORTRANGE("'+value[0][0]+'", "'+value[0][1]+'!'+value[0][2]+'");');    
    }
  }
  var test = Array.toString().replace(/;,/g, ";").replace(/;*$/, "");;
  var Formula_QueryString = 'QUERY({'+test+'},"select * where Col6 is not null",0)';
  
  var IsSortEnabled = Sheet.getRange(3,10).getValue();
  var IsAsc = Sheet.getRange(3,12).getValue();
  if (IsSortEnabled == true) {
    var SortByColumn = parseInt(Sheet.getRange(3,11).getValue());
      if (IsAsc == true) {
        var Formula_QueryString = 'SORT('+Formula_QueryString+','+SortByColumn+','+IsAsc+')';
      } else {
        var Formula_QueryString = 'SORT('+Formula_QueryString+','+SortByColumn+','+IsAsc+')';
      }
  }
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OrderLogSheetName).getRange('A2').setFormula(Formula_QueryString);
}
