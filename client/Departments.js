/** 
*   CORE VARIABLES
*/

var ControlSheetName = "Справочник";
var ControlSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName);

var VEHICHLE_ID_CLOUMN_TO_CHECK = 4;
var ADDRESS_CLOUMN_TO_CHECK = 6;
var FACT_CLOUMN_TO_CHECK = 10;
var DD_POLY1_CLOUMN_TO_CHECK = 14;
var VALUE_POLY1_CLOUMN_TO_CHECK = 16;

/** 
*   OFFSET TABLE (ON EDIT)
*   format: [row, column1, column2]
*   column1 for formula offset based on row with address (first input UX)
*   column2 for script offset based on edit cell
*/  

var DEPARTMENT_NAME = [0,-5];
var DATE = [0,-4];
var VEHICHLE_MARK = [0,-1];
var DD_VEHICLES = [0,-2];
var DD_WORKTYPE = [0,3];
var FORMULA_FACTFROMCLIENT = [0,5];
var DD_METRICSFACTFROMCLIENT = [0,6];
var DD_COUNTERPARTIES = [0,7,12];
var DD_METRICSPOLYGON = [0,9,14];
var FORMULA_FACTFROMPOLY = [0,11,16];
var FORMULA_UTILIZATIONTOTAL = [0,17];
var FORMULA_FUELEXPENSES = [0,22];
var FORMULA_EXPENSES = [0,29];
var FORMULA_TOTAL = [0,30];

var DropDownVehicleColumn = "A";
var DropDownCounterpartiesColumn = "E";
var DropDownWorktypeColumn = "G";

/**
*    MAIN THREAD
*/

function DOCNAME() {
  var docName = SpreadsheetApp.getActiveSpreadsheet().getName();
  return docName
}

function DropDownForm_Vehicles() {
  var VehicleArray = [];
  var VehiclesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName);
  var Vehicles_rows = VehiclesSheet.getDataRange();
  var Vehicles_numRows = Vehicles_rows.getNumRows();
  var Vehicles_values = Vehicles_rows.getValues();
  for (var i = 1; i <= Vehicles_numRows - 1; i++) {
    var Vehicles_rows = Vehicles_values[i];
    var value = VehiclesSheet.getRange(DropDownVehicleColumn+(parseInt(i)+1)).getValues()
    VehicleArray.push(value);    
  }
  return VehicleArray
}

function DropDownForm_Counterparties() {
  var CounterpartiesArray = [];
  var CounterpartiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName);
  var Counterparties_rows = CounterpartiesSheet.getDataRange();
  var Counterparties_numRows = Counterparties_rows.getNumRows();
  var Counterparties_values = Counterparties_rows.getValues();
  for (var i = 1; i <= Counterparties_numRows - 1; i++) {
    var Vehicles_rows =Counterparties_values[i];
    var value = CounterpartiesSheet.getRange(DropDownCounterpartiesColumn+(parseInt(i)+1)).getValues()
    CounterpartiesArray.push(value);    
  }
  return CounterpartiesArray
}

function DropDownForm_Worktype() {
  var WorktypeArray = [];
  var WorktypeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName);
  var Worktype_rows = WorktypeSheet.getDataRange();
  var Worktype_numRows = Worktype_rows.getNumRows();
  var Worktype_values = Worktype_rows.getValues();
  for (var i = 1; i <= Worktype_numRows - 1; i++) {
    var Vehicles_rows = Worktype_values[i];
    var value = WorktypeSheet.getRange(DropDownWorktypeColumn+(parseInt(i)+1)).getValues()
    WorktypeArray.push(value);    
  }
  return WorktypeArray
}

function PVLOOKUP(column, index, value) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName);
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(1,column,lastRow,column+index).getValues();
  for(i = 0; i < data.length; ++i){
    if (data[i][0] == value){
      return data[i][index];
    }
  }
}

function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var selectedCell = ss.getActiveCell();
  var range = e.range;
  var row = range.getRow();
  var column = range.getColumn();
  var search_value = range.getValues();
  if (sheet.getSheetName() == "Заказы") {
    if (row > 3) {
      if (selectedCell.getColumn() == ADDRESS_CLOUMN_TO_CHECK) { 
        var CheckOnInsert = sheet.getRange("B"+row);
        if (CheckOnInsert.isBlank() == true) {
          selectedCell.offset(DATE[0],DATE[1]).setValue(new Date());
        }
        var DepartmentName = SpreadsheetApp.getActiveSpreadsheet().getName();
        selectedCell.offset(DEPARTMENT_NAME[0],DEPARTMENT_NAME[1]).setValue(DepartmentName+'-'+row+'-'+Math.floor((new Date().getTime()/1000)).toString());
        
        var VehiclesArray = DropDownForm_Vehicles();
        var VEHICLES = SpreadsheetApp.newDataValidation().requireValueInList(VehiclesArray).build();
        selectedCell.offset(DD_VEHICLES[0],DD_VEHICLES[1]).setDataValidation(VEHICLES);
        
        var WorktypeArray = DropDownForm_Worktype();
        var WORKTYPE = SpreadsheetApp.newDataValidation().requireValueInList(WorktypeArray).build();
        selectedCell.offset(DD_WORKTYPE[0],DD_WORKTYPE[1]).setDataValidation(WORKTYPE);

        selectedCell.offset(FORMULA_FACTFROMCLIENT[0],FORMULA_FACTFROMCLIENT[1]).setFormula('=(E'+row+'*J'+row+')');

        var METRICS = SpreadsheetApp.newDataValidation().requireValueInList(['м3', 'тонны']).build();
        selectedCell.offset(DD_METRICSFACTFROMCLIENT[0],DD_METRICSFACTFROMCLIENT[1]).setDataValidation(METRICS);
        
        var CounterpartiesArray = DropDownForm_Counterparties();
        var COUNTERPARTIES = SpreadsheetApp.newDataValidation().requireValueInList(CounterpartiesArray).build();
        selectedCell.offset(DD_COUNTERPARTIES[0],DD_COUNTERPARTIES[1]).setDataValidation(COUNTERPARTIES);
        selectedCell.offset(DD_COUNTERPARTIES[0],DD_COUNTERPARTIES[2]).setDataValidation(COUNTERPARTIES);

        var METRICS = SpreadsheetApp.newDataValidation().requireValueInList(['м3', 'тонны', 'за машину']).build();
        selectedCell.offset(DD_METRICSPOLYGON[0],DD_METRICSPOLYGON[1]).setDataValidation(METRICS);
        selectedCell.offset(DD_METRICSPOLYGON[0],DD_METRICSPOLYGON[2]).setDataValidation(METRICS);

        selectedCell.offset(FORMULA_FACTFROMPOLY[0],FORMULA_FACTFROMPOLY[1]).setFormula('=(N'+row+'*P'+row+')')
        selectedCell.offset(FORMULA_FACTFROMPOLY[0],FORMULA_FACTFROMPOLY[2]).setFormula('=(U'+row+'*S'+row+')')

        selectedCell.offset(FORMULA_UTILIZATIONTOTAL[0],FORMULA_UTILIZATIONTOTAL[1]).setFormula('=(V'+row+'+Q'+row+')')

        selectedCell.offset(FORMULA_FUELEXPENSES[0],FORMULA_FUELEXPENSES[1]).setFormula('=(Z'+row+'*AA'+row+')')
        
        selectedCell.offset(FORMULA_EXPENSES[0],FORMULA_EXPENSES[1]).setFormula('=(AG'+(parseInt(row))+'+AF'+(parseInt(row))+'+AD'+(parseInt(row))+'+AC'+(parseInt(row))+'+AB'+(parseInt(row))+'+W'+(parseInt(row))+')')
        selectedCell.offset(FORMULA_TOTAL[0],FORMULA_TOTAL[1]).setFormula('=(AH'+(parseInt(row))+'-AI'+(parseInt(row))+')')
      }
      if (selectedCell.getColumn() == VEHICHLE_ID_CLOUMN_TO_CHECK) { 
        var RequestedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ControlSheetName);
        var ControlSheet_Argument = RequestedSheet.getRange(2,9).getValue();
        if (ControlSheet_Argument == "АВТОМАТИЧЕСКИ") {
          var lastRow = RequestedSheet.getLastRow();
          var data = RequestedSheet.getRange(1,1,lastRow,1+2).getValues();
          for(i = 0; i < data.length; ++i){
            if (data[i][0] == search_value){
              var marked = data[i][1];
              var cap = data[i][2];
              Logger.log(cap);
            }
          }
          var selectedCell = ss.getActiveCell();
          var DepartmentVehichleMarkCell = selectedCell.offset(0,-1);
          var DepartmentCapacityCell = selectedCell.offset(0,1);
          DepartmentVehichleMarkCell.setValue(marked)
          DepartmentCapacityCell.setValue(cap)
        }
        if (ControlSheet_Argument == "INDIRECT") {
          var DepartmentVehichleMarkCell = selectedCell.offset(VEHICHLE_MARK[0],VEHICHLE_MARK[1]);
          DepartmentVehichleMarkCell.setFormula('=INDIRECT("'+ControlSheetName+'!B"&MATCH(D'+(parseInt(row))+','+ControlSheetName+'!$A$1:$A,0))');
          selectedCell.offset(0,1).setFormula('=INDIRECT("'+ControlSheetName+'!C"&MATCH(D'+(parseInt(row))+','+ControlSheetName+'!$A$1:$A,0))');
        }
        if (ControlSheet_Argument == "pVLOOKUP") {
          var DepartmentVehichleMarkCell = selectedCell.offset(VEHICHLE_MARK[0],VEHICHLE_MARK[1]);
          DepartmentVehichleMarkCell.setFormula('=PVLOOKUP(1,1,D'+(parseInt(row))+')');
          selectedCell.offset(0,1).setFormula('=PVLOOKUP(1,2,D'+(parseInt(row))+')');
        }
        if (ControlSheet_Argument == "НЕ ИСПОЛЬЗОВАТЬ") {
          //NOTHING
        }
      }
    }
  }
}

