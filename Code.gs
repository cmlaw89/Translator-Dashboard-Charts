function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu('Wallace')
      .addItem('Trend Charts', 'viewCharts')
      .addToUi();
}

function viewCharts() {

  var html = HtmlService
  .createTemplateFromFile("Index")
  .evaluate()
  .setTitle("Test Chart")
  .setHeight(450)
  .setWidth(750)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  SpreadsheetApp.getUi().showModelessDialog(html, "Trend Charts")
}

function getDataTable(type) {
  
  var months = {1: "Jan",
                2: "Feb",
                3: "Mar",
                4: "Apr",
                5: "May",
                6: "Jun",
                7: "Jul",
                8: "Aug",
                9: "Sep",
                10: "Oct",
                11: "Nov",
                12: "Dec"}
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pivot_sheet = ss.getSheetByName("Quality History");

  
  //Extract the data based on the query request
  var translator = pivot_sheet.getRange("T1").getValue().split(" ")[1].slice(0, -1);
  var num_months = pivot_sheet.getRange("B1").getValue();
  var month_data = pivot_sheet.getRange(3, 1, num_months, 1).getValues();
  var TPR_data = pivot_sheet.getRange(3, 4, num_months, 1).getValues();
  var team_TPR_data = pivot_sheet.getRange(3, 14, num_months, 1).getValues();
  var TE_data = pivot_sheet.getRange(3, 5, num_months, 1).getValues();
  var team_TE_data = pivot_sheet.getRange(3, 15, num_months, 1).getValues();
  var cat_data = pivot_sheet.getRange(3, 6, num_months, 6).getValues();
  
  
  //Format the data for use in google.visualization.arrayToDataTable()
  if (type == "Cat") {
    var labels = [{label: 'Month', id: 'Month'}, 
                  {label: 'Omissions/additions', id: "O/a", type: 'number'},
                  {label: 'Terminology', id: "Term", type: 'number'},
                  {label: 'Mistranslation', id: "Mis", type: 'number'},
                  {label: 'Literal translation', id: "Lit", type: 'number'},
                  {label: 'Language', id: "Lan", type: 'number'},
                  {label: 'Non-compliance', id: "Non-c", type: 'number'},];
  }
  else {
    var labels = [{label: 'Month', id: 'Month'}, 
                  {label: translator, id: translator, type: 'number'},
                  {label: 'Team Average', id: "TAvg", type: 'number'}];
  }
  
  for (var i = 1; i < num_months; i++) {
    if (TPR_data[i][0].toString() == "") {
      TPR_data[i][0] = null;
    }
    if (team_TPR_data[i][0].toString() == "") {
      team_TPR_data[i][0] = null;
    }
    if (TE_data[i][0].toString() == "") {
      TE_data[i][0] = null;
    }
    if (team_TE_data[i][0].toString() == "") {
      team_TE_data[i][0] = null;
    }
  }

  var data_table = [labels];
  var start_month = 1
  if (num_months > 12) {
    start_month = num_months - 12;
  }  
  for (var i = start_month; i < num_months; i++) {
    if (type == "TPR") {
      data_table.push([months[month_data[i][0].getMonth() + 1] + " " + month_data[i][0].getFullYear().toString().slice(2,4),
                      TPR_data[i][0],
                      team_TPR_data[i][0]]);
    }
    else if (type == "TE") {
      data_table.push([months[month_data[i][0].getMonth() + 1] + " " + month_data[i][0].getFullYear().toString().slice(2,4),
                      TE_data[i][0],
                      team_TE_data[i][0]]);
    }
    else if (type == "Cat") {
      data_table.push([months[month_data[i][0].getMonth() + 1] + " " + month_data[i][0].getFullYear().toString().slice(2,4)]
                      .concat(cat_data[i]))
    }              
  }
  
  return data_table
}

function getDataTables(types) {
  
  var data_tables = [];
  for (var i = 0; i < types.length; i++) {
    data_tables.push(getDataTable(types[i]))
  }
  return data_tables
}


function include(filename) {
  //Adds stylesheet and javascript to Index.html
  
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}