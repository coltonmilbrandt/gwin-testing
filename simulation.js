var ss = SpreadsheetApp.getActiveSpreadsheet()
var sheet = SpreadsheetApp.getActive().getSheetByName('Simulation');
var randomizer = sheet.getRange('change');
var profitPercentage;
var tranches = [];
var profitPercentages = [];
var gainDifferentials = [];

function myFunction() {
  if (!ss.output) {
    var output = SpreadsheetApp.getActive().getSheetByName('Output');
  } else {
    var output = ss.insertSheet('Output');
  }
  tranches.push(sheet.getRange('tranches').getValues());
  var length = tranches.length * 21;
  console.log(length);
  output.getRange(1,2,1,21).setValues(tranches);
  count = 0;
  for (var i = 1; i <= 50; i++) {
    refresh();
    // DO NOT change any values on the sheet or the data won't match !!!!!!
    profitPercentage = sheet.getRange('profitPercentage').getValue();
    gainDifferentials.push(sheet.getRange('gainDifferential').getValues());
    // profitPercentages.push([i],[profitPercentage],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""]);
    profitPercentages.push([profitPercentage]);
    console.log(randomizer);
    console.log(profitPercentages);
    console.log(tranches);
    console.log(gainDifferentials);
    count++;
  }
  output.getRange(2,1,count).setValues(profitPercentages);
  output.getRange(2,2,count,21).setValues(gainDifferentials);
}

function refresh() {
  if(randomizer == false) {
    randomizer = true;
  } else {
    randomizer = false;
  }
  sheet.getRange('change').setValue(randomizer);
}

function setRange(column, row) {
  var range = column + row;
  return range;
}
