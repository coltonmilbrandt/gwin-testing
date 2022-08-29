var ss = SpreadsheetApp.getActiveSpreadsheet()
var sheet = SpreadsheetApp.getActive().getSheetByName('Simulation');
var randomizer = sheet.getRange('change');
var profitPercentage;
var tranches = [];
var profitPercentages = [];
var gainDifferentials = [];

function myFunction() {
  // Create new sheet if it doesn't exist already
  if (!ss.output) {
    var output = SpreadsheetApp.getActive().getSheetByName('Output');
  } else {
    var output = ss.insertSheet('Output');
  }

  // Get tranch values from Simulation Sheet
  tranches.push(sheet.getRange('tranches').getValues());
  var length = tranches.length * 21;
  console.log(length);
  output.getRange(1,2,1,21).setValues(tranches);
  // Count to determine cell ranges for when values are set on Output sheet
  count = 0;
  
  // For loop to run simulation i times
  for (var i = 1; i <= 50; i++) {
    refresh();
    // DO NOT change any values on the sheet or the data won't match !!!!!!
    
    // Retrieve values from Simulation Sheet
    profitPercentage = sheet.getRange('profitPercentage').getValue();
    gainDifferentials.push(sheet.getRange('gainDifferential').getValues());
    // profitPercentages.push([i],[profitPercentage],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""]);
    
    // push read profit percentage values to array
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

// Changes a checkbox on the spreadsheet
function refresh() {
  if(randomizer == false) {
    randomizer = true;
  } else {
    randomizer = false;
  }
  sheet.getRange('change').setValue(randomizer);
}

// Creates a range value
function setRange(column, row) {
  var range = column + row;
  return range;
}
