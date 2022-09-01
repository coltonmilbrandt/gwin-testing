var ss = SpreadsheetApp.getActiveSpreadsheet()
var sheet = SpreadsheetApp.getActive().getSheetByName('Simulation');
var randomizer = sheet.getRange('change');
var profitPercentage;
var longRate;
var tranches = [];
var profitPercentages = [];
var gainDifferentials = [];
var longRates = [];
var profitRange = [];

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
  // Print tranch header to Output sheet
  output.getRange(1,2,1,21).setValues(tranches);

  output.getRange(1,23).setValue('Long Rates');
  // Count to determine cell ranges for when values are set on Output sheet
  count = 0;
  
  // For loop to run simulation i times
  for (var i = 1; i <= 100; i++) {
    refresh();
    // DO NOT change any values on the sheet or the data won't match !!!!!!
    
    // Retrieve values from Simulation Sheet
    longRate = sheet.getRange('longRate').getValue();
    profitPercentage = sheet.getRange('profitPercentage').getValue();
    gainDifferentials.push(sheet.getRange('gainDifferential').getValues());
    // profitPercentages.push([i],[profitPercentage],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""]);
    
    // Push profit percentage values to array
    profitPercentages.push([profitPercentage]);

    longRates.push([longRate]);
    console.log(profitPercentages);
    console.log(tranches);
    console.log(gainDifferentials);
    count++;
  }
  // Print profit percentages array in row 2, column 1 as y-axis key for gain differentials
  output.getRange(2,1,count).setValues(profitPercentages);
  // Print gain differentials array starting at row 2, column 2
  output.getRange(2,2,count,21).setValues(gainDifferentials);
  // Print long rates array starting at row 2, column 23
  output.getRange(2,23,count).setValues(longRates);
}

// Changes a checkbox on the spreadsheet to refresh values
function refresh() {
  if(randomizer == false) {
    randomizer = true;
  } else {
    randomizer = false;
  }
  // Flip checkbox to refresh values
  sheet.getRange('change').setValue(randomizer);
}

// Creates a range value
function setRange(column, row) {
  var range = column + row;
  return range;
}

function profitsSimulation() {
  // Create new sheet if it doesn't exist already
  if (!ss.output) {
    var output = SpreadsheetApp.getActive().getSheetByName('Profit');
  } else {
    var output = ss.insertSheet('Profit');
  }

  // Get tranch values from Simulation Sheet
  tranches.push(sheet.getRange('tranches').getValues());
  var length = tranches.length * 21;
  console.log(length);
  // Print tranch header to Profit Sheet
  output.getRange(1,2,1,21).setValues(tranches);

  // Print the long rates to Profit Sheet
  output.getRange(1,23).setValue('Long Rates');

  // Count to determine cell ranges for when values are set on Output sheet
  count = 0;
  
  // For loop to run simulation i times
  for (var i = 1; i <= 100; i++) {
    refresh();
    // DO NOT change any values on the sheet or the data won't match !!!!!!
    
    // Retrieve values from Simulation Sheet
    longRate = sheet.getRange('longRate').getValue();
    profitPercentage = sheet.getRange('profitPercentage').getValue();
    profitRange.push(sheet.getRange('UsdPercentageGain').getValues());
    // profitPercentages.push([i],[profitPercentage],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""],[""]);
    
    // Push profit percentage values to array
    profitPercentages.push([profitPercentage]);

    longRates.push([longRate]);
    console.log(profitPercentages);
    console.log(tranches);
    console.log(gainDifferentials);
    count++;
  }
  // Print profit percentages array in row 2, column 1 as y-axis key for profit range
  output.getRange(2,1,count).setValues(profitPercentages);
  // Print profit range array starting at row 2, column 2
  output.getRange(2,2,count,21).setValues(profitRange);
  // Print long rates array starting at row 2, column 23
  output.getRange(2,23,count).setValues(longRates);
}
