// Google Apps Script for tracking new client meetings, using the "Date Added" parameter to check if the client is added within the quarter and has no previous meetings
// UPEI Research Services: SpringBoard requires metrics on the number of researcher, company, and industry meetings/events
// By: Alex O'Brien

// MAIN //

// @params: event
// @return: none
// Executes on edit
function onEdit(e) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // format dates (using date picker changes font for some reason)
  sh.getRange(5,2,sh.getMaxRows()).setFontFamily("Arial").setFontSize(10);

  var curCell = sh.getActiveCell();
  const data = initialize(sh, curCell);
  const quarters = data.quarter;
  const clients = data.client;
  var q = getQuarter(sh, quarters, curCell);

  if (e.value == null){
    var action = "del"
  }
  else{
    var action = "add"
  }

  if(action == "add"){
    if(dateCompare(clients.cDate, q.qStart, q.qEnd) && clients.cMeetingsCount == 1){
      updateQuarterTotal(q, 1);
    }
  }
  else{
    if(dateCompare(clients.cDate, q.qStart, q.qEnd) && clients.cMeetingsCount == 0 && q.qTotal.getValue() >= 1){
      updateQuarterTotal(q, -1);
    }
  }
}

// CONSTRUCTOR //

// @params: active sheet, active cell
// @return: {quarter object, client object}
// Sets up data to reference in the main function
function initialize(sh, cell){

  var client = {
    cName: sh.getRange(cell.getRow(), 4).getValue(),
    cDate: sh.getRange(cell.getRow(), 2).getValue(),
    cDetail: sh.getRange(cell.getRow(), 3).getValue(),
    cMeetingsCount: sh.getRange(cell.getRow(), 1).getValue()
  };

  var quarter = {
    quarter1 : {
      qStart : new Date("april 1 2023"),
      qEnd : new Date("june 30 2023"),
      qTotal : sh.getRange("I4"),
      qRange : sh.getRange("F5:H1000")
    },
    quarter2 : {
      qStart : new Date("july 1 2023"),
      qEnd : new Date("september 30 2023"),
      qTotal : sh.getRange("M4"),
      qRange : sh.getRange("J5:L1000")
    },
    quarter3 : {
      qStart : new Date("october 1 2023"),
      qEnd : new Date("december 31 2023"),
      qTotal : sh.getRange("Q4"),
      qRange : sh.getRange("N5:P1000")
    },
    quarter4 : {
      qStart : new Date("january 1 2024"),
      qEnd : new Date("march 31 2024"),
      qTotal : sh.getRange("U4"),
      qRange : sh.getRange("R5:T1000")
    },
    quarter5 : {
      qStart : new Date("april 1 2024"),
      qEnd : new Date("june 30 2024"),
      qTotal : sh.getRange("Y4"),
      qRange : sh.getRange("V5:X1000")
    },
    quarter6 : {
      qStart : new Date("july 1 2024"),
      qEnd : new Date("september 30 2024"),
      qTotal : sh.getRange("AC4"),
      qRange : sh.getRange("Z5:AB1000")
    },
    quarter7 : {
      qStart : new Date("october 1 2024"),
      qEnd : new Date("december 31 2024"),
      qTotal : sh.getRange("AG4"),
      qRange : sh.getRange("AD5:AF1000")
    },
    quarter8 : {
      qStart : new Date("january 1 2025"),
      qEnd : new Date("march 31 2025"),
      qTotal : sh.getRange("AK4"),
      qRange : sh.getRange("AH5:AJ1000")
    },
    quarter9 : {
      qStart : new Date("april 1 2025"),
      qEnd : new Date("june 30 2025"),
      qTotal : sh.getRange("AO4"),
      qRange : sh.getRange("AL5:AN1000")
    },
    quarter10 : {
      qStart : new Date("july 1 2025"),
      qEnd : new Date("september 30 2025"),
      qTotal : sh.getRange("AS4"),
      qRange : sh.getRange("AP5:AR1000")
    },
    quarter11 : {
      qStart : new Date("october 1 2025"),
      qEnd : new Date("december 31 2025"),
      qTotal : sh.getRange("AW4"),
      qRange : sh.getRange("AT5:AV1000")
    },
    quarter12 : {
      qStart : new Date("january 1 2026"),
      qEnd : new Date("march 31 2026"),
      qTotal : sh.getRange("BA4"),
      qRange : sh.getRange("AX5:AZ1000")
    }
  };
  return {quarter, client};
}

// UTILITY FUNCTIONS //

// @params: active sheet, quarters, active cell
// @return: quarter child object or null
// Takes in the active cell and returns which quarter was edited, or null
function getQuarter(sh, quarters, cell) {
  const quarterId = sh.getRange(2, cell.getColumn()).getValue();
  return quarters["quarter" + quarterId.substring(1)];
}

// @params: quarter object, increment value
// @return: none
// Updates the quarter's total with the given increment value
function updateQuarterTotal(quarter, incrementValue) {
  var totalCell = quarter.qTotal;
  var currentTotal = totalCell.getValue();
  totalCell.setValue(currentTotal + incrementValue);
}

// @params: input date, start date, end date
// @return: bool
// Takes in an input date, and returns whether or not it falls within the input quarter
function dateCompare(inD, qD1, qD2) {
  if (inD.valueOf() >= qD1.valueOf() && inD.valueOf() < qD2.valueOf()){
    return true;
  } 
  else{
    return false;
  }
}
