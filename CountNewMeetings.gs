// Google Apps Script for tracking new client meetings, using the "Date Added" parameter to check if the client is added within the quarter and has no previous meetings
// UPEI Research Services: SpringBoard requires metrics on the number of researcher, company, and industry meetings/events
// By: Alex O'Brien

function onEdit(e) {
  // define active sheet
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // if on "Researcher Meetings" sheet
  if(sh.getName() == "Engagements with Faculty"){
    
    // reference cells each containing new meeting tally
    const newQ1 = sh.getRange("I4");
    
    // define active cell
    var curCell = sh.getActiveCell();
    var curCol = curCell.getColumn();
    var curRow = curCell.getRow();

    // invisible meeting counter (Column A)
    var numMeetings = sh.getRange(curRow, 1).getValue();

    // store dates
    var addedDate = sh.getRange(curRow, 2).getValue();
    const q1Start = new Date("april 1 2023");

    // edited cell in Q1
    if(curCol >= 6 && curCol <= 8 && curRow >= 5 && addedDate.valueOf() > q1Start.valueOf() && numMeetings == 1){
      newQ1.setValue(newQ1.getValue()+1);}

    // format dates (using date picker changes text font for some reason)
    sh.getRange(5,2,sh.getMaxRows()).setFontFamily("Arial").setFontSize(10);
  }
}
