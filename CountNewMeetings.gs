// Google Apps Script for tracking new client meetings, using the "Date Added" parameter to check if the client is added within the quarter and has no previous meetings
// UPEI Research Services: SpringBoard requires metrics on the number of researcher, company, and industry meetings/events
// By: Alex O'Brien

function onEdit(e) {

  // DEFINITIONS: Could be a waste to save everything on each edit, but I feel the excess computing is arbitrary //

  // define active sheet
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // define current cell coordinates
  var curCell = sh.getCurrentCell();
  var curRow = curCell.getRow();
  var curCol = curCell.getColumn();
  var cellValue = curCell.getValue();

  // reference each cell containing new meeting tally
  const newQ1 = sh.getRange("I4");
  const newQ2 = sh.getRange("M4");
  const newQ3 = sh.getRange("Q4");
  const newQ4 = sh.getRange("U4");
  const newQ5 = sh.getRange("Y4");
  const newQ6 = sh.getRange("AC4");
  const newQ7 = sh.getRange("AG4");
  const newQ8 = sh.getRange("AK4");
  const newQ9 = sh.getRange("AO4");
  const newQ10 = sh.getRange("AS4");
  const newQ11 = sh.getRange("AW4");
  const newQ12 = sh.getRange("BA4");
  
  // invisible meeting counter (Column A)
  var numMeetings = sh.getRange(curRow, 1).getValue();

  // store dates
  var addedDate = sh.getRange(curRow, 2).getValue();
  const q1Start = new Date("april 1 2023");
  const q2Start = new Date("july 1 2023");
  const q3Start = new Date("october 1 2023");
  const q4Start = new Date("january 1 2024");
  const q5Start = new Date("april 1 2024");
  const q6Start = new Date("july 1 2024");
  const q7Start = new Date("october 1 2024");
  const q8Start = new Date("january 1 2025");
  const q9Start = new Date("april 1 2025");
  const q10Start = new Date("july 1 2025");
  const q11Start = new Date("october 1 2025");
  const q12Start = new Date("january 1 2026");

  // SECTION 1: Info Section is Sheet Dependent //

  // if on any sheet that uses "Added By" column
  if(sh.getName() == "Spin Off or Startup from the Institution w/o IP Assignment" || sh.getName() == "Total Companies Engaged" || sh.getName() == "Supporting Start-ups/Scale-ups" ||
    sh.getName() == "Ecosystem Partnerships and Projects" || sh.getName() == "Committees & Boards Supported by Members" || sh.getName() == "Network members Collaboration and Support" || sh.getName() == "Joint Network Member Events" || sh.getName() == "Workshops with Industry" || sh.getName() == "Inter-Member Collaborations"){

    // if editing company info
    if(curRow >= 5 && curCol >= 2 && curCol <= 4){
      if(sh.getRange(curRow, 4).getValue() != ""){sh.getRange(curRow, 3).setValue(Session.getActiveUser().getEmail());}
      else{sh.getRange(curRow, 3).setValue("");}
    }
  }

  // if on faculty engagements sheet
  if(sh.getName() == "Total Engagement with Faculty/Staff/Students"){
    // if editing researcher info
    if(curRow >= 5 && curCol >= 2 && curCol <= 4){
      if(sh.getRange(curRow, 4).getValue() != ""){
        sh.getRange(curRow, 3).setValue("Faculty");
      }
    }
  }

  // SECTION 2: Data Section is Sheet Independent //

  // edited cell in Q1
  if(curCol >= 6 && curCol <= 8 && curRow >= 5 && addedDate.valueOf() >= q1Start.valueOf() && addedDate.valueOf() < q2Start.valueOf()){
    if(cellValue == "" && numMeetings == 0){
      if(newQ1.getValue() >= 1){newQ1.setValue(newQ1.getValue()-1);}
    }
    else if(cellValue != "" && numMeetings == 1){newQ1.setValue(newQ1.getValue()+1);}
  }

  // edited cell in Q2
  if(curCol >= 10 && curCol <= 12 && curRow >= 5 && addedDate.valueOf() >= q2Start.valueOf() && addedDate.valueOf() < q3Start.valueOf()){
    if(cellValue == "" && numMeetings == 0){
      if(newQ2.getValue() >= 1){newQ2.setValue(newQ2.getValue()-1);}
    }
    else if(cellValue != "" && numMeetings == 1){newQ2.setValue(newQ2.getValue()+1);}
  }

  // edited cell in Q3
  if(curCol >= 14 && curCol <= 16 && curRow >= 5 && addedDate.valueOf() >= q3Start.valueOf() && addedDate.valueOf() < q4Start.valueOf()){
    if(cellValue == "" && numMeetings == 0){
      if(newQ3.getValue() >= 1){newQ3.setValue(newQ3.getValue()-1);}
    }
    else if(cellValue != "" && numMeetings == 1){newQ3.setValue(newQ3.getValue()+1);}
  }

  // edited cell in Q4
  if(curCol >= 18 && curCol <= 20 && curRow >= 5 && addedDate.valueOf() >= q4Start.valueOf() && addedDate.valueOf() < q5Start.valueOf()){
    if(cellValue == "" && numMeetings == 0){
      if(newQ4.getValue() >= 1){newQ4.setValue(newQ4.getValue()-1);}
    }
    else if(cellValue != "" && numMeetings == 1){newQ4.setValue(newQ4.getValue()+1);}
  }

  // edited cell in Q5
  if(curCol >= 22 && curCol <= 24 && curRow >= 5 && addedDate.valueOf() >= q5Start.valueOf() && addedDate.valueOf() < q6Start.valueOf()){
    if(cellValue == "" && numMeetings == 0){
      if(newQ5.getValue() >= 1){newQ5.setValue(newQ5.getValue()-1);}
    }
    else if(cellValue != "" && numMeetings == 1){newQ5.setValue(newQ5.getValue()+1);}
  }

  // edited cell in Q6
  if(curCol >= 26 && curCol <= 28 && curRow >= 5 && addedDate.valueOf() >= q6Start.valueOf() && addedDate.valueOf() < q7Start.valueOf()){
    if(cellValue == "" && numMeetings == 0){
      if(newQ6.getValue() >= 1){newQ6.setValue(newQ6.getValue()-1);}
    }
    else if(cellValue != "" && numMeetings == 1){newQ6.setValue(newQ6.getValue()+1);}
  }

  // edited cell in Q7
  if(curCol >= 30 && curCol <= 32 && curRow >= 5 && addedDate.valueOf() >= q7Start.valueOf() && addedDate.valueOf() < q8Start.valueOf()){
    if(cellValue == "" && numMeetings == 0){
      if(newQ7.getValue() >= 1){newQ7.setValue(newQ7.getValue()-1);}
    }
    else if(cellValue != "" && numMeetings == 1){newQ7.setValue(newQ7.getValue()+1);}
  }

  // edited cell in Q8
  if(curCol >= 34 && curCol <= 36 && curRow >= 5 && addedDate.valueOf() >= q8Start.valueOf() && addedDate.valueOf() < q9Start.valueOf()){
    if(cellValue == "" && numMeetings == 0){
      if(newQ8.getValue() >= 1){newQ8.setValue(newQ8.getValue()-1);}
    }
    else if(cellValue != "" && numMeetings == 1){newQ8.setValue(newQ8.getValue()+1);}
  }

  // edited cell in Q9
  if(curCol >= 38 && curCol <= 40 && curRow >= 5 && addedDate.valueOf() >= q9Start.valueOf() && addedDate.valueOf() < q10Start.valueOf()){
    if(cellValue == "" && numMeetings == 0){
      if(newQ9.getValue() >= 1){newQ9.setValue(newQ9.getValue()-1);}
    }
    else if(cellValue != "" && numMeetings == 1){newQ9.setValue(newQ9.getValue()+1);}
  }

  // edited cell in Q10
  if(curCol >= 42 && curCol <= 44 && curRow >= 5 && addedDate.valueOf() >= q10Start.valueOf() && addedDate.valueOf() < q11Start.valueOf()){
    if(cellValue == "" && numMeetings == 0){
      if(newQ10.getValue() >= 1){newQ10.setValue(newQ10.getValue()-1);}
    }
    else if(cellValue != "" && numMeetings == 1){newQ10.setValue(newQ10.getValue()+1);}
  }

  // edited cell in Q11
  if(curCol >= 46 && curCol <= 48 && curRow >= 5 && addedDate.valueOf() >= q11Start.valueOf() && addedDate.valueOf() < q12Start.valueOf()){
    if(cellValue == "" && numMeetings == 0){
      if(newQ11.getValue() >= 1){newQ11.setValue(newQ11.getValue()-1);}
    }
    else if(cellValue != "" && numMeetings == 1){newQ11.setValue(newQ11.getValue()+1);}
  }

  // edited cell in Q12
  if(curCol >= 50 && curCol <= 52 && curRow >= 5 && addedDate.valueOf() >= q12Start.valueOf()){
    if(cellValue == "" && numMeetings == 0){
      if(newQ12.getValue() >= 1){newQ12.setValue(newQ12.getValue()-1);}
    }
    else if(cellValue != "" && numMeetings == 1){newQ12.setValue(newQ12.getValue()+1);}
  }

  // format dates (using date picker changes text font for some reason)
  sh.getRange(5,2,sh.getMaxRows()).setFontFamily("Arial").setFontSize(10);

  return;

}
