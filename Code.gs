// to do:
// - auto-update when final assignment column is edited -- can use google sheets functions for that? idk how those work tbh
// - algo to assign RAs


function onOpen() {
  // add menu to spreadsheet to use active sheet function and ease of access
  // should only be run on individual month sheets and in chronological order
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("RA Menu").addItem("Calculate Points/Weekends", "countPts").addToUi();
}

// helper function to make points and weekends rows cumulative for each month's sheet
// gets name in string format of the previous month from the current sheet
function getPrevMonth(currMonth) {
  var months = ["August", "September", "October", "November", "December", "January", "February", "March", "April", "May"];
  var currentIndex = months.indexOf(currMonth);
  
  if (currentIndex === -1) {
    return null; // invalid month name
  }
  
  var prevIndex = currentIndex === 0 ? months.length - 1 : currentIndex - 1;
  return months[prevIndex];
}

// helper function to make points and weekends rows cumulative for each month's sheet
// gets index of cumulative points row, array indexed at 1 to work with spreadsheetapp datarange
// must include the text "cumulative pts"
function getLabelsRows(sheetData) {
  for(var i = 14; i < sheetData.length; i++) { // we know there will be at least 13 days in each duty month
    var row = sheetData[i];
    if(row[2].trim() === "cumulative pts") {
      return i+1;
    }
  }
  return -1;
}

// helper function to get weighted points and isWeekend for a duty night
// regular weeknight = 1 pt; thurs = 1.25
// weekend (including study days and preceding sunday before long weekend; marked by *) = 1.5
// long holiday (thanksgiving, spring break, easter, halloween; marked by **) = 2
function getPtsForDay(day) {
  var metricsArr = [1, 0]; // [pts, weekends]
  if(day.indexOf("**") > -1) {
    metricsArr[0] = 2;
    metricsArr[1] += 1;
  } else if(day.indexOf("*") > -1 || day === "Friday" || day === "Saturday") {
    metricsArr[0] = 1.5;
    metricsArr[1] += 1;
  } else if(day === "Thursday") {
    metricsArr[0] = 1.25;
  }
  return metricsArr;
}

// main function to loop through month availability sheet, count up points/weekends for each ra, and update stats of current sheet
function countPts() {
  var ss = SpreadsheetApp.getActiveSheet();
  var currSheetName = ss.getName().trim();

  if(currSheetName !== "Summary by NIGHT-Final") {
    var dataRange = ss.getDataRange().getValues();

    var pointsByRA = {}
    var weekendsByRA = {}
    row = 1 // will store index of cumulative points row, zero-indexed
    
    // initialize points and weekends dicts if not first month so stats will be cumulative
    var prevSheetName = currSheetName === "August" ? null : getPrevMonth(currSheetName);
    var prevSheet = prevSheetName ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(prevSheetName) : null;
    if(prevSheet) {
      var prevDataRange = prevSheet.getDataRange().getValues();
      var labelsRow = getLabelsRows(prevDataRange);
      if(labelsRow > -1) {
        for(var col = 3; col < prevSheet.getLastColumn(); col++) {
          var raName = prevDataRange[0][col].trim();
          var prevPoints = prevSheet.getRange(labelsRow, col+1).getValue();
          var prevWeekends = prevSheet.getRange(labelsRow+1, col+1).getValue();
          if(prevPoints) {
            pointsByRA[raName] = prevPoints;
          } else {
            Logger.log("Error: previous month's points for RA " + raName + " is empty. please run points/weekends counter in chronological order");
          }
          if(prevWeekends) {
            weekendsByRA[raName] = prevWeekends;
          } else {
            Logger.log("Error: previous month's weekends for RA " + raName + " is empty. please run points/weekends counter in chronological order");
          }
        }
      } else {
        Logger.log("Error: previous sheet label rows not found. text in FINAL column must be 'cumulative pts'");
        return;
      }
    }

    // count up points/weekends of current month 
    for(var i = 1; i < dataRange.length-10; i++) {
      var day = dataRange[i][1].toString().trim();
      var finalAssignment = dataRange[i][2].trim();
      if(day != "" && finalAssignment != "") {
        if(!pointsByRA[finalAssignment]) {
          pointsByRA[finalAssignment] = 0;
        }
        if(!weekendsByRA[finalAssignment]) {
          weekendsByRA[finalAssignment] = 0;
        }
        var ptsAndWeekends = getPtsForDay(day);
        var pts = ptsAndWeekends[0];
        var weekends = ptsAndWeekends[1];
        pointsByRA[finalAssignment] += pts;
        weekendsByRA[finalAssignment] += weekends;
        row++;
      }
    }

    //  fill in points and weekends rows on current sheet
    for(var j = 3; j < ss.getLastColumn(); j++) {
      var raName = dataRange[0][j].trim();
      if(pointsByRA[raName]) {
        ss.getRange(row+1, j+1).setValue(pointsByRA[raName]);
      } else {
        ss.getRange(row+1, j+1).setValue(0);
      }
      if(weekendsByRA[raName]) {
        ss.getRange(row+2, j+1).setValue(weekendsByRA[raName]);
      } else {
        ss.getRange(row+2, j+1).setValue(0);
      }
    }
  } else {
    Logger.log("Error: must run points/weekend counter on month availability sheet")
  }
}
