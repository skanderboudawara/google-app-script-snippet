DAY_MILLIS = 24 * 60 * 60 * 1000;

var logSheet;
var rowLog = 1;
var absencesSheet;
var rowAbsence = 1;

var begin;
var end;

var allCalendars = [];

function synchronizeOutOfOfficeFromCalendar() {
  /*
  // Describe this function
  // This function is used to synchronize the calendar of the user with the spreadsheet
  // The spreadsheet must have 3 sheets :
  // - calendar : the calendar of the user
  // - oOo : the out of office of the user
  // - log : the log of the process

  :param: none

  :return: none
  */

  // Only use english

  spreadSheet = SpreadsheetApp.getActiveSpreadsheet(); // Get the spreadsheet

  logSheet = spreadSheet.getSheetByName("log"); // Get the log sheet
  logSheet.clear(); // Clear the log sheet
  sheetLog("info", "Process started"); // Log the start of the process

  absencesSheet = spreadSheet.getSheetByName("oOo"); // Get the out of office sheet
  absencesSheet.clear(); // Clear the out of office sheet
  calSheet = spreadSheet.getSheetByName("calendar"); // Get the calendar sheet

  launcherAllCalendars = CalendarApp.getAllCalendars(); // Get all the calendars of the user
  for (var a = 0; a < launcherAllCalendars.length; a++) {
    // For each calendar of the user add it to the list of calendars of the user (allCalendars) and to the list of calendars to process (allCalendars)
    allCalendars.push(launcherAllCalendars[a].getId());
  }

  // Get the begin and the end of the period to process (the period is defined by the first line of the calendar sheet)
  var plage = calSheet
    .getRange("4:4")
    .getValues()[0]
    .filter(function (obj) {
      return obj instanceof Date;
    })
    .sort(function (a, b) {
      return a.getTime() - b.getTime();
    });

  // If the period is not defined, log an error and stop the process
  begin = plage[0];
  end = plage[plage.length - 1];

  // If the period is not defined, log an error and stop the process
  var ids = calSheet
    .getRange("B:B")
    .getValues()
    .filter(function (obj) {
      if (!(typeof obj[0] === "string")) {
        return false;
      }
      if (obj[0].indexOf("@") === -1) {
        return false;
      }
      return true;
    })
    .sort();

  for (var a = 0; a < ids.length; a++) {
    // For each calendar of the user add it to the list of calendars of the user (allCalendars) and to the list of calendars to process (allCalendars)  
    lookUpOneCalendar(ids[a]);
  }

  sheetLog("info", "Process ended"); // Log the end of the process
}

function lookUpOneCalendar(id) {
  /*
  This function is used to synchronize the calendar of the user with the spreadsheet
  
  :param: id: the id of the calendar to process

  :return: none
  */
  sheetLog("info", "Start of calendar process " + id); // Log the start of the process
  var cal; // The calendar to process
  var unsubscribe = false; // If the calendar is not subscribed, unsubscribe it at the end of the process

  if (!containsObject(id, allCalendars)) {
    sheetLog("info", "calendar " + id + " not found");
    try {
      cal = CalendarApp.subscribeToCalendar(id); // Subscribe the calendar
      unsubscribe = true; // Set the flag to unsubscribe the calendar at the end of the process
    } catch (err) {
      sheetLog("alert", "calendar " + id + " Error"); // Log an error
      sheetLog("info", "End of calendar process " + id); // Log the end of the process
      return;
    }
  } else {
    sheetLog("info", "calendar " + id + " subscribed"); // Log the subscription of the calendar
    cal = CalendarApp.getCalendarById(id); // Get the calendar
  }

  var events = cal.getEvents(begin, end, { search: "Absent du bureau" }); // Get the events of the calendar
  var events = events.concat(
    cal.getEvents(begin, end, { search: "Out of Office" })
  ); // Get the events of the calendar

  var event;
  var dates;
  var fraction;
  var duration;

  for (var i = 0; i < events.length; i++) {
    event = events[i]; // Get the event

    if (
      event.getTitle() == "Absent du bureau" ||
      event.getTitle() == "Out of office"
    ) {
      dates = createDateSpan(
        event.getStartTime(),
        new Date(event.getEndTime().getTime() - 1)
      ); // Get the dates of the event

      for (var j = 0; j < dates.length; j++) {
        absencesSheet.getRange(rowAbsence, 1).setValue(cal.getId()); // Set the calendar id
        absencesSheet.getRange(rowAbsence, 2).setValue(dateTrunc(dates[j])); // Set the date
        absencesSheet
          .getRange(rowAbsence, 3)
          .setFormulaR1C1(
            '=CONCAT(INDIRECT("R[0]C[-2]";FALSE);INDIRECT("R[0]C[-1]";FALSE))'
          ); // Set the title

        if (
          event.getTitle() != "Télétravail" &&
          event.getTitle() != "Home office"
        ) {
          fraction = 1; // Set the fraction to 1
          if (
            j === 0 &&
            event.getStartTime().getHours() * 60 +
              event.getStartTime().getMinutes() >
              0
          ) {
            duration =
              event.getEndTime().getHours() * 60 +
              event.getEndTime().getMinutes() -
              (event.getStartTime().getHours() * 60 +
                event.getStartTime().getMinutes());

            if (duration > 4 * 60) {
              fraction = 1;
            } else if (duration > 2 * 60) {
              fraction = 0.5;
            } else {
              break;
            }
          } 
          if (
            j === dates.length - 1 &&
            event.getEndTime().getHours() * 60 +
              event.getEndTime().getMinutes() >
              0
          ) {
            duration =
              event.getEndTime().getHours() * 60 +
              event.getEndTime().getMinutes() -
              (event.getStartTime().getHours() * 60 +
                event.getStartTime().getMinutes());

            if (duration > 4 * 60) {
              fraction = 1;
            } else if (duration > 3 * 60) {
              fraction = 0.5;
            } else break;
          }
          absencesSheet.getRange(rowAbsence, 4).setValue(fraction); // Set the fraction
        }
        rowAbsence++; // Increment the row of the out of office sheet
      }
    }
  }

  if (unsubscribe) {
    cal.unsubscribeFromCalendar();  // Unsubscribe the calendar
  }
  sheetLog("info", "End of calendar process " + id);  // Log the end of the process
}

// creaate a unit test for createDateSpan
function testCreateDateSpan() {
  var startDate = new Date(2019, 0, 1);
  var endDate = new Date(2019, 0, 3);
  var dates = createDateSpan(startDate, endDate);
  Logger.log(dates);
  var assert = require("assert");
  assert.equals(createDateSpan(startDate, endDate), [ startDate, endDate ]);
}
function createDateSpan(startDate, endDate) {
  /*
  This function is used to create a list of dates between two dates

  :param: startDate: the start date
  :param: endDate: the end date

  :return: dates: the list of dates between the two dates
  */
  if (startDate.getTime() > endDate.getTime()) {
    throw Error("startDate > endDate");
  }

  var dates = [];

  var curDate = new Date(startDate.getTime());
  while (!dateCompare(curDate, endDate)) {
    dates.push(curDate);
    curDate = new Date(curDate.getTime() + DAY_MILLIS);
  }
  dates.push(endDate);
  return dates;
}

function testDateCompare() {
  var date1 = new Date(2019, 0, 1);
  var date2 = new Date(2019, 0, 1);
  var assert = require("assert");
  assert.equals(dateCompare(date1, date2), true);
  var date3 = new Date(2019, 0, 2);
  assert.equals(dateCompare(date1, date3), false);
}
function dateCompare(a, b) {
  /*
  This function is used to compare two dates

  :param: a: the first date
  :param: b: the second date

  :return: true if the two dates are equals, false otherwise
  */
  return (
    a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate()
  );
}

function testDateTrunc() {
  var date1 = new Date(2019, 0, 1, 12, 0, 0);
  var date2 = new Date(2019, 0, 1, 0, 0, 0);
  var assert = require("assert");
  assert.equals(dateTrunc(date1), date2);
}
function dateTrunc(d) {
  /*
  This function is used to truncate a date

  :param: d: the date to truncate

  :return: the truncated date
  */
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function testSheetLog() {
  var assert = require("assert");
  assert.equals(sheetLog("info", "test", "T"), undefined);
}
function sheetLog(level, msg, cptl) {
  /*
  This function is used to log a message in the log sheet

  :param: level: the level of the message
  :param: msg: the message to log
  :param: cptl: the capital letter of the level

  :return: none
  */
  logSheet
    .getRange(rowLog, 1)
    .setValue(
      Utilities.formatDate(
        new Date(),
        spreadSheet.getSpreadsheetTimeZone(),
        "yyyy/MM/dd HH:mm:ss.SSS"
      )
    );
  if (level != undefined) {
    logSheet.getRange(rowLog, 2).setValue(level);
  }
  if (msg != undefined) {
    logSheet.getRange(rowLog, 3).setValue(msg);
  }
  if (cptl != undefined) {
    logSheet.getRange(rowLog, 4).setValue(cptl);
  }
  rowLog++;
}

function testContainsObject() {
  var assert = require("assert");
  assert.equals(containsObject("test", ["test", "test2"]), true);
  assert.equals(containsObject("test", ["test2"]), false);
}
function containsObject(obj, list) {
  /*
  This function is used to check if an object is in a list

  :param: obj: the object to check
  :param: list: the list to check

  :return: true if the object is in the list, false otherwise
  */
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i].toString() === obj.toString()) {
      return true;
    }
  }

  return false;
}

function testOnOpen() {
  var assert = require("assert");
  assert.equals(onOpen(), undefined);
}
function onOpen() {
  /*
  This function is used to add a menu to the spreadsheet

  :param: none

  :return: none
  */
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    {
      name: "Synchro Calendar",
      functionName: "synchronizeOutOfOfficeFromCalendar",
    },
  ];
  sheet.addMenu("Fonctions", menuEntries);
}
