const CALL_ID = "";
const GOOGLE_FORM_ID = "";
const TITLE_OF_DROPDOWN = "TITLE OF DROPDOWN";
var form = FormApp.openById(GOOGLE_FORM_ID);

function debug() {
  /*
  This is a function to debug the code

  :param: None

  :return: None
  */
  console.log(form.getItems());
}
function createTimeDrivenTriggers() {
  /*
  This function creates a trigger to update the form every day at 5 a.m.

  :param: None

  :return: None

  */
  // Trigger every day at 5 a.m.
  ScriptApp.newTrigger("updateForm")
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .create();
}

function updateForm() {
  /*
  This function updates the form with the events of the next 30 days

  :param: None

  :return: None

  */
  events_name = getEventsByName();
  var items = form.getItems();
  var titles = items.map(function (item) {
    return item.getTitle();
  });
  var pos = titles.indexOf(TITLE_OF_DROPDOWN);
  var item = items[pos];
  id_item = item.getId();

  updateDropdown(id_item, events_name);
}

function updateDropdown(id, values) {
  /*
  This function updates the dropdown with the events of the next 30 days

  :param: id: id of the dropdown
  :param: values: values to update the dropdown

  :return: None

  */
  var item = form.getItemById(id);
  item.asListItem().setChoiceValues(values);
}

// create unit testfor removeDuplicates
function testRemoveDuplicates() {
  var arr = ["a", "b", "c", "a", "b", "c", "d"];
  var arr2 = ["a", "b", "c", "d"];
  var arr3 = removeDuplicates(arr);
  Logger.log(arr3);
  Logger.log(arr2);
  var assert = require("assert");
  assert.deepEqual(arr2, arr3);
} 
function removeDuplicates(arr) {
  /*
  This function removes duplicates from an array

  :param: arr: array to remove duplicates

  :return: array without duplicates

  */
  return arr.filter((item, index) => arr.indexOf(item) === index);
}

function getEventsByName() {
  /*
  This function gets the events of the next 30 days

  :param: None

  :return: array with the events of the next 30 days

  */
  var now = new Date();
  var searchPeriodMin = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  var searchPeriodMax = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

  search_args = {
    timeMin: ISODateString(searchPeriodMin),
    timeMax: ISODateString(searchPeriodMax),
    singleEvents: true,
    showHiddenInvitations: true,
    maxResults: 2500,
    showDeleted: true,
  }; // search arguments

  // link for calendar api: https://developers.google.com/calendar/v3/reference/events/list
  var events = Calendar.Events.list(CALL_ID, search_args).items; // get events

  // Sort events per starting date
  events.sort(compareStartingDates);

  events_name = [];
  events.forEach(function (e) {
    /*
    This forEach will push the events that are not cancelled and that are in the next 30 days

    :param: e: event

    :return: None

    */
    start_date = new Date(e.start.dateTime);
    if (
      (start_date >= searchPeriodMin) & // start date should be supperior than actual tie
      (e.status != "cancelled") // event should not be cancelled
    ) {
      events_name.push(e.summary);
    }
  });

  events_name = removeDuplicates(events_name);
  return events_name;
}

function testCompareStartingDates() {
  var assert = require("assert");
  var a = "2020-10-10T10:00:00Z";
  var b = "2020-10-10T11:00:00Z";
  var c = "2020-10-10T10:00:00Z";
  var d = "2020-10-10T09:00:00Z";
  assert.equal(compareStartingDates(a, b), -1);
  assert.equal(compareStartingDates(a, c), 0);
  assert.equal(compareStartingDates(a, d), 1);
}
function compareStartingDates(a, b) {
  /*
  This function compares the starting dates of two events

  :param: a: event
  :param: b: event

  :return: 1 if a is greater than b, -1 otherwise

  */
  return new Date(a.start.dateTime).getTime() >
    new Date(b.start.dateTime).getTime()
    ? 1
    : -1;
}

function testISODateString() {
  var assert = require("assert");
  var d = new Date("2020-10-10T10:00:00Z");
  var d2 = new Date("2020-10-10T11:00:00Z");
  assert.equal(ISODateString(d), "2020-10-10T10:00:00Z");
  assert.equal(ISODateString(d2), "2020-10-10T11:00:00Z");
}
function ISODateString(d) {
  /*
  This function converts a date to ISO format

  :param: d: date

  :return: date in ISO format

  */
  function pad(n) {
    return n < 10 ? "0" + n : n;
  }
  return (
    d.getUTCFullYear() +
    "-" +
    pad(d.getUTCMonth() + 1) +
    "-" +
    pad(d.getUTCDate()) +
    "T" +
    pad(d.getUTCHours()) +
    ":" +
    pad(d.getUTCMinutes()) +
    ":" +
    pad(d.getUTCSeconds()) +
    "Z"
  );
}
