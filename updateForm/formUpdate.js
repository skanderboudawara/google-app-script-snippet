const CALL_ID = "";
const GOOGLE_FORM_ID = "";
const TITLE_OF_DROPDOWN = "TITLE OF DROPDOWN";
var form = FormApp.openById(GOOGLE_FORM_ID);

function debug() {
  console.log(form.getItems());
}
function createTimeDrivenTriggers() {
  // Trigger every day at 5 a.m.
  ScriptApp.newTrigger("updateForm")
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .create();
}

function updateForm() {
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
  var item = form.getItemById(id);
  item.asListItem().setChoiceValues(values);
}

function removeDuplicates(arr) {
  return arr.filter((item, index) => arr.indexOf(item) === index);
}

function getEventsByName() {
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
  };

  var events = Calendar.Events.list(CALL_ID, search_args).items;

  // Sort events per starting date
  events.sort(compareStartingDates);

  events_name = [];
  events.forEach(function (e) {
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
function compareStartingDates(a, b) {
  return new Date(a.start.dateTime).getTime() >
    new Date(b.start.dateTime).getTime()
    ? 1
    : -1;
}

function ISODateString(d) {
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
