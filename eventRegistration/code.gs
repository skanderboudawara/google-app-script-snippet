///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Main function
// Create the html page based on the file 'page.html'+ generate the call to the init function with parameters given in URL
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function testDoGet() {
  // Test doGet function
  var assert = require("assert");
  var e = {
    parameter: {
      calId: "calId",
      Text: "Text",
      Color: "Color",
      Days: "Days",
      ReplyTo: "ReplyTo",
      TimeZone: "TimeZone",
    },
  };
  assert.equal(doGet(e), "html" );
}
function doGet(e) {
  /*
  This function is called when the webapp is called with parameters
  This function will call the init function with the parameters given in the URL

  :param e: the parameters passed in the URL

  :return: the HTML page with the buttons
  */

  // these are the logs of the Session
  console.log(Session.getActiveUser().getEmail());
  console.log(Session.getActiveUser().getUserLoginId());
  // Build and display HTML page
  // These are the parameters passed in the URL
  var tmp = HtmlService.createTemplateFromFile("page");
  var html = tmp.evaluate().getContent();
  var numDays = 120;
  var color = "00205b";
  var replyTo = "email";
  var usedTimezone = "Europe/Paris";

  if (e.parameter.Color != undefined) {
    color = e.parameter.Color;
  }
  if (e.parameter.Days != undefined) {
    numDays = e.parameter.Days;
  }
  if (e.parameter.ReplyTo != undefined) {
    replyTo = e.parameter.ReplyTo;
  }
  if (e.parameter.TimeZone != undefined) {
    usedTimezone = e.parameter.TimeZone;
  }
  
  // Get the HTML code for the buttons
  return HtmlService.createTemplate(
    html +
      "<script>\n" +
      'init("' +
      e.parameter.calId +
      '","' +
      e.parameter.Text +
      '","' +
      color +
      '","' +
      numDays +
      '","' +
      replyTo +
      '","' +
      usedTimezone +
      '");\n</script> </body> </html>'
  )
    .evaluate()
    .setTitle("📅 Events");
}

function testIsBlank() {
  // Test isBlank function
  var assert = require("assert");
  assert.equal(isBlank(""), true);
  assert.equal(isBlank(" "), true);
  assert.equal(isBlank("  "), true);
  assert.equal(isBlank("a"), false);
  assert.equal(isBlank(" a"), false);
  assert.equal(isBlank("a "), false);
  assert.equal(isBlank(" a "), false);
  assert.equal(isBlank(" a b "), false);
}
function isBlank(str) {
  /*
  This function is used to check if a string is empty or contains only spaces

  :param str: the string to check

  :return: true if the string is empty or contains only spaces, false otherwise

  */
  return !str || /^\s*$/.test(str);
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Function initialisation
// Generate the HTML code for the buttons corresponding to the events matching the 'searchedText' in the title, in the 'calId'
// calendar. Set the 'color' background.
// Check first if the event session is full or not (if limit defined)
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function testInitialisation() {
  // Test initialisation function
  var assert = require("assert");
  var calId = "calId";
  var searchedText = "searchedText";
  var color = "color";
  var numberofDaysForSearch = 120;
  var usedTimezone = "Europe/Paris";
  var jsonData = {
    usedTimezone: usedTimezone,
    data: {},
  };
  assert.equal(
    initialisation(
      calId, searchedText, color, numberofDaysForSearch, usedTimezone 
    ), JSON.stringify(jsonData)
  );
}

function initialisation(
  calId,
  searchedText = null,
  color = "731b62",
  numberofDaysForSearch = 360,
  usedTimezone = "Europe/Paris"
) {
  /*
  This function is called when the webapp is called with parameters
  This function will call the init function with the parameters given in the URL

  :param calId: the calendar ID
  :param searchedText: the text to search in the title of the events
  :param color: the color of the buttons
  :param numberofDaysForSearch: the number of days to search for events
  :param usedTimezone: the timezone to use for the events

  :return: the HTML page with the buttons

  */

  // Get the calendar
  var cal = CalendarApp.getCalendarById(calId);

  // Get calendar events (webinars - next 3 months)
  // Get the current date
  var now = new Date();

  // Get the date in 3 months
  var searchPeriod = new Date(
    now.getTime() + numberofDaysForSearch * 24 * 60 * 60 * 1000
  );

  // Prepare search arguments
  search_args = {
    timeMin: ISODateString(now),
    timeMax: ISODateString(searchPeriod),
    singleEvents: true,
    showHiddenInvitations: true,
    maxResults: 2500,
    showDeleted: true,
  };

  // If a text is given, search for it in the title of the events
  if (isBlank(searchedText)) {
    search_args["q"] = searchedText;
  }

  // Get events or instances in case of recurring event
  // If the first event has a reccurence, we consider it's only instances of a recurring event
  var events = Calendar.Events.list(calId, search_args).items;

  // Sort events per starting date
  events.sort(compareStartingDates);

  // Build HTML string to be displayed
  jsonData = {
    usedTimezone: usedTimezone,
    data: {},
  };

  // Loop on events
  events.forEach(function (e) {
    // We first check if the event has a limit of participants and if this limit is reached
    // If yes, we don't display the button
    var isPossible = true;
    // Get the limit of participants (if any)
    // Get the description of the event
    var body = e.getDescription();
    var guests;
    var limit;
    var seatsAvailable = 0;

    if (body != "") {
      // There is a description
      // Search for the limit using a regular expression (if any)
      var regExp = new RegExp(/\[Limit: ?(\d+)\]/g);
      // Get the limit (if any) and the number of participants who accepted
      var res = regExp.exec(body);
      var limit;
      if (res != null) {
        // There is a limit!
        limit = res[1];
        // Check if the limit of participants who accepted is reached (if any limit)
        guests = e.attendees;

        if (guests != undefined) {
          // There are guests
          guests = guests.filter(function (g) {
            // We only keep guests who accepted or who didn't answer yet
            return (
              g.responseStatus == "accepted" ||
              g.responseStatus == "needsAction"
            );
          });

          if (guests.length >= limit) {
            // The limit is reached
            isPossible = false;
          } else {
            // The limit is not reached
            seatsAvailable = limit - guests.length;
          }
        } else {
          // There are no guests yet
          seatsAvailable = limit; // By default number of seats available is the limit, if no guests yet
        }
      }
    }

    // Depending on this result, color is requested color or grey, and button is active or not
    // Prepare HTML code for the button
    var finalColor;
    var btnActivity;
    var addMessage = "";

    if (isPossible == true) {
      // The limit is not reached
      // We can display the button with the requested color and the message "Session open" or "x seats remaining" (if limit)
      finalColor = color;
      btnActivity = "active";
      if (seatsAvailable > 0) {
        // There is a limit and there are seats available
        addMessage = " (" + seatsAvailable.toString() + " seat(s) remaining)";
      } else {
        // There is no limit or the limit is reached
        addMessage = " (Session open)";
      }
    } else {
      // The limit is reached or there is no limit and the event is full (no more seats available)
      finalColor = "808B96"; // grey
      btnActivity = "inactive";
      addMessage = " (Session full)";
    }
    try {
      // Prepare HTML code for the button (with the event title, the starting date and the ending date)
      event_data = {
        btnActivity: btnActivity,
        finalColor: finalColor,
        id: e.getId(),
        summary: e.summary,
        addMessage: addMessage,
        start_date: Utilities.formatDate(
          new Date(e.start.dateTime),
          usedTimezone,
          "MMMM dd, yyyy - HH:mm"
        ),
        end_date: Utilities.formatDate(
          new Date(e.end.dateTime),
          usedTimezone,
          "HH:mm"
        ),
      };
      
      // Add the event to the list of events to be displayed
      start_date = new Date(e.start.dateTime);
      if (
        (start_date >= new Date()) & // start date should be supperior than actual tie
        (e.status != "cancelled") // event should not be cancelled
      ) {
        jsonData["data"][e.summary + start_date] = event_data;
      }
    } catch (e) {
      Logger.log("skip");
    }
  });

  // Return the HTML code for the buttons to be displayed on the page (in the div with id 'events') and the JSON data for the calendar
  return JSON.stringify(jsonData);

}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Function registerUserToEvent
// Just add the user in the list of guests of the selected event
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// create unit test for registerUserToEvent
function testRegisterUserToEvent() {
  // Test registerUserToEvent function
  var assert = require("assert");
  var calId = "calId";
  var eventId = "eventId";
  var replyTo = "replyTo";
  var usedTimeZone = "usedTimeZone";
  assert.equal( registerUserToEvent(calId, eventId, replyTo, usedTimeZone), "OK_eventId" );
}
function registerUserToEvent(calId, eventId, replyTo, usedTimeZone) {
  /*
  This function is called when the user clicks on a button
  This function will add the user in the list of guests of the selected event

  :param calId: the calendar ID
  :param eventId: the event ID
  :param replyTo: the email address to reply to
  :param usedTimeZone: the timezone to use for the events

  :return: the HTML page with the buttons

  */

  // Used for Debug
  var message;

  // Used for Debug - Log the event in the Log sheet of the spreadsheet (if any)
  var logArray = [];
  logArray.push(new Date());
  logArray.push(calId);
  logArray.push(eventId);
  // End Debug

  try {
    // Get the event to update it with the new guest added (the user) and send the invitation to the user
    var event = Calendar.Events.get(calId, eventId);

    var guestEmail = Session.getActiveUser().getEmail(); // Get the email address of the user
    var aliases = GmailApp.getAliases(); // Get the aliases of the user

    // Guest added and confrmation message prepared

    var attendees = event.attendees;
    if (attendees) {
      // If there are already attendees, push the new attendee onto the list of attendees for the event
      attendees.push({ email: guestEmail });
      var resource = { attendees: attendees };
    } else {
      // If attendees[] doesn't exist, add it manually (there's probably a better way of doing this)
      var resource = { attendees: [{ email: guestEmail }] };
    }

    var args = { sendUpdates: "none" };
    // https://developers.google.com/calendar/api/v3/reference/events/patch for more info on args
    // args.sendUpdates = "none" will not send an email to the guest
    // args.sendUpdates = "all" will send an email to the guest
    // args.sendUpdates = "externalOnly" will send an email to the guest only if the guest is not in the same domain as the organizer
    try {
      // Update the event with the new guest added and send the invitation to the user
      // https://developers.google.com/calendar/api/v3/reference/events/patch for more info on Calendar.Events.patch
      var send_invitation = Calendar.Events.patch(
        resource, // the resource to patch
        calId, // the calendar ID
        eventId, // the event ID
        args // the arguments
      );
      Logger.log(send_invitation);
    } catch (err) {
      Logger.log(err);
      message = "KO_" + eventId; // Used for Debug
      return message;
    }

    // get recent updates
    var event = Calendar.Events.get(calId, eventId); // Get the event to update it with the new guest added (the user) and send the invitation to the user

    var eventURL =
      "https://www.google.com/calendar/render?action=VIEW&eid=" +
      Utilities.base64Encode(eventId + " " + guestEmail); // Create the URL to the event in the calendar (with the event ID and the email address of the user)
    if (eventURL.slice(-1) === "=") {
      eventURL = eventURL.slice(0, -1);
    } // remove trailing "="

    Logger.log(eventURL); // Used for Debug

    // For Debug
    logArray.push(guestEmail);
    logArray.push(event.summary);
    logArray.push(event.start.dateTime);
    logArray.push(event.end.dateTime);
    var logArrays = [];
    logArrays[0] = logArray;


    var emailBody = 
      "Dear participant, Your registration is confirmed. \
      We thank you for your interest.If you are a Google user, \
      we invite you to accept the event in your calendar directly. \
      If you are not a google user, please download the ICS file attached \
      to this email and then accept the invitation.&nbsp; \
      Would you have any question, feel free to answer to this email.";

    var guestName = guestEmail
      .split(".")[0]
      .toLowerCase()
      .replace(/^[a-z]/i, function (match) {
        return match.toUpperCase();
      });
    var eventTitle = event.summary;
    var eventStart = Utilities.formatDate(
      new Date(event.start.dateTime),
      usedTimeZone,
      "MMMM dd, yyyy HH:mm"
    );
    var eventEnd = Utilities.formatDate(
      new Date(event.end.dateTime),
      usedTimeZone,
      "HH:mm"
    );
    var eventTimespan =
      eventStart + " - " + eventEnd + " (timezone '" + usedTimeZone + "')";
    var eventDescription = event.description;
    var eventLocation = event.location;
    if (!eventLocation) {
      eventLocation = "Location not available. Please ask meeting organizer";
    }
    var htmlBody =
      "" +
      '<div style="font:small/1.5 Arial,Helvetica,sans-serif;direction:ltr;word-wrap:break-word;text-align:left;word-break:break-word;color:#3c4043!important;font-family:Roboto,sans-serif;font-style:normal;font-weight:400;font-size:14px;line-height:20px;letter-spacing:0.2px;">' +
      "<p>Dear " +
      guestName +
      ",<br><br>Your registration to this event is <strong>confirmed</strong>. Thank you for your interest.</p>" +
      "<p>Please click the below button to accept the event in your calendar directly." +
      "</div>" +
      '<div style="font:small/1.5 Arial,Helvetica,sans-serif;direction:ltr;word-wrap:break-word;text-align:left;word-break:break-word;color:#3c4043!important;font-family:Roboto,sans-serif;font-style:normal;font-weight:400;font-size:14px;line-height:20px;letter-spacing:0.2px;text-decoration:none;border:solid 1px #dadce0;border-radius:8px;padding:24px 32px;text-align:left;vertical-align:top;">' +
      "<div>" +
      '<div style="color:#222;font-size:140%">' +
      eventTitle +
      "</div>" +
      '<div style="margin-bottom:10px"><a style="color:#15c;cursor:pointer;text-decoration:none" href="' +
      eventURL +
      '" target="_blank">View on Google Calendar</a></div>' +
      "<table>" +
      "<tr>" +
      '<td style="color:#999;min-width:60px;vertical-align:top;">When</td>' +
      "<td>" +
      eventTimespan +
      "</td>" +
      "</tr>" +
      "<tr>" +
      '<td style="color:#999;vertical-align:top;">Where</td>' +
      "<td>" +
      eventLocation +
      "</td>" +
      "</tr>" +
      "<tr>" +
      '<td style="color:#999;vertical-align:top;">Who</td>' +
      "<td>Open calendar link to see guest list</td>" +
      "</tr>" +
      "</table>" +
      "<p></p>" +
      '<p><a style="background-color: #1a73e8; padding: 10px 20px; color: white; text-decoration:none;font-size:14px; font-family:Roboto,sans-serif;font-weight:700;border-radius:5px" href="' +
      eventURL +
      '" target="_blank">Open Calendar and Confirm Participation</a></p>' +
      "</div>" +
      "<p>&nbsp;</p>" +
      "<div>" +
      eventDescription +
      "</div>" +
      "</div>";

    var mail_options = {
      from: aliases[0],
      htmlBody: htmlBody,
      name: "Registration tool",
      noReply: false,
      replyTo: replyTo,
      subject: "Registration confirmed",
      //attachments: [icsFile]
    };
    GmailApp.sendEmail(
      guestEmail,
      "Registration confirmed",
      htmlBody,
      mail_options
    );
  } catch (err) {
    console.log(err);
    message = "KO_" + eventId;
    return message;
  } finally {
    console.log("ok-" + calId + "  " + eventId + "  " + replyTo);
    message = "OK_" + eventId;
    return message;
  }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Technical functions
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function include(filename) {
  /*
  This function is used to include the HTML file in the HTML page

  :param filename: the name of the HTML file to include

  :return: the HTML code of the file

  */
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// create unit test for ldapDate
function testLdapDate() {
  // Test ldapDate function
  var assert = require("assert");
  var d = new Date();
  assert.equal(ldapDate(d), Utilities.formatDate(d, "GMT+00:00", "yyyyMMdd") + "T" + Utilities.formatDate(d, "GMT+00:00", "HHmmss") + "Z" );
}
function ldapDate(d) {
  /*
  This function is used to format a date in the LDAP format

  :param d: the date to format

  :return: the date in the LDAP format

  */
  var formattedDate =
    Utilities.formatDate(d, "GMT+00:00", "yyyyMMdd") +
    "T" +
    Utilities.formatDate(d, "GMT+00:00", "HHmmss");
  return formattedDate + "Z";
}

// create unit test for rfc3339Date
function testRfc3339Date() {
  // Test rfc3339Date function
  var assert = require("assert");
  var d = new Date();
  assert.equal(rfc3339Date(d), Utilities.formatDate(d, "GMT+00:00", "yyyy-MM-dd") + "T" + Utilities.formatDate(d, "GMT+00:00", "HH:mm:ss") + "Z" );
}
function rfc3339Date(d) {
  /*
  This function is used to format a date in the RFC3339 format

  :param d: the date to format

  :return: the date in the RFC3339 format

  */
  var formattedDate =
    Utilities.formatDate(d, "GMT+00:00", "yyyy-MM-dd") +
    "T" +
    Utilities.formatDate(d, "GMT+00:00", "HH:mm:ss");
  return formattedDate + "Z";
}

function testMakeICS() {
  // Test makeICS function
  var assert = require("assert");
  var e = {
    guests: [ {guestName: "guestName", guestEmail: "guestEmail"} ],
    start: {dateTime: "dateTime"},
    end: {dateTime: "dateTime"},
    iCalUID: "iCalUID",
    id: "id",
    created: "created",
    description: "description",
    location: "location",
    summary: "summary",
  };
  assert.equal(makeICS(e), "icsFile" );
}
function makeICS(event) {
  /*
  This function is used to create the ICS file to be attached to the email sent to the user

  :param event: the event to create the ICS file for

  :return: the ICS file

  */
  var e = event;
  var attendees = [];
  var guests = e.guests;

  for (g in guests) {
    var guest = guests[g];
    var atendee = [
      "ATTENDEE;",
      "CUTYPE=INDIVIDUAL;",
      "ROLE=REQ-PARTICIPANT;",
      "PARTSTAT=NEEDS-ACTION;",
      "RSVP=TRUE;",
      "CN=" + guest.guestName + ";",
      "X-NUM-GUESTS=0:mailto:" + guest.guestEmail,
    ].join("");
    attendees.push(atendee);
  }

  var vcal = [
    "BEGIN:VCALENDAR",
    "PRODID:-//Google Inc//Google Calendar 70.9054//EN",
    "VERSION:2.0",
    "CALSCALE:GREGORIAN",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    "DTSTART:" + ldapDate(new Date(e.start.dateTime)),
    "DTEND:" + ldapDate(new Date(e.end.dateTime)),
    "DTSTAMP:" + ldapDate(new Date(Date.now())),
    "ORGANIZER;CN=" +
      CalendarApp.getCalendarById(e.iCalUID).getName() +
      ":mailto:" +
      e.iCalUID,
    "UID:" + e.id,
    attendees.join("\n"),
    "CREATED:" + ldapDate(new Date(e.created)),
    "DESCRIPTION:" + e.description,
    "LAST-MODIFIED:" + ldapDate(new Date(Date.now())), // although if I wasn't changing things as I issue this it'd be e.getLastUpdated()
    "LOCATION:" + e.location,
    "SEQUENCE:" + Date.now(), //this is a horrible hack, but it ensures that this change will overrule all other changes.
    "STATUS:CONFIRMED",
    "SUMMARY:" + e.summary,
    "TRANSP:OPAQUE",
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\n");

  var icsFile = Utilities.newBlob(vcal, "text/calendar", "invite.ics");
  return icsFile;
}

function testCompareStartingDates() {
  // Test compareStartingDates function
  var assert = require("assert");
  var a = {start: {dateTime: "dateTime"}};
  var b = {start: {dateTime: "dateTime"}};
  assert.equal(compareStartingDates(a, b), 1 );
  assert.equal(compareStartingDates(b, a), -1 );
  assert.equal(compareStartingDates(a, a), 0 );
}
function compareStartingDates(a, b) {
  /*
  This function is used to compare the starting dates of two events

  :param a: the first event
  :param b: the second event

  :return: 1 if the starting date of the first event is after the starting date of the second event, -1 otherwise

  */
  return new Date(a.start.dateTime).getTime() >
    new Date(b.start.dateTime).getTime()
    ? 1
    : -1;
}

function testISODateString() {
  // Test ISODateString function
  var assert = require("assert");
  var d = new Date();
  assert.equal(ISODateString(d), Utilities.formatDate(d, "GMT+00:00", "yyyy-MM-dd") + "T" + Utilities.formatDate(d, "GMT+00:00", "HH:mm:ss") + "Z" );
  assert.equal(ISODateString(d), rfc3339Date(d) );
  assert.equal(ISODateString(d), ISODate(d) + "T" + Utilities.formatDate(d, "GMT+00:00", "HH:mm:ss") + "Z" );
}
function ISODateString(d) {
  /*
  This function is used to format a date in the ISO format

  :param d: the date to format

  :return: the date in the ISO format

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

function testISODate() {
  // Test ISODate function
  var assert = require("assert");
  var d = new Date();
  assert.equal(ISODate(d), Utilities.formatDate(d, "GMT+00:00", "yyyy-MM-dd") );
  assert.equal(ISODate(d), ISODateString(d).split("T")[0] );
  assert.equal(ISODate(d), rfc3339Date(d).split("T")[0] );
}
function ISODate(d) {
  /*
  This function is used to format a date in the ISO format

  :param d: the date to format

  :return: the date in the ISO format

  */
  function pad(n) {
    return n < 10 ? "0" + n : n;
  }
  return (
    d.getUTCFullYear() +
    "-" +
    pad(d.getUTCMonth() + 1) +
    "-" +
    pad(d.getUTCDate())
  );
}

function testCompareDates() {
  // Test compareDates function
  var assert = require("assert");
  var a = "a";
  var b = "b";
  assert.equal(compareDates(a, b), -1 );
  var a = "12/31/2020";
  var b = "01/01/2021";
  assert.equal(compareDates(a, b), -1 );
  var a = "01/01/2021";
  var b = "12/31/2020";
  assert.equal(compareDates(a, b), 1 );
}
function compareDates(a, b) {
  /*
  This function is used to compare two dates

  :param a: the first date
  :param b: the second date

  :return: 1 if the first date is after the second date, -1 otherwise

  */
  return new Date(a).getTime() > new Date(b).getTime() ? 1 : -1;
}

function testContainsObject() {
  // Test containsObject function
  var assert = require("assert");
  var obj = "obj";
  var list = ["obj"];
  assert.equal(containsObject(obj, list), true );

  var obj = "obj";
  var list = ["obj1"];
  assert.equal(containsObject(obj, list), false );
}
function containsObject(obj, list) {
  /*
  This function is used to check if an object is in a list

  :param obj: the object to check
  :param list: the list to check

  :return: true if the object is in the list, false otherwise

  */
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i] === obj) {
      return true;
    }
  }

  return false;
}

function testSheetLog() {
  // Test sheetLog function
  var assert = require("assert");
  var type = "type";
  var message = "message";
  assert.equal(sheetLog(type, message), undefined );
}
function sheetLog(type, message) {
  /*
  This function is used to log an event in the Log sheet of the spreadsheet

  :param type: the type of the event
  :param message: the message of the event

  :return: None

  */
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
  var lastRow = sheet.getLastRow();

  var now = new Date();
  var logArray = [];
  logArray.push(now);
  logArray.push(type);
  logArray.push(message);

  var logArrays = [];
  logArrays[0] = logArray;

  var range = sheet.getRange(lastRow + 1, 1, 1, logArrays[0].length);
  range.setValues(logArrays);
}