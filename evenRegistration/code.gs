/**************************************************************/
/* EVENT REGISTRATION                                         */
/* This script generates buttons corresponding to calendar    */
/* events allowing users to register automatically to sessions*/
/* See full description here:                                 */
/* https://sites.google.com/airbus.com/demoeventreg/home      */
/*                                                            */
/*                                                            */
/* Versions:                                                  */
/* V16 - Feb 5th 2021..                                       */
/* V17 - Apr 13th 2021: Use of Advanced calendar API to send  */
/*         an email with event details                        */
/* V18 - May 27th 2021: Change of the mailing mechanism.      */
/*         Google users are automatically invited             */
/*         Non Google users receive an ics file for the event */
/*         Adding of a parameter 'ReplyTo' with email address */
/*         of the one who manages the event                   */
/*         Message is now sent from a generic email           */
/* V19 - June 1st - Minor corrections                         */
/* V20 - June 4th - Corrections:                              */
/*        - Sessions full are now in grey and disabled        */
/*        - Confirmation message modified                     */
/*       Improvements:                                        */
/*        - Email is enriched with title of the event, date   */
/*          and the link to the event in user's calendar      */
/* V21 - June 9th - Minor changes in the messages sent        */
/* V22 - Sept 13th - Correction of the date format in ICS file*/
/*          (Adding of a 'T' between date and hour)           */
/* V23 - Sept 16th - Modification of parameter calId (it shall*/
/*          be now the complete address)                      */
/* V24 to V28: Deployments for specific project. No change    */
/* V29 - Oct 22nd: Timezone is precised in the email sent     */
/* V32 - Instrumentalisation                                  */
/* V33 - August 2022: Adding of a new parameter: timezone     */
/*         Plus the possibility for the user to change it     */
/* V34 - Sept 2022-E.Doumenc - Usage via generic MR calendar  */
/* Vxx - Oct 2022: Color is now optional (Airbus dark blue by */
/*        default)                                            */
/*        The link in the sent email opens a frame instead of */
/*        a full page                                         */
/* V35 - 21/11/2022 Adding event list and search button       */
/*                                                            */
/**************************************************************/

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Main function
// Create the html page based on the file 'page.html'+ generate the call to the init function with parameters given in URL
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function doGet(e) {
  console.log(Session.getActiveUser().getUserLoginId());
  // Build and display HTML page
  var tmp = HtmlService.createTemplateFromFile("page");
  var html = tmp.evaluate().getContent();
  var numDays = 120;
  var color = "00205b";
  var replyTo = "noreply@airbus.com";
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
  ).evaluate().setTitle('📅 Events');
}

function isBlank(str) {
  return (!str || /^\s*$/.test(str));
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Function initialisation
// Generate the HTML code for the buttons corresponding to the events matching the 'searchedText' in the title, in the 'calId'
// calendar. Set the 'color' background.
// Check first if the event session is full or not (if limit defined)
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function initialisation(calId, searchedText = null, color = "731b62", numberofDaysForSearch = 360, usedTimezone = "Europe/Paris") {
  var cal = CalendarApp.getCalendarById(calId);

  // Get calendar events (webinars - next 3 months)
  var now = new Date();
  var searchPeriod = new Date(now.getTime() + numberofDaysForSearch * 24 * 60 * 60 * 1000);

  search_args = {
    timeMin: ISODateString(now),
    timeMax: ISODateString(searchPeriod),
    singleEvents: false,
    maxResults: 2500,
    showDeleted: true
  }

  if (isBlank(searchedText)) {
    Logger.log("skander", searchedText)
    search_args['q'] = searchedText
  }

  // Get events or instances in case of recurring event
  var events = Calendar.Events.list(
    calId,
    search_args,
  ).items;

  Logger.log(events)
  // If the first event has a reccurence, we consider it's only instances of a recurring event
  /*if (events.length > 0) {
    if (events[0].recurringEventId != null) {
      var recurringId = events[0].recurringEventId;
      var events = Calendar.Events.instances(calId, recurringId, {
        timeMin: ISODateString(now),
        timeMax: ISODateString(searchPeriod),
      }).items;
    }
  }*/

  // Sort events per starting date
  events.sort(compareStartingDates);

  // Build HTML string to be displayed
  jsonData = {
    usedTimezone: usedTimezone,
    data: {},
  };

  events.forEach(function (e) {
    // We first check if the event has a limit of participants and if this limit is reached

    var isPossible = true;
    // Get the limit of participants (if any)
    var body = e.getDescription();
    var guests;
    var limit;
    var seatsAvailable = 0;

    if (body != "") {

      var regExp = new RegExp(/\[Limit: ?(\d+)\]/g);
      var res = regExp.exec(body);
      var limit;
      if (res != null) {
        // There is a limit!
        limit = res[1];
        // Check if the limit of participants who accepted is reached (if any limit)
        guests = e.attendees;

        if (guests != undefined) {
          guests = guests.filter(function (g) {
            return g.responseStatus == "accepted" || g.responseStatus == "needsAction";
          });

          if (guests.length >= limit) {
            isPossible = false;
          } else {
            seatsAvailable = limit - guests.length;
          }
        } else {
          seatsAvailable = limit; // By default number of seats available is the limit, if no guests yet
        }
      }
    }

    // Depending on this result, color is requested color or grey, and button is active or not

    var finalColor;
    var btnActivity;
    var addMessage = "";

    if (isPossible == true) {
      finalColor = color;
      btnActivity = "active";
      if (seatsAvailable > 0) {
        addMessage = " (" + seatsAvailable.toString() + " seat(s) remaining)";
      } else {
        addMessage = " (Session open)"
      }
    } else {
      finalColor = "808B96"; // grey
      btnActivity = "inactive";
      addMessage = " (Session full)";
    }
    try {
      // Prepare HTML code for the button
      event_data = {
        'btnActivity': btnActivity,
        'finalColor': finalColor,
        'id': e.getId(),
        'summary': e.summary,
        'addMessage': addMessage,
        'start_date': Utilities.formatDate(new Date(e.start.dateTime), usedTimezone, "MMMM dd, yyyy - HH:mm"),
        'end_date': Utilities.formatDate(new Date(e.end.dateTime), usedTimezone, "HH:mm"),
      };

      start_date = new Date(e.start.dateTime)
      if (
        (start_date >= (new Date())) & // start date should be supperior than actual tie
        (e.status != "cancelled") // event should not be cancelled
      ) {
        jsonData["data"][e.summary + start_date] = event_data;
      }
    }
    catch (e) {
      Logger.log('skip')
    }
  });

  return JSON.stringify(jsonData);

  //return htmlString;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Function registerUserToEvent
// Just add the user in the list of guests of the selected event
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function registerUserToEvent(calId, eventId, replyTo, usedTimeZone) {
  var message;

  // Used for Debug
  var logArray = [];
  logArray.push(new Date());
  logArray.push(calId);
  logArray.push(eventId);
  // End Debug


  try {

    var event = Calendar.Events.get(calId, eventId);

    var guestEmail = Session.getActiveUser().getEmail();
    var aliases = GmailApp.getAliases();

    // Guest added and confrmation message prepared

    var attendees = event.attendees;
    if (attendees) {
      // If there are already attendees, push the new attendee onto the list
      attendees.push({ email: guestEmail });
      var resource = { attendees: attendees };
    } else {
      // If attendees[] doesn't exist, add it manually (there's probably a better way of doing this)
      var resource = { attendees: [{ email: guestEmail }] }
    }

    var args = { sendUpdates: "none" };
    // https://developers.google.com/calendar/api/v3/reference/events/patch
    try {
      var send_invitation = Calendar.Events.patch(resource, calId, eventId, args);
      Logger.log(send_invitation)
    } catch (err) {
      Logger.log(err);
      message = "KO_" + eventId;
      return message;
    }

    // get recent updates
    var event = Calendar.Events.get(calId, eventId);

    //var splitEventId = event.iCalUID.split('@');

    var eventURL = "https://www.google.com/calendar/render?action=VIEW&eid=" + Utilities.base64Encode(eventId + ' ' + guestEmail);
    if (eventURL.slice(-1) === "=") {
      eventURL = eventURL.slice(0, -1)
    }

    Logger.log(eventURL)

    // var icsFile=makeICS(event);

    // For Debug
    logArray.push(guestEmail);
    logArray.push(event.summary);
    logArray.push(event.start.dateTime);
    logArray.push(event.end.dateTime);
    var logArrays = [];
    logArrays[0] = logArray;

    /* try
    {var logSheet = SpreadsheetApp.openById("11tFWVcDgWAj8s6cJKKKahVkoNwfYjSSgyEN2KF3trMM").getSheetByName("Log");
    //Logger.log("LOG: "+logArrays);
    //Logger.log(logSheet.getSheetName());
    logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, 7).setValues(logArrays);
    // End debug
    } catch (e) {
      Logger.log('skip')
    }*/

    var emailBody = "Dear participant, Your registration is confirmed. We thank you for your interest.If you are a Google user, we invite you to accept the event in your calendar directly. If you are not a google user, please download the ICS file attached to this email and then accept the invitation.&nbsp; Would you have any question, feel free to answer to this email.";

    var guestName = guestEmail.split(".")[0].toLowerCase().replace(/^[a-z]/i, function (match) { return match.toUpperCase() });;
    var eventTitle = event.summary;
    var eventStart = Utilities.formatDate(new Date(event.start.dateTime), usedTimeZone, 'MMMM dd, yyyy HH:mm');
    var eventEnd = Utilities.formatDate(new Date(event.end.dateTime), usedTimeZone, 'HH:mm');
    var eventTimespan = eventStart + " - " + eventEnd + " (timezone '" + usedTimeZone + "')";
    var eventDescription = event.description;
    var eventLocation = event.location;
    if (!eventLocation) {
      eventLocation = "Location not available. Please ask meeting organizer";
    }
    var htmlBody = "" +
      "<div style=\"font:small/1.5 Arial,Helvetica,sans-serif;direction:ltr;word-wrap:break-word;text-align:left;word-break:break-word;color:#3c4043!important;font-family:Roboto,sans-serif;font-style:normal;font-weight:400;font-size:14px;line-height:20px;letter-spacing:0.2px;\">" +
      "<p>Dear " + guestName + ",<br><br>Your registration to this event is <strong>confirmed</strong>. Thank you for your interest.</p>" +
      "<p>Please click the below button to accept the event in your calendar directly." +
      "</div>" +
      "<div style=\"font:small/1.5 Arial,Helvetica,sans-serif;direction:ltr;word-wrap:break-word;text-align:left;word-break:break-word;color:#3c4043!important;font-family:Roboto,sans-serif;font-style:normal;font-weight:400;font-size:14px;line-height:20px;letter-spacing:0.2px;text-decoration:none;border:solid 1px #dadce0;border-radius:8px;padding:24px 32px;text-align:left;vertical-align:top;\">" +
      "<div>" +
      "<div style=\"color:#222;font-size:140%\">" + eventTitle + "</div>" +
      "<div style=\"margin-bottom:10px\"><a style=\"color:#15c;cursor:pointer;text-decoration:none\" href=\"" + eventURL + "\" target=\"_blank\">View on Google Calendar</a></div>" +
      "<table>" +
      "<tr>" +
      "<td style=\"color:#999;min-width:60px;vertical-align:top;\">When</td>" +
      "<td>" + eventTimespan + "</td>" +
      "</tr>" +
      "<tr>" +
      "<td style=\"color:#999;vertical-align:top;\">Where</td>" +
      "<td>" + eventLocation + "</td>" +
      "</tr>" +
      "<tr>" +
      "<td style=\"color:#999;vertical-align:top;\">Who</td>" +
      "<td>Open calendar link to see guest list</td>" +
      "</tr>" +
      "</table>" +
      "<p></p>" +
      "<p><a style=\"background-color: #1a73e8; padding: 10px 20px; color: white; text-decoration:none;font-size:14px; font-family:Roboto,sans-serif;font-weight:700;border-radius:5px\" href=\"" + eventURL + "\" target=\"_blank\">Open Calendar and Confirm Participation</a></p>" +
      "</div>" +
      "<p>&nbsp;</p>" +
      "<div>" + eventDescription + "</div>" +
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
    GmailApp.sendEmail(guestEmail, "Registration confirmed", htmlBody, mail_options);
  }

  catch (err) {
    console.log(err);
    message = "KO_" + eventId;
    return message;
  }
  finally {
    console.log("ok-" + calId + "  " + eventId + "  " + replyTo);
    message = "OK_" + eventId;
    return message;

  }
}




///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Technical functions
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
//function putCache(key, array) {cache.put(key, JSON.stringify(array) );}
//function getCache(key) {return JSON.parse(cache.get(key));}

function ldapDate(d) {
  var formattedDate = Utilities.formatDate(d, "GMT+00:00", "yyyyMMdd") + "T" + Utilities.formatDate(d, "GMT+00:00", "HHmmss");
  return formattedDate + "Z";
}
function rfc3339Date(d) {
  var formattedDate = Utilities.formatDate(d, "GMT+00:00", "yyyy-MM-dd") + "T" + Utilities.formatDate(d, "GMT+00:00", "HH:mm:ss");
  return formattedDate + "Z";
}
function makeICS(event) {
  var e = event;
  var attendees = [];
  var guests = e.guests;

  for (g in guests) {
    var guest = guests[g];
    var atendee = [
      "ATTENDEE;", "CUTYPE=INDIVIDUAL;", "ROLE=REQ-PARTICIPANT;",
      "PARTSTAT=NEEDS-ACTION;", "RSVP=TRUE;", "CN=" + guest.guestName + ";",
      "X-NUM-GUESTS=0:mailto:" + guest.guestEmail
    ].join("");
    attendees.push(atendee);
  }

  var vcal = ["BEGIN:VCALENDAR",
    "PRODID:-//Google Inc//Google Calendar 70.9054//EN",
    "VERSION:2.0",
    "CALSCALE:GREGORIAN",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    "DTSTART:" + ldapDate(new Date(e.start.dateTime)),
    "DTEND:" + ldapDate(new Date(e.end.dateTime)),
    "DTSTAMP:" + ldapDate(new Date(Date.now())),
    "ORGANIZER;CN=" + CalendarApp.getCalendarById(e.iCalUID).getName() + ":mailto:" + e.iCalUID,
    "UID:" + e.id,
    attendees.join("\n"),
    "CREATED:" + ldapDate(new Date(e.created)),
    "DESCRIPTION:" + e.description,
    "LAST-MODIFIED:" + ldapDate(new Date(Date.now())), // although if I wasn't changing things as I issue this it'd be e.getLastUpdated()
    "LOCATION:" + e.location,
    "SEQUENCE:" + Date.now(),//this is a horrible hack, but it ensures that this change will overrule all other changes.
    "STATUS:CONFIRMED",
    "SUMMARY:" + e.summary,
    "TRANSP:OPAQUE",
    "END:VEVENT",
    "END:VCALENDAR"
  ].join("\n");

  var icsFile = Utilities.newBlob(vcal, 'text/calendar', 'invite.ics');
  return icsFile;
}
function compareStartingDates(a, b) {
  return (new Date(a.start.dateTime).getTime() > new Date(b.start.dateTime).getTime() ? 1 : -1);
}


function ISODateString(d) {
  function pad(n) { return n < 10 ? '0' + n : n }
  return d.getUTCFullYear() + '-'
    + pad(d.getUTCMonth() + 1) + '-'
    + pad(d.getUTCDate()) + 'T'
    + pad(d.getUTCHours()) + ':'
    + pad(d.getUTCMinutes()) + ':'
    + pad(d.getUTCSeconds()) + 'Z'
}
