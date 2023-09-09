/*
 *
 * Updated by Skander BOUDAWARA
 * V2 : will take into account subtasks when they are in the title
 *
 */

const COLUMN_NEEDED = [
  "Type of Demand",
  "ID Request",
  "Name & First name",
  "Title",
  "Stream",
  "Topic",
  "Status",
  "Référent",
  "project Assignment",
  "Progress status",
  "Workload",
  "Comments",
  "Creation week",
  "Previsional delivery week",
  "Closing week",
  "DELIVERY WEEK EXPECTED ELSE CLIENT",
  "Begin date",
  "Delivery date",
];

const ACTUAL_SPREADSHEET = SpreadsheetApp.getActive(); // Get the current spreadsheet

const LISTE_SHEET = ACTUAL_SPREADSHEET.getSheetByName("Listes"); // Get the sheet "Listes"
const FEEDER_SHEET = ACTUAL_SPREADSHEET.getSheetByName("Feeder"); // Get the sheet "Feeder"
const ANSWER_SHEET = ACTUAL_SPREADSHEET.getSheetByName("Requests");   // Get the sheet "Requests"
const WEBHOOK_REQUEST_CHAT = ""; // Get the webhook link of the chat
const WEBHOOK_ALTEN_CHAT = ""; // Get the webhook link of the chat

const ANSWER_COLUMN_NAMES = ANSWER_SHEET.getRange(1, 1, 1, 69)
  .getValues()
  .flat(); // Get the column names of the sheet "Requests"

const ID_project_COLUMN = ANSWER_COLUMN_NAMES.indexOf("ID Request"); // Get the column number of the column "ID Request"
const REQUESTOR_COLUMN = ANSWER_COLUMN_NAMES.indexOf("Name & First name"); // Get the column number of the column "Name & First name"
const TITLE_COLUMN = ANSWER_COLUMN_NAMES.indexOf("Title"); // Get the column number of the column "Title"

const regExpSubtask = new RegExp("^\\[project-\\d+-\\d+\\]"); // Get the regex of the subtask
const regExpLastDigits = new RegExp("[0-9]+$"); // Get the regex of the last digits

function onOpen() {
  /*
    * This function will be executed when the spreadsheet is opened
    * It will create a menu in the spreadsheet
    * The menu will contain the function "correctPrevious"
    * The function "correctPrevious" will be executed when the user will click on the menu
    * The function "correctPrevious" will correct the previous items  
  
  :param: none

  :return: none
  */
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("✈️ project")
    .addItem("✍️ Correct items", "correctPrevious")
    .addToUi();
}

function verifNameChanges() {
  /*
  // This function will verify if the column names have been modified
  // If the column names have been modified, it will send a message in the chat
  // The message will contain the name of the column that has been modified

  :param: none

  :return: true if the column names have not been modified, false otherwise
  */
  COLUMN_NEEDED_length = COLUMN_NEEDED.length;
  for (var indexColumn = 0; indexColumn < COLUMN_NEEDED_length; indexColumn++) {
    if (!ANSWER_COLUMN_NAMES.includes(COLUMN_NEEDED[indexColumn])) {
      //notifyViaChat('``` Someone has modified the column previously named : ' + COLUMN_NEEDED[indexColumn] + '```', WEBHOOK_REQUEST_CHAT);
      return false;
    }
  }
  return true;
}

Object.prototype.get1stNonEmptyRowFromBottom = function (
  columnNumber,
  offsetRow = 1
) {
  /*
  // This function will get the 1st non empty row from the bottom of the sheet
  // It will search from the bottom of the sheet to the top
  // It will search in the column "columnNumber"
  // It will search from the row "offsetRow"

  :param: columnNumber: the column number of the column to search in
  :param: offsetRow: the row number of the row to start the search from

  :return: the row number of the 1st non empty row from the bottom of the sheet
  */
  const search = this.getRange(offsetRow, columnNumber, this.getMaxRows())
    .createTextFinder(".")
    .useRegularExpression(true)
    .findPrevious();
  return search ? search.getRow() : offsetRow;
};


function getLastRowFromColumn(sheet, colNumber) {
  /*
  // This function will get the last row from the column "colNumber" of the sheet "sheet"

  :param: sheet: the sheet to search in
  :param: colNumber: the column number of the column to search in

  :return: the row number of the last row from the column "colNumber" of the sheet "sheet"
  */
  return ANSWER_SHEET.get1stNonEmptyRowFromBottom(colNumber);
  //return sheet.getRange(sheet.getMaxRows(), colNumber).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
}

function notifyViaChat(messageToSend, webhookLink) {
  /*
  // This function will send a message in the chat
  // The message will contain the message "messageToSend"
  // The message will be sent in the chat "webhookLink"

  :param: messageToSend: the message to send in the chat
  :param: webhookLink: the webhook link of the chat

  :return: none
  */
  var message = { text: messageToSend };
  var payload = JSON.stringify(message);
  var options = {
    method: "POST",
    contentType: "application/json",
    payload: payload,
  };

  var response = UrlFetchApp.fetch(webhookLink, options).getContentText(); // Send the message in the chat
}

function countTiret(text) {
  /*
  // This function will count the number of tiret in the text "text"

  :param: text: the text to count the number of tiret in

  :return: the number of tiret in the text "text"
  */
  return (text.match(new RegExp("-", "g")) || []).length;
}

function onFormSubmit(e) {
  /*
  // This function will be executed when a form is submitted
  // It will get the values of the form

  :param: e: the event that triggered the function

  :return: none
  */
  Logger.log(e.values);
  var answerRowIndex = e.range.getRow();
  registerNecessaryInfo(answerRowIndex);
}

function correctPrevious() {
  /*
  // This function will be executed when the user will click on the menu
  
  :param: none

  :return: none
  */
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "It Should be used in case of bug",
    "is there any bug?",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() == ui.Button.OK) {
    var lastRowprojectId =
      getLastRowFromColumn(ANSWER_SHEET, ID_project_COLUMN + 1) + 1;
    console.log(lastRowprojectId);
    var answerSheetLastRow = getLastRowFromColumn(ANSWER_SHEET, 1) + 1;
    console.log(answerSheetLastRow);
    for (
      var indexRow = lastRowprojectId;
      indexRow <= answerSheetLastRow;
      indexRow++
    ) {
      if (ANSWER_SHEET.getRange(indexRow, 1).getValue() != "") {
        registerNecessaryInfo(indexRow, false);
      }
    }
  }
}

function registerNecessaryInfo(indexRow, sendMessage = true) {
  /*
  // This function will register the necessary info in the sheet "Requests"
  // It will get the values of the row "indexRow"

  :param: indexRow: the row number of the row to get the values of
  :param: sendMessage: true if the function should send a message in the chat, false otherwise

  :return: none
  */
  answerRowIndex = indexRow;
  var lastprojectIdNumber = "ERROR";
  var projectIdNumber = "ERROR";
  if (!verifNameChanges()) {
    return;
  }
  var titleAnswer = ANSWER_SHEET.getRange(
    answerRowIndex,
    TITLE_COLUMN + 1
  ).getValue();

  if (regExpSubtask.test(titleAnswer)) {
    // If the title contains a subtask
    projectIdNumber = regExpSubtask
      .exec(titleAnswer)[0]
      .replace("[", "")
      .replace("]", "");
    titleAnswer = titleAnswer.substring(
      ("[" + projectIdNumber + "] ").length,
      titleAnswer.length
    );
  } else {
    var lastRowprojectId = getLastRowFromColumn(ANSWER_SHEET, ID_project_COLUMN + 1);
    var allprojectId = ANSWER_SHEET.getRange(
      1,
      ID_project_COLUMN + 1,
      lastRowprojectId,
      1
    )
      .getValues()
      .flat();
    var allprojectId_length = allprojectId.length;

    for (var i = allprojectId_length - 1; i > 0; i--) {
      if (countTiret(allprojectId[i]) == 1) {
        lastprojectIdNumber = allprojectId[i];
        break;
      }
    }
    projectIdNumber =
      "project-" + (parseInt(regExpLastDigits.exec(lastprojectIdNumber)[0]) + 1);
  }

  ANSWER_SHEET.getRange(answerRowIndex, ID_project_COLUMN + 1)
    .setValue(projectIdNumber)
    .setHorizontalAlignment("center");
  ANSWER_SHEET.getRange(
    answerRowIndex,
    ANSWER_COLUMN_NAMES.indexOf("Stream") + 1
  ).setValue("Waiting stream");
  ANSWER_SHEET.getRange(
    answerRowIndex,
    ANSWER_COLUMN_NAMES.indexOf("Topic") + 1
  ).setValue("Waiting topic");
  ANSWER_SHEET.getRange(
    answerRowIndex,
    ANSWER_COLUMN_NAMES.indexOf("Status") + 1
  ).setValue("To affect");
  ANSWER_SHEET.getRange(
    answerRowIndex,
    ANSWER_COLUMN_NAMES.indexOf("Référent") + 1
  ).setValue("Waiting assignment");
  ANSWER_SHEET.getRange(
    answerRowIndex,
    ANSWER_COLUMN_NAMES.indexOf("project Assignment") + 1
  ).setValue("Waiting assignment");
  ANSWER_SHEET.getRange(
    answerRowIndex,
    ANSWER_COLUMN_NAMES.indexOf("Progress status") + 1
  ).setValue("0%");
  ANSWER_SHEET.getRange(
    answerRowIndex,
    ANSWER_COLUMN_NAMES.indexOf("Workload") + 1
  ).setValue("Waiting level");
  ANSWER_SHEET.getRange(answerRowIndex, TITLE_COLUMN + 1).setValue(titleAnswer);

  messageToSendInChat =
    "```🆕 <users/all> \n\nA new item [" +
    projectIdNumber +
    "] - " +
    titleAnswer +
    "\nfrom: " +
    ANSWER_SHEET.getRange(answerRowIndex, REQUESTOR_COLUMN + 1).getValue() +
    " has been added \n=========================================```";
  try {
    sendMessage
      ? notifyViaChat(messageToSendInChat, WEBHOOK_REQUEST_CHAT)
      : Logger.log("pass");
  } catch {
    Logger.log("pass");
  }
}

function onEdit(e) {
  /*
  // This function will be executed when a cell is edited

  :param: e: the event that triggered the function

  :return: none
  */
  /*if (Session.getActiveUser().getEmail() == '') {
    return
  }*/
  try {
    var ui = SpreadsheetApp.getUi();
    if (e.source.getActiveSheet().getName() != "Requests") {
      return;
    }
    newValue = e.value;
    newValues = e.range.getValues();
    oldValue = e.oldValue;
    editRowIndex = e.range.getRow();
    if (
      ANSWER_SHEET.getRange(editRowIndex, 1).getValue() == "" ||
      (newValue == oldValue && newValue != null)
    ) {
      return;
    }
    if (e.range.getRowIndex() == 1) {
      if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
        ANSWER_SHEET.getRange(1, 1, 1, ANSWER_SHEET.getLastColumn()).setValues(
          FEEDER_SHEET.getRange(
            1,
            2,
            1,
            ANSWER_SHEET.getLastColumn()
          ).getValues()
        );
        try {
          ui.alert("You have no right to batch edit the header column");
        } catch {
          Logger.log("ERROR IN RANGE COLUMN EDITION");
        }
        return;
      }
    }
    if (editRowIndex == 1) {
      if (oldValue == null) {
        oldValue = FEEDER_SHEET.getRange(1, e.range.getColumn() + 1).getValue();
      }
      if (newValue == null) {
        newValue = newValues.flat()[0];
      }
      try {
        var response = ui.alert(
          "Are you sure?",
          `Are you sure you want to edit the header column?\n\n  Your want to rename [${oldValue}] to → [${newValue}] \n\n`,
          ui.ButtonSet.OK_CANCEL
        );
        if (response == ui.Button.CANCEL) {
          ANSWER_SHEET.getRange(1, e.range.getColumn()).setValue(oldValue);
        } else if (response == ui.Button.OK) {
          try {
            var response2 = ui.alert(
              "Be careful !",
              `This will impact DATASTUDIO & APP SCRIPTS`,
              ui.ButtonSet.OK_CANCEL
            );
            if (response2 == ui.Button.CANCEL) {
              ANSWER_SHEET.getRange(1, e.range.getColumn()).setValue(oldValue);
            } else if (response2 == ui.Button.OK) {
              FEEDER_SHEET.getRange(1, e.range.getColumn() + 1).setValue(
                oldValue
              );
              messageToSendInChat =
                "```🚨 <users/all>\n\n" +
                Session.getActiveUser().getEmail() +
                "\nhas changed the column: [" +
                oldValue +
                "] to → [" +
                newValue +
                "] ```";
              notifyViaChat(messageToSendInChat, WEBHOOK_ALTEN_CHAT);
            } else {
              ANSWER_SHEET.getRange(1, e.range.getColumn()).setValue(oldValue);
            }
          } catch {
            ANSWER_SHEET.getRange(1, e.range.getColumn()).setValue(oldValue);
          }
        } else {
          ANSWER_SHEET.getRange(1, e.range.getColumn()).setValue(oldValue);
        }
      } catch {
        ANSWER_SHEET.getRange(1, e.range.getColumn()).setValue(oldValue);
      }

      return;
    }

    ANSWER_SHEET.getRange(
      editRowIndex,
      ANSWER_COLUMN_NAMES.indexOf("LAST PERSON TO UPDATE") + 1,
      e.range.getNumRows(),
      1
    ).setValue(Session.getActiveUser().getEmail());
    ANSWER_SHEET.getRange(
      editRowIndex,
      ANSWER_COLUMN_NAMES.indexOf("LAST UPDATE") + 1,
      e.range.getNumRows(),
      1
    ).setValue(new Date());

    if (newValues.flat().includes("Done")) {
      ANSWER_SHEET.getRange(
        editRowIndex,
        ANSWER_COLUMN_NAMES.indexOf("Progress status") + 1,
        e.range.getNumRows(),
        1
      ).setValue("100%");
      if (
        ANSWER_SHEET.getRange(
          editRowIndex,
          ANSWER_COLUMN_NAMES.indexOf("Closing Date ") + 1,
          e.range.getNumRows(),
          1
        ).getValue() == ""
      ) {
        ANSWER_SHEET.getRange(
          editRowIndex,
          ANSWER_COLUMN_NAMES.indexOf("Closing Date ") + 1,
          e.range.getNumRows(),
          1
        ).setValue(new Date());
      }
    }
  } catch {
    Logger.log("Couldn't perform the action");
  }
}
