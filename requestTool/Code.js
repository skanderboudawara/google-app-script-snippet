/*
 *
 * Code Created By Stephan
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
  "XBOM Assignment",
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

const ACTUAL_SPREADSHEET = SpreadsheetApp.getActive();

const LISTE_SHEET = ACTUAL_SPREADSHEET.getSheetByName("Listes");
const FEEDER_SHEET = ACTUAL_SPREADSHEET.getSheetByName("Feeder");
const ANSWER_SHEET = ACTUAL_SPREADSHEET.getSheetByName("Requests");
const WEBHOOK_REQUEST_CHAT = "";
const WEBHOOK_ALTEN_CHAT = "";

const ANSWER_COLUMN_NAMES = ANSWER_SHEET.getRange(1, 1, 1, 69)
  .getValues()
  .flat();

const ID_XBOM_COLUMN = ANSWER_COLUMN_NAMES.indexOf("ID Request");
const REQUESTOR_COLUMN = ANSWER_COLUMN_NAMES.indexOf("Name & First name");
const TITLE_COLUMN = ANSWER_COLUMN_NAMES.indexOf("Title");

const regExpSubtask = new RegExp("^\\[XBOM-\\d+-\\d+\\]");
const regExpLastDigits = new RegExp("[0-9]+$");

/*function test() {
	titleAnswer = '[XBOM-1992-1] Skander';
	if (regExpSubtask.test(titleAnswer)) {
		xbomIdNumber = regExpSubtask.exec(titleAnswer)[0].replace('[', '').replace(']', '');
		Logger.log(xbomIdNumber);
		Logger.log(titleAnswer.substring(('[' + xbomIdNumber + '] ').length, titleAnswer.length));
	}
}*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("✈️ XBOM")
    .addItem("✍️ Correct items", "correctPrevious")
    .addToUi();
}

function verifNameChanges() {
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
  const search = this.getRange(offsetRow, columnNumber, this.getMaxRows())
    .createTextFinder(".")
    .useRegularExpression(true)
    .findPrevious();
  return search ? search.getRow() : offsetRow;
};

// Please run this function.
/*function main() {
  const res = ANSWER_SHEET.get1stNonEmptyRowFromBottom(ID_XBOM_COLUMN + 1);
  Logger.log(res); // Retrieve the 1st non empty row of column "C" by searching from BOTTOM of sheet.
}

function test(){
  var lastRowXbomId = getLastRowFromColumn(ANSWER_SHEET, ID_XBOM_COLUMN + 1) + 1;
  Logger.log(lastRowXbomId)
}*/
function getLastRowFromColumn(sheet, colNumber) {
  return ANSWER_SHEET.get1stNonEmptyRowFromBottom(colNumber);
  //return sheet.getRange(sheet.getMaxRows(), colNumber).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
}

function notifyViaChat(messageToSend, webhookLink) {
  var message = { text: messageToSend };
  var payload = JSON.stringify(message);
  var options = {
    method: "POST",
    contentType: "application/json",
    payload: payload,
  };

  var response = UrlFetchApp.fetch(webhookLink, options).getContentText();
}

function countTiret(text) {
  return (text.match(new RegExp("-", "g")) || []).length;
}

function onFormSubmit(e) {
  Logger.log(e.values);
  var answerRowIndex = e.range.getRow();
  registerNecessaryInfo(answerRowIndex);
}

function correctPrevious() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "It Should be used in case of bug",
    "is there any bug?",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() == ui.Button.OK) {
    var lastRowXbomId =
      getLastRowFromColumn(ANSWER_SHEET, ID_XBOM_COLUMN + 1) + 1;
    console.log(lastRowXbomId);
    var answerSheetLastRow = getLastRowFromColumn(ANSWER_SHEET, 1) + 1;
    console.log(answerSheetLastRow);
    for (
      var indexRow = lastRowXbomId;
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
  answerRowIndex = indexRow;
  var lastXbomIdNumber = "ERROR";
  var xbomIdNumber = "ERROR";
  if (!verifNameChanges()) {
    return;
  }
  var titleAnswer = ANSWER_SHEET.getRange(
    answerRowIndex,
    TITLE_COLUMN + 1
  ).getValue();

  if (regExpSubtask.test(titleAnswer)) {
    xbomIdNumber = regExpSubtask
      .exec(titleAnswer)[0]
      .replace("[", "")
      .replace("]", "");
    titleAnswer = titleAnswer.substring(
      ("[" + xbomIdNumber + "] ").length,
      titleAnswer.length
    );
  } else {
    var lastRowXbomId = getLastRowFromColumn(ANSWER_SHEET, ID_XBOM_COLUMN + 1);
    var allXbomId = ANSWER_SHEET.getRange(
      1,
      ID_XBOM_COLUMN + 1,
      lastRowXbomId,
      1
    )
      .getValues()
      .flat();
    var allXbomId_length = allXbomId.length;

    for (var i = allXbomId_length - 1; i > 0; i--) {
      if (countTiret(allXbomId[i]) == 1) {
        lastXbomIdNumber = allXbomId[i];
        break;
      }
    }
    xbomIdNumber =
      "XBOM-" + (parseInt(regExpLastDigits.exec(lastXbomIdNumber)[0]) + 1);
  }

  ANSWER_SHEET.getRange(answerRowIndex, ID_XBOM_COLUMN + 1)
    .setValue(xbomIdNumber)
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
    ANSWER_COLUMN_NAMES.indexOf("XBOM Assignment") + 1
  ).setValue("Waiting assignment");
  ANSWER_SHEET.getRange(
    answerRowIndex,
    ANSWER_COLUMN_NAMES.indexOf("Progress status") + 1
  ).setValue("0%");
  ANSWER_SHEET.getRange(
    answerRowIndex,
    ANSWER_COLUMN_NAMES.indexOf("Workload") + 1
  ).setValue("Waiting level");
  /*ANSWER_SHEET.getRange(answerRowIndex, ANSWER_COLUMN_NAMES.indexOf('Comments') + 1).setValue('None');
	ANSWER_SHEET.getRange(answerRowIndex, ANSWER_COLUMN_NAMES.indexOf('Creation week') + 1).setFormula('=WEEKNUM(A' + answerRowIndex + ',21)');
	ANSWER_SHEET.getRange(answerRowIndex, ANSWER_COLUMN_NAMES.indexOf('Previsional delivery week') + 1).setFormula('=WEEKNUM(AE' + answerRowIndex + ',21)');
	ANSWER_SHEET.getRange(answerRowIndex, ANSWER_COLUMN_NAMES.indexOf('Closing week') + 1).setFormula('=WEEKNUM(AF' + answerRowIndex + ',21)');
	ANSWER_SHEET.getRange(answerRowIndex, ANSWER_COLUMN_NAMES.indexOf('DELIVERY WEEK EXPECTED ELSE CLIENT') + 1).setFormula('=IF(ISBLANK(AE' + answerRowIndex + ')=TRUE,WEEKNUM(I' + answerRowIndex + ',21),WEEKNUM(AE' + answerRowIndex + ',21))');
	ANSWER_SHEET.getRange(answerRowIndex, ANSWER_COLUMN_NAMES.indexOf('Begin date') + 1).setFormula('=IF(AD' + answerRowIndex + '<>"",AD' + answerRowIndex + ',AL' + answerRowIndex + '-15)');
	ANSWER_SHEET.getRange(answerRowIndex, ANSWER_COLUMN_NAMES.indexOf('Delivery date') + 1).setFormula('=IF(AE' + answerRowIndex + '<>"",AE' + answerRowIndex + ',I' + answerRowIndex + ')');*/
  ANSWER_SHEET.getRange(answerRowIndex, TITLE_COLUMN + 1).setValue(titleAnswer);

  messageToSendInChat =
    "```🆕 <users/all> \n\nA new item [" +
    xbomIdNumber +
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
