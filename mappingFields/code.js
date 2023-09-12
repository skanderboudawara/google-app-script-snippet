MENU_SHEET_NAME = "Name of sheet containing all the dataset names";
DICT_SHEET_NAME = "The stored data";

function getDataFromDictionnary() {
  /*
  // This functions aims to get the data from the Sheet Menu the dataset names
  It will then loop through all the dataset names and get the data from the sheet
  containing the dataset name. It will then store the data in a dictionnary

  It will then compute a matrix where the field appears in the column of each dataset
  
  :param: None

  :return: None
  */

  actualSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // To get the current spreadsheet

  // To store all the data eg: {"dataset_name" : [["field1", "description"], ["field2", "description"], etc],}
  var datasetMetadata = {}; 
  // To store all the data eg: {"field1" : ["dataset_name1", "dataset_name2"], "field2" : ["dataset_name1"], etc}
  var fieldMapping = {}; 

  menuSheet = actualSpreadsheet.getSheetByName(MENU_SHEET_NAME); // To get the sheet containing all the dataset names
  lastRowMenu = menuSheet.getLastRow();

  // getRange(row, column, optNumRows, optNumColumns)
  rangeDataSetsMenu = menuSheet
    .getRange(2, 4, lastRowMenu, 1)
    .getValues()
    .flat(); // To get all the dataset names
  rangeDataSetsMenu = removeDuplicates(rangeDataSetsMenu); // To remove duplicates

  // To gather all field with description and build the datasetMetadata
  rangeDataSetsMenu.forEach(function (datasetName) {
    /*
    // 1. get the sheet of the dataset
    // 2. get the last row of the sheet
    // 3. get the values of the sheet
    // 4. store the values in the datasetMetadata

    :param datasetName: the name of the dataset

    :return: None
    */
    try {
      sheetOfWork = actualSpreadsheet.getSheetByName(datasetName);
      lastRowsheetOfWork = sheetOfWork.getLastRow();
      valuesOfWork = sheetOfWork
        .getRange(2, 2, lastRowsheetOfWork, 2)
        .getValues();
      datasetMetadata[datasetName] = valuesOfWork;

      var resultArray = [];

      valuesOfWork.forEach((subArray) => {
        resultArray.push(subArray[0]);
      });

      resultArray.forEach(function (fieldName) {
        /*
        // 1. check if the field is already in the fieldMapping
        // 2. if yes, add the datasetName to the array
        // 3. if no, create a new array with the datasetName

        :param fieldName: the name of the field

        :return: None
        */
        if (fieldMapping.hasOwnProperty(fieldName)) {
          oldValues = fieldMapping[fieldName];
          oldValues.push(datasetName);
          fieldMapping[fieldName] = oldValues;
        } else {
          fieldMapping[fieldName] = [datasetName];
        }
      });
      Logger.log(`Processed ${datasetName}`);
    } catch (e) {
      // If the sheet is not found
      Logger.log(`${datasetName} sheet was not found`);
    }
  });

  // Concat all Field & Descriptions in a single Array
  var allFieldWithDescriptions = [];

  Object.values(datasetMetadata).forEach((fieldWithDescription) => {
    /*
    // 1. get the values of the datasetMetadata
    // 2. push the values in the allFieldWithDescriptions

    :param fieldWithDescription: the values of the datasetMetadata

    :return: None
    */
    allFieldWithDescriptions.push(fieldWithDescription);
  });

  allFieldWithDescriptions = removeDuplicates(allFieldWithDescriptions.flat());

  // Get Existing Data in Dictionnary
  dictionnarySheet = actualSpreadsheet.getSheetByName(DICT_SHEET_NAME); // To get the sheet containing all the dataset names
  lastRowDictionnary = dictionnarySheet.getLastRow();
  lastColDictionnary = dictionnarySheet.getLastColumn();
  dictionnaryDataFieldDescription = removeDuplicates(
    dictionnarySheet.getRange(2, 2, lastRowDictionnary, 2).getValues()
  );
  dictionnaryDataDataSetName = dictionnarySheet
    .getRange(1, 4, 1, lastColDictionnary)
    .getValues()
    .flat();

  // Add Non Existing Field Descriptions
  newFieldsDescription = findValuesNotInArray(
    dictionnaryDataFieldDescription,
    allFieldWithDescriptions
  ); // To get all the field with description that are not in the dictionnary
  if (newFieldsDescription.length > 0) {
    dictionnarySheet
      .getRange(lastRowDictionnary + 1, 2, newFieldsDescription.length, 2)
      .setValues(newFieldsDescription); // To add the new field with description in the dictionnary
    lastRowDictionnary = lastRowDictionnary + newFieldsDescription.length - 1;
    dictionnarySheet
      .getRange(2, 1, lastRowDictionnary, 1)
      .setFormula(
        'IF(AND(COUNTIF(B:B,B340)>1,COUNTIF(C340,C:C)=1),"several description","ok")'
      ); // To add the formula to check if the field with description is unique
  } else {
    // If there is no new field with description
    Logger.log("No New Fields with Description will be Added");
  }

  // Add Non Existing Dataset
  newDatasetsNames = findValuesNotInArray(
    dictionnaryDataDataSetName,
    rangeDataSetsMenu
  ); // To get all the dataset names that are not in the dictionnary
  if (newDatasetsNames.length > 0) {
    dictionnarySheet
      .getRange(1, lastColDictionnary + 1, 1, newDatasetsNames.length)
      .setValues([newDatasetsNames]); // To add the new dataset names in the dictionnary
    lastColDictionnary = lastColDictionnary + newDatasetsNames.length - 1;
  } else {
    // If there is no new dataset names
    Logger.log("No New Dataset will be added");
  }

  // Update data
  Logger.log("Processing Matrix Computation"); // To log the start of the process
  lastRowDictionnary = dictionnarySheet.getLastRow();
  lastColDictionnary = dictionnarySheet.getLastColumn();
  dictionnaryDataField = dictionnarySheet
    .getRange(2, 2, lastRowDictionnary, 1)
    .getValues()
    .flat(); // To get all the field with description in the dictionnary
  dictionnaryDataDataSetName = dictionnarySheet
    .getRange(1, 4, 1, lastColDictionnary)
    .getValues()
    .flat(); // To get all the dataset names in the dictionnary

  newMatrix = [];
  /*
  The main idea is that through the mapping we can detect the  occurence
  [
    [x , null, null , null],
    [x , null, null , x],
  ]
  */
  dictionnaryDataField.forEach(function (fieldName) {
    /*
    // 1. get the field name
    // 2. get the array of dataset name
    // 3. for each dataset name, check if the field name is in the array
    // 4. if yes, add "x" to the newMatrixValue
    // 5. if no, add null to the newMatrixValue
    // 6. push the newMatrixValue to the newMatrix

    :param fieldName: the name of the field

    :return: None
    */
    var newMatrixValue = [];
    arrayToTest = fieldMapping[fieldName];
    if (arrayToTest == undefined) {
      arrayToTest = [];
    }
    dictionnaryDataDataSetName.forEach(function (dataSetName) {
      /*
      // 1. get the dataset name
      // 2. check if the dataset name is in the array
      // 3. if yes, add "x" to the newMatrixValue
      // 4. if no, add null to the newMatrixValue

      :param dataSetName: the name of the dataset

      :return: None
      */
      if (arrayToTest.indexOf(dataSetName) > -1) {
        newMatrixValue.push("x");
      } else {
        newMatrixValue.push(null);
      }
    });
    newMatrix.push(newMatrixValue);
  });

  // Pop the last array full of empty
  newMatrix.pop(); // To remove the last array full of empty
  dictionnarySheet
    .getRange(2, 4, newMatrix.length, newMatrix[0].length)
    .setValues(newMatrix); // To add the new matrix in the dictionnary
  
  Logger.log("Process finished");
}

function ManualUpdateDictionnary() {
  /*
  // This functions will update the dictionnary manually

  :param: None

  :return: None

  */
  getDataFromDictionnary();
  var ui = SpreadsheetApp.getUi();
  ui.alert("Dictionnary Updated");
}

function onOpen(){
  /*
  // This functions will add a menu to the spreadsheet

  :param: None

  :return: None

  */
  getDataFromDictionnary();
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    {name: "Force Update Dictionnary", functionName: "ManualUpdateDictionnary"}
  ];
  sheet.addMenu("✈️ Functions", menuEntries);
}