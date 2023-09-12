
MENU_SHEET_NAME = "Name of sheet containing all the dataset names"
DICT_SHEET_NAME = "The stored data"

function getDataFromDictionnary() {
  /*
  This functions aims to get the data from the Sheet Dictionnary
  */

  actualSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // To store all the data eg:
  // {
  //  "dataset_name" : [["field1", "description"], ["field2", "description"], etc], 
  //}
  var datasetMetadata = {};
  var fieldMapping = {};

  menuSheet = actualSpreadsheet.getSheetByName(MENU_SHEET_NAME);
  lastRowMenu = menuSheet.getLastRow();

  // getRange(row, column, optNumRows, optNumColumns)
  rangeDataSetsMenu = menuSheet.getRange(2, 4, lastRowMenu, 1).getValues().flat();
  rangeDataSetsMenu = removeDuplicates(rangeDataSetsMenu);

  // To gather all field with description and build the datasetMetadata
  rangeDataSetsMenu.forEach(function (datasetName) {
    try {
      sheetOfWork = actualSpreadsheet.getSheetByName(datasetName);
      lastRowsheetOfWork = sheetOfWork.getLastRow();
      valuesOfWork = sheetOfWork.getRange(2, 2, lastRowsheetOfWork, 2).getValues();
      datasetMetadata[datasetName] = valuesOfWork;

      var resultArray = [];

      valuesOfWork.forEach((subArray) => {
        resultArray.push(subArray[0]);
      });

      resultArray.forEach(function (fieldName) {
        if (fieldMapping.hasOwnProperty(fieldName)) {
          oldValues = fieldMapping[fieldName]
          oldValues.push(datasetName)
          fieldMapping[fieldName] = oldValues;
        } else {
          fieldMapping[fieldName] = [datasetName];
        }
        Logger.log(`Processed ${datasetName}`)

      });
    } catch (e) {
      Logger.log(`${datasetName} sheet was not found`);
    }
  })

  // Concat all Field & Descriptions in a single Array
  var allFieldWithDescriptions = [];

  Object.values(datasetMetadata).forEach((fieldWithDescription) => {
    allFieldWithDescriptions.push(fieldWithDescription);
  });

  allFieldWithDescriptions = removeDuplicates(allFieldWithDescriptions.flat())

  // Get Existing Data in Dictionnary 
  dictionnarySheet = actualSpreadsheet.getSheetByName(DICT_SHEET_NAME);
  lastRowDictionnary = dictionnarySheet.getLastRow();
  lastColDictionnary = dictionnarySheet.getLastColumn();
  dictionnaryDataFieldDescription = removeDuplicates(dictionnarySheet.getRange(2, 2, lastRowDictionnary, 2).getValues());
  dictionnaryDataDataSetName = dictionnarySheet.getRange(1, 4, 1, lastColDictionnary).getValues().flat();

  // Add Non Existing Field Descriptions
  newFieldsDescription = findValuesNotInArray(dictionnaryDataFieldDescription, allFieldWithDescriptions)
  if (newFieldsDescription.length > 0) {
    dictionnarySheet.getRange(lastRowDictionnary + 1, 2, newFieldsDescription.length, 2).setValues(newFieldsDescription)
    lastRowDictionnary = lastRowDictionnary + newFieldsDescription.length - 1
    dictionnarySheet.getRange(2, 1, lastRowDictionnary, 1).setFormula('IF(AND(COUNTIF(B:B,B340)>1,COUNTIF(C340,C:C)=1),"several description","ok")')
  } else {
    Logger.log("No New Fields with Description will be Added")
  };

  // Add Non Existing Dataset
  newDatasetsNames = findValuesNotInArray(dictionnaryDataDataSetName, rangeDataSetsMenu);
  if (newDatasetsNames.length > 0) {
    dictionnarySheet.getRange(1, lastColDictionnary + 1, 1, newDatasetsNames.length).setValues([newDatasetsNames])
    lastColDictionnary = lastColDictionnary + newDatasetsNames.length - 1
  } else {
    Logger.log("No New Dataset will be added")
  };

  // Update data
  Logger.log("Processing Matrix Computation");
  lastRowDictionnary = dictionnarySheet.getLastRow();
  lastColDictionnary = dictionnarySheet.getLastColumn();
  dictionnaryDataField = dictionnarySheet.getRange(2, 2, lastRowDictionnary, 1).getValues().flat();
  dictionnaryDataDataSetName = dictionnarySheet.getRange(1, 4, 1, lastColDictionnary).getValues().flat();

  newMatrix = []
  /*
  The main idea is that through the mapping we can detect the  occurence
  [
    [x , null, null , null],
    [x , null, null , x],
  ]
  */
  dictionnaryDataField.forEach(function (fieldName) {
    var newMatrixValue = []
    arrayToTest = fieldMapping[fieldName]
    if (arrayToTest == undefined) {
      arrayToTest = []
    }
    dictionnaryDataDataSetName.forEach(function (dataSetName) {
      if (arrayToTest.indexOf(dataSetName) > -1) {
        newMatrixValue.push('x')
      } else {
        newMatrixValue.push(null)
      }
    });
    newMatrix.push(newMatrixValue)
  });

  // Pop the last array full of empty
  newMatrix.pop()
  dictionnarySheet.getRange(2, 4, newMatrix.length, newMatrix[0].length).setValues(newMatrix)
  Logger.log("Process finished");
}

