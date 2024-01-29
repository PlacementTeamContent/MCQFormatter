const ss = SpreadsheetApp.getActiveSpreadsheet();
const number_of_options = 4;
const number_of_tags = 3;
const subSheetName = "Sheet44"
 
function createSubSheetWithData() {
  //  var ss = SpreadsheetApp.openById("1_Ld78mycoMNNSzV8r0oJGBHBY5NUhAlI1s0FDDi3HW4");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define the source and target sheets
  var sourceSheet = ss.getSheetByName(subSheetName);
  var targetSheetName = "Questions";
  
  // Create a new sub-sheet
  var targetSheet = ss.insertSheet(targetSheetName);
  
  // Get data from the source sheet
  var data = sourceSheet.getDataRange().getValues()
  // Logger.log(data)
 
 
  // Write data to the target sheet
  targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}
 
function addBlankRowsAfterEachRow() {
  const sheet = ss.getSheetByName("Questions");
  
   // Replace with the name of your sheet
  var lastRow = sheet.getLastRow();
 
  // Loop through the rows from row 2 to the last active row
  row = 2
  while (row <= lastRow) {
    // Insert blank rows after the current row
    var number_of_rows_to_add = number_of_tags + number_of_options;
    sheet.insertRowsAfter(row, number_of_rows_to_add);
    
    // Increment the row counter to skip the newly added blank rows
    row += number_of_rows_to_add + 1;
    lastRow = sheet.getLastRow();
  }
}
 
function transposeRowsToColumns() {
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Questions"); // Replace with the actual name of your sheet
  var numRows = sheet.getLastRow();
  var startRow = 2; // Starting row for your data
  var target_start_row = (startRow + number_of_tags + 1);
 
  if (sheet) {
    i = 0;
    while (startRow + i * 8 <= numRows) {
      var sourceRange = sheet.getRange(startRow + i * 8, 3, 1, 4); // Adjust as needed
      var rowData = sourceRange.getValues()[0];
 
      // Create a target range in a 4x1 column (e.g., B6:B9, B14:B17, B22:B25, etc.)
      var targetRange = sheet.getRange(target_start_row, 2, 4, 1);
 
      // Set the transposed data to the target range
      targetRange.setValues(rowData.map(function(value) { return [value]; }));
      i += 1;
      target_start_row += 8;
    }
  } else {
    Logger.log("Sheet not found with the specified name.");
  }
}
 
function setFalseInGColumn() {
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Questions"); // Replace with the actual name of your sheet
  var numRows = sheet.getLastRow();
 
  if (sheet) {
    for (var i = 2; i <= numRows; i++) {
      var cellA = sheet.getRange("A" + i).getValue();
      var cellB = sheet.getRange("B" + i).getValue();
      var cellH = sheet.getRange("H" + i);
 
      if (cellA === "" && cellB !== "") {
        cellH.setValue("FALSE");
      }
    }
  } else {
    Logger.log("Sheet not found with the specified name.");
  }
}
 
function placeTrueOption() {
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Questions"); // Replace with the actual name of your sheet
  var numRows = sheet.getLastRow();
  var startRow = 2; // Starting row for your data
  var target_start_row = 6;
 
  if (sheet) {
    var i = 0;
    while (startRow + i * 8 <= numRows) {
      var sourceRange = sheet.getRange(startRow + i * 8, 8, 1, 1); // Adjust as needed
      var value = sourceRange.getValue();
 
      // Create a target range in a 4x1 column (e.g., G6:G9, G14:G17, G22:G25, etc.)
      var targetRange;
 
      // Set the "True" or "False" value to the target range
      if (value.toLowerCase() === "Option A".toLowerCase()) {
        targetRange = sheet.getRange(target_start_row, 8, 1, 1);
      } else if (value.toLowerCase() === "Option B".toLowerCase()){
        targetRange = sheet.getRange(target_start_row + 1, 8, 1, 1);
      } else if (value.toLowerCase() === "Option C".toLowerCase()){
        targetRange = sheet.getRange(target_start_row + 2, 8, 1, 1);
      } else if (value.toLowerCase() === "Option D".toLowerCase()){
        targetRange = sheet.getRange(target_start_row + 3, 8, 1, 1);
      }
      targetRange.setValue("TRUE");
      i += 1;
      target_start_row += 8;
    }
  } else {
    Logger.log("Sheet not found with the specified name.");
  }
}
 
function copyColumnDataToAnotherColumn() {
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sourceSheet = ss.getSheetByName("Questions");
  
  var sourceColumn = 2; // Source column number (Column B)
  var targetColumn = 14; // Target column number (Column N)
 
  // Get the last active row in the source column
  var lastRow = sourceSheet.getLastRow();
  
  // Get the source data range
  var sourceDataRange = sourceSheet.getRange(2, sourceColumn, lastRow - 1, 1);
  var sourceData = sourceDataRange.getValues();
  
  // Get the target data range
  var targetDataRange = sourceSheet.getRange(2, targetColumn, lastRow - 1, 1);
  
  // Copy data from the source column to the target column
  targetDataRange.setValues(sourceData);
}
 
function copyExplanationData() {
  var sourceSheet = ss.getSheetByName("Questions");
  var sourceColumn = 7; // Source column number (Column G)
  var targetColumn = 25; // Target column number (Column Y)
  
  // Get the last active row in the source column
  var lastRow = sourceSheet.getLastRow();
 
  // Get the source data range
  var sourceDataRange = sourceSheet.getRange(2, sourceColumn, lastRow, 1);
  var sourceData = sourceDataRange.getValues();
 
  // Get the target data range
  var targetDataRange = sourceSheet.getRange(2, targetColumn, lastRow, 1);
  
  // Copy data from the source column to the target column
  targetDataRange.setValues(sourceData);
}
 
function addQuestionDefaultColumn() {
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Questions"); // Replace with the name of your sheet
  var data = sheet.getRange("A2:A").getValues(); // Get data in column A starting from row 2
  var question_type = [];
  var answer_count = [];
  var tag_name_count = [];
  var explanation_content_type = [];
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== "") {
      question_type.push(["MULTIPLE_CHOICE"]);
      answer_count.push([4]);
      tag_name_count.push([3]);
      explanation_content_type.push(["TEXT"]);
    } else {
      question_type.push([""]);
      answer_count.push([""]);
      tag_name_count.push([""]);
      explanation_content_type.push([""]);
    }
  }
  
  sheet.getRange(2, 13, question_type.length, 1).setValues(question_type);
  sheet.getRange(2, 21, answer_count.length, 1).setValues(answer_count);
  sheet.getRange(2, 23, tag_name_count.length, 1).setValues(tag_name_count);
  sheet.getRange(2, 26, explanation_content_type.length, 1).setValues(explanation_content_type);
}
 
function copyHToU() {
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Questions"); // Replace with the actual name of your sheet
  var numRows = sheet.getLastRow();
 
  if (sheet) {
    for (var i = 2; i <= numRows; i++) {
      var cellH = sheet.getRange("H" + i).getValue();
      var cellU = sheet.getRange("U" + i);
 
      if (cellH === true || cellH === false) {
        cellU.setValue(cellH);
      }
    }
  } else {
    Logger.log("Sheet not found with the specified name.");
  }
}
 
function addTags() {
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Questions"); // Replace with the actual name of your sheet
  var numRows = sheet.getLastRow();
 
  if (sheet) {
    for (var i = 2; i <= numRows; i++) {
      var tag_name_count = sheet.getRange("W" + i).getValue();
      var tag_name_1 = sheet.getRange("X" + (i + 1));
 
      var topic = sheet.getRange("I" + i).getValue();
      var tag_name_2 = sheet.getRange("X" + (i + 2));
 
      var difficulty = sheet.getRange("J" + i).getValue();
      var tag_name_3 = sheet.getRange("X" + (i + 3));
 
      if (tag_name_count === 3) {
        tag_name_1.setValue("POOL_1");
      }
 
      if (topic) {
        tag_name_2.setValue("TOPIC_" + topic + "_MCQ");
      }
 
      if (difficulty) {
        tag_name_3.setValue("DIFFICULTY_" + difficulty);
      }
    }
  } else {
    Logger.log("Sheet not found with the specified name.");
  }
}
 
function moveColumnImageToLast() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Questions"); // Replace with the actual name of your sheet
  var lastColumn = sheet.getLastColumn();
 
  if (sheet) {
    var dataRange = sheet.getRange("K1:K" + sheet.getLastRow());
    var values = dataRange.getValues();
 
    // Clear column K
    dataRange.clearContent();
 
    // Append the values to the last column (e.g., column Z)
    var targetRange = sheet.getRange(1, lastColumn + 1, values.length, 1);
    targetRange.setValues(values);
  } else {
    Logger.log("Sheet not found with the specified name.");
  }
}
 
function deleteColumnsAK() {
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Questions"); // Replace with the actual name of your sheet
  
  if (sheet) {
    sheet.deleteColumns(1, 11); // Delete 11 columns starting from column A
  } else {
    Logger.log("Sheet not found with the specified name.");
  }
}
 
function addDataToColumn() {
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Questions"); // Replace with the name of your sheet
  var data = sheet.getRange("C2:C").getValues(); // Get data in column C starting from row 2
  var question_id = [];
  var multimedia_count = [];
  var Language = [];
  var content_type = [];
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== "") {
      question_id.push([Utilities.getUuid()]);
      multimedia_count.push([0]);
      Language.push(["ENGLISH"]);
      content_type.push(["TEXT"]);
    } else {
      question_id.push([""]);
      multimedia_count.push([""]);
      Language.push([""]);
      content_type.push([""]);
    }
  }
  
  sheet.getRange(2, 1, question_id.length, 1).setValues(question_id);
  sheet.getRange(2, 5, multimedia_count.length, 1).setValues(multimedia_count);
  sheet.getRange(2, 9, Language.length, 1).setValues(Language);
  sheet.getRange(2, 11, content_type.length, 1).setValues(content_type);
}
 
function boldHeaderRow() {
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Questions"); // Replace with the actual name of your sheet
 
  if (sheet) {
    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setFontWeight("bold");
  } else {
    Logger.log("Sheet not found with the specified name.");
  }
}
 
function mcqFormatter() {
  createSubSheetWithData();
 
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var targetSheet = ss.getSheetByName("Questions");
  
  addBlankRowsAfterEachRow();
  transposeRowsToColumns();
  setFalseInGColumn();
  placeTrueOption();
 
  var newColumnHeaders = ["question_id",  "question_type", "question_content", "short_text",  "multimedia_count", "multimedia_format",  "multimedia_url", "thumbnail_url",  "Language", "answer_count", "content_type", "tag_name_count", "tag_names",  "answer_explanation_content", "explanation_content_type"];
  
  var existingHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  
  // Concatenate the existing headers and the new headers
  var combinedHeaders = existingHeaders.concat(newColumnHeaders);
  
  // Write the combined headers to the first row of the sub-sheet
  targetSheet.getRange(1, 1, 1, combinedHeaders.length).setValues([combinedHeaders]);
 
  copyColumnDataToAnotherColumn();
  copyExplanationData();
  addQuestionDefaultColumn();
  copyHToU();
  addTags();
  moveColumnImageToLast();
  deleteColumnsAK();
  addDataToColumn();
  boldHeaderRow();
}