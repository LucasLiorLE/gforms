function onFormSubmit(e) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var responsesSheet1 = spreadsheet.getSheetByName("Form Responses 1");
  var responsesSheet2 = spreadsheet.getSheetByName("Form Responses 2");
  var levelSystemSheet = spreadsheet.getSheetByName("Level System");
  var slapCalculatorSheet = spreadsheet.getSheetByName("Slap Calculator");

  handleFormResponses1(responsesSheet1, levelSystemSheet);

  handleFormResponses2(responsesSheet2, slapCalculatorSheet);
}

function handleFormResponses1(responsesSheet, levelSystemSheet) {
  var lastRowResponses = responsesSheet.getLastRow();
  var name = responsesSheet.getRange(lastRowResponses, 2).getValue();
  var currentLevel = parseInt(responsesSheet.getRange(lastRowResponses, 3).getValue());
  var currentEXP = parseInt(responsesSheet.getRange(lastRowResponses, 4).getValue());
  var timeActive = [
    parseInt(responsesSheet.getRange(lastRowResponses, 5).getValue()), 
    parseInt(responsesSheet.getRange(lastRowResponses, 6).getValue()), 
    parseInt(responsesSheet.getRange(lastRowResponses, 7).getValue()), 
    parseInt(responsesSheet.getRange(lastRowResponses, 8).getValue()), 
    parseInt(responsesSheet.getRange(lastRowResponses, 9).getValue()), 
    parseInt(responsesSheet.getRange(lastRowResponses, 10).getValue()), 
    parseInt(responsesSheet.getRange(lastRowResponses, 11).getValue()) 
  ];
  var targetLevel = parseInt(responsesSheet.getRange(lastRowResponses, 12).getValue());
  var luckRating = parseInt(responsesSheet.getRange(lastRowResponses, 13).getValue());
  var chatFrequency = parseInt(responsesSheet.getRange(lastRowResponses, 14).getValue());

  var expRequired = calculateEXPRequired(currentLevel, targetLevel);

  var expPerMinute = 20; 
  if (luckRating === 1) {
    expPerMinute = 15;
  } else if (luckRating === 3) {
    expPerMinute = 25;
  }

  var totalMinutesPerWeek = timeActive.reduce(function(a, b) { return a + b; }, 0);
  var expPerDay = (totalMinutesPerWeek / 7) * expPerMinute;
  var estimatedDays = expPerDay > 0 ? Math.ceil(expRequired / expPerDay) : "N/A";

  var chatMultiplier = 2; 
  if (chatFrequency === 2) {
    chatMultiplier = 1.8;
  } else if (chatFrequency === 3) {
    chatMultiplier = 1.6;
  } else if (chatFrequency === 4) {
    chatMultiplier = 1.4;
  } else if (chatFrequency === 5) {
    chatMultiplier = 1.2;
  }

  var estimatedMessages = Math.ceil((expRequired / 20) * chatMultiplier);

  var levelSystemLastRow = levelSystemSheet.getLastRow() + 1;
  levelSystemSheet.getRange(levelSystemLastRow, 1).setValue(name);
  levelSystemSheet.getRange(levelSystemLastRow, 2).setValue(estimatedMessages);
  levelSystemSheet.getRange(levelSystemLastRow, 3).setValue(estimatedDays);
  levelSystemSheet.getRange(levelSystemLastRow, 4).setValue(expRequired - currentEXP);

  var completionDate = new Date();
  if (typeof estimatedDays === "number" && estimatedDays > 0) {
    completionDate.setDate(completionDate.getDate() + estimatedDays);
    levelSystemSheet.getRange(levelSystemLastRow, 5).setValue(completionDate.toDateString());
  } else {
    levelSystemSheet.getRange(levelSystemLastRow, 5).setValue("N/A");
  }

  Logger.log('New submission received for ' + name);
}

function handleFormResponses2(responsesSheet, slapCalculatorSheet) {
  var lastRowResponses = responsesSheet.getLastRow();
  var name = responsesSheet.getRange(lastRowResponses, 2).getValue();
  var currentSlaps = parseInt(responsesSheet.getRange(lastRowResponses, 3).getValue());
  var slapMultiplier = parseInt(responsesSheet.getRange(lastRowResponses, 4).getValue());
  var slapsPerDay = parseInt(responsesSheet.getRange(lastRowResponses, 5).getValue());
  var targetSlaps = parseInt(responsesSheet.getRange(lastRowResponses, 6).getValue());

  var slapsWithMultiplier = Math.ceil((targetSlaps - currentSlaps) / slapMultiplier);
  var estimatedDays = slapsPerDay > 0 ? Math.ceil((targetSlaps - currentSlaps) / slapsPerDay) : "N/A";

  var slapCalculatorLastRow = slapCalculatorSheet.getLastRow() + 1;
  slapCalculatorSheet.getRange(slapCalculatorLastRow, 1).setValue(name);
  slapCalculatorSheet.getRange(slapCalculatorLastRow, 2).setValue(estimatedDays);
  slapCalculatorSheet.getRange(slapCalculatorLastRow, 3).setValue(slapsWithMultiplier);

  var completionDate = new Date();
  if (typeof estimatedDays === "number" && estimatedDays > 0) {
    completionDate.setDate(completionDate.getDate() + estimatedDays);
    slapCalculatorSheet.getRange(slapCalculatorLastRow, 4).setValue(completionDate.toDateString());
  } else {
    slapCalculatorSheet.getRange(slapCalculatorLastRow, 4).setValue("N/A");
  }

  Logger.log('New submission received for ' + name);
}

function calculateEXPRequired(currentLevel, targetLevel) {
  var expRequired = 0;
  for (var lvl = currentLevel + 1; lvl <= targetLevel; lvl++) {
    expRequired += (5 * (lvl ** 2)) + (50 * lvl) + 100;
  }
  return expRequired;
}

function isNameInLevelSystem(sheet, name) {
  var lastRow = sheet.getLastRow();
  for (var row = 1; row <= lastRow; row++) {
    if (sheet.getRange(row, 1).getValue() === name) {
      return true;
    }
  }
  return false;
}
