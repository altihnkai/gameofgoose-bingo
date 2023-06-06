--------------
macro's.gs
--------------

/** @OnlyCurrentDoc */

function Naamlozemacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
};

--------------
naamloos.gs
--------------

function onSelectionChange() {
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("assignments");
  var targetRange = targetSheet.getRange("A1:A5");
  var targetValues = targetRange.getValues();
  
  var webhookUrl = "<webhook url>";
  
  // Generate a random index between 0 and the number of values - 1
  var randomIndex = Math.floor(Math.random() * targetValues.length);
  
  var value = targetValues[randomIndex][0];
  var payload = {
    content: value
  };
  var options = {
    method: "post",
    headers: {
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch(webhookUrl, options);
}

---------------
automatedrolls.gs
---------------

function moveTeam1AndPostToDiscord() {
  moveTeamAndPostToDiscord("Team 1", 2);
}

function moveTeam2AndPostToDiscord() {
  moveTeamAndPostToDiscord("Team 2", 3);
}

function moveTeam3AndPostToDiscord() {
  moveTeamAndPostToDiscord("Team 3", 4);
}

function moveTeam4AndPostToDiscord() {
  moveTeamAndPostToDiscord("Team 4", 5);
}

function moveTeam5AndPostToDiscord() {
  moveTeamAndPostToDiscord("Team 5", 6);
}

function moveTeamAndPostToDiscord(team, columnIndex) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var updateTableSheet = spreadsheet.getSheetByName("UpdateTable");
  var gooseBoardSheet = spreadsheet.getSheetByName("Game of Goose Board");

  var updateTableRange = updateTableSheet.getDataRange();
  var updateTableValues = updateTableRange.getValues();

  for (var i = 0; i < updateTableValues.length; i++) {
    if (updateTableValues[i][columnIndex - 1] === team) {
      var updateTableCell = updateTableSheet.getRange(i + 1, columnIndex);
      var currentValue = updateTableCell.getValue();

      var randomAmount = Math.floor(Math.random() * 6) + 1;
      var newRow = i + 1 + randomAmount;

      // Additional conditions for row adjustments
      if (newRow === 15) {
        newRow = 14;
      } else if (newRow === 28) {
        newRow = 31;
      } else if (newRow === 47) {
        newRow = 42;
      

      //end of the board returns
      } else if (newRow === 66) {
        newRow = 64;
      } else if (newRow === 67) {
        newRow = 63;
      } else if (newRow === 68) {
        newRow = 62;
      } else if (newRow === 69) {
        newRow = 61;
      } else if (newRow === 70) {
        newRow = 60;
      } else if (newRow === 71) {
        newRow = 59;
      } else if (newRow === 72) {
        newRow = 58;
      }
      
      // Limit newRow to a maximum of 65 (can not exceed tile 63)
      //newRow = Math.min(newRow, 65);

      updateTableSheet.getRange(newRow, columnIndex).setValue(currentValue);
      updateTableCell.clearContent();

      var tileValue = updateTableValues[newRow - 1][0];
      var tileBelowValue = getTileBelowValue(gooseBoardSheet, tileValue);
      var discordMessage =
        team +
        " has rolled " +
        randomAmount +
        " tiles and is currently on tile " +
        tileValue +
        ". You will need to obtain: " +
        tileBelowValue;

      postToDiscord(discordMessage);
      break;
    }
  }
}


function getTileBelowValue(sheet, tileValue) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] == tileValue && i < values.length - 1) {
        return values[i + 1][j];
      }
    }
  }
  return "";
}

function postToDiscord(message) {
  var webhookUrl = "<webhook url>";

  var payload = {
    content: message
  };

  var params = {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch(webhookUrl, params);
}

--------------
EventModeration.gs
--------------

//Move teams +1 tiles - Event Mod action from Event Moderation sheet

function moveTeamOneDown() {
  MoveDown("Team 1", 2);
}

function moveTeamTwoDown() {
  MoveDown("Team 2", 3);
}

function moveTeamThreeDown() {
  MoveDown("Team 3", 4);
}

function moveTeamFourDown() {
  MoveDown("Team 4", 5);
}

function moveTeamFiveDown() {
  MoveDown("Team 5", 6);
}

function MoveDown(team, columnIndex) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var updateTableSheet = spreadsheet.getSheetByName("UpdateTable");
  var gooseBoardSheet = spreadsheet.getSheetByName("Game of Goose Board");

  var updateTableRange = updateTableSheet.getDataRange();
  var updateTableValues = updateTableRange.getValues();
  
  for (var i = 0; i < updateTableValues.length; i++) {
    if (updateTableValues[i][columnIndex - 1] === team) {
      var updateTableCell = updateTableSheet.getRange(i + 1, columnIndex);
      var currentValue = updateTableCell.getValue();

      var newRow = i + 1 + 1;

  

      updateTableSheet.getRange(newRow, columnIndex).setValue(currentValue);
      updateTableCell.clearContent();

      var tileValue = updateTableValues[newRow - 1][0];
      var tileBelowValue = getTileBelowValue(gooseBoardSheet, tileValue);
      var webhookUrl = '<webhook url>';
    var message = 'Due to landing on the same tile again, ' + team + ' has been moved an extra tile by an Event Moderator. Your new assignment is: ' + tileBelowValue;

    var payload = {
      content: message
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    };

    UrlFetchApp.fetch(webhookUrl, options);
    }
  }
}


// Team 1 rerolls

function moveTeamFourUpByOneRow() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("B:B");
  var values = range.getValues();

  var foundRow = -1;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 1') {
      foundRow = i;
      break;
    }
  }

  if (foundRow >= 1) {
    var teamValue = sheet.getRange(foundRow + 1, 2).getValue();
    sheet.getRange(foundRow + 1, 2).clearContent();
    sheet.getRange(foundRow, 2).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 1. The team has been moved back 1 tile.`);
  }
}

function moveTeamFourUpByTwoRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("B:B");
  var values = range.getValues();

  var foundRow = -1;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 1') {
      foundRow = i;
      break;
    }
  }

  if (foundRow >= 2) {
    var teamValue = sheet.getRange(foundRow + 1, 2).getValue();
    sheet.getRange(foundRow + 1, 2).clearContent();
    sheet.getRange(foundRow - 1, 2).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 1. The team has been moved back 2 tiles.`);
  }
}

function moveTeamFourUpByThreeRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("B:B");
  var values = range.getValues();

  var foundRow = -1;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 1') {
      foundRow = i;
      break;
    }
  }

  if (foundRow >= 3) {
    var teamValue = sheet.getRange(foundRow + 1, 2).getValue();
    sheet.getRange(foundRow + 1, 2).clearContent();
    sheet.getRange(foundRow - 2, 2).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 1. The team has been moved back 3 tiles.`);
  }
}

// Team 2 rerolls

function moveTeamTwoUpByOneRow() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("C:C");
  var values = range.getValues();

  var foundRow = -1;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 2') {
      foundRow = i;
      break;
    }
  }

  if (foundRow > 0) {
    var teamValue = sheet.getRange(foundRow + 1, 3).getValue();
    sheet.getRange(foundRow + 1, 3).clearContent();
    sheet.getRange(foundRow, 3).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 2. The team has been moved back 1 tile.`);
  }
}

function moveTeamTwoUpByTwoRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("C:C");
  var values = range.getValues();

  var foundRow = -2;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 2') {
      foundRow = i;
      break;
    }
  }

  if (foundRow > 0) {
    var teamValue = sheet.getRange(foundRow + 1, 3).getValue();
    sheet.getRange(foundRow + 1, 3).clearContent();
    sheet.getRange(foundRow - 1, 3).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 2. The team has been moved back 2 tiles.`);
  }
}

function moveTeamTwoUpByThreeRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("C:C");
  var values = range.getValues();

  var foundRow = -3;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 2') {
      foundRow = i;
      break;
    }
  }

  if (foundRow > 0) {
    var teamValue = sheet.getRange(foundRow + 1, 3).getValue();
    sheet.getRange(foundRow + 1, 3).clearContent();
    sheet.getRange(foundRow - 2, 3).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 2. The team has been moved back 3 tiles.`);
  }
}

// Team 3 rerolls

function moveTeamThreeUpByOneRow() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("D:D");
  var values = range.getValues();

  var foundRow = -1;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 3') {
      foundRow = i;
      break;
    }
  }

  if (foundRow > 0) {
    var teamValue = sheet.getRange(foundRow + 1, 4).getValue();
    sheet.getRange(foundRow + 1, 4).clearContent();
    sheet.getRange(foundRow, 4).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 3. The team has been moved back 1 tile.`);
  }
}

function moveTeamThreeUpByTwoRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("D:D");
  var values = range.getValues();

  var foundRow = -2;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 3') {
      foundRow = i;
      break;
    }
  }

  if (foundRow > 0) {
    var teamValue = sheet.getRange(foundRow + 1, 4).getValue();
    sheet.getRange(foundRow + 1, 4).clearContent();
    sheet.getRange(foundRow - 1, 4).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 3. The team has been moved back 2 tiles.`);
  }
}

function moveTeamThreeUpByThreeRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("D:D");
  var values = range.getValues();

  var foundRow = -5;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 3') {
      foundRow = i;
      break;
    }
  }

  if (foundRow > 0) {
    var teamValue = sheet.getRange(foundRow + 1, 4).getValue();
    sheet.getRange(foundRow + 1, 4).clearContent();
    sheet.getRange(foundRow - 2, 4).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 3. The team has been moved back 3 tiles.`);
  }
}

// Team 4 rerolls

function moveTeamFourUpByOneRow() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("E:E");
  var values = range.getValues();

  var foundRow = -1;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 4') {
      foundRow = i;
      break;
    }
  }

  if (foundRow >= 1) {
    var teamValue = sheet.getRange(foundRow + 1, 5).getValue();
    sheet.getRange(foundRow + 1, 5).clearContent();
    sheet.getRange(foundRow, 5).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 4. The team has been moved back 1 tile.`);
  }
}

function moveTeamFourUpByTwoRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("E:E");
  var values = range.getValues();

  var foundRow = -1;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 4') {
      foundRow = i;
      break;
    }
  }

  if (foundRow >= 2) {
    var teamValue = sheet.getRange(foundRow + 1, 5).getValue();
    sheet.getRange(foundRow + 1, 5).clearContent();
    sheet.getRange(foundRow - 1, 5).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 4. The team has been moved back 2 tiles.`);
  }
}

function moveTeamFourUpByThreeRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("E:E");
  var values = range.getValues();

  var foundRow = -1;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 4') {
      foundRow = i;
      break;
    }
  }

  if (foundRow >= 3) {
    var teamValue = sheet.getRange(foundRow + 1, 5).getValue();
    sheet.getRange(foundRow + 1, 5).clearContent();
    sheet.getRange(foundRow - 2, 5).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 4. The team has been moved back 3 tiles.`);
  }
}

// Team 5 rerolls

function moveTeamFiveUpByOneRow() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("F:F");
  var values = range.getValues();

  var foundRow = -1;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 5') {
      foundRow = i;
      break;
    }
  }

  if (foundRow >= 1) {
    var teamValue = sheet.getRange(foundRow + 1, 6).getValue();
    sheet.getRange(foundRow + 1, 6).clearContent();
    sheet.getRange(foundRow, 6).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 5. The team has been moved back 1 tile.`);
  }
}

function moveTeamFiveUpByTwoRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("F:F");
  var values = range.getValues();

  var foundRow = -1;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 5') {
      foundRow = i;
      break;
    }
  }

  if (foundRow >= 2) {
    var teamValue = sheet.getRange(foundRow + 1, 6).getValue();
    sheet.getRange(foundRow + 1, 6).clearContent();
    sheet.getRange(foundRow - 1, 6).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 5. The team has been moved back 2 tiles.`);
  }
}

function moveTeamFiveUpByThreeRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('UpdateTable');

  var range = sheet.getRange("F:F");
  var values = range.getValues();

  var foundRow = -1;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === 'Team 5') {
      foundRow = i;
      break;
    }
  }

  if (foundRow >= 3) {
    var teamValue = sheet.getRange(foundRow + 1, 6).getValue();
    sheet.getRange(foundRow + 1, 6).clearContent();
    sheet.getRange(foundRow - 2, 6).setValue(teamValue);
    sendDiscordMessage(`Reroll has been assigned to Team 5. The team has been moved back 3 tiles.`);
  }
}


function sendDiscordMessage(message) {
  var webhookUrl = '<webhook url>';

  var payload = {
    content: message
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch(webhookUrl, options);
}


