// Click the Run button ^^^

/*
ID of Google Document you are writing to (just double click the red text and ctrl+v the id):
*/
var docID = "";

/*
ID of Google Spreadsheet you want updated (just double click the red text and ctrl+v the id):
*/
var sheetID = "";


// Do not touch beyond this line


function main() {
  var docLines = DocumentApp.openById(docID).getBody().getParagraphs();
  
  var shortTimeString = "";
  var bonusStandardsMap = new Map(); // Key:Value format: "Name":[amount,[reasons]]
  var mode = '';

  for (let i = 0; i < docLines.length; i++) { // Begin scan of doc
    let line = docLines[i].getText();

    let longTimeMatch = /(.*) ([0-3]?[0-9])...?, ([0-9]{4})/.exec(line); // Finds if a line is formatted in the form of "January 1st, 2012"
    if (longTimeMatch && !shortTimeString) { // If the line is matched and the shortTimeString has not already been set
      shortTimeString = shortenToTimeString(longTimeMatch);
      //console.log(shortTimeString);
    }
    
    // Enter fines mode

    // Enter demerits mode

    if (line == ("Bonus Standards"))// Enter bonus standards mode
    { mode = 'B'; continue; }

    if (docLines[i].getIndentStart() == 36)// Reset mode // This (36) MIGHT be changed in the future, would be nice if it was made dynamic
    { mode = ''; }
    

    if (mode == 'B')
    { readBonusStandards(bonusStandardsMap, line); }
  } // End scan of doc

  updateCopySheet(bonusStandardsMap); // (TODO: NEEDS TO CREATE NEW SHEET) Creates a new sheet for this minutes with all of the info on it
  updateLiveSheet(bonusStandardsMap, shortTimeString); // Updates the live sheet for this minutes and removes entries from bonusStandardsMap
  
  printUnaddedEntries(bonusStandardsMap); // Notifies operator via console of unadded keys and values.

  return;
}

function shortenToTimeString(longTimeMatch) {
  let monthMap = new Map([
    ["January","1"],
    ["February","2"],
    ["March","3"],
    ["April","4"],
    ["May","5"],
    ["June","6"],
    ["July","7"],
    ["August","8"],
    ["September","9"],
    ["October","10"],
    ["November","11"],
    ["December","12"]
  ]);
  return monthMap.get(longTimeMatch[1]) + "/" + longTimeMatch[2];
}

function readBonusStandards(bonusStandardsMap, line) {
  let matches = /(-?[0-9]+) (Bonus Standards? |BS )?to (.*) for (.*)/.exec(line); // Finds all text matching the RegEx and organizes into groups
  
  let amount = parseInt(matches[1]);
  let members = matches[3].replaceAll(/ ?(, and|and|,) ?/g, ",") // Replaces ", and " and "and " with ","
                      .replaceAll(/(Brothers? |Associates? )/g, "") //Removes all Brother and Associate titles, as well as whitespace
                      .split(","); // Splits the string into a list at every ","
  let reason = matches[4].replaceAll(/- .*/g,""); // Removes everything after a hyphen

  /*console.log(amount);
  console.log(members);
  console.log(reason);
  */

  for (let m = 0; m < members.length; m++) {
    if (!bonusStandardsMap.has(members[m])) {
      bonusStandardsMap.set(members[m],[0,[]]); // If map entry is empty, set it to default values
    }

    oldAmount = bonusStandardsMap.get(members[m])[0];
    oldReasons = bonusStandardsMap.get(members[m])[1];

    oldAmount += amount;
    oldReasons.push(reason);

    bonusStandardsMap.set(members[m],[oldAmount,oldReasons]);
  }
}

function updateCopySheet(bonusStandardsMap) {
  let spreadsheet = SpreadsheetApp.openById(sheetID);
  let sheetActive = spreadsheet.getSheetByName("Sheet1");

  // Clear sheet
  sheetActive.getRange(1, 1, 99, 26).setValue(""); // Sheets starts at 1,1 instead of 0,0 like normal, also in format of y, x
  // Write to sheet
  sheetActive.getRange(1, 1).setValue("Names");
  sheetActive.getRange(1, 2).setValue("Bonus Standards");
  sheetActive.getRange(1, 3).setValue("Reason(s)...");

  let rowIndex = 2;
  let colIndex = 1;
  let maxColIndex = 0;
  
  bonusStandardsMap.forEach((value, key) => {
    sheetActive.getRange(rowIndex, colIndex).setValue(key); // Set name
    colIndex += 1;
    sheetActive.getRange(rowIndex, colIndex).setValue(value[0]); // Set unit amount
    for (let r = 0; r < value[1].length; r++) {
      colIndex += 1;
      sheetActive.getRange(rowIndex, colIndex).setValue(value[1][r]);

      if (colIndex > maxColIndex) { // Update max column index
        maxColIndex = colIndex;
      }
    }
    
    rowIndex += 1; // Go down a row
    colIndex = 1; // Reset column index
  });

  for (let c = 1; c <= maxColIndex; c++) { // Resize all columns to match content
    spreadsheet.autoResizeColumn(c);
    spreadsheet.setColumnWidth(c, spreadsheet.getColumnWidth(c) + 4); // Matches default padding on the left to be on the right
  }
}

function updateLiveSheet(bonusStandardsMap, shortTimeString) {
  let spreadsheet = SpreadsheetApp.openById(sheetID);
  let sheetActive = spreadsheet.getSheetByName("Other");

  let dateIndex = 1 + sheetActive.getRange(2,1,1,26).getValues()[0].indexOf("Exec " + shortTimeString); // Index of column that matches date on doc
  let nameIndex = 1; // Index of column containing the names of the members
  
  let namesArray = sheetActive.getRange(2, nameIndex, 99, 1).getValues();

  // Update sheet
  for (let rowIndex = 1; namesArray[rowIndex][0].length !== 0; rowIndex++) { // Loops until there is no name in the name column
    // Match name to bonusStandardsMap key
    let currentNameSplit = namesArray[rowIndex][0].split(' ');
    let lastName = currentNameSplit[1];
    let firstInitial = currentNameSplit[0].charAt(0) + ". " + currentNameSplit[1];
    let nameKey = undefined;

    // See if last name isn't an entry in bonusStandardsMap
    if (typeof(bonusStandardsMap.get(lastName)) !== "undefined") { // If row name's last name matches an entry in bonusStandardsMap
      nameKey = lastName;
    } else if (typeof(bonusStandardsMap.get(firstInitial)) !== "undefined") { // If row name's first initial + last name matches an entry in bonusStandardsMap
      nameKey = firstInitial;
    } else {
      sheetActive.getRange(rowIndex+2, dateIndex).setValue('0'); // Reset to zero if their name isn't on the bonusStandardsMap
      continue;
    }

    // Set name's bonus standard amount to value in bonusStandardsMap
    sheetActive.getRange(rowIndex+2, dateIndex).setValue(bonusStandardsMap.get(nameKey)[0]);
    // Remove entry from bonusStandardsMap
    bonusStandardsMap.delete(nameKey);
  }
}

function printUnaddedEntries(bonusStandardsMap, copySheetID) {
  console.log("These names were unable to be found from the rows on the Live Operating Standards Spreadsheet, please fill in their corresponding entries manually:")
  // Prints out all of the names and corresponding bonus standards
  bonusStandardsMap.forEach((value, key) => {
    console.log(key + ": " + value[0] + " Bonus Standards");
  })
  
  // Highlight all missing entries on the copy spreadsheet
  
}

/*function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var ui = DocumentApp.getUi();
  ui.createAddonMenu()
    .addItem('Run main', 'main')
    .addToUi();
}
*/