// Click the Run button ^^^

/*
Reset Live Doc to a template version after running?
*/
var resetLiveDocToTemplate = false;

/*
ID of Google Document you are writing to (just double click the red text and ctrl+v the id):
*/
var docID = "";

/*
ID of Google Spreadsheet you want updated:
*/
var sheetID = "";

/*
[OPTIONAL] ID of Google Drive folder you want the copy sheet and doc to go into:
Otherwise it will go into your main Google Drive folder.
*/
var folderID = "";

/*
Create copies of the current doc and spreadsheet information?
*/
var createCopies = false;

// Do not touch beyond this line

function main() {
  var docLines = DocumentApp.openById(docID).getBody().getParagraphs();
  
  var shortTimeString = "";
  var yearString = "";
  var bonusStandardsMap = new Map(); // Key:Value format: "Name":[amount,[reasons]]
  var mode = '';

  for (let i = 0; i < docLines.length; i++) { // Begin scan of doc
    let line = docLines[i].getText();

    let longTimeMatch = /(.*) ([0-3]?[0-9])...?, ([0-9]{4})/.exec(line); // Finds if a line is formatted in the form of "January 1st, 2012"
    if (longTimeMatch && !shortTimeString) { // If the line is matched and the shortTimeString has not already been set
      shortTimeString = shortenToTimeString(longTimeMatch);
      yearString = longTimeMatch[3];
    }

    // Enter fines mode (TBD)

    // Enter demerits mode (TBD)

    if (/Bonus Standards.*/.exec(line) && mode == '') // Enter bonus standards mode
    { mode = 'B'; continue; }

    if (docLines[i].getIndentStart() == 36) // Reset mode // This (36) MIGHT be changed in the future, would be nice if it was made dynamic
    { mode = ''; }
    

    if (mode == 'B')
    { readBonusStandards(bonusStandardsMap, docLines, i); }
  } // End scan of doc

  console.log("Updating Operating Standards information for " + shortTimeString + "...");

  if (createCopies) {
    createCopyDoc(shortTimeString + "/" + yearString);
    createCopySheet(bonusStandardsMap, shortTimeString + "/" + yearString); // Creates a new sheet for this minutes with all of the info on it
  }

  updateLiveSheet(bonusStandardsMap, shortTimeString); // Updates the live sheet for this minutes and removes entries from bonusStandardsMap
  
  printUnaddedEntries(bonusStandardsMap); // Notifies operator via console of unadded keys and values.

  if (resetLiveDocToTemplate) {
    resetLiveDoc(docLines); // Reset live doc to match template
  }

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

function readBonusStandards(bonusStandardsMap, docLines, lineNum) {
  let line = docLines[lineNum].getText();
  let matches = /(-?[0-9]+) (Bonus Standards? |BS )?to (.*?) for (.*)/.exec(line); // Finds all text matching the RegEx and organizes into groups

  // If the line doesnt fit the format (usually something like "Passes, automatic, etc, skip reading it")
  if (!matches) {
    return;
  }

  let amount = parseInt(matches[1]);
  let members = matches[3].replaceAll(/ ?(, and|and |,) ?/g, ",") // Replaces ", and " and "and " with ","
                      .replaceAll(/(Brothers? |Associates? )/g, "") //Removes all Brother and Associate titles, as well as whitespace
                      .split(","); // Splits the string into a list at every ","
  let reason = matches[4];

  // Detect if the bonus standard(s) are tabled or not passed. If so, set value to 0
  console.log(members + ": " + reason.toLowerCase());
  if (reason.toLowerCase().includes("not pass") ||
      reason.toLowerCase().includes("tabled") || 
      docLines[lineNum].getIndentStart() < docLines[lineNum+1].getIndentStart() && ( // If next line is a descriptor for current line
        docLines[lineNum+1].getText().toLowerCase().includes("not pass") ||
        docLines[lineNum+1].getText().toLowerCase().includes("tabled"))) {
    amount = 0;
  }

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

function createCopyDoc(shortTimeString) {
  // Copy doc and save id
  let copyDocID = DriveApp.getFileById(docID).makeCopy('Copy Doc ' + shortTimeString).getId();

  // See if folder is given/available
  var folder;
  try {
    folder = DriveApp.getFolderById(folderID);
  } catch(e) {}

  if (folder) {
    // Move Copy Doc to operating standards folder
    let file = DriveApp.getFileById(copyDocID);
    
    file.moveTo(folder);
  }

  DocumentApp.openById(copyDocID).getHeader().setText(""); // Clear header after copying
}

function createCopySheet(bonusStandardsMap, shortTimeString) {
  let spreadsheet = SpreadsheetApp.create('Copy Sheet ' + shortTimeString);
  let sheetActive = spreadsheet.getActiveSheet();

  // Clear sheet
  // sheetActive.getRange(1, 1, 99, 26).setValue(""); // Sheets starts at 1,1 instead of 0,0 like normal, also in format of y, x
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
  // See if folder is given/available
  var folder;
  try {
    folder = DriveApp.getFolderById(folderID);
  } catch(e) {}

  if (folder) {
    // Move Copy Sheet to operating standards folder
    let file = DriveApp.getFileById(spreadsheet.getId());
    
    file.moveTo(folder);
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
  
  // TODO: Highlight all missing entries on the copy spreadsheet
  
}

function resetLiveDoc(docLines) {
  // Reset date
  let fontSize = docLines[4].editAsText().getFontSize();
  docLines[4].setText("LongMonth Dayth, Year");
  docLines[4].editAsText().setFontSize(fontSize);

  // Reset time start
  fontSize = docLines[5].editAsText().getFontSize();
  docLines[5].setText("Call to Order: HH:MM");
  docLines[5].editAsText().setFontSize(fontSize);

  // Reset body
  let sectionCount = 0;
  //i = 6 to keep the title, date, and call to order lines
  for (let i = 6; i < docLines.length-1; i++) {
    if (docLines[i].getIndentStart() > 36) {
      if (sectionCount < 1) {
        docLines[i].clear(); // Keep the paragraph indent of the current section of text to preserve formatting
      } else {
        docLines[i].removeFromParent(); // Remove all lines after the first indent
      }
      sectionCount += 1; 
    } else {
      sectionCount = 0; // Reset section count when a header is read
    }
  }

  // Reset time end
  fontSize = docLines[docLines.length-1].editAsText().getFontSize();
  docLines[docLines.length-1].setText("Adjournment: HH:MM");
  docLines[docLines.length-1].editAsText().setFontSize(fontSize);
}
