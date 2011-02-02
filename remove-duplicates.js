var Columns = 6;
var Debug   = true;

function onOpen() {  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  
  menuEntries.push({name: "Show duplicates", functionName: "showDuplicates"});
  menuEntries.push({name: "Remove duplicates", functionName: "removeDuplicates"});
  spreadsheet.addMenu("Transactions", menuEntries);
}

function logError(message) {
  Logger.log('ERROR: ' + message);
};

function logInfo(message) {
  Logger.log('INFO: ' + message);
};

function logDebug(message) {
  if (Debug)
    Logger.log('DEBUG: ' + message);
};

function formatDate(date) {
  return Utilities.formatDate(date, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
};

function clearBackgroundColours() {
  var spreadsheet  = SpreadsheetApp.getActiveSpreadsheet();
  var transactions = spreadsheet.getSheetByName('transactions');
  
  transactions.getDataRange().setBackgroundRGB(255, 255, 255);
};

function duplicateTransactions(tran1, tran2) {
  logDebug(tran1);
  logDebug(tran2);

  // Set the transaction dates to their string equivalents to allow comparison
  var tran1Date = formatDate(tran1[0]);
  var tran2Date = formatDate(tran2[0]);
  tran1[0] = tran1Date;
  tran2[0] = tran2Date;
  
  // Are the two transactions equal?
  var duplicateTransaction = true;
  for (var colIndex = 0; colIndex < Columns; colIndex++) {
    if (tran1[colIndex] != tran2[colIndex]) {
      logInfo('Transactions differ: ' + tran1[colIndex] + ' / ' + tran2[colIndex]);
      duplicateTransaction = false;
      break;
    };
  };
  
  return duplicateTransaction;
};
  
function eachDuplicate(transactions, callback) {
  var numberOfRows = transactions.getLastRow();
  
  // Sort by Transaction ID
  transactions.sort(2, true);
  
  // Reverse loop through transactions and delete duplicates
  for (var rowIndex = numberOfRows; rowIndex > 1; rowIndex--) {
    var thisRow = rowIndex;
    var prevRow = rowIndex - 1;
    
    logDebug('Row: ' + thisRow);
    
    // Don't try to compare the header row
    if (prevRow == 1)
      break;
    
    var tran1 = transactions.getRange(thisRow, 1, 1, Columns).getValues()[0];
    var tran2 = transactions.getRange(prevRow, 1, 1, Columns).getValues()[0];
    var duplicateTransaction = duplicateTransactions(tran1, tran2);
    
    if (duplicateTransaction) {
      callback(thisRow, prevRow);
    };
  };
};
  
function removeDuplicates() {
  var spreadsheet  = SpreadsheetApp.getActiveSpreadsheet();
  var transactions = spreadsheet.getSheetByName('transactions');
  
  eachDuplicate(transactions, function(thisRow, prevRow) {
    var tran1Note = transactions.getRange(thisRow, Columns+1, 1, 1).getValue();
    var tran2Note = transactions.getRange(prevRow, Columns+1, 1, 1).getValue();
    if (tran1Note && tran2Note) {
      var msg = 'Rows ' + prevRow + ' and ' + thisRow + ' contain duplicate data but both contain a note.';
      msg += 'Remove one of the notes before running the script again.';
      logError(msg);
    } else if (tran1Note) {
      logInfo('Deleting row ' + prevRow);
      transactions.deleteRow(prevRow);
    } else {
      logInfo('Deleting row ' + thisRow);
      transactions.deleteRow(thisRow);
    };
  });
};

function showDuplicates() {
  var spreadsheet  = SpreadsheetApp.getActiveSpreadsheet();
  var transactions = spreadsheet.getSheetByName('transactions');
  
  clearBackgroundColours();
  
  eachDuplicate(transactions, function(thisRow, prevRow) {
    var lastColumn = transactions.getLastColumn();
    var tran1Note = transactions.getRange(thisRow, Columns+1, 1, 1).getValue();
    var tran2Note = transactions.getRange(prevRow, Columns+1, 1, 1).getValue();
    if (tran1Note && tran2Note) {
      var range = transactions.getRange(prevRow, 1, 1, lastColumn);
      range.setBackgroundRGB(255, 255, 0);
      var range = transactions.getRange(thisRow, 1, 1, lastColumn);
      range.setBackgroundRGB(255, 255, 0);
    } else if (tran1Note) {
      var range = transactions.getRange(prevRow, 1, 1, lastColumn);
      range.setBackgroundRGB(189, 230, 225);
    } else {
      var range = transactions.getRange(thisRow, 1, 1, lastColumn);
      range.setBackgroundRGB(189, 230, 225);
    };
  });
};
​