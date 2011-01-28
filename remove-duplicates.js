var Columns = 6;
var Debug   = true;

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
  
function removeDuplicates() {
  var spreadsheet  = SpreadsheetApp.getActiveSpreadsheet();
  var transactions = spreadsheet.getSheetByName('transactions');
  
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
      Logger.log('Deleting row ' + thisRow);
      transactions.deleteRow(thisRow);
    };
  };
};
​