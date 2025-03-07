
// Utils.gs

function formatDate(date) {
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function formatTimestamp(timestamp) {
  return Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'M/d/yyyy HH:mm:ss');
}

function logToSheet(message) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Logs");
    sheet.appendRow(["Timestamp", "Log Message"]);
  }
  sheet.appendRow([new Date(), message]);
}