function formatDate(date) {
  const formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  return formattedDate;
}

function formatTimestamp(timestamp) {
  const formattedTimestamp = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'M/d/yyyy HH:mm:ss');
  return formattedTimestamp;
}

function logToSheet(message) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Logs");
    sheet.appendRow(["Timestamp", "Log Message"]);
  }
  sheet.appendRow([new Date(), message]); // Log the message with timestamp
}
