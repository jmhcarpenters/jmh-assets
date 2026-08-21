// DEV Apps Script configuration.
// The DEV code/deployment is isolated from production but intentionally uses the JMH data spreadsheet.
const JMH_SPREADSHEET_ID = '1-EfAlj16IpIwYn3pKUEZHD3eB2wCk-RTwaw40lPoCUs';

function getJmhSpreadsheet_() {
  return SpreadsheetApp.openById(JMH_SPREADSHEET_ID);
}
