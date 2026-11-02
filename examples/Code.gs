// A2A bridge doorbell — bound Apps Script for the "A2A - Results" sheet.
// Install: in the sheet, Extensions > Apps Script, paste this into Code.gs,
// save, then Triggers > + Add Trigger:
//   function: onBridgeChange | source: From spreadsheet | type: On change
// Replace <YOUR_EMAIL> with the Google account your Gemini Spark watches.

function onBridgeChange(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestsSheet = ss.getSheetByName("Requests");
  if (!requestsSheet) return;

  // Check if there are data rows below the header row
  const lastRow = requestsSheet.getLastRow();
  if (lastRow <= 1) return;

  // 15-second debounce to avoid duplicate alerts during batch row inserts
  const cache = CacheService.getScriptCache();
  if (cache.get("RECENT_ALERT")) return;
  cache.put("RECENT_ALERT", "true", 15);

  MailApp.sendEmail({
    to: "<YOUR_EMAIL>",
    subject: "A2A-Event",
    body: "A new pending request was appended to the Requests tab in A2A - Results."
  });
}
