// A2A bridge doorbell — bound Apps Script for the "A2A - Results" sheet.
// Install: in the sheet, Extensions > Apps Script, paste this into Code.gs,
// save, then Triggers > + Add Trigger:
//   function: onBridgeChange | source: From spreadsheet | type: On change
// Replace <YOUR_EMAIL> with the Google account your Gemini Spark watches.
//
// History: a business.call routing upgrade (2026-11-03) briefly added a
// read-only kind scan that flagged business.call rows in the alert body
// for a Phone Call Agent. RETIRED 2026-11-05: Muse gained its own native
// business calling, so business calls no longer route through this bridge
// at all, and the doorbell is back to its original form below — it only
// sends the alert email; it never writes to the sheet, never creates or
// deletes triggers, and never polls.

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
