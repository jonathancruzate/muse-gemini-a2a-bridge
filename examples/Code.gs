// A2A bridge doorbell — bound Apps Script for the "A2A - Results" sheet.
// Install: in the sheet, Extensions > Apps Script, paste this into Code.gs,
// save, then Triggers > + Add Trigger:
//   function: onBridgeChange | source: From spreadsheet | type: On change
// Replace <YOUR_EMAIL> with the Google account your Gemini Spark watches.
//
// UPGRADE 2026-11-03 (business.call routing): this script stays a doorbell
// only — it never writes to the sheet, never creates/deletes triggers, and
// never polls. The one change from the original: it reads the pending
// Requests rows (columns A-F: appended_at | request_id | token | kind |
// params_json | context — there is NO status column) and names any
// business.call rows in the alert email body, so the receiving schedule
// routes those rows to the Phone Call Agent instead of treating them as
// read/query-only requests. Subject and recipient are unchanged, so the
// existing email-triggered schedule filter (subject:"A2A-Event") still
// matches exactly.

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

  // Read-only scan: count pending business.call rows (kind = column D).
  var businessCalls = 0;
  var kinds = [];
  try {
    var values = requestsSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (var i = 0; i < values.length; i++) {
      var kind = values[i][3];
      if (kind) kinds.push(kind);
      if (kind === "business.call") businessCalls++;
    }
  } catch (err) {
    // Scan is best-effort — the alert must still fire without it.
  }

  var body = "A new pending request was appended to the Requests tab in A2A - Results.";
  if (kinds.length) body += "\nPending kinds: " + kinds.join(", ") + ".";
  if (businessCalls > 0) {
    body += "\n" + businessCalls + " business.call request(s) pending: route to @Phone Call Agent per the business.call branch of the processing schedule. Do not treat business.call rows as read/query-only.";
  }

  MailApp.sendEmail({
    to: "<YOUR_EMAIL>",
    subject: "A2A-Event",
    body: body
  });
}

// Optional MANUAL re-ring for business.call rows already sitting in
// Requests (e.g. after a schedule-prompt upgrade). Run by hand from the
// Apps Script editor only. It sends the same A2A-Event alert and writes
// nothing. NEVER wire this to a time-driven trigger and NEVER delete
// existing triggers to add one — the On change trigger above is the
// standing wake-up path.
function checkBusinessCallRequestsManual() {
  onBridgeChange(null);
}
