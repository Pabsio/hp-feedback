// ============================================================
// APPS SCRIPT — pega esto en Extensions > Apps Script del Sheet
// Luego: Deploy > New deployment > Web app > Anyone
// Copia la URL y pégala en SCRIPT_URL del HTML
// ============================================================

const SHEET_NAME = "Feedback";

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const rows = data.map(r => ({
    timestamp: r[0],
    tool: r[1],
    issues: r[2] ? r[2].split("||") : [],
    rating: r[3],
    comment: r[4],
    name: r[5],
    newIssue: r[6] || ""
  }));

  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Timestamp", "Tool", "Feedback Items", "Rating", "Comment", "Name", "New Feedback"
    ]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
  }

  const body = JSON.parse(e.postData.contents);

  sheet.appendRow([
    new Date().toISOString(),
    body.tool,
    (body.issues || []).join("||"),
    body.rating || "",
    body.comment || "",
    body.name || "Anónimo",
    body.newIssue || ""
  ]);

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
