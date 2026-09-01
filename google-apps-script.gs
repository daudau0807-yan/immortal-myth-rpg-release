const SHEET_NAME = "玩家期望";

function doPost(event) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const payload = JSON.parse(event.postData.contents || "{}");
    if (!payload.submissionId || !payload.answer) {
      return jsonResponse({ ok: false, error: "missing_fields" });
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["提交編號", "匿名玩家編號", "玩家期望", "提交時間", "來源", "章節"]);
      sheet.setFrozenRows(1);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
      if (existingIds.includes(String(payload.submissionId))) {
        return jsonResponse({ ok: true, duplicate: true });
      }
    }

    sheet.appendRow([
      String(payload.submissionId),
      String(payload.anonymousId || ""),
      String(payload.answer).slice(0, 5000),
      String(payload.submittedAt || new Date().toISOString()),
      String(payload.source || ""),
      String(payload.story || "")
    ]);

    return jsonResponse({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
