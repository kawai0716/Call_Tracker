function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    const spreadsheetId = String(payload.spreadsheetId || "").trim();
    const sheetName = String(payload.sheetName || "").trim();
    const day = Number(payload.day);
    const values = payload.values || {};

    if (!spreadsheetId) {
      return jsonResponse({ ok: false, error: "Missing spreadsheetId" });
    }

    if (!sheetName) {
      return jsonResponse({ ok: false, error: "Missing sheetName" });
    }

    if (!day) {
      return jsonResponse({ ok: false, error: "Missing day" });
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      return jsonResponse({ ok: false, error: "Sheet not found" });
    }

    const dataRange = sheet.getDataRange();
    const displayValues = dataRange.getDisplayValues();

    const columnMap = findActualColumns(displayValues);

    if (!columnMap.calls || !columnMap.connections || !columnMap.sampleSent || !columnMap.introductions) {
      return jsonResponse({ ok: false, error: "Required columns not found" });
    }

    const rowIndex = findDayRow(displayValues, day);

    if (!rowIndex) {
      return jsonResponse({ ok: false, error: "Day row not found" });
    }

    sheet.getRange(rowIndex, columnMap.calls).setValue(Number(values.calls || 0));
    sheet.getRange(rowIndex, columnMap.connections).setValue(Number(values.connections || 0));
    sheet.getRange(rowIndex, columnMap.sampleSent).setValue(Number(values.sampleSent || 0));
    sheet.getRange(rowIndex, columnMap.introductions).setValue(Number(values.introductions || 0));

    return jsonResponse({
      ok: true,
      row: rowIndex,
      updated: {
        calls: Number(values.calls || 0),
        connections: Number(values.connections || 0),
        sampleSent: Number(values.sampleSent || 0),
        introductions: Number(values.introductions || 0),
      },
    });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error) });
  }
}

function normalizeHeader(value) {
  return String(value || "")
    .replace(/\s+/g, "")
    .replace(/[()（）]/g, "")
    .trim();
}

function matchesHeader(value, aliases) {
  const normalized = normalizeHeader(value);
  return aliases.some((alias) => normalizeHeader(alias) === normalized);
}

function findActualColumns(displayValues) {
  const result = {
    calls: 0,
    connections: 0,
    sampleSent: 0,
    introductions: 0,
  };

  const aliases = {
    calls: ["コール数実"],
    connections: ["担当者接続数実"],
    sampleSent: ["サンプル送付数実"],
    introductions: ["導入数新規成約実", "導入数実"],
  };

  for (let row = 0; row < displayValues.length; row += 1) {
    for (let col = 0; col < displayValues[row].length; col += 1) {
      const cell = displayValues[row][col];

      if (!result.calls && matchesHeader(cell, aliases.calls)) {
        result.calls = col + 1;
      }

      if (!result.connections && matchesHeader(cell, aliases.connections)) {
        result.connections = col + 1;
      }

      if (!result.sampleSent && matchesHeader(cell, aliases.sampleSent)) {
        result.sampleSent = col + 1;
      }

      if (!result.introductions && matchesHeader(cell, aliases.introductions)) {
        result.introductions = col + 1;
      }
    }
  }

  return result;
}

function findDayRow(displayValues, day) {
  const target = String(day);

  for (let row = 0; row < displayValues.length; row += 1) {
    const firstColumn = String(displayValues[row][0] || "").trim();

    if (firstColumn === target) {
      return row + 1;
    }
  }

  return 0;
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
