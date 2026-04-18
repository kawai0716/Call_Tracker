function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    const spreadsheetId = String(payload.spreadsheetId || "").trim();
    const sheetName = String(payload.sheetName || "").trim();
    const year = Number(payload.year);
    const month = Number(payload.month);
    const day = Number(payload.day);
    const values = payload.values || {};

    if (!spreadsheetId) {
      return jsonResponse({ ok: false, error: "Missing spreadsheetId" });
    }

    if (!sheetName) {
      return jsonResponse({ ok: false, error: "Missing sheetName" });
    }

    if (!year || !month || !day) {
      return jsonResponse({ ok: false, error: "Missing date parts" });
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      return jsonResponse({ ok: false, error: "Sheet not found" });
    }

    const dataRange = sheet.getDataRange();
    const displayValues = dataRange.getDisplayValues();
    const monthBlockTitle = buildMonthBlockTitle(year, month);
    const monthBlock = findMonthBlock(displayValues, monthBlockTitle);

    if (!monthBlock) {
      return jsonResponse({ ok: false, error: "Month block not found", monthBlockTitle: monthBlockTitle });
    }

    const headerInfo = findActualColumnsInBlock(displayValues, monthBlock.startRow, monthBlock.endRow);

    if (
      !headerInfo.columns.calls ||
      !headerInfo.columns.secondCalls ||
      !headerInfo.columns.connections ||
      !headerInfo.columns.sampleSent ||
      !headerInfo.columns.introductions
    ) {
      return jsonResponse({ ok: false, error: "Required columns not found", monthBlockTitle: monthBlockTitle });
    }

    const rowIndex = findDayRowInBlock(displayValues, headerInfo.headerRow, monthBlock.endRow, day);

    if (!rowIndex) {
      return jsonResponse({ ok: false, error: "Day row not found", monthBlockTitle: monthBlockTitle, day: day });
    }

    writeIntegerValue(sheet, rowIndex, headerInfo.columns.calls, values.calls);
    writeIntegerValue(sheet, rowIndex, headerInfo.columns.secondCalls, values.secondCalls);
    writeIntegerValue(sheet, rowIndex, headerInfo.columns.connections, values.connections);
    writeIntegerValue(sheet, rowIndex, headerInfo.columns.sampleSent, values.sampleSent);
    writeIntegerValue(sheet, rowIndex, headerInfo.columns.introductions, values.introductions);

    return jsonResponse({
      ok: true,
      monthBlockTitle: monthBlockTitle,
      row: rowIndex,
      updated: {
        calls: Number(values.calls || 0),
        secondCalls: Number(values.secondCalls || 0),
        connections: Number(values.connections || 0),
        sampleSent: Number(values.sampleSent || 0),
        introductions: Number(values.introductions || 0),
      },
    });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error) });
  }
}

function buildMonthBlockTitle(year, month) {
  return year + "年" + month + "月｜Q" + getQuarterForMonth(month);
}

function getQuarterForMonth(month) {
  if (month >= 4 && month <= 6) {
    return 1;
  }

  if (month >= 7 && month <= 9) {
    return 2;
  }

  if (month >= 10 && month <= 12) {
    return 3;
  }

  return 4;
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\s+/g, "")
    .replace(/\|/g, "｜")
    .trim();
}

function normalizeHeader(value) {
  return normalizeText(value).replace(/[()（）]/g, "");
}

function rowIncludesValue(rowValues, target) {
  const normalizedTarget = normalizeText(target);

  for (let col = 0; col < rowValues.length; col += 1) {
    if (normalizeText(rowValues[col]) === normalizedTarget) {
      return true;
    }
  }

  return false;
}

function isMonthBlockTitleRow(rowValues) {
  for (let col = 0; col < rowValues.length; col += 1) {
    if (/^\d{4}年\d{1,2}月｜Q[1-4]$/.test(normalizeText(rowValues[col]))) {
      return true;
    }
  }

  return false;
}

function findMonthBlock(displayValues, monthBlockTitle) {
  for (let row = 0; row < displayValues.length; row += 1) {
    if (!rowIncludesValue(displayValues[row], monthBlockTitle)) {
      continue;
    }

    let endRow = displayValues.length - 1;

    for (let nextRow = row + 1; nextRow < displayValues.length; nextRow += 1) {
      if (isMonthBlockTitleRow(displayValues[nextRow])) {
        endRow = nextRow - 1;
        break;
      }
    }

    return { startRow: row, endRow: endRow };
  }

  return null;
}

function matchesHeader(value, aliases) {
  const normalized = normalizeHeader(value);
  return aliases.some(function (alias) {
    return normalizeHeader(alias) === normalized;
  });
}

function findActualColumnsInBlock(displayValues, startRow, endRow) {
  const aliases = {
    calls: ["コール数実"],
    secondCalls: ["2回目架電数実"],
    connections: ["担当者接続数実"],
    sampleSent: ["サンプル送付数実"],
    introductions: ["導入数新規成約実", "導入数実"],
  };

  for (let row = startRow; row <= endRow; row += 1) {
    const columns = {
      calls: 0,
      secondCalls: 0,
      connections: 0,
      sampleSent: 0,
      introductions: 0,
    };

    for (let col = 0; col < displayValues[row].length; col += 1) {
      const cell = displayValues[row][col];

      if (!columns.calls && matchesHeader(cell, aliases.calls)) {
        columns.calls = col + 1;
      }

      if (!columns.secondCalls && matchesHeader(cell, aliases.secondCalls)) {
        columns.secondCalls = col + 1;
      }

      if (!columns.connections && matchesHeader(cell, aliases.connections)) {
        columns.connections = col + 1;
      }

      if (!columns.sampleSent && matchesHeader(cell, aliases.sampleSent)) {
        columns.sampleSent = col + 1;
      }

      if (!columns.introductions && matchesHeader(cell, aliases.introductions)) {
        columns.introductions = col + 1;
      }
    }

    if (columns.calls && columns.secondCalls && columns.connections && columns.sampleSent && columns.introductions) {
      return { headerRow: row + 1, columns: columns };
    }
  }

  return {
    headerRow: 0,
    columns: {
      calls: 0,
      secondCalls: 0,
      connections: 0,
      sampleSent: 0,
      introductions: 0,
    },
  };
}

function findDayRowInBlock(displayValues, headerRow, endRow, day) {
  const target = String(day);

  for (let row = headerRow; row <= endRow; row += 1) {
    const firstColumn = String(displayValues[row - 1][0] || "").trim();

    if (firstColumn === target) {
      return row;
    }

    if (firstColumn === "月計" || firstColumn === "達成率") {
      break;
    }
  }

  return 0;
}

function writeIntegerValue(sheet, rowIndex, columnIndex, value) {
  sheet
    .getRange(rowIndex, columnIndex)
    .setNumberFormat("0")
    .setValue(Number(value || 0));
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
