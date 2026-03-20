// ============================================================
// No Location Reported - Google Apps Script Backend v2
// Paste this in Google Apps Script and deploy as Web App
// (Execute as: Me, Access: Anyone)
//
// CHANGES FROM v1:
// - Added updateTaskStatus action (supports v5 app Task Status toggle)
// - Added "Task Status" column to setupSheet headers and seed data
// ============================================================

const SHEET_NAME = "No Location Reported";

function doGet(e) {
  var action = e.parameter.action || "read";

  if (action === "read") {
    return readData();
  } else if (action === "updateComment") {
    return updateComment(e.parameter.stockNum, e.parameter.comment);
  } else if (action === "updateTaskStatus") {
    return updateTaskStatus(e.parameter.stockNum, e.parameter.taskStatus);
  } else if (action === "addRow") {
    return addRow(e);
  } else if (action === "clearAndReload") {
    return clearAndReload(e);
  }

  return jsonResponse({ status: "error", message: "Unknown action" });
}

function readData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ status: "error", message: "Sheet not found" });

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    rows.push(row);
  }

  return jsonResponse({ status: "ok", data: rows });
}

function updateComment(stockNum, comment) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ status: "error", message: "Sheet not found" });

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var stockCol = headers.indexOf("Stock Num");
  var commentCol = headers.indexOf("Comments");

  if (stockCol === -1 || commentCol === -1) {
    return jsonResponse({ status: "error", message: "Required columns not found" });
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][stockCol]) === String(stockNum)) {
      sheet.getRange(i + 1, commentCol + 1).setValue(comment);
      return jsonResponse({ status: "ok", message: "Comment updated" });
    }
  }

  return jsonResponse({ status: "error", message: "Stock number not found" });
}

// NEW in v2 - saves the Task Status value for a given stock number
function updateTaskStatus(stockNum, taskStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ status: "error", message: "Sheet not found" });

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var stockCol = headers.indexOf("Stock Num");
  var tsCol = headers.indexOf("Task Status");

  if (stockCol === -1) {
    return jsonResponse({ status: "error", message: "Stock Num column not found" });
  }
  if (tsCol === -1) {
    return jsonResponse({ status: "error", message: "Task Status column not found - please add it to the sheet manually" });
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][stockCol]) === String(stockNum)) {
      sheet.getRange(i + 1, tsCol + 1).setValue(taskStatus);
      return jsonResponse({ status: "ok", message: "Task status updated" });
    }
  }

  return jsonResponse({ status: "error", message: "Stock number not found" });
}

function addRow(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ status: "error", message: "Sheet not found" });

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRow = [];

  for (var i = 0; i < headers.length; i++) {
    var key = headers[i];
    newRow.push(e.parameter[key] || "");
  }

  sheet.appendRow(newRow);
  return jsonResponse({ status: "ok", message: "Row added" });
}

// Called by n8n or automation: wipes data rows and reloads fresh batch
// Pass rows as JSON string in parameter "rows"
function clearAndReload(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ status: "error", message: "Sheet not found" });

  var rowsJson = e.parameter.rows;
  if (!rowsJson) return jsonResponse({ status: "error", message: "No rows provided" });

  var rows;
  try {
    rows = JSON.parse(rowsJson);
  } catch (err) {
    return jsonResponse({ status: "error", message: "Invalid JSON in rows parameter" });
  }

  // Keep header row, delete all data rows
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  // Append new rows
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < rows.length; i++) {
    var newRow = [];
    for (var j = 0; j < headers.length; j++) {
      newRow.push(rows[i][headers[j]] || "");
    }
    sheet.appendRow(newRow);
  }

  return jsonResponse({ status: "ok", message: rows.length + " rows loaded", count: rows.length });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// ONE-TIME SETUP: Run this function manually to create the
// sheet with headers and seed data from the 3/12/26 report
// NOTE: Now includes Task Status column (added in v2)
// ============================================================
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  } else {
    sheet.clearContents();
  }

  var headers = [
    "Stock Num", "Year", "Vehicle Description", "Last Scan Date",
    "Last Scan Location", "Zone Description", "Days in Color Zone",
    "VIN", "Product Status", "Days In Status", "Comments", "Task Status"
  ];

  sheet.appendRow(headers);

  // Seed data from the 03/12/26 report image
  var seedData = [
    ["28375174", "2015", "FORD ESCAPE 4D SPORT UTILIT SE", "3/11/26", "WZA52", "WHOLESALE ZONE A", "1", "1FMCU9GX4FUC90572", "WHOLESALE RESERV", "1", "", ""],
    ["28442662", "2016", "SUBARU OUTBACK 4D WAGON 3.6R LIMITED", "3/11/26", "WZE61", "WHOLESALE ZONE E", "2", "4S4BSENC4G3239770", "VALUMAX RECON", "1", "", ""],
    ["28554432", "2017", "AUDI Q3 4D SPORT UTILIT PREMIUM", "3/11/26", "WZE64", "WHOLESALE ZONE E", "3", "WA1ECCFS9HR017310", "WHOLESALE RESERV", "1", "", ""]
  ];

  for (var i = 0; i < seedData.length; i++) {
    sheet.appendRow(seedData[i]);
  }

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#1a1a2e");
  headerRange.setFontColor("#ffffff");

  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);

  SpreadsheetApp.getUi().alert("Sheet created successfully with " + seedData.length + " rows!");
}
