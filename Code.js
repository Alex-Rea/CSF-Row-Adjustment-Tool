const SHEET_ID = "14G38aqZ69FdbIHYwCwxIu0jq21Qw_Uo4-yB_2tOemkY";
const EXCLUDED_TABS = ["StockTracker", "Directory", "Sender", "All Products Selector"];
const TEMP_SHEET_NAME = '_RowMoverTemp';
const ZONE_SIZE = 50;
const MAX_ZONES = 10;

function findAvailableZone(sheet) {
  for (let z = 0; z < MAX_ZONES; z++) {
    const flagCell = sheet.getRange(ZONE_SIZE * z + 1, 1); // flag marker
    if (!flagCell.getValue()) {
      flagCell.setValue('IN USE');
      return z;
    }
  }
  throw new Error("No copy zones available — try again shortly.");
}

function releaseZone(sheet, zoneIndex) {
  const flagCell = sheet.getRange(ZONE_SIZE * zoneIndex + 1, 1);
  flagCell.clearContent();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Row Adjuster Tool")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheetsList() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheets()
    .filter(sheet => !EXCLUDED_TABS.includes(sheet.getName()))
    .map(sheet => sheet.getName());
}

function getSheetData(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return [];

  // Col C styling and names
  const names = sheet.getRange(1, 3, lastRow, 1).getValues();
  const bgColors = sheet.getRange(1, 3, lastRow, 1).getBackgrounds();
  const fontColors = sheet.getRange(1, 3, lastRow, 1).getFontColors();

  return names.map((row, i) => ({
    index: i + 1,
    name: row[0],
    bgColor: bgColors[i][0],
    fontColor: fontColors[i][0]
  }));
}

function moveSelectedRows(sheetName, selectedIndices, direction, newPosition = null) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || !selectedIndices || !selectedIndices.length) {
    return { success: false, message: "No sheet or rows selected." };
  }

  // Normalize & basics
  selectedIndices.sort((a, b) => a - b);
  const totalCols = sheet.getLastColumn();
  const numRows   = selectedIndices.length;
  const originalMin = selectedIndices[0];
  const originalMax = selectedIndices[selectedIndices.length - 1];

  // === Temp sheet & zone ===
  let tempSheet = ss.getSheetByName(TEMP_SHEET_NAME);
  if (!tempSheet) tempSheet = ss.insertSheet(TEMP_SHEET_NAME);
  else tempSheet.clear();

  const zoneIndex   = findAvailableZone(tempSheet);
  const zoneStartRow = zoneIndex * ZONE_SIZE + 2;

  // Backup selected rows
  for (let i = 0; i < numRows; i++) {
    sheet.getRange(selectedIndices[i], 1, 1, totalCols)
         .copyTo(tempSheet.getRange(zoneStartRow + i, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  }

  // Decide target strategy
  const isGoTo = (newPosition !== null && newPosition !== undefined);
  let targetStart = null; // final start row AFTER we make space

  if (isGoTo) {
    targetStart = Math.max(1, Number(newPosition) || 1);
  } else if (direction === "up") {
    targetStart = Math.max(1, originalMin - 1);        // one row up
  } else if (direction === "down") {
    // We'll handle "down" specially after deletion via insertRowsAfter(originalMin)
    // to avoid off-by-one confusion. targetStart will become originalMin + 1.
  } else {
    releaseZone(tempSheet, zoneIndex);
    return { success: false, message: "No move direction/position provided." };
  }

  // Delete originals (bottom → top)
  for (let i = numRows - 1; i >= 0; i--) {
    sheet.deleteRow(selectedIndices[i]);
  }

  // Ensure capacity (different for "down" vs others)
  if (!isGoTo && direction === "down") {
    // We want block to start at originalMin + 1 (post-deletion index).
    const needed = originalMin + numRows; // final end row
    const maxRows = sheet.getMaxRows();
    if (needed > maxRows) {
      sheet.insertRowsAfter(maxRows, needed - maxRows);
    }

    // Make space AFTER the row now at originalMin.
    sheet.insertRowsAfter(originalMin, numRows);
    targetStart = originalMin + 1; // space we just made

  } else {
    // Go-to and Up use insertRowsBefore at targetStart
    const requiredRows = (targetStart + numRows - 1);
    const currentRows = sheet.getMaxRows();
    if (requiredRows > currentRows) {
      sheet.insertRowsAfter(currentRows, requiredRows - currentRows);
    }
    sheet.insertRowsBefore(targetStart, numRows);
  }

  // Paste rows into the new space
  for (let i = 0; i < numRows; i++) {
    const target = sheet.getRange(targetStart + i, 1, 1, totalCols);
    const source = tempSheet.getRange(zoneStartRow + i, 1, 1, totalCols);
    source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  }

  // Cleanup
  tempSheet.getRange(zoneStartRow, 1, numRows, totalCols).clearContent();
  releaseZone(tempSheet, zoneIndex);

  // Report names (Col C) for re-highlighting
  const movedNames = [];
  for (let i = 0; i < numRows; i++) {
    movedNames.push(sheet.getRange(targetStart + i, 3).getValue());
  }

  return { success: true, movedNames };
}
