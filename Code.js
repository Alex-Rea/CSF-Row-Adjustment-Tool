const SHEET_ID = "14G38aqZ69FdbIHYwCwxIu0jq21Qw_Uo4-yB_2tOemkY";
const EXCLUDED_TABS = ["StockTracker", "Directory", "Sender", "All Products Selector", "Walmart #1584 (Rainbow / Spring Mountain)", "Walmart #1838 (Rainbow / Cheyenne)", "Walmart #2838 (Marks / Sunset)", "Walmart #3350 (Boulder / Nellis)", "Walmart #3355 (Lamb / Charleston)", "Walmart #3356 (Eastern / Warm Springs)", "Walmart #3728 (Lake Mead / Rancho)", "Walmart #4356 (Near Rainbow / 215)", "Walmart #4557 (Tropicana / Mcleod)", "Walmart #5269 (Silverado / Bermuda)"];
const TEMP_SHEET_NAME = '_RowMoverTemp';
const ZONE_SIZE = 50;
const MAX_ZONES = 10;

function findAvailableZone(sheet) {
  for (let z = 0; z < MAX_ZONES; z++) {
    const flagCell = sheet.getRange(ZONE_SIZE * z + 1, 1); // assume flag marker
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

function findAvailableZone(sheet) {
  for (let z = 0; z < MAX_ZONES; z++) {
    const flagCell = sheet.getRange(ZONE_SIZE * z + 1, 1); // assume flag marker
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

function getSheetData(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const names = sheet.getRange(1, 3, lastRow).getValues(); // Col C
  const bgColors = sheet.getRange(1, 3, lastRow).getBackgrounds(); // Col C
  const fontColors = sheet.getRange(1, 3, lastRow).getFontColors(); // Col C

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
  if (!sheet || !selectedIndices.length) return { success: false };

  selectedIndices.sort((a, b) => a - b);
  const totalCols = sheet.getLastColumn();
  const numRows = selectedIndices.length;
  const originalMin = selectedIndices[0];
  const originalMax = selectedIndices[selectedIndices.length - 1];

  // === STEP 1: Prepare temp sheet and zone ===
  let tempSheet = ss.getSheetByName(TEMP_SHEET_NAME);
  if (!tempSheet) tempSheet = ss.insertSheet(TEMP_SHEET_NAME);
  else tempSheet.clear();

  const zoneIndex = findAvailableZone(tempSheet);
  const zoneStartRow = zoneIndex * ZONE_SIZE + 2;

  // === STEP 2: Backup rows to temp zone ===
  for (let i = 0; i < numRows; i++) {
    sheet.getRange(selectedIndices[i], 1, 1, totalCols)
      .copyTo(tempSheet.getRange(zoneStartRow + i, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  }

  // === STEP 3: Determine insertIndex ===
  let insertIndex;
  if (newPosition !== null) {
    insertIndex = Math.max(1, newPosition);
    // Because we use insertRowsBefore(), insert one row after
    insertIndex += 1;
  } else if (direction === "up") {
    insertIndex = Math.max(1, originalMin - 1);
  } else if (direction === "down") {
    insertIndex = originalMax + 2;
  }

  // === STEP 4: Delete original rows ===
  for (let i = numRows - 1; i >= 0; i--) {
    sheet.deleteRow(selectedIndices[i]);
  }

  // === STEP 5: Adjust insertIndex after deletion ===
  // ONLY adjust if target was after the original position
  if (direction === "down" && insertIndex > originalMax) {
    insertIndex -= numRows;
  } else if (newPosition !== null && newPosition > originalMax) {
    insertIndex -= numRows;
  }

  // === STEP 6: Ensure sheet has enough rows ===
  const requiredRows = insertIndex + numRows - 1;
  const currentRows = sheet.getMaxRows();
  if (requiredRows > currentRows) {
    sheet.insertRowsAfter(currentRows, requiredRows - currentRows);
  }

  // === STEP 7: Make space and reinsert rows ===
  sheet.insertRowsBefore(insertIndex, numRows);

  for (let i = 0; i < numRows; i++) {
    const target = sheet.getRange(insertIndex + i, 1, 1, totalCols);
    const source = tempSheet.getRange(zoneStartRow + i, 1, 1, totalCols);
    source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  }

  // === STEP 8: Clean up ===
  tempSheet.getRange(zoneStartRow, 1, numRows, totalCols).clearContent();
  releaseZone(tempSheet, zoneIndex);

  const movedNames = [];
  for (let i = 0; i < numRows; i++) {
    movedNames.push(sheet.getRange(insertIndex + i, 3).getValue()); // Col C
  }

  return {
    success: true,
    movedNames: movedNames
  };
}
