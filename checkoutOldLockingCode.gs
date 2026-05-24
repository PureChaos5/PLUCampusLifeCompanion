function protectSheetExceptWhiteCells() {
  const allowedEditors = ["marie.tucker@plu.edu", "brianna.brum@plu.edu", "cavichouser@plu.edu", "yclemons@plu.edu", "gek@plu.edu", "mettler@plu.edu", "mnash@plu.edu", "olie.santossiqueira@plu.edu", "rmitchell@plu.edu", "swhite@plu.edu", "seth.setu@plu.edu", "titus.malaga@plu.edu", "gongzl@plu.edu"];
  let targetSheetNames;
  let currentSheets;
  const now = new Date();
  const hour = now.getHours(); 
  const day = now.getDate();

  if ((day === 17 && hour >= 17) || (day === 18 && hour < 17)) {
    targetSheetNames = ["MON - 5/19", "TUES - 5/20", "WED - 5/21", "THURS - 5/22", "FRI - 5/23"];
    currentSheets = ["SUN - 5/18"];
  } else if ((day === 18 && hour >= 17) || (day === 19 && hour < 17)) {
    targetSheetNames = ["TUES - 5/20", "WED - 5/21", "THURS - 5/22", "FRI - 5/23"];
    currentSheets = ["SUN - 5/18", "MON - 5/19"];
  } else if ((day === 19 && hour >= 17) || (day === 20 && hour < 17)) {
    targetSheetNames = ["WED - 5/21", "THURS - 5/22", "FRI - 5/23"];
    currentSheets = ["SUN - 5/18", "MON - 5/19", "TUES - 5/20"];
  } else if ((day === 20 && hour >= 17) || (day === 21 && hour < 17)) {
    targetSheetNames = ["THURS - 5/22", "FRI - 5/23"];
    currentSheets = ["SUN - 5/18", "MON - 5/19", "TUES - 5/20", "WED - 5/21"];
  } else if ((day === 21 && hour >= 17) || (day === 22 && hour < 17)) {
    targetSheetNames = ["FRI - 5/23"];
    currentSheets = ["SUN - 5/18", "MON - 5/19", "TUES - 5/20", "WED - 5/21", "THURS - 5/22"];
  } else if (day >= 22 && hour >= 17) {
    targetSheetNames = [];
    currentSheets = ["SUN - 5/18", "MON - 5/19", "TUES - 5/20", "WED - 5/21", "THURS - 5/22", "FRI - 5/23"];
  } else {
    targetSheetNames = ["SUN - 5/18", "MON - 5/19", "TUES - 5/20", "WED - 5/21", "THURS - 5/22", "FRI - 5/23"];
    currentSheets = [];
  }

  const START_ROW = 5;
  const END_ROW = 44;
  const START_COL = 3;  // Column C
  const END_COL = 26;   // Column Z

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  currentSheets.forEach(sheetName => {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return;

  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);

  // Check if "Close the sheet" protection already exists
  const hasCloseProtection = protections.some(p => p.getDescription() === "Close the sheet");

  // Remove all other protections
  protections.forEach(p => {
    if (p.getDescription() !== "Close the sheet") {
      p.remove();
    }
  });

  // If "Close the sheet" doesn't exist, create it
  if (!hasCloseProtection) {
    const sheetProtection = sheet.protect().setDescription("Close the sheet");
    sheetProtection.removeEditors(sheetProtection.getEditors());
    sheetProtection.addEditors(allowedEditors);
  }
  });

  targetSheetNames.forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;

    // Remove existing sheet protections (optional but safe)
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(p => {
      p.remove();
    });

    // Apply full sheet protection
    const sheetProtection = sheet.protect().setDescription("Protect entire sheet except unprotected white blocks");
    sheetProtection.removeEditors(sheetProtection.getEditors());
    sheetProtection.addEditors(allowedEditors);

    const numRows = END_ROW - START_ROW + 1;
    const numCols = END_COL - START_COL + 1;

    const bgColors = sheet.getRange(START_ROW, START_COL, numRows, numCols).getBackgrounds();
    const unprotectedRanges = [];

    // Loop through paired columns (C&D, E&F, ...)
    for (let colOffset = 0; colOffset < numCols; colOffset += 2) {
      const col1 = colOffset;
      const col2 = colOffset + 1;
      let startRow = null;

      for (let row = 0; row <= numRows; row++) {
        const isWhitePair =
          row < numRows &&
          bgColors[row][col1]?.toLowerCase() !== "#000000" &&
          bgColors[row][col2]?.toLowerCase() !== "#000000";

        if (isWhitePair && startRow === null) {
          startRow = row;
        } else if (!isWhitePair && startRow !== null) {
          // Add unprotected block
          const range = sheet.getRange(
            START_ROW + startRow,
            START_COL + col1,
            row - startRow,
            2
          );
          unprotectedRanges.push(range);
          startRow = null;
        }
      }
    }

    // Apply unprotected ranges
    try {
      sheetProtection.setUnprotectedRanges(unprotectedRanges);
    } catch (e) {
      console.error(`Failed to set unprotected ranges on ${sheetName}: ${e}`);
    }

  });
}
