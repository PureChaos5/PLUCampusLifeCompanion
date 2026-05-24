function dayLockHourly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');
  const schedSheet = ss.getSheetByName('SCHEDULE HERE');

  if (!dataSheet || !schedSheet) return;

  // Check if global lock is active. If not, stop immediately.
  if (dataSheet.getRange('P22').getValue() !== true) {
    return;
  }

  // 1. Set our 24-hour lock horizon
  const now = new Date();
  const lockHorizon = new Date(now.getTime() + 24 * 60 * 60 * 1000); 
  const me = Session.getEffectiveUser();

  // Read the Schedule grid to calculate exact shift times (matching the calendar script)
  const schedLastRow = Math.max(schedSheet.getLastRow(), 46);
  const schedGrid = schedSheet.getRange(1, 1, schedLastRow, 24).getValues(); 

  // Get all current protections so we can update them without making messy duplicates
  const protections = schedSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  // 2. Loop through the 11 columns 
  // These map to SCHEDULE HERE Cols C, E, G, I, K, M, O, Q, S, U, W
  for (let c = 0; c < 11; c++) {
    const schedColStart = (c * 2) + 3; // Index 3 is Col C
    const dateVal = schedGrid[3][schedColStart - 1]; // Row 4 (index 3) contains the date

    if (!dateVal || isNaN(new Date(dateVal).getTime())) continue;

    let lockUntilRow = -1;

    // 3. Check every shift slot in this column (Rows 7 to 46 / Array indices 6 to 45)
    for (let r = 6; r < 46; r++) {
      const timeStr = schedGrid[r][1]; // Column B has the times
      const shiftTime = combineDateAndTime(dateVal, timeStr);

      // If the shift time is valid and is happening within the next 24 hours (or in the past)
      if (shiftTime && shiftTime <= lockHorizon) {
        lockUntilRow = r + 1; // Convert array index back to Sheet row number (r=6 -> Row 7)
      } else {
        // Since time is sequential top-to-bottom, once we hit a future shift, we can stop checking this column
        break; 
      }
    }

    // 4. Apply the lock if we found shifts that need it
    if (lockUntilRow >= 7) {
      const numRowsToLock = lockUntilRow - 7 + 1;
      const rangeToLock = schedSheet.getRange(7, schedColStart, numRowsToLock, 2);
      const lockDesc = `Rolling Lock: Col ${schedColStart}`;

      // Check if we've already locked this exact block to prevent duplicate ranges
      let alreadyLocked = false;
      let obsoleteProtections = [];

      for (let i = 0; i < protections.length; i++) {
        const p = protections[i];
        if (p.getDescription() === lockDesc) {
          if (p.getRange().getNumRows() === numRowsToLock) {
            alreadyLocked = true; // The lock is perfectly up to date
          } else {
            obsoleteProtections.push(p); // It's an older, smaller lock. We will delete it.
          }
        }
      }

      if (!alreadyLocked) {
        // Remove the old, smaller locks for this column so it stays clean
        obsoleteProtections.forEach(p => p.remove());

        // Create the new, expanded lock block
        const p = rangeToLock.protect().setDescription(lockDesc);
        p.addEditor(me);
        p.removeEditors(p.getEditors());
        if (p.canDomainEdit()) p.setDomainEdit(false);
      }
    }
  }
}

/**
 * Helper to merge a Date object and a Time string (e.g., "8:00 AM")
 * (You may already have this in your script file, but it's required for dayLock to work!)
 */
function combineDateAndTime(dateObj, timeStr) {
  if (!timeStr) return null;
  const date = new Date(dateObj);
  const time = new Date(timeStr); 
  
  if (isNaN(time.getTime())) {
    const match = timeStr.match(/(\d+):(\d+)\s*(AM|PM)/i);
    if (!match) return null;
    let hours = parseInt(match[1]);
    const mins = parseInt(match[2]);
    const ampm = match[3].toUpperCase();
    if (ampm === "PM" && hours < 12) hours += 12;
    if (ampm === "AM" && hours === 12) hours = 0;
    date.setHours(hours, mins, 0, 0);
  } else {
    date.setHours(time.getHours(), time.getMinutes(), 0, 0);
  }
  return date;
}