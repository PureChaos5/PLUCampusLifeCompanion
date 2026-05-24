/**
 *  Desk Menu Code
 * 
 * onOpen() - makes the Campus Life Companion menu
 * showHelpAlert() & openLinkInNewTab() - open the manual
 * setupDeskScheduleTrigger() - sets up hourly triggers for lock and calander sync
 * removeAllTriggers() - removes previously set up triggers
 * unlockSheets() & lockSheets() - to lock all but the student scheduling portion for only CD access
 * clearSchedule() - clears the RA schedule
 * 
 **/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Campus Life Companion')
      .addItem('Generate RA Schedule', 'scheduleRACheckouts')
      .addItem('Shuffle RA Schedule', 'randomizeShiftColumns')
      .addItem('Clear RA Schedule', 'clearSchedule')
      .addSeparator()
      .addItem('Lock RA Schedule', 'lockSheets')
      .addItem('Unlock RA Schedule', 'unlockSheets')
      .addSeparator()
      .addItem('Manually Sync Calander', 'fullCheckoutsToCalendar')
      .addItem('Set Up Triggers', 'setupDeskScheduleTrigger')
      .addItem('Remove All Triggers', 'removeAllTriggers')
      .addSeparator()
      .addItem('Campus Life Companion Manual', 'showHelpAlert') 
      .addToUi();
}

function showHelpAlert() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Show the Yes/No prompt
  const response = ui.alert(
    'Open Manual', 
    'Would you like to open the Campus Life Companion Manual in a new tab?', 
    ui.ButtonSet.YES_NO
  );

  // 2. If they click Yes, run the HTML workaround to open the link
  if (response === ui.Button.YES) {
    // REPLACE THIS LINK with your actual Google Doc URL
    const docUrl = 'https://docs.google.com/document/d/1saQrB2yQXdNqI3cAnF16vhgJgT4qhD_NvUOIEHVtcZc/edit?usp=sharing'; 
    openLinkInNewTab(docUrl);
  }
}

// Helper function to force the browser to open a URL
function openLinkInNewTab(url) {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <script>
        // Open the URL in a new tab
        window.open('${url}', '_blank');
        // Instantly close this little popup box
        google.script.host.close();
      </script>
      <body>
        <p>Opening manual...</p>
      </body>
    </html>
  `)
  .setWidth(250)
  .setHeight(50);

  SpreadsheetApp.getUi().showModalDialog(html, 'Redirecting...');
}

function setupDeskScheduleTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');
  
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('Error: "Data" sheet not found.');
    return;
  }

  const currentUserEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  const lockedByEmail = dataSheet.getRange('P23').getValue();

    // Check Authorization
    if (lockedByEmail !== "" && currentUserEmail !== lockedByEmail) {
      SpreadsheetApp.getUi().alert(`Access Denied: These sheets were locked by ${lockedByEmail}. Only they can set the trigger.`);
      return; 
    }

  var functionToTrigger1 = 'dayLockHourly';
  var functionToTrigger2 = 'fullCheckoutsToCalendar';
  
  // 1. Find and delete any existing triggers for BOTH functions
  var allTriggers = ScriptApp.getProjectTriggers();
  
  for (var i = 0; i < allTriggers.length; i++) {
    var currentHandler = allTriggers[i].getHandlerFunction();
    
    // Check if the trigger matches EITHER of your functions
    if (currentHandler === functionToTrigger1 || currentHandler === functionToTrigger2) {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
  
  // 2. Now build the fresh triggers safely
  ScriptApp.newTrigger(functionToTrigger1)
    .timeBased()
    .everyHours(1)
    .create();

  ScriptApp.newTrigger(functionToTrigger2)
    .timeBased()
    .everyHours(1)
    .create();
    
  SpreadsheetApp.getUi().alert('Old triggers cleared. New hourly triggers set successfully.');
}

function removeAllTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');
  
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('Error: "Data" sheet not found.');
    return;
  }

  const currentUserEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  const lockedByEmail = dataSheet.getRange('P23').getValue();

    // Check Authorization
    if (lockedByEmail !== "" && currentUserEmail !== lockedByEmail) {
      SpreadsheetApp.getUi().alert(`Access Denied: These sheets were locked by ${lockedByEmail}. Only they can set the trigger.`);
      return; 
    }
  
  // 1. Find and delete any existing triggers for BOTH functions
  var allTriggers = ScriptApp.getProjectTriggers();
  
  for (var i = 0; i < allTriggers.length; i++) {
    var currentHandler = allTriggers[i].getHandlerFunction();
    
    // Check if the trigger matches EITHER of your functions
    if (currentHandler === functionToTrigger1 || currentHandler === functionToTrigger2) {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }

  SpreadsheetApp.getUi().alert('Old triggers cleared.');
}

function randomizeShiftColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');
  
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('Error: "Data" sheet not found.');
    return;
  }

  const lockedByEmail = dataSheet.getRange('P23').getValue();

  // Check Authorization
  if (lockedByEmail !== "") {
    SpreadsheetApp.getUi().alert(`Access Denied: You must unlock the sheet before you can make a new schedule.`);
    return; 
  }

  // Target the schedule grid on the Data sheet: C4:M43 (40 rows, 11 columns)
  const schedRange = dataSheet.getRange(4, 3, 40, 11);
  const values = schedRange.getValues();
  const backgrounds = schedRange.getBackgrounds();

  // Define the column indices for days with multiple shifts.
  // Col C (Monday) is index 0, so we skip it.
  const dayGroups = [
    [1, 2],       // Tuesday (Cols D, E)
    [3, 4],       // Wednesday (Cols F, G)
    [5, 6, 7],    // Thursday (Cols H, I, J)
    [8, 9, 10]    // Friday (Cols K, L, M)
  ];

  // Loop through every 15-minute timeslot (40 rows)
  for (let r = 0; r < 40; r++) {
    
    // For each timeslot, loop through our multi-column days
    for (let group of dayGroups) {
      
      let shufflableIndices = [];
      let shufflableData = [];

      // 1. Identify only the cells that are NOT blacked out
      group.forEach(colIdx => {
        const bg = backgrounds[r][colIdx].toLowerCase();
        if (bg !== '#000000' && bg !== 'black') {
          shufflableIndices.push(colIdx);
          shufflableData.push({
            val: values[r][colIdx],
            bg: backgrounds[r][colIdx]
          });
        }
      });

      // 2. Randomize the order of ONLY the valid shifts (Fisher-Yates Shuffle)
      for (let i = shufflableData.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [shufflableData[i], shufflableData[j]] = [shufflableData[j], shufflableData[i]];
      }

      // 3. Put the shuffled values and colors back into the specific non-black columns
      shufflableIndices.forEach((colIdx, index) => {
        values[r][colIdx] = shufflableData[index].val;
        backgrounds[r][colIdx] = shufflableData[index].bg;
      });
    }
  }

  // Write the freshly shuffled names and colors back to the sheet all at once
  schedRange.setValues(values);
  schedRange.setBackgrounds(backgrounds);
  
  SpreadsheetApp.getUi().alert('Success: Active shift orders for Tuesday through Friday have been randomized!');
}

function unlockSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');
  const schedSheet = ss.getSheetByName('SCHEDULE HERE');
  
  if (!dataSheet || !schedSheet) {
    SpreadsheetApp.getUi().alert('Error: Make sure both "Data" and "SCHEDULE HERE" sheets exist.');
    return;
  }

  try {
    const currentUserEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    const lockedByEmail = dataSheet.getRange('P23').getValue();

    // Check Authorization
    if (lockedByEmail !== "" && currentUserEmail !== lockedByEmail) {
      SpreadsheetApp.getUi().alert(`Access Denied: These sheets were locked by ${lockedByEmail}. Only they can unlock them.`);
      return; 
    }

    // --- UNLOCK DATA SHEET ---
    const existingDataProtections = dataSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    existingDataProtections.forEach(p => p.remove());

    dataSheet.getRange('P22').setValue(false);
    dataSheet.getRange('P23').clearContent();

    // --- UNLOCK RA SHEETS (RA1 to RA13) ---
    let unlockedRACount = 0;
    for (let i = 1; i <= 13; i++) {
      const raSheet = ss.getSheetByName(`RA${i}`);
      if (raSheet) {
        const existingRAProtections = raSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        existingRAProtections.forEach(p => p.remove());
        unlockedRACount++;
      }
    }

    // --- UNLOCK AND RE-COLOR 'SCHEDULE HERE' SHEET ---
    // 1. Remove all specific range protections (This clears the unstaffed shifts AND the structural boundaries)
    const existingSchedProtections = schedSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    existingSchedProtections.forEach(p => p.remove());

    // 2. Reset the grid (C7:X46) to the updated day colors
    // UPDATED: Thursday (index 3) is now #c9daf8
    const dayColors = ['#fce5cd', '#fff2cc', '#d9ead3', '#c9daf8', '#d9d2e9'];
    const colToDay = [0, 1, 1, 2, 2, 3, 3, 3, 4, 4, 4]; 
    
    const schedRange = schedSheet.getRange(7, 3, 40, 22);
    const schedBackgrounds = schedRange.getBackgrounds();

    for (let c = 0; c < 11; c++) {
      for (let r = 0; r < 40; r++) {
        let color = dayColors[colToDay[c]];
        schedBackgrounds[r][c * 2] = color;
        schedBackgrounds[r][c * 2 + 1] = color;
      }
    }
    
    schedRange.setBackgrounds(schedBackgrounds);

    SpreadsheetApp.getUi().alert(`Success: Data sheet, ${unlockedRACount} RA sheets, and "SCHEDULE HERE" unlocked.\nP22 is set to FALSE.`);

  } catch (error) {
    Logger.log(error.stack);
    SpreadsheetApp.getUi().alert('Error unlocking sheets: ' + error.message);
  }
}


function lockSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');
  const schedSheet = ss.getSheetByName('SCHEDULE HERE');
  
  if (!dataSheet || !schedSheet) {
    SpreadsheetApp.getUi().alert('Error: Make sure both "Data" and "SCHEDULE HERE" sheets exist.');
    return;
  }

  try {
    const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();

    // 1. Set P22 to TRUE and log the user in P23 on the Data sheet
    dataSheet.getRange('P22').setValue(true);
    dataSheet.getRange('P23').setValue(userEmail);

    const me = Session.getEffectiveUser();

    // --- LOCK DATA SHEET ---
    const existingDataProtections = dataSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    existingDataProtections.forEach(p => p.remove());
    
    const dataProtection = dataSheet.protect().setDescription('Locked Data Sheet');
    dataProtection.addEditor(me);
    dataProtection.removeEditors(dataProtection.getEditors());
    if (dataProtection.canDomainEdit()) dataProtection.setDomainEdit(false);

    // --- LOCK RA SHEETS (RA1 to RA13) ---
    let lockedRACount = 0;
    for (let i = 1; i <= 13; i++) {
      const raSheet = ss.getSheetByName(`RA${i}`);
      if (raSheet) {
        const existingRAProtections = raSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        existingRAProtections.forEach(p => p.remove());

        const raProtection = raSheet.protect().setDescription(`Locked RA${i} Sheet`);
        raProtection.addEditor(me);
        raProtection.removeEditors(raProtection.getEditors());
        if (raProtection.canDomainEdit()) raProtection.setDomainEdit(false);
        
        lockedRACount++;
      }
    }

    // --- SYNC & LOCK 'SCHEDULE HERE' SHEET ---
    // Clear any leftover range protections to prevent duplicate locks stacking up
    const existingSchedProtections = schedSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    existingSchedProtections.forEach(p => p.remove());

    // NEW: Lock the static structural boundaries (Rows 1:6, 47, Cols A:B, Y)
    const staticRanges = ['1:6', '47:47', 'A:B', 'Y:Y'];
    staticRanges.forEach(rangeStr => {
      const boundaryRange = schedSheet.getRange(rangeStr);
      const boundaryProtection = boundaryRange.protect().setDescription('Structural Boundaries');
      boundaryProtection.addEditor(me);
      boundaryProtection.removeEditors(boundaryProtection.getEditors());
      if (boundaryProtection.canDomainEdit()) boundaryProtection.setDomainEdit(false);
    });

    // Map and lock the black unfilled shifts
    const dataBackgrounds = dataSheet.getRange(4, 3, 40, 11).getBackgrounds();
    const schedRange = schedSheet.getRange(7, 3, 40, 22); // C7:X46
    const schedBackgrounds = schedRange.getBackgrounds();

    for (let c = 0; c < 11; c++) {
      let startRow = null;
      
      for (let r = 0; r <= 40; r++) {
        let isBlack = (r < 40) && (dataBackgrounds[r][c] === '#000000');
        
        if (isBlack) {
          schedBackgrounds[r][c * 2] = '#000000';
          schedBackgrounds[r][c * 2 + 1] = '#000000';
          
          if (startRow === null) startRow = r; 
        } else {
          if (startRow !== null) {
            let chunkRange = schedSheet.getRange(startRow + 7, (c * 2) + 3, r - startRow, 2);
            let p = chunkRange.protect().setDescription('Unstaffed Shift');
            p.addEditor(me);
            p.removeEditors(p.getEditors());
            if (p.canDomainEdit()) p.setDomainEdit(false);
            
            startRow = null; 
          }
        }
      }
    }
    
    schedRange.setBackgrounds(schedBackgrounds);

    SpreadsheetApp.getUi().alert(`Success: Data sheet, ${lockedRACount} RA sheets, structural boundaries, and unstaffed slots on "SCHEDULE HERE" are locked.\nP22 set to TRUE.\nRun by: ${userEmail}`);

  } catch (error) {
    Logger.log(error.stack);
    SpreadsheetApp.getUi().alert('Error locking sheets: ' + error.message);
  }
}


function clearSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');
  
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('Error: "Data" sheet not found.');
    return;
  }

  const lockedByEmail = dataSheet.getRange('P23').getValue();

  // Check Authorization
  if (lockedByEmail !== "") {
    SpreadsheetApp.getUi().alert(`Access Denied: You must unlock the sheet before you can clear it.`);
    return; 
  }

  // Capacities: Mon(1), Tue(2), Wed(2), Thu(3), Fri(3)
  const capacities = [1, 2, 2, 3, 3];
  const dayColors = ['#fce5cd', '#fff2cc', '#d9ead3', '#c9daf8', '#d9d2e9'];
  
  const outValues = [];
  const outColors = [];

  // Build the clean 40x11 grid
  for (let slot = 0; slot < 40; slot++) {
    const rowVals = [];
    const rowColors = [];
    
    for (let day = 0; day < 5; day++) {
      for (let c = 0; c < capacities[day]; c++) {
        rowVals.push(""); // Empty string to clear the name
        rowColors.push(dayColors[day]); // Reset to the day's background color
      }
    }
    outValues.push(rowVals);
    outColors.push(rowColors);
  }

  // Write to Data Sheet (Row 4, Col 3 = C4. Dimensions: 40 rows, 11 total columns)
  const targetRange = dataSheet.getRange(4, 3, 40, 11);
  targetRange.setValues(outValues);
  targetRange.setBackgrounds(outColors);
  
  SpreadsheetApp.getUi().alert('Schedule cleared and colors reset!');
}