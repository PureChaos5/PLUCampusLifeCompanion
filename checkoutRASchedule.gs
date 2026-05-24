function scheduleRACheckouts() {
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

  try {
    // 1. EXTRACT RA DATA
    const ras = [];
    for (let i = 1; i <= 13; i++) {
      const sheet = ss.getSheetByName(`RA${i}`);
      if (!sheet) continue;

      const nameCell = sheet.getRange('I3').getValue();
      const name = String(nameCell).trim();

      // Skip this sheet if I3 is blank OR if it matches your template's placeholder text
      if (!name || name === "" || name === "RA Name" || name === "First Last") {
        continue; 
      }

      // Read C3:G42 (40 timeslots, 5 days)
      const grid = [];
      const backgrounds = sheet.getRange('C3:G42').getBackgrounds();
      let totalOpenSlots = 0; 
      
      for (let day = 0; day < 5; day++) {
        grid[day] = [];
        for (let slot = 0; slot < 40; slot++) {
          const bg = backgrounds[slot][day].toLowerCase();
          if (bg === '#000000') {
            grid[day][slot] = 'unavailable';
          } else if (bg === '#ffffff') {
            grid[day][slot] = 'preferred';
            totalOpenSlots++; 
          } else {
            grid[day][slot] = 'available'; 
            totalOpenSlots++; 
          }
        }
      }

      ras.push({
        name: name,
        availability: grid,
        shiftCount: 0,
        totalOpen: totalOpenSlots 
      });
    }

    if (ras.length === 0) {
      SpreadsheetApp.getUi().alert('No RAs found. Check I3 names on RA sheets.');
      return;
    }

    // 2. CALCULATE CAPS & SPLIT RA POOLS (UPDATED FOR 3 TIERS)
    const totalShifts = 440;
    const maxShiftsPerRA = Math.ceil(totalShifts / ras.length);
    
    // Split into Highly Restricted (<100), Moderately Restricted (100-119), and Flexible (>=120)
    const highlyRestrictedRAs = ras.filter(ra => ra.totalOpen < 100);
    const moderatelyRestrictedRAs = ras.filter(ra => ra.totalOpen >= 100 && ra.totalOpen < 101);
    const flexibleRAs = ras.filter(ra => ra.totalOpen >= 101);

    Logger.log(`Total RAs: ${ras.length} | Max shifts capped at: ${maxShiftsPerRA}`);
    Logger.log(`Highly Restricted (<70): ${highlyRestrictedRAs.length} | Moderately Restricted (70-119): ${moderatelyRestrictedRAs.length} | Flexible (>=120): ${flexibleRAs.length}`);

    // 3. INITIALIZE SCHEDULE GRID
    const capacities = [1, 2, 2, 3, 3]; 
    const schedule = [];
    for (let day = 0; day < 5; day++) {
      schedule[day] = [];
      for (let slot = 0; slot < 40; slot++) {
        schedule[day][slot] = [];
      }
    }

    // 4. SCHEDULING ENGINE
    function assignShifts(preferredOnly, raPool) {
      if (raPool.length === 0) return; 
      
      let madeAssignment = true;
      
      while (madeAssignment) {
        madeAssignment = false;
        
        for (let i = raPool.length - 1; i > 0; i--) {
          const j = Math.floor(Math.random() * (i + 1));
          [raPool[i], raPool[j]] = [raPool[j], raPool[i]]; 
        }
        
        raPool.sort((a, b) => a.shiftCount - b.shiftCount);

        for (let ra of raPool) {
          if (ra.shiftCount >= maxShiftsPerRA) {
            continue; 
          }

          let validSlots = [];
          
          for (let day = 0; day < 5; day++) {
            for (let slot = 0; slot < 40; slot++) {
              if (schedule[day][slot].length < capacities[day] && !schedule[day][slot].includes(ra.name)) {
                let avail = ra.availability[day][slot];
                if (avail === 'preferred' || (!preferredOnly && avail === 'available')) {
                  validSlots.push({ day, slot });
                }
              }
            }
          }

          if (validSlots.length > 0) {
            let pick = validSlots[Math.floor(Math.random() * validSlots.length)];
            schedule[pick.day][pick.slot].push(ra.name);
            ra.shiftCount++;
            madeAssignment = true;
            break; 
          }
        }
      }
    }

    // PHASE 1: HIGHLY RESTRICTED RAs (< 70 slots)
    assignShifts(true, highlyRestrictedRAs);
    assignShifts(false, highlyRestrictedRAs);

    // PHASE 2: MODERATELY RESTRICTED RAs (70 - 119 slots)
    assignShifts(true, moderatelyRestrictedRAs);
    assignShifts(false, moderatelyRestrictedRAs);

    // PHASE 3: FLEXIBLE RAs (>= 120 slots)
    assignShifts(true, flexibleRAs);
    assignShifts(false, flexibleRAs);

    // 5. FORMAT OUTPUT FOR 'Data' SHEET
    const dayColors = ['#fce5cd', '#fff2cc', '#d9ead3', '#c9daf8', '#d9d2e9'];
    const outValues = [];
    const outColors = [];

    for (let slot = 0; slot < 40; slot++) {
      const rowVals = [];
      const rowColors = [];
      for (let day = 0; day < 5; day++) {
        const assignedRAs = schedule[day][slot];
        
        for (let c = 0; c < capacities[day]; c++) {
          if (c < assignedRAs.length) {
            rowVals.push(assignedRAs[c]);
            rowColors.push(dayColors[day]); 
          } else {
            rowVals.push("");
            rowColors.push('#000000'); 
          }
        }
      }
      outValues.push(rowVals);
      outColors.push(rowColors);
    }

    const targetRange = dataSheet.getRange(4, 3, 40, 11);
    targetRange.setValues(outValues);
    targetRange.setBackgrounds(outColors);

    // 6. GENERATE REPORT (SORTED BY LEAST AVAILABLE SHIFTS)
    ras.sort((a, b) => a.totalOpen - b.totalOpen); // Sort ascending by open slots
    
    let report = `Checkout Scheduling Complete!\nMax shift cap: ${maxShiftsPerRA}\n\nShift Counts (15-min blocks):\n`;
    ras.forEach(ra => {
      let tag = "";
      if (ra.totalOpen < 100) {
        tag = "[Highly Restricted]";
      } else if (ra.totalOpen < 120) {
        tag = "[Moderately Restricted]";
      } else {
        tag = "[Flexible]";
      }
      report += `${ra.name} ${tag}: ${ra.shiftCount} assigned / ${ra.totalOpen} avaiable\n`;
    });

    SpreadsheetApp.getUi().alert(report);

  } catch (error) {
    Logger.log(error.stack);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}