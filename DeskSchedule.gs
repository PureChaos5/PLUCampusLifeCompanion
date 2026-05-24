/**
 *  RA Desk Scheduler - 3 Phases x 4 Loops Architecture
 * 
 **/

function sDDEFG4() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const config = readConfig2(ss);
    const ras = readRAData2(ss, config);
    
    if (ras.length === 0) {
      SpreadsheetApp.getUi().alert('No RAs found.');
      return;
    }
    
    clearSchedules2(ss, config.buildings);
    const schedules = {};
    config.buildings.forEach(b => schedules[b] = createEmpty2());

    // PHASE 1: Non-Priority RAs in their OWN Non-Priority Buildings
    Logger.log('=== PHASE 1: Non-Pri RAs in Own Buildings ===');
    const nonPriBuildings = config.buildings.filter(b => !config.priorityBuildings.includes(b));
    const nonPriRAs = ras.filter(r => !r.isPriority);
    runAllLoops4(1, config, nonPriRAs, schedules, nonPriBuildings);

    // PHASE 2: Priority RAs in their OWN Priority Buildings  
    Logger.log('=== PHASE 2: Pri RAs in Own Buildings ===');
    const priRAs = ras.filter(r => r.isPriority);
    const priBuildings = config.priorityBuildings;
    runAllLoops4(2, config, priRAs, schedules, priBuildings);

    // PHASE 3: Non-Priority RAs Remaining Hours in Priority Buildings (sorted by RA count)
    Logger.log('=== PHASE 3: Non-Pri RAs Remaining in Pri Buildings ===');
    const priBuildingsSorted = [...priBuildings]
      .map(b => ({ name: b, count: ras.filter(r => r.building === b).length }))
      .sort((a, b) => a.count - b.count)
      .map(obj => obj.name);
    runAllLoops4(3, config, nonPriRAs, schedules, priBuildingsSorted);

    //optimizeBlockPreferences(config, ras, schedules);

    writeSchedules2(ss, schedules, config);
    applyPriorityFormatting(ss, config);
    
    let report = 'Scheduling complete!\n\nRA Hours:\n';
    ras.forEach(p => {
      const status = p.scheduledHours >= config.totalHoursPerWeek ? '✓' : '⚠';
      report += `${status} ${p.name}: ${p.scheduledHours}/${config.totalHoursPerWeek} hrs\n`;
    });
    SpreadsheetApp.getUi().alert(report);

  } catch (error) {
    Logger.log(error.stack);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Runs all 4 loops for a given phase
 */
function runAllLoops4(phase, config, rasSubset, schedules, buildingList) {
  
  // --- NEW: 4-Hour Priority Pre-Pass (Two-Step) ---
  // Condition: Max RAs > 1 AND (High Split hours OR High Daily hours)
  const highSplit = config.nonPriorityRAOwnBuildingHours === null || config.nonPriorityRAOwnBuildingHours >= 4;
  const highDaily = config.maxNonPriorityHoursPerDay === null || config.maxNonPriorityHoursPerDay >= 4;
  
  if (config.maxRAsOnDesk > 1 && (highSplit || highDaily)) {
    Logger.log('  >> Pre-Pass: Priority Scheduling for 4-Hour RAs (Strict Mode)...');
    const ras4Hour = rasSubset.filter(r => r.blockPreference === '4 Hours per Shift');
    
    if (ras4Hour.length > 0) {
      // Step 1: Preferred Hours Only (Loop 2 Mode)
      // isStrictPass = true (rejects partial shifts)
      Logger.log('    >> Step 1: Preferred Hours');
      runSchedulingLoop4(phase, 2, config, ras4Hour, schedules, buildingList, true);

      // Step 2: Any Available Hours (Loop 3 Mode)
      // isStrictPass = true (rejects partial shifts)
      Logger.log('    >> Step 2: Available Hours');
      runSchedulingLoop4(phase, 3, config, ras4Hour, schedules, buildingList, true);
    }
  }
  // ------------------------------------------------

  for (let loopNum = 1; loopNum <= 4; loopNum++) {
    Logger.log(`  Loop ${loopNum} starting...`);
    
    // Capture hours before loop
    const hoursBefore = {};
    rasSubset.forEach(ra => hoursBefore[ra.name] = ra.scheduledHours);
    
    // Run the loop once for all RAs (isStrictPass = false)
    runSchedulingLoop4(phase, loopNum, config, rasSubset, schedules, buildingList, false);
    
    // Report changes
    let assignmentsMade = 0;
    rasSubset.forEach(ra => {
      if (ra.scheduledHours > hoursBefore[ra.name]) {
        assignmentsMade++;
        Logger.log(`    ${ra.name}: ${hoursBefore[ra.name]} -> ${ra.scheduledHours} hrs`);
      }
    });
    Logger.log(`  Loop ${loopNum} complete: ${assignmentsMade} RAs scheduled`);
  }
}

/**
 * CORE SCHEDULING ENGINE - Single Loop Execution
 */
function runSchedulingLoop4(phase, loopNum, config, rasSubset, schedules, buildingList, isStrictPass = false) {
  buildingList.forEach(building => {
    const isPriBldg = config.priorityBuildings.includes(building);
    const isNonPriBldg = !isPriBldg;
    
    // Track daily hours for non-priority buildings
    let dailyHoursThisBuilding = [0, 0, 0, 0, 0];
    for (let d = 0; d < 5; d++) {
      for (let s = 0; s < 26; s++) {
        dailyHoursThisBuilding[d] += schedules[building][d][s].length * 0.5;
      }
    }

    for (let day = 0; day < 5; day++) {
      // Check daily limit for non-priority buildings
      const maxDaily = isNonPriBldg && config.maxNonPriorityHoursPerDay 
        ? config.maxNonPriorityHoursPerDay 
        : 999;
      
      for (let slot = config.deskTimeStart; slot < config.deskTimeEnd; ) {
        // Stop if daily limit reached for non-priority building
        if (isNonPriBldg && dailyHoursThisBuilding[day] >= maxDaily) break;
        
        // Skip if slot already filled
        if (schedules[building][day][slot].length >= config.maxRAsOnDesk) { 
          slot++; 
          continue; 
        }

        // Loop 1: Only priority time slots (Skipped during Pre-Pass because Pre-Pass uses loopNum 2 & 3)
        if (loopNum === 1 && !isSlotInPriorityTime4(building, day, slot, config)) { 
          slot++; 
          continue; 
        }

        let bestRA = null;
        let bestLen = 0;

        // Find eligible RAs for this slot
        rasSubset.forEach(ra => {
          // Phase-specific eligibility
          if (phase === 1 && ra.building !== building) return; 
          if (phase === 2 && ra.building !== building) return; 
          if (phase === 3 && ra.building === building) return; 
          
          // Hour limits
          if (ra.scheduledHours >= config.totalHoursPerWeek) return;
          
          // Non-pri RA own building hour limit
          if (phase === 1 && config.nonPriorityRAOwnBuildingHours !== null) {
            if (ra.ownBuildingHours >= config.nonPriorityRAOwnBuildingHours) return;
          }
          
          // One shift per day - ALWAYS enforced
          const hasShiftToday = ra.assignments.some(a => a.day === day);
          if (hasShiftToday) return;

          // Availability based on loop
          const avail = ra.availability[day][slot];
          if (avail === 'unavailable') return;
          
          // --- AVAILABILITY CHECK ---
          // Loop 2 enforces 'preferred' only. Loop 3 allows 'available' or 'preferred'.
          if (loopNum === 2 && avail !== 'preferred') return;

          // Determine max allowed shift length
          const prefBlock = getMaxShift2(ra.blockPreference);
          const needed = (config.totalHoursPerWeek - ra.scheduledHours) * 2;
          
          let maxAllowed;
          if (loopNum < 4) {
            maxAllowed = Math.min(prefBlock, needed); 
          } else {
            maxAllowed = needed; 
          }

          // Cap for non-pri RA in own building
          if (phase === 1 && config.nonPriorityRAOwnBuildingHours !== null) {
            const ownRemaining = (config.nonPriorityRAOwnBuildingHours - ra.ownBuildingHours) * 2;
            maxAllowed = Math.min(maxAllowed, ownRemaining);
          }

          // Find longest contiguous block
          let possibleLen = 0;
          for (let len = 2; len <= maxAllowed; len += 2) {
            if (slot + len > config.deskTimeEnd) break;
            
            // Loop 1: stay within priority time
            if (loopNum === 1 && !isSlotInPriorityTime4(building, day, slot + len - 1, config)) break;

            let canFit = true;
            for (let i = 0; i < len; i++) {
              const checkSlot = slot + i;
              
              // Slot availability
              if (schedules[building][day][checkSlot].length >= config.maxRAsOnDesk || 
                  ra.availability[day][checkSlot] === 'unavailable') {
                canFit = false;
                break;
              }
              
              // Not working elsewhere at this time
              const workingElsewhere = ra.assignments.some(a => 
                a.day === day && a.startSlot <= checkSlot && (a.startSlot + a.length) > checkSlot
              );
              if (workingElsewhere) {
                canFit = false;
                break;
              }
            }
            
            if (canFit) possibleLen = len;
            else break;
          }

          // --- LOGIC: Avoid 1 Hour Shifts ---
          if (config.maxRAsOnDesk > 1 && possibleLen === 2) {
            const prefersOneHour = ra.blockPreference === '1 Hour per shift';
            const finishingWeekly = (config.totalHoursPerWeek - ra.scheduledHours) * 2 <= 2;
            
            let finishingSplit = false;
            if (phase === 1 && config.nonPriorityRAOwnBuildingHours !== null) {
               const remainingSplit = (config.nonPriorityRAOwnBuildingHours - ra.ownBuildingHours) * 2;
               if (remainingSplit <= 2) finishingSplit = true;
            }

            if (!prefersOneHour && !finishingWeekly && !finishingSplit) {
               possibleLen = 0; 
            }
          }

          // --- LOGIC: Strict 4-Hour Check ---
          // If we are in the Pre-Pass (isStrictPass=true), we reject anything less than 4 hours (8 slots)
          // UNLESS the RA physically cannot work more (needs < 4 hours to finish weekly quota).
          if (isStrictPass && ra.blockPreference === '4 Hours per Shift') {
             const prefSlots = 8;
             if (possibleLen < prefSlots && needed > possibleLen) {
                possibleLen = 0; 
             }
          }
          // ----------------------------------------

          // Select RA with longest valid block (prioritize own building)
          if (possibleLen >= 2) {
            if (!bestRA || possibleLen > bestLen || 
                (possibleLen === bestLen && ra.building === building && bestRA.building !== building)) {
              bestRA = ra;
              bestLen = possibleLen;
            }
          }
        });

        // Assign the best RA found
        if (bestRA) {
          for (let i = 0; i < bestLen; i++) {
            schedules[building][day][slot + i].push(bestRA.name);
          }
          const hours = bestLen * 0.5;
          bestRA.scheduledHours += hours;
          dailyHoursThisBuilding[day] += hours;
          
          if (bestRA.building === building) {
            bestRA.ownBuildingHours += hours;
          }
          
          bestRA.assignments.push({ building, day, startSlot: slot, length: bestLen });
          slot += bestLen;
        } else {
          slot++;
        }
      }
    }
  });
}

/**
 * Helper to compare two RAs based on standard rules (Length, then Building Match)
 */
function compareStandard(ra, len, currentBest, bestLen, currentBuilding) {
  if (len > bestLen) return true;
  if (len === bestLen) {
    // If lengths are equal, prefer the RA whose home building is this building
    if (ra.building === currentBuilding && currentBest.building !== currentBuilding) return true;
  }
  return false;
}

/**
 * Check if slot falls within any priority time for this building
 */
function isSlotInPriorityTime4(building, day, slot, config) {
  return config.priorityTimes.some(pt => {
    const buildingMatch = !pt.building || pt.building === building;
    return buildingMatch && slot >= pt.start && slot < pt.end;
  });
}

// --- DATA READING & PARSING ---

function readConfig2(ss) {
  const sheet = ss.getSheetByName('Info');
  const config = {
    buildings: parseMulti2(sheet.getRange('C15').getValue()),
    priorityBuildings: parseMulti2(sheet.getRange('C17').getValue()),
    deskTimeStart: parseTime2(sheet.getRange('C19').getValue()),
    deskTimeEnd: parseTime2(sheet.getRange('D19').getValue()),
    priorityTimes: [],
    maxRAsOnDesk: sheet.getRange('G16').getValue(),
    maxNonPriorityHoursPerDay: parseFloat(sheet.getRange('G19').getValue()) || null,
    nonPriorityRAOwnBuildingHours: parseFloat(sheet.getRange('G23').getValue()) || null,
    totalHoursPerWeek: parseFloat(sheet.getRange('G26').getValue())
  };

  [21, 23, 25].forEach(row => {
    const start = sheet.getRange(`C${row}`).getValue();
    const end = sheet.getRange(`D${row}`).getValue();
    const bldg = sheet.getRange(`E${row}`).getValue();
    if (start && end) {
      config.priorityTimes.push({ 
        start: parseTime2(start), 
        end: parseTime2(end), 
        building: bldg || null 
      });
    }
  });
  
  return config;
}

function readRAData2(ss, config) {
  const ras = [];
  for (let i = 1; i <= 13; i++) {
    const sheet = ss.getSheetByName(`RA${i}`);
    if (!sheet) continue;
    
    const nameCell = sheet.getRange('J3').getValue();
    if (!nameCell) continue;
    
    const name = String(nameCell).split(/[,;]/)[0].trim();
    const building = sheet.getRange('J6').getValue();
    
    ras.push({
      name: name,
      building: building,
      blockPreference: sheet.getRange('J9').getValue(),
      availability: readAvail2(sheet),
      isPriority: config.priorityBuildings.includes(building),
      scheduledHours: 0,
      ownBuildingHours: 0,
      assignments: []
    });
  }
  return ras;
}

function readAvail2(sheet) {
  const grid = [];
  const backgrounds = sheet.getRange('D3:H28').getBackgrounds();
  
  for (let day = 0; day < 5; day++) {
    grid[day] = [];
    for (let slot = 0; slot < 26; slot++) {
      const bg = backgrounds[slot][day].toLowerCase();
      if (bg === '#000000') {
        grid[day][slot] = 'unavailable';
      } else if (bg === '#ffffff') {
        grid[day][slot] = 'preferred';
      } else {
        grid[day][slot] = 'available'; // grey/gray
      }
    }
  }
  return grid;
}

function clearSchedules2(ss, buildings) {
  for (let i = 1; i <= 3; i++) {
    const sheet = ss.getSheetByName(`Desk Schedule ${i}`);
    if (sheet) {
      sheet.getRange('D3:H28').clearContent().setBackground('white');
      sheet.getRange('J19').clearContent();
    }
  }
}

function writeSchedules2(ss, schedules, config) {
  const buildingNames = Object.keys(schedules);
  
  for (let i = 0; i < Math.min(3, buildingNames.length); i++) {
    const bName = buildingNames[i];
    const sheet = ss.getSheetByName(`Desk Schedule ${i + 1}`);
    if (!sheet) continue;
    
    sheet.getRange('J19').setValue(bName);
    
    const cols = ['D', 'E', 'F', 'G', 'H'];
    const output = [];
    
    for (let slot = 0; slot < 26; slot++) {
      const row = [];
      for (let day = 0; day < 5; day++) {
        row.push(schedules[bName][day][slot].join(', '));
      }
      output.push(row);
    }
    
    sheet.getRange('D3:H28').setValues(output);
    
    // Black out times outside desk hours
    for (let slot = 0; slot < 26; slot++) {
      if (slot < config.deskTimeStart || slot >= config.deskTimeEnd) {
        sheet.getRange(3 + slot, 4, 1, 5).setBackground('#d9d9d9');
      }
    }
  }
}

function applyPriorityFormatting(ss, config) {
  Logger.log('--- COLORING START ---');

  // 1. Check if the priority array exists and has items
  if (!config.priorityTimes || config.priorityTimes.length === 0) {
    Logger.log('No priority times found in config.');
    return;
  }

  const COLOR_PRIORITY = '#d9ead3'; // Light Green

  // 2. Loop through every Priority Time block defined in the Info sheet
  config.priorityTimes.forEach((pt, index) => {
    Logger.log(`Block ${index + 1}: Slots ${pt.start} to ${pt.end} (Building: ${pt.building || 'All'})`);

    // Basic validation
    if (typeof pt.start !== 'number' || typeof pt.end !== 'number' || pt.start >= pt.end) {
      return; 
    }

    const startRow = 3 + pt.start; // Row 3 corresponds to Slot 0
    const numRows = pt.end - pt.start;

    // 3. Apply to the 3 Schedule Sheets
    for (let i = 1; i <= 3; i++) {
      const sheet = ss.getSheetByName(`Desk Schedule ${i}`);
      if (!sheet) continue;

      // Check which building is on this sheet (written in J19 by the previous function)
      const currentBuilding = sheet.getRange('J19').getValue();

      // Apply Green if:
      // a) The rule applies to ALL buildings (pt.building is null/empty)
      // b) OR The rule matches this specific building
      if (!pt.building || pt.building === currentBuilding) {
        sheet.getRange(startRow, 4, numRows, 5).setBackground(COLOR_PRIORITY);
        Logger.log(`  -> Applied Green to ${sheet.getName()} (Rows ${startRow}-${startRow + numRows - 1})`);
      }
    }
  });

  Logger.log('--- COLORING COMPLETE ---');
}

function parseMulti2(value) {
  if (!value) return [];
  return String(value).split(/[,\n]/).map(s => s.trim()).filter(s => s);
}

function parseTime2(timeValue) {
  if (!timeValue || !(timeValue instanceof Date)) return 0;
  const hours = timeValue.getHours();
  const minutes = timeValue.getMinutes();
  return (hours - 7) * 2 + (minutes >= 30 ? 1 : 0);
}

function getMaxShift2(preference) {
  const map = {
    '1 Hour per shift': 2,
    '2 Hours per Shift': 4,
    '3 Hour Shift': 6,
    '4 Hours per Shift': 8,
    'No Preference': 8
  };
  return map[preference] || 8;
}

function createEmpty2() {
  const schedule = [];
  for (let day = 0; day < 5; day++) {
    schedule[day] = [];
    for (let slot = 0; slot < 26; slot++) {
      schedule[day][slot] = [];
    }
  }
  return schedule;
}