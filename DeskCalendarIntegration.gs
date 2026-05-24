/**
 *  Calendar Code
 * 
 * calUploadDesk() - for adding to the calendar
 * calDeleteDesk() - for clearing the calendar
 * 
 **/

function calUploadDesk() {
  const ss = SpreadsheetApp.getActive();
  const info = ss.getSheetByName('Info');

  function norm(str) { return (str || '').toString().trim(); }
  const parseDate = s => { if (!s) return null; const d = new Date(s); return isNaN(d) ? null : d; };

  const startDate = parseDate(info.getRange('C29').getValue());
  const endDate = parseDate(info.getRange('E29').getValue());
  const scheduleMode = norm(info.getRange('E31').getValue());
  const calendarId = norm(info.getRange('C33').getValue());
  const calendarId2 = norm(info.getRange('C35').getValue());

  const buildingsRaw = norm(info.getRange('C15').getValue());
  const hasMultipleBuildings = buildingsRaw.split(/,|&|\n/).length > 1;
  
  if (!startDate || !endDate) {
    SpreadsheetApp.getUi().alert('Invalid Start or End Date on Info sheet.');
    return;
  }

  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    SpreadsheetApp.getUi().alert('Error: Primary Calendar not found.');
    return;
  }

  // Map for recurrence days
  const weekdayMap = [
    CalendarApp.Weekday.MONDAY,
    CalendarApp.Weekday.TUESDAY,
    CalendarApp.Weekday.WEDNESDAY,
    CalendarApp.Weekday.THURSDAY,
    CalendarApp.Weekday.FRIDAY
  ];

  // 1. LOAD RA EMAILS
  const raData = {};
  for (let i = 1; i <= 13; i++) {
    const sh = ss.getSheetByName('RA' + i);
    if (!sh) continue;
    let rawName = norm(sh.getRange('J3').getValue());
    if (!rawName) continue;
    let name = rawName.split(/\s+/)[0].toLowerCase();
    let email = norm(sh.getRange('J12').getValue()); 
    if (email) raData[name] = email;
  }

  // 2. DETERMINE WHICH SCHEDULES TO RUN
  let sheetsToRun = [];
  if (scheduleMode === "All Schedules") {
    sheetsToRun = ['Desk Schedule 1', 'Desk Schedule 2', 'Desk Schedule 3'];
  } else if (scheduleMode.startsWith("Desk Schedule")) {
    sheetsToRun = [scheduleMode];
  }

  let seriesCreated = 0;

  // 3. LOOP THROUGH SELECTED SHEETS
  sheetsToRun.forEach(sheetName => {
    const deskSheet = ss.getSheetByName(sheetName);
    if (!deskSheet) return;

    const buildingName = norm(deskSheet.getRange('J19').getValue());
    const gridValues = deskSheet.getRange('D3:H28').getValues();

    // 4. LOOP THROUGH THE 5 COLUMNS (Mon-Fri) ONCE
    for (let colIdx = 0; colIdx < 5; colIdx++) {
      let activeShifts = {}; 
      
      // Calculate the first occurrence of this specific weekday
      let firstOccurrence = new Date(startDate);
      let targetDay = colIdx + 1; // Mon=1, Tue=2...
      let currentDay = firstOccurrence.getDay();
      let daysToWait = (targetDay - currentDay + 7) % 7;
      firstOccurrence.setDate(firstOccurrence.getDate() + daysToWait);

      // Create the recurrence rule for this specific weekday
      let recurrence = CalendarApp.newRecurrence()
        .addWeeklyRule()
        .onlyOnWeekday(weekdayMap[colIdx])
        .until(endDate);

      // 5. LOOP THROUGH 30-MIN INCREMENTS IN THE COLUMN
      for (let r = 0; r <= 26; r++) { 
        let cellNames = [];
        if (r < 26) {
          let cellValue = norm(gridValues[r][colIdx]);
          if (cellValue) {
            cellNames = cellValue.split(',').map(n => norm(n).toLowerCase()).filter(n => n);
          }
        }

        // A. Shift End Logic
        for (let activeName in activeShifts) {
          if (!cellNames.includes(activeName)) {
            let startRow = activeShifts[activeName];
            let endRow = r; 

            let eventStart = new Date(firstOccurrence);
            eventStart.setHours(7 + Math.floor(startRow / 2), (startRow % 2) * 30, 0, 0);

            let eventEnd = new Date(firstOccurrence);
            eventEnd.setHours(7 + Math.floor(endRow / 2), (endRow % 2) * 30, 0, 0);

            let guests = [];
            if (raData[activeName]) guests.push(raData[activeName]);
            if (calendarId2) guests.push(calendarId2);

            let displayName = activeName.charAt(0).toUpperCase() + activeName.slice(1);
            let title = `Desk Shift: ${displayName}`;
            if (hasMultipleBuildings) {
                let bName = buildingName ? buildingName : sheetName.slice(-1);
                title += ` (${bName})`; 
            }

            // CREATE RECURRING SERIES
            try {
              calendar.createEventSeries(title, eventStart, eventEnd, recurrence, {
                guests: guests.join(','),
                sendInvites: false 
              });
              seriesCreated++;
            } catch (e) {
              Logger.log(`Error creating series: ${e.message}`);
            }

            delete activeShifts[activeName];
          }
        }

        // B. Shift Start Logic
        cellNames.forEach(name => {
          if (!(name in activeShifts)) {
            activeShifts[name] = r; 
          }
        });
      }
    }
  });

  SpreadsheetApp.getUi().alert(`Success! Created ${seriesCreated} recurring weekly shifts.`);
}


function calDeleteDesk() {
  const ss = SpreadsheetApp.getActive();
  const info = ss.getSheetByName('Info');

  function norm(str) { return (str || '').toString().trim(); }
  const parseDate = s => { if (!s) return null; const d = new Date(s); return isNaN(d) ? null : d; };

  const calendarId = norm(info.getRange('C33').getValue());
  const startDate = parseDate(info.getRange('C29').getValue());
  const endDate = parseDate(info.getRange('E29').getValue());
  
  if (!startDate || !endDate) {
    SpreadsheetApp.getUi().alert('Invalid start or end date in Info (C29/E29).');
    return;
  }

  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    SpreadsheetApp.getUi().alert('Calendar not found.');
    return;
  }

  let endSearch = new Date(endDate);
  endSearch.setDate(endSearch.getDate() + 1);
  
  let events = calendar.getEvents(startDate, endSearch);
  let deletedCount = 0;
  
  // Use a Set to keep track of Series IDs we have already deleted
  // This prevents the script from crashing when it hits the "next week" instance of a series it just removed.
  let deletedSeriesIds = new Set();

  events.forEach(ev => {
    let title = ev.getTitle();
    
    if (title.startsWith('Desk Shift: ')) {
      let series = ev.getEventSeries();
      
      if (series) {
        let sId = series.getId();
        if (!deletedSeriesIds.has(sId)) {
          series.deleteEventSeries();
          deletedSeriesIds.add(sId);
          deletedCount++;
        }
      } else {
        // Fallback for single events that aren't part of a series
        ev.deleteEvent();
        deletedCount++;
      }
    }
  });

  SpreadsheetApp.getUi().alert(`Success! Removed ${deletedCount} shift series from the calendar.`);
}