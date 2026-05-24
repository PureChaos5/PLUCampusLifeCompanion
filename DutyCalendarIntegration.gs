/**
 *  Calendar Code
 * 
 * calUploadDuty() - for adding to the calendar
 * calDeleteDuty() - for clearing the calendar
 * 
 **/

function calUploadDuty() {
  const ss = SpreadsheetApp.getActive();
  const info = ss.getSheetByName('Info');
  const out = ss.getSheetByName('Duty Schedule');

  function norm(str) { return (str || '').toString().trim(); }
  const parseDate = s => { if (!s) return null; const d = new Date(s); return isNaN(d) ? null : d; };

  const calendarId = norm(info.getRange('C26').getValue());
  const calendarId2 = norm(info.getRange('C28').getValue());
  const startDate = parseDate(info.getRange('C17').getValue());
  const endDate = parseDate(info.getRange('F17').getValue());
  
  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    Logger.log('Calendar not found. Check the ID and permissions.');
    SpreadsheetApp.getUi().alert('Error: Primary Calendar not found. Check ID and permissions.');
    return;
  }

  // 1. LOAD RA EMAILS
  // Assuming the Name is in C2 and the Email is in C3. Adjust 'C3' if needed.
  const raData = {};
  for (let i = 1; i <= 13; i++) {
    const sh = ss.getSheetByName('RA' + i);
    if (!sh) continue;
    let rawName = norm(sh.getRange('C2').getValue());
    if (!rawName) continue;
    let name = rawName.split(/\s+/)[0].toLowerCase();
    let email = norm(sh.getRange('G2').getValue()); 
    if (email) raData[name] = email;
  }

  // 2. READ THE GRID
  const weekdayToCol = {'Sunday':2,'Monday':3,'Tuesday':4,'Wednesday':5,'Thursday':6,'Friday':8,'Saturday':9};
  const weekdayNames = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const gridValues = out.getRange('B2:I55').getValues();
  
  const anchorSunday = new Date(startDate); 
  anchorSunday.setDate(startDate.getDate() - startDate.getDay());

  let eventsCreated = 0;

  // 3. LOOP THROUGH DATES AND CREATE EVENTS
  for (let dt = new Date(startDate); dt <= endDate; dt.setDate(dt.getDate() + 1)) {
    const dayName = weekdayNames[dt.getDay()];
    const colIdx = {2:0, 3:1, 4:2, 5:3, 6:4, 8:6, 9:7}[weekdayToCol[dayName]];
    const daysDiff = Math.round((dt - anchorSunday) / 86400000);
    const rowIdx = Math.floor(daysDiff / 7) * 3;
    
    if (colIdx === undefined || rowIdx >= gridValues.length - 2) continue;

    // Extract names, stripping the "P: " and "S: " prefixes
    let pRaw = norm(gridValues[rowIdx + 1][colIdx]).replace(/^P:\s*/i, '');
    let sRaw = norm(gridValues[rowIdx + 2][colIdx]).replace(/^S:\s*/i, '');
    
    // If BOTH are empty, do not create an event
    if (!pRaw && !sRaw) continue;

    let pName = pRaw || "none";
    let sName = sRaw || "none";
    let title = `P: ${pName} S: ${sName}`;

    // Compile Guest List
    let guests = [];
    if (pRaw && raData[pRaw.toLowerCase()]) guests.push(raData[pRaw.toLowerCase()]);
    if (sRaw && raData[sRaw.toLowerCase()]) guests.push(raData[sRaw.toLowerCase()]);
    if (calendarId2) guests.push(calendarId2);

    let guestString = guests.join(',');

    // Create the event
    try {
      calendar.createAllDayEvent(title, new Date(dt), {
        guests: guestString,
        sendInvites: false
      });
      eventsCreated++;
    } catch (e) {
      Logger.log(`Failed to create event for ${dt}: ${e.message}`);
    }
  }

  SpreadsheetApp.getUi().alert(`Success! Created ${eventsCreated} duty shifts.`);
}

function calDeleteDuty() {
  const ss = SpreadsheetApp.getActive();
  const info = ss.getSheetByName('Info');

  function norm(str) { return (str || '').toString().trim(); }
  const parseDate = s => { if (!s) return null; const d = new Date(s); return isNaN(d) ? null : d; };

  const calendarId = norm(info.getRange('C26').getValue());
  const startDate = parseDate(info.getRange('C17').getValue());
  const endDate = parseDate(info.getRange('F17').getValue());
  
  if (!startDate || !endDate) {
    SpreadsheetApp.getUi().alert('Invalid start or end date in Info.');
    return;
  }

  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    SpreadsheetApp.getUi().alert('Calendar not found. Check the ID in C26.');
    return;
  }

  // Define the search range. 
  // We add 1 day to the end date because the getEvents search is "exclusive" of the end time.
  let endSearch = new Date(endDate);
  endSearch.setDate(endSearch.getDate() + 1);
  
  // Fetch all events in the range
  let events = calendar.getEvents(startDate, endSearch);
  let deletedCount = 0;

  events.forEach(ev => {
    let title = ev.getTitle();
    
    // SAFETY CHECK: Only delete if the title follows our "P: ... S: ..." format.
    // This ensures you don't accidentally delete other events on the calendar.
    if (title.startsWith('P: ') && title.indexOf(' S: ') !== -1) {
      ev.deleteEvent();
      deletedCount++;
    }
  });

  SpreadsheetApp.getUi().alert(`Success! Removed ${deletedCount} duty shifts from the calendar.`);
}
