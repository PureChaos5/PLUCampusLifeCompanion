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
