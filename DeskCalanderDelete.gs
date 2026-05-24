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