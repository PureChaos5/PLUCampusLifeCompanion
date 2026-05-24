/**
 *  Duty Menu Code
 * 
 * onOpen() - makes the Campus Life Companion menu
 * showHelpAlert() & openLinkInNewTab - open the manual
 * clearDuty() - clears the scheduler
 * baseDutySchedule() - creates a blank schedule
 * 
 **/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Campus Life Companion')
      .addItem('Generate Schedule', 'scheduleDutyChange2')
      .addItem('Balance Schedule', 'balanceScheduleChange7')
      .addItem('Add All To Calendar', 'calUploadDuty')
      .addItem('Remove All From Calendar', 'calDeleteDuty')
      .addSeparator()
      .addItem('Clear Schedule', 'clearDuty') 
      .addItem('Create A Blank Schedule', 'baseDutySchedule') 
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

function clearDuty() {
  const ss = SpreadsheetApp.getActive();
  const out = ss.getSheetByName('Duty Schedule');
  out.getRange('B2:I55').clearContent().setBackground(null).setFontColor('#000000');
  out.getRange('G2:G55').setBackground('#000000').clearContent();
  SpreadsheetApp.getUi().alert(`Success! Cleared the schedule.`);
}

function baseDutySchedule() {
  const ss = SpreadsheetApp.getActive();
  const info = ss.getSheetByName('Info');
  const out = ss.getSheetByName('Duty Schedule');

  // Clear previous content and reset colors for the whole grid
  out.getRange('B2:I55').clearContent().setBackground(null).setFontColor('#000000').setFontWeight('normal');
  out.getRange('G2:G55').setBackground('#000000').clearContent();

  // --- Helper Functions ---
  const parseDate = s => {
    if (!s) return null;
    const d = new Date(s);
    if (isNaN(d)) return null;
    // FIX 1: Set the time to Noon so DST midnight shifts don't affect day counting
    d.setHours(12, 0, 0, 0); 
    return d;
  };
  function formatDateKey(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
  function formatDisplayDate(d) { return (d.getMonth() + 1) + '/' + d.getDate(); }

  const weekdayToCol = {'Sunday':2, 'Monday':3, 'Tuesday':4, 'Wednesday':5, 'Thursday':6, 'Friday':8, 'Saturday':9};
  const weekdayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const monthColors = ['#bf9000', '#0b5394', '#990000', '#38761d', '#351c75'];

  // 1. SETUP
  const startDate = parseDate(info.getRange('C17').getValue());
  const endDate = parseDate(info.getRange('F17').getValue());
  
  // Manual Primary Dates (C20) - for highlighting yellow
  const skipDatesRaw = (info.getRange('C20').getValue() || '').toString();
  const skipDates = new Set(skipDatesRaw.split(/\s+/).map(d => {
    const dd = parseDate(d); 
    return dd ? formatDateKey(dd) : null;
  }).filter(Boolean));

  if (!startDate || !endDate) {
    SpreadsheetApp.getUi().alert('Start or end date invalid in Info!');
    return;
  }

  // 2. BUILD THE GRID
  let distinctMonths = [];
  const anchorSunday = new Date(startDate); 
  anchorSunday.setDate(startDate.getDate() - startDate.getDay());

  for (let dt = new Date(startDate); dt <= endDate; dt.setDate(dt.getDate() + 1)) {
    const dKey = formatDateKey(dt);
    const mKey = dt.getMonth() + '-' + dt.getFullYear(); // Unique ID for each month
    if (!distinctMonths.includes(mKey)) distinctMonths.push(mKey);

    const dayName = weekdayNames[dt.getDay()];
    
    // FIX 2: Find total days safely with Math.round, then divide by 7
    const daysDiff = Math.round((dt - anchorSunday) / (24 * 60 * 60 * 1000));
    const weekIdx = Math.floor(daysDiff / 7);
    
    const row = 2 + weekIdx * 3;
    const col = weekdayToCol[dayName];
    
    // Safety check so we don't write outside the grid
    if (!col || row > 53) continue; 

    const isManual = skipDates.has(dKey);
    const dateCell = out.getRange(row, col);
    
    // Write the date
    dateCell.setValue(formatDisplayDate(dt));

    // Apply Colors
    if (isManual) {
      // Color the whole 3-cell block yellow for manual dates
      out.getRange(row, col, 3, 1).setBackground('#ffff00').setFontColor('#000000');
    } else {
      // Apply the specific month color to the date header cell
      const mIdx = distinctMonths.indexOf(mKey) % monthColors.length;
      dateCell.setBackground(monthColors[mIdx]).setFontColor('#ffffff').setFontWeight('bold');
      
      // Ensure the assignment cells below it are explicitly blank
      out.getRange(row + 1, col).setValue("");
      out.getRange(row + 2, col).setValue("");
    }
  }
  SpreadsheetApp.getUi().alert(`Success! Created a blank schedule.`);
}