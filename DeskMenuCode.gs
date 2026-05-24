/**
 *  Desk Menu Code
 * 
 * onOpen() - makes the Campus Life Companion menu
 * showHelpAlert() & openLinkInNewTab() - open the manual
 * clearDesk() - clears the scheduler
 * baseDeskSchedule() - creates a blank schedule
 * 
 **/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Campus Life Companion')
      .addItem('Generate Schedule', 'sDDEFG4')
      .addItem('Add To Calendar', 'calUploadDesk')
      .addItem('Remove From Calendar', 'calDeleteDesk')
      .addSeparator()
      .addItem('Clear All Schedules', 'clearDesk') 
      .addItem('Create Blank Schedules', 'baseDeskSchedule') 
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

function clearDesk() {
  const ss = SpreadsheetApp.getActive();
  const out1 = ss.getSheetByName('Desk Schedule 1');
  const out2 = ss.getSheetByName('Desk Schedule 2');
  const out3 = ss.getSheetByName('Desk Schedule 3');
  out1.getRange('D3:H28').clearContent().setBackground(null).setFontColor('#000000');
  out2.getRange('D3:H28').clearContent().setBackground(null).setFontColor('#000000');
  out3.getRange('D3:H28').clearContent().setBackground(null).setFontColor('#000000');

  out1.getRange('J19:K19').clearContent().setBackground(null).setFontColor('#000000');
  out2.getRange('J19:K19').clearContent().setBackground(null).setFontColor('#000000');
  out3.getRange('J19:K19').clearContent().setBackground(null).setFontColor('#000000');

  SpreadsheetApp.getUi().alert(`Success! Cleared all schedules.`);
}

function baseDeskSchedule() {
  const ss = SpreadsheetApp.getActive();
  const info = ss.getSheetByName('Info');
  const out1 = ss.getSheetByName('Desk Schedule 1');
  const out2 = ss.getSheetByName('Desk Schedule 2');
  const out3 = ss.getSheetByName('Desk Schedule 3');

  // Clear previous content and reset colors for the whole grid
  out1.getRange('D3:H28').clearContent().setBackground(null).setFontColor('#000000');
  out2.getRange('D3:H28').clearContent().setBackground(null).setFontColor('#000000');
  out3.getRange('D3:H28').clearContent().setBackground(null).setFontColor('#000000');

  out1.getRange('J19:K19').clearContent().setBackground(null).setFontColor('#000000');
  out2.getRange('J19:K19').clearContent().setBackground(null).setFontColor('#000000');
  out3.getRange('J19:K19').clearContent().setBackground(null).setFontColor('#000000');

  const COLOR_PRIORITY = '#d9ead3'; // Light green

  // --- Helper: Convert time to slot index ---
  function timeToSlot(timeValue) {
    if (!timeValue || !(timeValue instanceof Date)) return null;
    const hours = timeValue.getHours();
    const minutes = timeValue.getMinutes();
    return (hours - 7) * 2 + (minutes >= 30 ? 1 : 0);
  }

  // --- Helper: Convert time to slot index ---
  function timeToSlot(timeValue) {
    if (!timeValue || !(timeValue instanceof Date)) return 0;
    const hours = timeValue.getHours();
    const minutes = timeValue.getMinutes();
    return (hours - 7) * 2 + (minutes >= 30 ? 1 : 0);
  }

  const startSlot = timeToSlot(info.getRange('C19').getValue());
  const endSlot = timeToSlot(info.getRange('D19').getValue());

  // --- Assign buildings ---
  const raw = info.getRange('C15').getValue();
  if (!raw) return;

  const buildings = String(raw)
    .split(/[, \n]/)
    .map(s => s.trim())
    .filter(s => s);

  SpreadsheetApp.getUi().alert(`Success! Created blank schedule(s) for: ${buildings}`);

  if (buildings.length === 0) return;

  for (let i = 0; i < 3; i++) {
    const sheet = ss.getSheetByName(`Desk Schedule ${i + 1}`);
    if (!sheet) continue;

    const buildingName = buildings[i] || '';
    sheet.getRange('J19').setValue(buildingName);

    for (let slot = 0; slot < 26; slot++) {
      if (slot < startSlot || slot >= endSlot) {
        sheet.getRange(3 + slot, 4, 1, 5).setBackground('#d9d9d9');
      }
    }
  }

  // --- Read priority time blocks from Info sheet ---
  const priorityRows = [21, 23, 25];
  const priorityTimes = [];

  priorityRows.forEach(row => {
    const startVal = info.getRange(`C${row}`).getValue();
    const endVal = info.getRange(`D${row}`).getValue();
    const building = info.getRange(`E${row}`).getValue();

    const start = timeToSlot(startVal);
    const end = timeToSlot(endVal);

    if (start !== null && end !== null && start < end) {
      priorityTimes.push({
        start,
        end,
        building: building || null
      });
    }
  });

  if (priorityTimes.length === 0) return;

  // --- Apply coloring to schedule sheets ---
  for (let i = 1; i <= 3; i++) {
    const sheet = ss.getSheetByName(`Desk Schedule ${i}`);
    if (!sheet) continue;

    const currentBuilding = sheet.getRange('J19').getValue();

    priorityTimes.forEach(pt => {
      if (!pt.building || pt.building === currentBuilding) {
        const startRow = 3 + pt.start;
        const numRows = pt.end - pt.start;

        sheet.getRange(startRow, 4, numRows, 5).setBackground(COLOR_PRIORITY);
      }
    });

    if (currentBuilding === '') {
        sheet.getRange('D3:H28').clearContent().setBackground(null).setFontColor('#000000');
    }
    
  }
}