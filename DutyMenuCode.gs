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