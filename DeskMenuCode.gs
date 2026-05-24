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
