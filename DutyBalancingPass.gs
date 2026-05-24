function balanceScheduleChange7() {
  const ss = SpreadsheetApp.getActive();
  const info = ss.getSheetByName('Info');
  const out = ss.getSheetByName('Duty Schedule');
  const startTime = new Date().getTime();
  const MAX_RUNTIME_MS = 120000; 

  function norm(str){return (str||'').toString().trim();}
  function formatDateKey(d){return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');}
  function formatDisplayDate(d) { return (d.getMonth() + 1) + '/' + d.getDate(); }
  
  const weekdayToCol = {'Sunday':2,'Monday':3,'Tuesday':4,'Wednesday':5,'Thursday':6,'Friday':8,'Saturday':9};
  const weekdayNames = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const monthColors = ['#bf9000', '#0b5394', '#990000', '#38761d', '#351c75'];

  // 1. DATA LOADING
  const scheduleMode = norm(info.getRange('D24').getValue());
  const startDate = new Date(info.getRange('C17').getValue());
  const endDate = new Date(info.getRange('F17').getValue());
  const skipDates = new Set((info.getRange('C20').getValue() || '').toString().split(/\s+/).map(d => {
    const dd = new Date(d); return isNaN(dd) ? null : formatDateKey(dd);
  }).filter(Boolean));

  const ras = [];
  for (let i = 1; i <= 13; i++) {
    const sh = ss.getSheetByName('RA' + i);
    if (!sh) continue;
    let rawName = norm(sh.getRange('C2').getValue());
    if (!rawName) continue;
    let name = rawName.split(/\s+/)[0];

    const getDaySet = (range, days) => {
      const set = new Set();
      const vals = sh.getRange(range).getValues()[0];
      vals.forEach((v, idx) => { if (v.toString().toLowerCase().includes('unavailable')) set.add(days[idx]); });
      return set;
    };

    ras.push({
      name: name,
      uWeek: getDaySet('B5:F5', ['sunday','monday','tuesday','wednesday','thursday']),
      uWknd: getDaySet('H5:I5', ['friday','saturday']),
      specificDates: new Set((norm(sh.getRange('D7').getValue()) || '').split(/\s+/).map(d=>{
        const dt = new Date(d); return isNaN(dt) ? null : formatDateKey(dt);
      }).filter(Boolean)),
      maxConsec: parseInt(sh.getRange('H13').getValue()) || 1,
      powerWkndOk: norm(sh.getRange('H15').getValue()) !== "I am not ok with it",
      wdP: 0, wdS: 0, weP: 0, weS: 0, assignments: [] 
    });
  }

  // 2. READ GRID & INITIALIZE ASSIGNMENTS
  const anchorSunday = new Date(startDate); 
  anchorSunday.setDate(startDate.getDate() - startDate.getDay());
  const gridValues = out.getRange('B2:I55').getValues();
  let allSlots = [];
  let distinctMonths = [];

  for (let dt = new Date(startDate); dt <= endDate; dt.setDate(dt.getDate() + 1)) {
    const dKey = formatDateKey(dt);
    const mKey = (dt.getMonth() + 1) + '-' + dt.getFullYear();
    if (!distinctMonths.includes(mKey)) distinctMonths.push(mKey);

    const dayName = weekdayNames[dt.getDay()];
    const isWknd = (dt.getDay() === 5 || dt.getDay() === 6);
    const colIdx = {2:0, 3:1, 4:2, 5:3, 6:4, 8:6, 9:7}[weekdayToCol[dayName]];
    const daysDiff = Math.round((dt - anchorSunday) / 86400000);
    const rowIdx = Math.floor(daysDiff / 7) * 3;
    if (colIdx === undefined) continue;

    let p = norm(gridValues[rowIdx + 1][colIdx]).replace(/^P:\s*/i, '');
    let s = norm(gridValues[rowIdx + 2][colIdx]).replace(/^S:\s*/i, '');

    allSlots.push({ date: new Date(dt), dateKey: dKey, dayName, isWknd, isManual: skipDates.has(dKey), primary: p, secondary: s, monthKey: mKey });
    
    [p, s].forEach((n, idx) => {
      let ra = ras.find(r => r.name.toLowerCase() === n.toLowerCase());
      if (ra) {
        const role = idx === 0 ? 'P' : 'S';
        const metric = isWknd ? (role ==='P'?'weP':'weS') : (role ==='P'?'wdP':'wdS');
        ra[metric]++;
        if (!ra.assignments.includes(dKey)) ra.assignments.push(dKey);
      }
    });
  }

  // 3. THE UNIVERSAL CONSECUTIVE ENGINE
  const attemptTransfer = (donor, receiver, roleKey, isWeekend) => {
    const donorSlots = allSlots.filter(s => {
      if (s.isWknd !== isWeekend || s.isManual) return false;
      const val = (roleKey === 'P' ? s.primary : s.secondary);
      return val.toLowerCase() === donor.name.toLowerCase();
    });

    donorSlots.sort(() => Math.random() - 0.5);

    for (const slot of donorSlots) {
      const dLower = slot.dayName.toLowerCase();
      
      // 1. Basic Unavailability
      if (receiver.specificDates.has(slot.dateKey)) continue;
      if (isWeekend ? receiver.uWknd.has(dLower) : receiver.uWeek.has(dLower)) continue;
      
      // 2. Same-Day Double Booking (Universal P + S check)
      if (slot.primary.toLowerCase() === receiver.name.toLowerCase() || 
          slot.secondary.toLowerCase() === receiver.name.toLowerCase()) continue;

      // 3. Power Weekend Boundary (H15)
      if (!receiver.powerWkndOk && isWeekend) {
        const offset = (dLower === 'friday') ? 86400000 : -86400000;
        const otherDay = formatDateKey(new Date(slot.date.getTime() + offset));
        if (receiver.assignments.includes(otherDay)) continue;
      }

      // 4. THE CHAIN CHECK (H13) - Critical Fix
      let chain = 1;
      // Look Backward
      let bDate = new Date(slot.date);
      while (true) {
        bDate.setDate(bDate.getDate() - 1);
        if (receiver.assignments.includes(formatDateKey(bDate))) chain++;
        else break;
      }
      // Look Forward
      let fDate = new Date(slot.date);
      while (true) {
        fDate.setDate(fDate.getDate() + 1);
        if (receiver.assignments.includes(formatDateKey(fDate))) chain++;
        else break;
      }
      // If adding THIS shift makes the streak too long, REJECT
      if (chain > receiver.maxConsec) continue;

      // SUCCESS - EXECUTE
      if (roleKey === 'P') slot.primary = receiver.name; else slot.secondary = receiver.name;
      const metric = isWeekend ? (roleKey==='P'?'weP':'weS') : (roleKey==='P'?'wdP':'wdS');
      
      donor[metric]--;
      // Only remove date from donor if they have NO other role that day
      const dCheck = allSlots.find(s => s.dateKey === slot.dateKey && (s.primary === donor.name || s.secondary === donor.name));
      if (!dCheck) donor.assignments.splice(donor.assignments.indexOf(slot.dateKey), 1);
      
      receiver[metric]++;
      if (!receiver.assignments.includes(slot.dateKey)) receiver.assignments.push(slot.dateKey);
      
      return true;
    }
    return false;
  };

  // 4. DEEP RIPPLE PASSES
  const deepBalance = (isWeekend) => {
    const roles = ['P', 'S'];
    roles.forEach(roleKey => {
      let progress = true;
      let metric = isWeekend ? (roleKey==='P'?'weP':'weS') : (roleKey==='P'?'wdP':'wdS');
      let target = Math.floor(allSlots.filter(s => !s.isManual && s.isWknd === isWeekend).length / ras.length);

      while (progress && (new Date().getTime() - startTime < MAX_RUNTIME_MS)) {
        progress = false;
        ras.sort((a,b) => a[metric] - b[metric]);
        const receiver = ras[0];
        if (receiver[metric] >= target) break;

        const potentialDonors = ras.filter(r => r.name !== receiver.name && r[metric] > 0).sort((a,b) => b[metric] - a[metric]);
        for (let donor of potentialDonors) {
          if (attemptTransfer(donor, receiver, roleKey, isWeekend)) { progress = true; break; }
        }
      }
    });

    // Total Weekend/Weekday Volume Pass
    let totalProgress = true;
    while (totalProgress && (new Date().getTime() - startTime < MAX_RUNTIME_MS)) {
      totalProgress = false;
      const getTot = (r) => isWeekend ? (r.weP + r.weS) : (r.wdP + r.wdS);
      const targetTot = Math.floor((allSlots.filter(s => !s.isManual && s.isWknd === isWeekend).length * 2) / ras.length);
      ras.sort((a,b) => getTot(a) - getTot(b));
      const receiver = ras[0];
      if (getTot(receiver) >= targetTot) break;
      const potentialDonors = ras.filter(r => r.name !== receiver.name && getTot(r) > 0).sort((a,b) => getTot(b) - getTot(a));
      for (let donor of potentialDonors) {
        if (attemptTransfer(donor, receiver, 'P', isWeekend) || attemptTransfer(donor, receiver, 'S', isWeekend)) {
          totalProgress = true; break;
        }
      }
    }
  };

  // Execute Weekday Balancing if "All Days", "Only Weekdays", or if the cell is accidentally left blank
  if (scheduleMode === "Only Weekdays (Sun-Thurs)" || scheduleMode === "All Days" || scheduleMode === "") {
    deepBalance(false); // Weekdays
  }

  // Execute Weekend Balancing if "All Days", "Only Weekends", or if the cell is accidentally left blank
  if (scheduleMode === "Only Weekends (Fri-Sat)" || scheduleMode === "All Days" || scheduleMode === "") {
    deepBalance(true);  // Weekends
  }

  // 5. FINAL OUTPUT
  out.getRange('B2:I55').setBackground(null).setFontColor('#000000').clearContent();
  out.getRange('G2:G55').setBackground('#000000'); 
  allSlots.forEach(slot => {
    const daysDiff = Math.round((slot.date - anchorSunday) / 86400000); // 86400000 ms in a day
    const weekIdx = Math.floor(daysDiff / 7);
    const row = 2 + weekIdx * 3;
    const col = weekdayToCol[slot.dayName];
    if (!col || row > 47) return;
    if (!slot.isManual) {
      const mIdx = distinctMonths.indexOf(slot.monthKey) % monthColors.length;
      out.getRange(row, col).setBackground(monthColors[mIdx]).setFontColor('#ffffff').setFontWeight('bold');
    } else { out.getRange(row, col, 3, 1).setBackground('#ffff00'); }
    out.getRange(row, col).setValue(formatDisplayDate(slot.date));
    out.getRange(row + 1, col).setValue(slot.primary ? 'P: ' + slot.primary : "");
    out.getRange(row + 2, col).setValue(slot.secondary ? 'S: ' + slot.secondary : "");
  });
  SpreadsheetApp.getUi().alert(`Success! Balanced the schedule.`);
}