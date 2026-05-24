function scheduleDutyChange2() { 
  const ss = SpreadsheetApp.getActive();
  const info = ss.getSheetByName('Info');
  const out = ss.getSheetByName('Duty Schedule');
  
  // Clear previous content and reset colors
  out.getRange('B2:I55').clearContent().setBackground(null).setFontColor('#000000');
  out.getRange('G2:G55').setBackground('#000000').clearContent();

  // --- Helper Functions ---
  const parseDate = s => {
    if (!s) return null;
    const d = new Date(s);
    return isNaN(d) ? null : d;
  };
  function norm(str){return (str||'').toString().trim();}
  function formatDateKey(d){return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');}
  function formatDisplayDate(d){return (d.getMonth()+1)+'/'+d.getDate();}
  
  const weekdayToCol = {'Sunday':2,'Monday':3,'Tuesday':4,'Wednesday':5,'Thursday':6,'Friday':8,'Saturday':9};
  const weekdayNames = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const monthColors = ['#bf9000', '#0b5394', '#990000', '#38761d', '#351c75'];

  // 1. SETUP
  const scheduleMode = norm(info.getRange('D24').getValue());
  const isTargetMode = (isWknd) => {
    if (scheduleMode === "Only Weekdays (Sun-Thurs)") return !isWknd;
    if (scheduleMode === "Only Weekends (Fri-Sat)") return isWknd;
    return true; // Defaults to All Days
  };
  const startDate = parseDate(info.getRange('C17').getValue());
  const endDate = parseDate(info.getRange('F17').getValue());
  
  // Manual Primary Dates (C20)
  const skipDatesRaw = (info.getRange('C20').getValue() || '').toString();
  const skipDates = new Set(skipDatesRaw.split(/\s+/).map(d => {
    const dd = parseDate(d); 
    return dd ? formatDateKey(dd) : null;
  }).filter(Boolean));

  if (!startDate || !endDate) {
    SpreadsheetApp.getUi().alert('Start or end date invalid in Info!');
    return;
  }

  // 2. LOAD RAs
  const ras = [];
  for (let i = 1; i <= 13; i++) {
    const sh = ss.getSheetByName('RA' + i);
    if (!sh) continue; 
    let rawName = norm(sh.getRange('C2').getValue());
    if (!rawName) continue;
    const name = rawName.split(/\s+/)[0]; // First name only

    // STRICT PREFERENCE CHECKING (Case Insensitive)
    const wdPrefsRaw = sh.getRange('B5:F5').getValues()[0];
    const daysWD = ['sunday','monday','tuesday','wednesday','thursday'];
    const uWeek = new Set(), pWeek = new Set();
    
    wdPrefsRaw.forEach((val, idx) => {
      let v = val.toString().toLowerCase();
      if (v.includes('unavailable')) uWeek.add(daysWD[idx]);
      if (v.includes('preferred')) pWeek.add(daysWD[idx]);
    });

    const wePrefsRaw = sh.getRange('H5:I5').getValues()[0];
    const daysWE = ['friday','saturday'];
    const uWknd = new Set(), pWknd = new Set();
    
    wePrefsRaw.forEach((val, idx) => {
      let v = val.toString().toLowerCase();
      if (v.includes('unavailable')) uWknd.add(daysWE[idx]);
      if (v.includes('preferred')) pWknd.add(daysWE[idx]);
    });

    const specificUnavail = new Set((norm(sh.getRange('D7').getValue()) || '').split(/\s+/).map(d=>{
      const dd=parseDate(d); return dd ? formatDateKey(dd):null;
    }).filter(Boolean));
    
    ras.push({
      name: name,
      uWeek, uWeekends: uWknd,
      pWeek, pWeekends: pWknd,
      specificDates: specificUnavail,
      maxConsec: parseInt(sh.getRange('H13').getValue()) || 1,
      powerWeekendPref: norm(sh.getRange('H15').getValue()).toLowerCase(),
      semesterPref: norm(sh.getRange('H17').getValue()).toLowerCase(),
      wdP: 0, wdS: 0, weP: 0, weS: 0, 
      assignments: [] // Stores dateKeys
    });
  }

  // 3. SLOTS & MONTH COLORING MAP
  let allSlots = [];
  let distinctMonths = [];
  
  for (let dt = new Date(startDate); dt <= endDate; dt.setDate(dt.getDate() + 1)) {
    const dKey = formatDateKey(dt);
    const mKey = dt.getMonth() + '-' + dt.getFullYear(); // Unique ID for each month
    if (!distinctMonths.includes(mKey)) distinctMonths.push(mKey);

    const isManual = skipDates.has(dKey);
    const dayName = weekdayNames[dt.getDay()];
    const isWeekend = (dt.getDay() === 5 || dt.getDay() === 6);
    const progress = (dt.getTime() - startDate.getTime()) / (endDate.getTime() - startDate.getTime());
    
    allSlots.push({
      date: new Date(dt),
      dateKey: dKey,
      dayName: dayName,
      isWeekend: isWeekend,
      phase: progress < 0.5 ? 'beginning' : 'end',
      primary: null,
      secondary: null,
      monthKey: mKey,
      isManual: isManual
    });
  }

  // 4. QUOTAS
  const activeSlots = allSlots.filter(s => !s.isManual);
  const counts = {
    wd: activeSlots.filter(s => !s.isWeekend).length,
    we: activeSlots.filter(s => s.isWeekend).length
  };

  ras.forEach(r => {
    r.qWdP = Math.floor(counts.wd / ras.length);
    r.qWdS = Math.floor(counts.wd / ras.length);
    r.qWeP = Math.floor(counts.we / ras.length);
    r.qWeS = Math.floor(counts.we / ras.length);
  });
  // Distribute remainders...
  for(let i=0; i<(counts.wd % ras.length); i++) ras[i].qWdP++;
  for(let i=0; i<(counts.wd % ras.length); i++) ras[(i+2)%ras.length].qWdS++;
  for(let i=0; i<(counts.we % ras.length); i++) ras[i].qWeP++;
  for(let i=0; i<(counts.we % ras.length); i++) ras[(i+2)%ras.length].qWeS++;

  // 5. PHASE 1: POWER WEEKENDS
  if (isTargetMode(true)) {
    allSlots.filter(s => s.dayName === 'Friday' && !s.isManual).forEach(fri => {
      const sat = allSlots.find(s => s.dateKey === formatDateKey(new Date(fri.date.getTime() + 86400000)) && !s.isManual);
      if (!sat) return;

      let candidates = ras.filter(r => {
        if (r.powerWeekendPref !== 'i would prefer it') return false;
        if (r.uWeekends.has('friday') || r.uWeekends.has('saturday')) return false;
        if (r.specificDates.has(fri.dateKey) || r.specificDates.has(sat.dateKey)) return false;
        return (r.weP < r.qWeP && r.weS < r.qWeS);
      });

      if (candidates.length > 0) {
        const winner = candidates[0];
        fri.primary = winner.name; winner.weP++;
        sat.secondary = winner.name; winner.weS++;
        winner.assignments.push(fri.dateKey, sat.dateKey);
      }
    });
  }

  // 6. PHASE 2: FILL REMAINING
  ['primary', 'secondary'].forEach(role => {
    allSlots.filter(s => !s.isManual && isTargetMode(s.isWeekend)).forEach(slot => {
      if (role === 'primary' && slot.primary) return;
      if (role === 'secondary' && slot.secondary) return;

      const isWknd = slot.isWeekend;
      const dLower = slot.dayName.toLowerCase();

      let eligible = ras.filter(r => {
        // Strict Hard Blocks
        if (r.specificDates.has(slot.dateKey)) return false;
        if (isWknd ? r.uWeekends.has(dLower) : r.uWeek.has(dLower)) return false;
        if (slot.primary === r.name || slot.secondary === r.name) return false;

        // Max Consecutive (H13)
        let consecutive = 0;
        let cDate = new Date(slot.date);
        for (let k = 0; k < 5; k++) {
          cDate.setDate(cDate.getDate() - 1);
          if (r.assignments.includes(formatDateKey(cDate))) consecutive++;
          else break;
        }
        if (consecutive >= r.maxConsec) return false;

        return true;
      });

      // Prefer those under quota
      let underQuota = eligible.filter(r => {
        const cur = isWknd ? (role === 'primary' ? r.weP : r.weS) : (role === 'primary' ? r.wdP : r.wdS);
        const quo = isWknd ? (role === 'primary' ? r.qWeP : r.qWeS) : (role === 'primary' ? r.qWdP : r.qWdS);
        return cur < quo;
      });

      let pool = underQuota.length > 0 ? underQuota : eligible;
      if (pool.length > 0) {
        pool.sort((a,b) => (b.semesterPref === slot.phase ? 1 : 0) - (a.semesterPref === slot.phase ? 1 : 0) || Math.random() - 0.5);
        const winner = pool[0];
        if (role === 'primary') { slot.primary = winner.name; isWknd ? winner.weP++ : winner.wdP++; }
        else { slot.secondary = winner.name; isWknd ? winner.weS++ : winner.wdS++; }
        winner.assignments.push(slot.dateKey);
      }
    });
  });

  // 7. OUTPUT
  const anchorSunday = new Date(startDate); 
  anchorSunday.setDate(startDate.getDate() - startDate.getDay());

  allSlots.forEach(slot => {
    const daysDiff = Math.round((slot.date - anchorSunday) / 86400000); // 86400000 ms in a day
    const weekIdx = Math.floor(daysDiff / 7);
    const row = 2 + weekIdx * 3;
    const col = weekdayToCol[slot.dayName];
    if (!col || row > 47) return;

    const dateCell = out.getRange(row, col);
    dateCell.setValue(formatDisplayDate(slot.date));

    if (slot.isManual) {
      out.getRange(row, col, 3, 1).setBackground('#ffff00').setFontColor('#000000');
    } else {
      // 1. ALWAYS apply Month Color to the Date Cell
      const mIdx = distinctMonths.indexOf(slot.monthKey) % monthColors.length;
      dateCell.setBackground(monthColors[mIdx]).setFontColor('#ffffff').setFontWeight('bold');
      
      // 2. ONLY print names if this day matches your D24 dropdown selector
      if (isTargetMode(slot.isWeekend)) {
        if (slot.primary) out.getRange(row + 1, col).setValue('P: ' + slot.primary);
        if (slot.secondary) out.getRange(row + 2, col).setValue('S: ' + slot.secondary);
      }
    }
  });
  SpreadsheetApp.getUi().alert(`Success! Created the schedule.`);
}