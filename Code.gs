// ============================================================
// ShowOps v3.0 — Google Apps Script Backend
// ============================================================
// 1. Set SPREADSHEET_ID to your Google Sheet ID below.
// 2. Set OWNER_EMAIL to your Google account email.
// 3. Deploy → New deployment → Web app → Execute as: Me
//    → Who has access: Anyone (or "Anyone in your domain" for Workspace)
// 4. Copy the web app URL into index.html SHOWOPS_CONFIG.apiUrl
// ============================================================

const SPREADSHEET_ID = '1fHV_8KDZjUD1B2cl4eHRl92Hl0DgVxoB3fj4RPWH_Tg';
const OWNER_EMAIL    = 'con-mayank.mandal@pocketfm.com';

const SHEETS = {
  SHOWS:       'Shows_Master',
  TASKS:       'Tasks',
  ISSUES:      'Issues',
  PEOPLE:      'People_Access',
  ACTIVITY:    'Activity_Log',
  SETTINGS:    'Settings',
  TASK_UPDATES:'Task_Updates',
  REQUESTS:    'Requests',
  PERFORMANCE: 'Weekly_Performance'
};

// ── Entry points ──────────────────────────────────────────────

function doGet(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    out.setContent(JSON.stringify(getState()));
  } catch(err) {
    out.setContent(JSON.stringify({ error: err.message }));
  }
  return out;
}

function doPost(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    const payload = JSON.parse(e.postData.contents);
    out.setContent(JSON.stringify(handleAction(payload)));
  } catch(err) {
    out.setContent(JSON.stringify({ error: err.message }));
  }
  return out;
}

// ── Read full state ───────────────────────────────────────────

function getState() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const shows = sheetToObjects(ss, SHEETS.SHOWS).map(r => ({
    id:       r['Show_ID'],
    name:     r['Show_Name'],
    type:     r['Show_Type'],
    genre:    r['Genre'],
    language: r['Language'],
    active:   toBool(r['Active'])
  })).filter(s => s.id);

  const tasks = sheetToObjects(ss, SHEETS.TASKS).map(r => ({
    id:        r['Task_ID'],
    showId:    r['Show_ID'],
    task:      r['Task'],
    hm:        r['HML'],
    p:         r['P_Level'],
    high:      toBool(r['High_Priority']),
    assignee:  r['Assigned_To'] || '',
    status:    r['Status'],
    deadline:  r['Deadline'],
    created:   r['Created_At'],
    updated:   r['Updated_At'],
    requester: r['Raised_By'],
    message:   r['Latest_Message']
  })).filter(t => t.id);

  const issues = sheetToObjects(ss, SHEETS.ISSUES).map(r => ({
    id:       r['Issue_ID'],
    showId:   r['Show_ID'],
    taskId:   r['Task_ID'],
    issue:    r['Issue'],
    severity: r['Severity'],
    status:   r['Status'],
    raisedBy: r['Raised_By'],
    owner:    r['Owner'],
    created:  r['Created_At']
  })).filter(i => i.id);

  const people = sheetToObjects(ss, SHEETS.PEOPLE).map(r => ({
    id:     r['User_ID'],
    name:   r['Name'],
    email:  r['Email'],
    role:   r['Role'],
    after:  toBool(r['After_Hours']),
    active: toBool(r['Active'])
  })).filter(p => p.id);

  const activity = sheetToObjects(ss, SHEETS.ACTIVITY).map(r => ({
    t:    r['Date_Time'],
    u:    r['User'],
    a:    r['Action'],
    show: r['Show'],
    d:    r['Details']
  })).filter(a => a.t).reverse();

  // Read settings directly (bypass sheetToObjects) to get raw Date objects
  // before they are String()-ified and lose their time information
  const settings = { start: '06:00', end: '18:30', override: [] };
  const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (settingsSheet) {
    const sData = settingsSheet.getDataRange().getValues();
    if (sData.length > 1) {
      const hdrs = sData[0].map(h => String(h).trim());
      const ki = hdrs.indexOf('Key'), vi = hdrs.indexOf('Value');
      if (ki >= 0 && vi >= 0) {
        for (let i = 1; i < sData.length; i++) {
          const key = String(sData[i][ki]).trim();
          const raw = sData[i][vi];
          if (key === 'Start_Time')            settings.start    = parseSettingTime(raw);
          if (key === 'End_Time')              settings.end      = parseSettingTime(raw);
          if (key === 'After_Hours_Override')  settings.override = String(raw||'').split(',').map(s=>s.trim()).filter(Boolean);
        }
      }
    }
  }

  return { data: { shows, tasks, issues, people, activity, settings } };
}

// ── Write actions ─────────────────────────────────────────────

function handleAction(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  switch (p.action) {
    case 'addShow':         return actAddShow(ss, p);
    case 'archiveShow':     return actArchiveShow(ss, p);
    case 'addTask':         return actAddTask(ss, p);
    case 'updateTask':      return actUpdateTask(ss, p);
    case 'assignTask':      return actAssignTask(ss, p);
    case 'addIssue':        return actAddIssue(ss, p);
    case 'addPerson':       return actAddPerson(ss, p);
    case 'togglePerson':    return actTogglePerson(ss, p);
    case 'saveSettings':    return actSaveSettings(ss, p);
    case 'log':             return actLog(ss, p);
    default: return { error: 'Unknown action: ' + p.action };
  }
}

function actAddShow(ss, p) {
  ss.getSheetByName(SHEETS.SHOWS)
    .appendRow([p.id, p.name, p.type, p.genre||'—', p.language||'—', 'TRUE']);
  actLog(ss, { u: p.user, a: 'Show Added', show: p.name, d: 'Type: '+p.type });
  return { ok: true };
}

function actArchiveShow(ss, p) {
  updateRowWhere(ss.getSheetByName(SHEETS.SHOWS), 'Show_ID', p.id, { Active: 'FALSE' });
  actLog(ss, { u: p.user, a: 'Show Archived', show: p.name, d: 'Removed from active Show Master' });
  return { ok: true };
}

function actAddTask(ss, p) {
  ss.getSheetByName(SHEETS.TASKS).appendRow([
    p.id, '', p.showId, p.task, p.requester, p.assignee||'',
    p.status, p.hm, p.p, p.high ? 'TRUE' : 'FALSE',
    p.deadline, p.created, p.updated, p.message
  ]);
  actLog(ss, { u: p.requester, a: 'Request Raised', show: p.showName, d: p.task });
  if (p.assignee) actLog(ss, { u: p.user, a: 'Task Assigned', show: p.showName, d: 'Assigned to '+p.assignee });
  return { ok: true };
}

function actUpdateTask(ss, p) {
  updateRowWhere(ss.getSheetByName(SHEETS.TASKS), 'Task_ID', p.id, {
    Status: p.status, Latest_Message: p.message,
    Updated_At: p.updated, Assigned_To: p.assignee
  });
  const tu = ss.getSheetByName(SHEETS.TASK_UPDATES);
  if (tu) tu.appendRow([p.id, p.status, p.message, p.user, p.updated]);
  actLog(ss, { u: p.user, a: p.logAction || 'Task Updated', show: p.showName, d: p.message });
  return { ok: true };
}

function actAssignTask(ss, p) {
  updateRowWhere(ss.getSheetByName(SHEETS.TASKS), 'Task_ID', p.id, {
    Assigned_To: p.assignee, Status: p.status, Updated_At: p.updated
  });
  actLog(ss, { u: p.user, a: 'Task Assigned', show: p.showName, d: p.assignee||'Unassigned' });
  return { ok: true };
}

function actAddIssue(ss, p) {
  ss.getSheetByName(SHEETS.ISSUES)
    .appendRow([p.id, p.showId, p.taskId, p.issue, p.severity, p.status, p.raisedBy, p.owner, p.created]);
  actLog(ss, { u: p.raisedBy, a: 'Issue Raised', show: p.showName, d: p.issue });
  return { ok: true };
}

function actAddPerson(ss, p) {
  ss.getSheetByName(SHEETS.PEOPLE)
    .appendRow([p.id, p.name, p.email, p.role, p.after ? 'TRUE' : 'FALSE', 'TRUE']);
  actLog(ss, { u: p.user, a: 'Person Added', show: p.name, d: p.role });
  return { ok: true };
}

function actTogglePerson(ss, p) {
  updateRowWhere(ss.getSheetByName(SHEETS.PEOPLE), 'User_ID', p.id, {
    Active: p.active ? 'TRUE' : 'FALSE'
  });
  actLog(ss, { u: p.user, a: p.active ? 'Person Activated' : 'Person Deactivated', show: p.name, d: p.role });
  return { ok: true };
}

function actSaveSettings(ss, p) {
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);
  upsertSetting(sheet, 'Start_Time',           p.start);
  upsertSetting(sheet, 'End_Time',             p.end);
  upsertSetting(sheet, 'After_Hours_Override', (p.override||[]).join(','));
  return { ok: true };
}

function actLog(ss, p) {
  const sheet = ss.getSheetByName(SHEETS.ACTIVITY);
  sheet.appendRow([p.t || new Date().toISOString(), p.u||'System', p.a, p.show||'', p.d||'']);
  return { ok: true };
}

// ── Helpers ───────────────────────────────────────────────────

function sheetToObjects(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
    return obj;
  });
}

function updateRowWhere(sheet, keyCol, keyVal, updates) {
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const keyIdx  = headers.indexOf(keyCol);
  if (keyIdx < 0) return;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyIdx]) === String(keyVal)) {
      Object.entries(updates).forEach(([col, val]) => {
        const ci = headers.indexOf(col);
        if (ci >= 0) sheet.getRange(i + 1, ci + 1).setValue(val);
      });
      break;
    }
  }
}

function upsertSetting(sheet, key, value) {
  const sVal = String(value);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === key) {
      const cell = sheet.getRange(i + 1, 2);
      cell.setNumberFormat('@');   // force plain text — prevents time auto-format
      cell.setValue(sVal);
      return;
    }
  }
  const nr = sheet.getLastRow() + 1;
  sheet.getRange(nr, 1).setValue(key);
  const cell = sheet.getRange(nr, 2);
  cell.setNumberFormat('@');
  cell.setValue(sVal);
}

function toBool(v) {
  return v === 'TRUE' || v === true || v === '1' || v === 'true';
}

// Convert a raw Sheets cell value (may be Date object, time serial, or "HH:MM" string) to "HH:MM"
function parseSettingTime(raw) {
  if (!raw && raw !== 0) return '06:00';
  if (raw instanceof Date) {
    return String(raw.getHours()).padStart(2,'0') + ':' + String(raw.getMinutes()).padStart(2,'0');
  }
  const s = String(raw).trim();
  // Already "HH:MM" or "H:MM"
  const hm = s.match(/^(\d{1,2}):(\d{2})/);
  if (hm) return String(parseInt(hm[1])).padStart(2,'0') + ':' + hm[2];
  // Sheets time serial stored as plain number (fraction of a day)
  const n = parseFloat(s);
  if (!isNaN(n) && n >= 0 && n < 1) {
    const totalMin = Math.round(n * 1440);
    return String(Math.floor(totalMin/60)).padStart(2,'0') + ':' + String(totalMin%60).padStart(2,'0');
  }
  return s;
}

// ── One-time setup: run this ONCE from the Apps Script editor ──
// Extensions → Apps Script → select setupSheets → Run

function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const tabs = [
    { name: SHEETS.SHOWS,        headers: ['Show_ID','Show_Name','Show_Type','Genre','Language','Active'] },
    { name: SHEETS.TASKS,        headers: ['Task_ID','Request_ID','Show_ID','Task','Raised_By','Assigned_To','Status','HML','P_Level','High_Priority','Deadline','Created_At','Updated_At','Latest_Message'] },
    { name: SHEETS.ISSUES,       headers: ['Issue_ID','Show_ID','Task_ID','Issue','Severity','Status','Raised_By','Owner','Created_At'] },
    { name: SHEETS.PEOPLE,       headers: ['User_ID','Name','Email','Role','After_Hours','Active'] },
    { name: SHEETS.ACTIVITY,     headers: ['Date_Time','User','Action','Show','Details'] },
    { name: SHEETS.SETTINGS,     headers: ['Key','Value'] },
    { name: SHEETS.TASK_UPDATES, headers: ['Task_ID','Status','Message','Updated_By','Updated_At'] },
    { name: SHEETS.REQUESTS,     headers: ['Request_ID','Show_ID','Task','Raised_By','Assigned_To','Priority','P_Level','High_Priority','Deadline','Created_At'] },
    { name: SHEETS.PERFORMANCE,  headers: ['Week','Assigned','Completed','Blocked','Completion_Rate','Avg_TAT'] }
  ];

  tabs.forEach(({ name, headers }) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#091224')
           .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  });

  // Seed Settings defaults
  const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.appendRow(['Start_Time',           '06:00']);
    settingsSheet.appendRow(['End_Time',             '18:30']);
    settingsSheet.appendRow(['After_Hours_Override', 'Mayank']);
  }

  Logger.log('ShowOps sheet setup complete. All tabs created.');
}
