/**
 * IDS PMC Tool — backend
 * Standalone Apps Script project, separate scriptId and separate Google Sheet
 * from the DPR productivity system. Sheet is created by setup() and its ID
 * stored in Script Properties under PMC_SHEET_ID.
 */

var STAGES = ['Design Freeze', 'BOQ', 'Selection', 'Order Placement', 'Delivery', 'Installation'];

var SUBITEMS = [
  { category: 'Flooring', subItem: 'Flooring' },
  { category: 'False Ceiling', subItem: 'False Ceiling' },
  { category: 'Wall Panelling', subItem: 'Carpentry' },
  { category: 'Wall Panelling', subItem: 'Stone Work' },
  { category: 'Wall Panelling', subItem: 'Paint' },
  { category: 'Loose Furniture', subItem: 'Loose Furniture' },
  { category: 'Lighting', subItem: 'Lighting' },
  { category: 'Electrical', subItem: 'Conduiting' },
  { category: 'Electrical', subItem: 'Wiring' },
  { category: 'Electrical', subItem: 'Switch & Sockets' },
  { category: 'Electrical', subItem: 'Appliances' },
  { category: 'Air Conditioning', subItem: 'Copper Piping' },
  { category: 'Air Conditioning', subItem: 'Drain Channel' },
  { category: 'Air Conditioning', subItem: 'Machine Installation' },
  { category: 'Plumbing', subItem: 'Piping & Concealed Fittings' },
  { category: 'Plumbing', subItem: 'Sanitary Ware' },
  { category: 'Automation', subItem: 'Internal Wiring' },
  { category: 'Automation', subItem: 'Machine Installation' }
];

var CATEGORIES = (function () {
  var seen = {}, out = [];
  SUBITEMS.forEach(function (s) { if (!seen[s.category]) { seen[s.category] = true; out.push(s.category); } });
  return out;
})();

var DEFAULT_SPLIT = { 'Order Placement': 40, 'Delivery': 40, 'Installation': 20 };

var SHEET_PROP_KEY = 'PMC_SHEET_ID';

var TABS = {
  PROJECTS: { name: 'PROJECTS', headers: ['ProjectID', 'Name', 'Address', 'StartDate', 'TargetMoveIn', 'Budget', 'ClientName', 'CreatedAt'] },
  SPACES: { name: 'SPACES', headers: ['SpaceID', 'ProjectID', 'Name', 'SortOrder'] },
  TRACKER: { name: 'TRACKER', headers: ['TrackerID', 'ProjectID', 'SpaceID', 'Category', 'SubItem', 'StagesJSON', 'RollupStatus', 'UpdatedAt'] },
  CASHFLOW: { name: 'CASHFLOW', headers: ['CashflowID', 'ProjectID', 'SpaceID', 'Category', 'BOQValue', 'SplitJSON', 'Invoiced', 'Received', 'DueDate', 'PaymentStatus', 'UpdatedAt'] },
  DAILY_LOG: { name: 'DAILY_LOG', headers: ['LogID', 'Date', 'ProjectID', 'SpaceID', 'Entry', 'LoggedBy', 'HasBlocker', 'CreatedAt'] }
};

// ---------- Sheet plumbing ----------

function getSS_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(SHEET_PROP_KEY);
  if (!id) {
    setup();
    id = props.getProperty(SHEET_PROP_KEY);
  }
  return SpreadsheetApp.openById(id);
}

function getTab_(ss, key) {
  var t = TABS[key];
  var sh = ss.getSheetByName(t.name);
  if (!sh) {
    sh = ss.insertSheet(t.name);
    sh.appendRow(t.headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function rowsToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row.join('') === '') continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = row[j];
    obj._row = i + 1;
    out.push(obj);
  }
  return out;
}

function findRowById_(sheet, idCol, id) {
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var colIdx = headers.indexOf(idCol);
  for (var i = 1; i < values.length; i++) {
    if (values[i][colIdx] === id) return i + 1;
  }
  return -1;
}

function nextId_(prefix, existingIds) {
  var max = 0;
  existingIds.forEach(function (id) {
    var m = /(\d+)$/.exec(String(id));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return prefix + '-' + String(max + 1).padStart(3, '0');
}

// ---------- One-time setup ----------

function setup() {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty(SHEET_PROP_KEY);
  var ss;
  if (existing) {
    try { ss = SpreadsheetApp.openById(existing); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('IDS PMC Data');
    props.setProperty(SHEET_PROP_KEY, ss.getId());
    var defaultSheet = ss.getSheets()[0];
    ss.deleteSheet(defaultSheet);
  }
  Object.keys(TABS).forEach(function (key) { getTab_(ss, key); });

  // Seed a demo project only if PROJECTS is empty
  var projSheet = getTab_(ss, 'PROJECTS');
  if (rowsToObjects_(projSheet).length === 0) {
    seedDemoProject_(ss);
  }
  Logger.log('PMC sheet ready: ' + ss.getUrl());
  return ss.getUrl();
}

function seedDemoProject_(ss) {
  var now = new Date();
  var projectId = 'P-001';
  var projSheet = getTab_(ss, 'PROJECTS');
  projSheet.appendRow([
    projectId, 'Sharma Residence (Demo Project)', 'Kolar Road, Bhopal',
    daysAgoISO_(70), daysFromNowISO_(40), 3200000, 'Rohit & Priya Sharma', now
  ]);

  var spaceSheet = getTab_(ss, 'SPACES');
  var trackerSheet = getTab_(ss, 'TRACKER');
  var cashflowSheet = getTab_(ss, 'CASHFLOW');
  var logSheet = getTab_(ss, 'DAILY_LOG');

  var spaces = [
    { id: 'S-001', name: 'Drawing Room' },
    { id: 'S-002', name: 'Master Bedroom' }
  ];
  spaces.forEach(function (sp, idx) {
    spaceSheet.appendRow([sp.id, projectId, sp.name, idx + 1]);
    createTrackerRowsForSpace_(trackerSheet, projectId, sp.id, idx === 0 ? 'ahead' : 'behind');
    createCashflowRowsForSpace_(cashflowSheet, projectId, sp.id, idx === 0 ? 0.55 : 0.3);
  });

  var engineers = ['Deepak Soni', 'Achal Rathore'];
  var sampleLogs = [
    { d: -1, sp: 'S-001', by: engineers[0], entry: 'Flooring tiles laid in living area, grouting pending tomorrow. False ceiling frame work started.', blocker: false },
    { d: -1, sp: 'S-002', by: engineers[1], entry: 'Electrical conduiting completed for master bedroom. Waiting on switch plate selection from client before wiring.', blocker: true },
    { d: -3, sp: 'S-001', by: engineers[0], entry: 'Wall panelling carpentry work in progress — TV unit framing done.', blocker: false },
    { d: -5, sp: 'S-002', by: engineers[1], entry: 'AC copper piping installed, drain channel work delayed due to material shortage from vendor.', blocker: true },
    { d: -7, sp: 'S-001', by: engineers[0], entry: 'Site measurement re-verified for loose furniture layout. BOQ finalized with client.', blocker: false }
  ];
  sampleLogs.forEach(function (l) {
    logSheet.appendRow([
      Utilities.getUuid().slice(0, 8), daysAgoISO_(-l.d), projectId, l.sp, l.entry, l.by, l.blocker, new Date()
    ]);
  });
}

function daysAgoISO_(n) {
  var d = new Date();
  d.setDate(d.getDate() - n);
  return Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
}
function daysFromNowISO_(n) { return daysAgoISO_(-n); }

function createTrackerRowsForSpace_(trackerSheet, projectId, spaceId, mode) {
  var existingIds = rowsToObjects_(trackerSheet).map(function (r) { return r.TrackerID; });
  var counter = 0;
  SUBITEMS.forEach(function (si, idx) {
    counter++;
    var trackerId = nextId_('TR', existingIds.concat(['TR-' + String(counter).padStart(4, '0')]));
    trackerId = 'TR-' + Utilities.getUuid().slice(0, 6);
    var stages = buildDemoStages_(mode, idx);
    var rollup = computeRollup_(stages);
    trackerSheet.appendRow([trackerId, projectId, spaceId, si.category, si.subItem, JSON.stringify(stages), rollup, new Date()]);
  });
}

function buildDemoStages_(mode, idx) {
  // Produce a realistic mixed spread of statuses across the 6 stages.
  var pattern;
  if (mode === 'ahead') {
    var patterns = [
      ['Done', 'Done', 'Done', 'Done', 'Done', 'Done'],
      ['Done', 'Done', 'Done', 'Done', 'Done', 'In Progress'],
      ['Done', 'Done', 'Done', 'Done', 'In Progress', 'Not Started'],
      ['Done', 'Done', 'Done', 'In Progress', 'Not Started', 'Not Started'],
      ['Done', 'Done', 'Delayed', 'Not Started', 'Not Started', 'Not Started']
    ];
    pattern = patterns[idx % patterns.length];
  } else {
    var patterns2 = [
      ['Done', 'Done', 'In Progress', 'Not Started', 'Not Started', 'Not Started'],
      ['Done', 'Delayed', 'Not Started', 'Not Started', 'Not Started', 'Not Started'],
      ['Done', 'Done', 'Done', 'In Progress', 'Not Started', 'Not Started'],
      ['Done', 'In Progress', 'Not Started', 'Not Started', 'Not Started', 'Not Started'],
      ['Not Started', 'Not Started', 'Not Started', 'Not Started', 'Not Started', 'Not Started']
    ];
    pattern = patterns2[idx % patterns2.length];
  }
  var stages = [];
  STAGES.forEach(function (stageName, i) {
    var status = pattern[i];
    // Stage 0 target ~60 days ago, stage 5 target ~15 days from now (project timeline).
    var daysAgoForTarget = 60 - i * 15;
    var target = daysAgoISO_(daysAgoForTarget);
    var actual = (status === 'Done') ? daysAgoISO_(daysAgoForTarget + 2) : '';
    stages.push({ stage: stageName, status: status, target: target, actual: actual, note: '' });
  });
  return stages;
}

function computeRollup_(stages) {
  var hasDelayed = stages.some(function (s) { return s.status === 'Delayed'; });
  if (hasDelayed) return 'Delayed';
  var allDone = stages.every(function (s) { return s.status === 'Done' || s.status === 'N/A'; });
  if (allDone) return 'Done';
  var anyProgress = stages.some(function (s) { return s.status === 'In Progress'; });
  if (anyProgress) return 'In Progress';
  return 'Not Started';
}

function createCashflowRowsForSpace_(cashflowSheet, projectId, spaceId, receivedFraction) {
  CATEGORIES.forEach(function (cat) {
    var boq = Math.round((150000 + Math.random() * 250000) / 1000) * 1000;
    var invoiced = Math.round(boq * Math.min(1, receivedFraction + 0.15));
    var received = Math.round(boq * receivedFraction);
    var status = received >= invoiced ? 'Paid' : (invoiced > 0 ? 'Partial' : 'Pending');
    var cashflowId = 'CF-' + Utilities.getUuid().slice(0, 6);
    cashflowSheet.appendRow([
      cashflowId, projectId, spaceId, cat, boq, JSON.stringify(DEFAULT_SPLIT),
      invoiced, received, daysFromNowISO_(14), status, new Date()
    ]);
  });
}

// ---------- Web app entry ----------
// Frontend is hosted separately on GitHub Pages (pmc.ideaformdesignstudio.com) and
// calls this as a plain JSON API — same pattern as the DPR backend. GET for reads
// (?action=...), POST with a JSON body (routed by action) for writes. POST bodies
// must be sent with Content-Type: text/plain from the browser to avoid a CORS
// preflight (Apps Script doesn't answer OPTIONS requests).

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function safeRespond_(fn) {
  try { return respond_({ ok: true, data: fn() }); }
  catch (e) { return respond_({ ok: false, error: String(e && e.message || e) }); }
}

function doGet(e) {
  var p = e && e.parameter ? e.parameter : {};
  if (p.action === 'getAllData') return safeRespond_(getAllData);
  if (p.action === 'getSchema') return safeRespond_(getSchema);
  return respond_({ ok: false, error: 'Unknown action: ' + p.action });
}

function doPost(e) {
  var raw = e && e.postData ? e.postData.contents : '{}';
  var data;
  try { data = JSON.parse(raw || '{}'); } catch (err) { return respond_({ ok: false, error: 'Bad JSON body' }); }
  var routes = {
    addProject: addProject, updateProject: updateProject, addSpace: addSpace,
    updateStage: updateStage, updateCashflow: updateCashflow, addDailyLog: addDailyLog
  };
  var fn = routes[data.action];
  if (!fn) return respond_({ ok: false, error: 'Unknown action: ' + data.action });
  return safeRespond_(function () { return fn(data.payload || {}); });
}

// ---------- Client-callable API ----------

function getSchema() {
  return { stages: STAGES, subItems: SUBITEMS, categories: CATEGORIES, defaultSplit: DEFAULT_SPLIT };
}

function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Kolkata', 'yyyy-MM-dd');
  return v;
}

function getAllData() {
  var ss = getSS_();
  var projects = rowsToObjects_(getTab_(ss, 'PROJECTS')).map(function (r) {
    r.StartDate = normDate_(r.StartDate);
    r.TargetMoveIn = normDate_(r.TargetMoveIn);
    return r;
  });
  var spaces = rowsToObjects_(getTab_(ss, 'SPACES'));
  var tracker = rowsToObjects_(getTab_(ss, 'TRACKER')).map(function (r) {
    r.Stages = JSON.parse(r.StagesJSON || '[]').map(function (s) {
      s.target = normDate_(s.target);
      s.actual = normDate_(s.actual);
      return s;
    });
    return r;
  });
  var cashflow = rowsToObjects_(getTab_(ss, 'CASHFLOW')).map(function (r) {
    r.Split = JSON.parse(r.SplitJSON || '{}');
    r.DueDate = normDate_(r.DueDate);
    return r;
  });
  var dailyLog = rowsToObjects_(getTab_(ss, 'DAILY_LOG')).map(function (r) {
    r.Date = normDate_(r.Date);
    return r;
  });
  dailyLog.sort(function (a, b) { return new Date(b.Date) - new Date(a.Date) || new Date(b.CreatedAt) - new Date(a.CreatedAt); });

  return {
    schema: getSchema(),
    projects: projects,
    spaces: spaces,
    tracker: tracker,
    cashflow: cashflow,
    dailyLog: dailyLog
  };
}

function addProject(payload) {
  var ss = getSS_();
  var projSheet = getTab_(ss, 'PROJECTS');
  var existingIds = rowsToObjects_(projSheet).map(function (r) { return r.ProjectID; });
  var projectId = nextId_('P', existingIds);
  projSheet.appendRow([
    projectId, payload.name, payload.address || '', payload.startDate || '',
    payload.targetMoveIn || '', Number(payload.budget) || 0, payload.clientName || '', new Date()
  ]);

  var spaceSheet = getTab_(ss, 'SPACES');
  var trackerSheet = getTab_(ss, 'TRACKER');
  var cashflowSheet = getTab_(ss, 'CASHFLOW');
  var spaceNames = (payload.spaces && payload.spaces.length) ? payload.spaces : ['Drawing Room'];
  spaceNames.forEach(function (name, idx) {
    var spaceId = addSpaceInternal_(ss, spaceSheet, trackerSheet, cashflowSheet, projectId, name, idx + 1);
  });
  return { projectId: projectId };
}

function updateProject(payload) {
  var ss = getSS_();
  var sheet = getTab_(ss, 'PROJECTS');
  var row = findRowById_(sheet, 'ProjectID', payload.projectId);
  if (row < 0) throw new Error('Project not found');
  var headers = sheet.getDataRange().getValues()[0];
  var current = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  var map = { Name: 'name', Address: 'address', StartDate: 'startDate', TargetMoveIn: 'targetMoveIn', Budget: 'budget', ClientName: 'clientName' };
  headers.forEach(function (h, i) {
    if (map[h] && payload[map[h]] !== undefined && payload[map[h]] !== '') {
      current[i] = (h === 'Budget') ? Number(payload[map[h]]) : payload[map[h]];
    }
  });
  sheet.getRange(row, 1, 1, headers.length).setValues([current]);
  return { ok: true };
}

function addSpaceInternal_(ss, spaceSheet, trackerSheet, cashflowSheet, projectId, name, sortOrder) {
  var existingIds = rowsToObjects_(spaceSheet).map(function (r) { return r.SpaceID; });
  var spaceId = nextId_('S', existingIds);
  spaceSheet.appendRow([spaceId, projectId, name, sortOrder]);

  var existingTrackerIds = [];
  SUBITEMS.forEach(function (si) {
    var trackerId = 'TR-' + Utilities.getUuid().slice(0, 6);
    var stages = STAGES.map(function (stageName) {
      return { stage: stageName, status: 'Not Started', target: '', actual: '', note: '' };
    });
    trackerSheet.appendRow([trackerId, projectId, spaceId, si.category, si.subItem, JSON.stringify(stages), 'Not Started', new Date()]);
  });
  CATEGORIES.forEach(function (cat) {
    var cashflowId = 'CF-' + Utilities.getUuid().slice(0, 6);
    cashflowSheet.appendRow([cashflowId, projectId, spaceId, cat, 0, JSON.stringify(DEFAULT_SPLIT), 0, 0, '', 'Pending', new Date()]);
  });
  return spaceId;
}

function addSpace(payload) {
  var ss = getSS_();
  var spaceSheet = getTab_(ss, 'SPACES');
  var trackerSheet = getTab_(ss, 'TRACKER');
  var cashflowSheet = getTab_(ss, 'CASHFLOW');
  var existing = rowsToObjects_(spaceSheet).filter(function (r) { return r.ProjectID === payload.projectId; });
  var sortOrder = existing.length + 1;
  var spaceId = addSpaceInternal_(ss, spaceSheet, trackerSheet, cashflowSheet, payload.projectId, payload.name, sortOrder);
  return { spaceId: spaceId };
}

function updateStage(payload) {
  // payload: { trackerId, stage, status, target, actual, note }
  var ss = getSS_();
  var sheet = getTab_(ss, 'TRACKER');
  var row = findRowById_(sheet, 'TrackerID', payload.trackerId);
  if (row < 0) throw new Error('Tracker row not found');
  var headers = sheet.getDataRange().getValues()[0];
  var stagesColIdx = headers.indexOf('StagesJSON');
  var rollupColIdx = headers.indexOf('RollupStatus');
  var updatedColIdx = headers.indexOf('UpdatedAt');
  var stagesRaw = sheet.getRange(row, stagesColIdx + 1).getValue();
  var stages = JSON.parse(stagesRaw || '[]');
  var stageObj = stages.filter(function (s) { return s.stage === payload.stage; })[0];
  if (!stageObj) throw new Error('Stage not found');
  if (payload.status !== undefined) stageObj.status = payload.status;
  if (payload.target !== undefined) stageObj.target = payload.target;
  if (payload.actual !== undefined) stageObj.actual = payload.actual;
  if (payload.note !== undefined) stageObj.note = payload.note;

  var rollup = computeRollup_(stages);
  sheet.getRange(row, stagesColIdx + 1).setValue(JSON.stringify(stages));
  sheet.getRange(row, rollupColIdx + 1).setValue(rollup);
  sheet.getRange(row, updatedColIdx + 1).setValue(new Date());
  return { rollup: rollup, stages: stages };
}

function updateCashflow(payload) {
  // payload: { cashflowId, boqValue, split, invoiced, received, dueDate, status }
  var ss = getSS_();
  var sheet = getTab_(ss, 'CASHFLOW');
  var row = findRowById_(sheet, 'CashflowID', payload.cashflowId);
  if (row < 0) throw new Error('Cashflow row not found');
  var headers = sheet.getDataRange().getValues()[0];
  var current = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });
  if (payload.boqValue !== undefined) current[idx.BOQValue] = Number(payload.boqValue);
  if (payload.split !== undefined) current[idx.SplitJSON] = JSON.stringify(payload.split);
  if (payload.invoiced !== undefined) current[idx.Invoiced] = Number(payload.invoiced);
  if (payload.received !== undefined) current[idx.Received] = Number(payload.received);
  if (payload.dueDate !== undefined) current[idx.DueDate] = payload.dueDate;
  if (payload.status !== undefined) current[idx.PaymentStatus] = payload.status;
  current[idx.UpdatedAt] = new Date();
  sheet.getRange(row, 1, 1, headers.length).setValues([current]);
  return { ok: true };
}

function addDailyLog(payload) {
  // payload: { date, projectId, spaceId, entry, loggedBy, hasBlocker }
  var ss = getSS_();
  var sheet = getTab_(ss, 'DAILY_LOG');
  var logId = 'LOG-' + Utilities.getUuid().slice(0, 8);
  sheet.appendRow([
    logId, payload.date, payload.projectId, payload.spaceId || '',
    payload.entry, payload.loggedBy, !!payload.hasBlocker, new Date()
  ]);
  return { logId: logId };
}
