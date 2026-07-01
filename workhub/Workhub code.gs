/* ════════════════════════════════════════════════════════════════
   WORK HUB API — separate, private Apps Script project.
   Deploy this as ITS OWN Web App (its own URL), completely
   independent from your main platform's Apps Script/deployment.

   It NEVER modifies your main spreadsheet. It only reads from it.
   All private Work Hub data lives in a separate spreadsheet.
   ════════════════════════════════════════════════════════════════ */

/* ─── CONFIGURATION ─── */

// Your existing main platform spreadsheet (READ-ONLY from here).
var MAIN_SHEET_ID = '1U_UULtDDI9gbBSbORba-0I9z5068vXGC2_ntpoHcGRA';

// NEW, separate spreadsheet for private Work Hub data.
// Create a blank Google Sheet, copy its ID from the URL, paste it here.
var WORK_SHEET_ID = 'PASTE_YOUR_NEW_PRIVATE_SHEET_ID_HERE';

/* Access token check.
   Set this in Apps Script: Project Settings → Script Properties → add
   key "WORK_HUB_TOKEN" with a long random value. Never hardcode it here. */
function getValidToken() {
  return PropertiesService.getScriptProperties().getProperty('WORK_HUB_TOKEN');
}

/* ─── HELPERS ─── */

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function denied() {
  return jsonOut({ error: 'unauthorized' });
}

function checkToken(token) {
  var valid = getValidToken();
  return valid && token && token === valid;
}

function mainSheet(name) {
  return SpreadsheetApp.openById(MAIN_SHEET_ID).getSheetByName(name);
}

function workSheet(name, headers) {
  var ss = SpreadsheetApp.openById(WORK_SHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readRows(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = { _row: i + 1 };
    for (var j = 0; j < headers.length; j++) row[headers[j]] = data[i][j];
    rows.push(row);
  }
  return rows;
}

/* Work Hub tab definitions (auto-created on first use) */
var WORK_TABS = {
  'My Tasks':         ['ID','Title','Status','Note','Due Date','Created','Updated','Completed Date'],
  'Layout Work':       ['ID','Branch Code','Item','Status','Notes','ETA','Created','Updated'],
  'Receiving Work':    ['ID','Branch Code','Asset','Status','Notes','ETA','Created','Updated'],
  'Notes':             ['ID','Branch Code','Note','Created'],
  'Activity Log':      ['ID','Timestamp','Branch Code','Action Type','Details'],
  'Time Tracking':     ['ID','Branch Code','Task','Start','End','Duration Minutes','Notes'],
  'Personal Reports':  ['ID','Period','Branch Code','Summary','Created'],
  'Procurement Followup': ['ID','Branch Code','Asset Group','Owner','Request Number','Request Date','Status','ETA','Remarks','Created','Updated'],
  'Delivery Notes Followup': ['ID','Branch Code','D.N. Link','Status','Notes','Created','Updated']
};

/* ─── doGet ─── */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!checkToken(p.token)) return denied();

  var action = p.action || '';

  // Read-only pull from MAIN platform: branch list + status + links
  if (action === 'getBranches') {
    var branchSheet = mainSheet('Branch Codes');
    var branches = readRows(branchSheet); // Branch Code / Branch Status / Receiving Date / Live Date
    var linksSheet = mainSheet('Links');
    var links = readRows(linksSheet);     // Branch Code / D.N. Folder / Layout link / Asset Link
    var linksByCode = {};
    links.forEach(function (l) { linksByCode[String(l['Branch Code']).trim()] = l; });

    var merged = branches.map(function (b) {
      var code = String(b['Branch Code']).trim();
      var link = linksByCode[code] || {};
      return {
        branchCode: code,
        status: b['Branch Status'] || '',
        receivingDate: b['Receiving Date'] || '',
        liveDate: b['Live Date'] || '',
        dnFolder: link['D.N. Folder'] || '',
        layoutLink: link['Layout link'] || '',
        assetLink: link['Asset Link'] || ''
        // city / region not present in main sheet yet — add columns there
        // and they will need a matching field added here.
      };
    });
    return jsonOut({ rows: merged });
  }

  // Generic read-only passthrough to any MAIN platform tab (e.g. Receiving Checklist)
  if (action === 'getMainTab') {
    var mTab = p.tab;
    var mSheet = mainSheet(mTab);
    if (!mSheet) return jsonOut({ error: 'Main sheet not found: ' + mTab });
    var mRows = readRows(mSheet);
    if (p.branch) {
      var branchCol = 'Branch Code';
      mRows = mRows.filter(function (r) { return String(r[branchCol]).trim() === String(p.branch).trim(); });
    }
    return jsonOut({ rows: mRows });
  }

  // Generic private tab read
  if (action === 'getTab') {
    var tabName = p.tab;
    if (!WORK_TABS[tabName]) return jsonOut({ error: 'Unknown private tab: ' + tabName });
    var sheet = workSheet(tabName, WORK_TABS[tabName]);
    var rows = readRows(sheet);
    if (p.branch) rows = rows.filter(function (r) { return String(r['Branch Code']).trim() === String(p.branch).trim(); });
    return jsonOut({ rows: rows });
  }

  return jsonOut({ error: 'Unknown action: ' + action });
}

/* ─── doPost ─── */

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    if (!checkToken(payload.token)) return denied();

    var action = payload.action || '';
    if (action === 'saveRow')   return saveRow(payload);
    if (action === 'deleteRow') return deleteRow(payload);
    return jsonOut({ error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

/* Generic upsert into any private tab.
   payload = { tab, id (optional, blank = new), fields: {col: value, ...} } */
function saveRow(payload) {
  var tabName = payload.tab;
  if (!WORK_TABS[tabName]) return jsonOut({ error: 'Unknown private tab: ' + tabName });
  var headers = WORK_TABS[tabName];
  var sheet = workSheet(tabName, headers);
  var now = new Date().toLocaleString();

  var fields = payload.fields || {};
  if (headers.indexOf('Updated') !== -1) fields['Updated'] = now;

  if (payload.id) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(payload.id).trim()) {
        headers.forEach(function (h, j) {
          if (fields.hasOwnProperty(h)) sheet.getRange(i + 1, j + 1).setValue(fields[h]);
        });
        return jsonOut({ success: true, id: payload.id });
      }
    }
  }

  // New row
  var id = tabName.replace(/\s+/g, '').toUpperCase() + '_' + Date.now();
  var newRow = headers.map(function (h) {
    if (h === 'ID') return id;
    if (h === 'Created') return now;
    if (h === 'Timestamp') return now;
    return fields.hasOwnProperty(h) ? fields[h] : '';
  });
  sheet.appendRow(newRow);
  return jsonOut({ success: true, id: id, created: true });
}

function deleteRow(payload) {
  var tabName = payload.tab;
  if (!WORK_TABS[tabName]) return jsonOut({ error: 'Unknown private tab: ' + tabName });
  var sheet = workSheet(tabName, WORK_TABS[tabName]);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.id).trim()) {
      sheet.deleteRow(i + 1);
      return jsonOut({ success: true });
    }
  }
  return jsonOut({ error: 'Not found' });
}
