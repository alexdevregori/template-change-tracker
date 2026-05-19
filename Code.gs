// ============================================================
// Productboard Feature Hierarchy Tracker
// Generated against reference/entities.yaml on 2026-05-19
//
// SETUP:
//   1. Open your Google Sheet > Extensions > Apps Script
//   2. Paste this file as Code.gs
//   3. Project Settings > Script Properties > Add property:
//      Key: PB_TOKEN   Value: <your Productboard API token>
//   4. Set up two time-driven triggers:
//      - exportHierarchy()  → runs daily (e.g. 2–3 AM)
//      - buildChangelog()   → runs daily, a few minutes after exportHierarchy
//
// SHEETS CREATED AUTOMATICALLY:
//   "Current"   — latest hierarchy snapshot (built-in + custom field columns)
//   "Previous"  — snapshot from the prior run
//   "Changelog" — cumulative log of every detected change
// ============================================================

// --------------- MENU ---------------

/**
 * Runs automatically when the spreadsheet opens.
 * Adds a "Productboard" menu so all functions are accessible without opening Apps Script.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Productboard')
    .addItem('▶ Run Full Sync',          'runDailySync')
    .addSeparator()
    .addItem('Export Hierarchy',         'exportHierarchy')
    .addItem('Build Changelog',          'buildChangelog')
    .addSeparator()
    .addItem('⚙ Setup / Validate Token', 'setup')
    .addItem('🔍 Debug Diff',            'debugDiff')
    .addToUi();
}


// --------------- CONFIGURATION ---------------

var CONFIG = {
  TOKEN_KEY:       'PB_TOKEN',
  SHEET_CURRENT:   'Current',
  SHEET_PREVIOUS:  'Previous',
  SHEET_CHANGELOG: 'Changelog',
  API_BASE:        'https://api.productboard.com/v2',
  ENTITY_TYPES:    ['product', 'component', 'feature', 'subfeature'],
};

// Built-in columns always written first, in this order
var BASE_HEADERS = [
  'ID', 'Type', 'Name', 'Status', 'Owner Email',
  'Tags', 'Timeframe Start', 'Timeframe End',
  'Parent ID', 'Parent Type', 'Parent Name',
  'Archived', 'Created At', 'Updated At', 'PB URL'
];

// These columns are excluded from change detection
var DIFF_EXCLUDE = ['ID', 'Created At', 'Updated At', 'PB URL'];

var CHANGELOG_HEADERS = [
  'Logged At', 'Entity ID', 'Entity Name', 'Entity Type',
  'Change Type', 'Field', 'Old Value', 'New Value'
];

// UUID pattern — fields matching this are custom fields
var UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;


// --------------- SETUP ---------------

/**
 * Run once manually before scheduling triggers.
 * 1. Validates the PB_TOKEN script property against the live API.
 * 2. Creates Current, Previous, and Changelog sheets with correct headers.
 * Shows a dialog with the result so you don't have to read the logs.
 */
function setup() {
  var ui = SpreadsheetApp.getUi();

  // 1. Check token exists
  var token = PropertiesService.getScriptProperties().getProperty(CONFIG.TOKEN_KEY);
  if (!token) {
    ui.alert(
      'Setup failed',
      'No PB_TOKEN found in Script Properties.\n\n' +
      'Go to Project Settings → Script Properties and add:\n' +
      '  Key: PB_TOKEN\n  Value: <your Productboard API token>',
      ui.ButtonSet.OK
    );
    return;
  }

  // 2. Validate token with a lightweight API call
  var testUrl = CONFIG.API_BASE + '/entities/configurations?' +
    CONFIG.ENTITY_TYPES.map(function(t) { return 'type[]=' + encodeURIComponent(t); }).join('&');

  var response, body;
  try {
    response = apiGet_(testUrl, token);
    body = JSON.parse(response.getContentText());
  } catch (e) {
    ui.alert('Setup failed', 'API request threw an error:\n' + e.message, ui.ButtonSet.OK);
    return;
  }

  var code = response.getResponseCode();
  if (code === 401) {
    ui.alert('Setup failed',
      'Token was rejected by Productboard (HTTP 401).\n' +
      'Check that your PB_TOKEN is correct and has not expired.',
      ui.ButtonSet.OK);
    return;
  }
  if (code === 403) {
    ui.alert('Setup failed',
      'Token is valid but lacks permission to read entity configurations (HTTP 403).\n' +
      'Ensure the token has at least read access to entities.',
      ui.ButtonSet.OK);
    return;
  }
  if (code < 200 || code >= 300) {
    var detail = (body && body.errors && body.errors[0]) ? body.errors[0].detail : JSON.stringify(body);
    ui.alert('Setup failed', 'Unexpected API response ' + code + ':\n' + detail, ui.ButtonSet.OK);
    return;
  }

  // 3. Build field map and headers from the validated config response
  var fieldMap = parseFieldMap_(body);
  var headers  = buildHeaders_(fieldMap);

  // 4. Initialize sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  initSheet_(ss, CONFIG.SHEET_CURRENT,   headers);
  initSheet_(ss, CONFIG.SHEET_PREVIOUS,  headers);
  initSheet_(ss, CONFIG.SHEET_CHANGELOG, CHANGELOG_HEADERS);

  var customCount = Object.keys(fieldMap).length;
  ui.alert(
    'Setup complete',
    '✓ Token validated successfully.\n' +
    '✓ ' + customCount + ' custom field(s) discovered.\n' +
    '✓ Sheets initialized: Current, Previous, Changelog.\n\n' +
    'Next: add two daily time-driven triggers:\n' +
    '  1. exportHierarchy()\n' +
    '  2. buildChangelog()  (a few minutes after #1)',
    ui.ButtonSet.OK
  );
}

/**
 * Writes headers to a sheet, creating it if needed.
 * Does NOT clear existing data rows — safe to re-run after setup.
 */
function initSheet_(ss, sheetName, headers) {
  var sheet = getOrCreateSheet_(ss, sheetName);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Widen columns so headers are readable without manual resizing
  for (var i = 1; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 160);
  }
}


// --------------- ENTRY POINTS ---------------

/**
 * Trigger daily. On first run, bootstraps Previous so changelog starts from a
 * clean baseline (0 changes). On subsequent runs, rotates Current → Previous first.
 * Custom fields are discovered automatically from /entities/configurations.
 */
function exportHierarchy() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var firstRun   = isPreviousEmpty_(ss);
  var token      = getToken_();
  var fieldMap   = fetchFieldMap_(token);
  var headers    = buildHeaders_(fieldMap);

  if (!firstRun) {
    rotateSheets_(ss);
  }

  var entities = fetchAllEntities_(token);
  var rows     = entitiesToRows_(entities, fieldMap, headers);
  writeSheet_(ss, CONFIG.SHEET_CURRENT, headers, rows);

  if (firstRun) {
    bootstrapPrevious_(ss);
    Logger.log('First run: Previous bootstrapped as baseline. Changelog will show 0 changes.');
  }

  Logger.log('exportHierarchy complete: ' + rows.length + ' entities, ' +
             Object.keys(fieldMap).length + ' custom fields.');
}

/**
 * Trigger daily, a few minutes after exportHierarchy.
 * Diffs Current vs Previous and appends changes to Changelog.
 */
function buildChangelog() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var current  = readSheet_(ss, CONFIG.SHEET_CURRENT);
  var previous = readSheet_(ss, CONFIG.SHEET_PREVIOUS);

  if (previous.length === 0) {
    Logger.log('No Previous data yet — skipping changelog on first run.');
    return;
  }

  var changes = diffSheets_(current, previous);
  if (changes.length > 0) {
    appendChangelog_(ss, changes);
    Logger.log('buildChangelog: logged ' + changes.length + ' change(s).');
  } else {
    Logger.log('buildChangelog: no changes detected.');
  }
}

/**
 * Convenience wrapper: runs both steps in sequence from a single trigger.
 */
function runDailySync() {
  exportHierarchy();
  SpreadsheetApp.flush(); // ensure all sheet writes are committed before reading in buildChangelog
  buildChangelog();
}

/**
 * Diagnostic tool: logs exactly what the diff sees without writing to Changelog.
 * Run this from the menu if changes aren't appearing — check View → Logs afterwards.
 */
function debugDiff() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var current  = readSheet_(ss, CONFIG.SHEET_CURRENT);
  var previous = readSheet_(ss, CONFIG.SHEET_PREVIOUS);

  Logger.log('Current rows: ' + current.length + ', Previous rows: ' + previous.length);

  if (current.length === 0 || previous.length === 0) {
    Logger.log('One or both sheets are empty — nothing to compare.');
    return;
  }

  var diffFields = Object.keys(current[0]).filter(function(h) {
    return DIFF_EXCLUDE.indexOf(h) === -1;
  });
  Logger.log('Diff fields: ' + diffFields.join(', '));

  // Log Status values for every entity where the two sheets disagree on any field
  var currentById  = indexById_(current);
  var previousById = indexById_(previous);
  var mismatches   = 0;

  Object.keys(currentById).forEach(function(id) {
    if (!previousById[id]) return;
    var curr = currentById[id];
    var prev = previousById[id];
    diffFields.forEach(function(field) {
      var oldVal = prev[field] !== undefined ? prev[field] : '';
      var newVal = curr[field] !== undefined ? curr[field] : '';
      if (oldVal !== newVal) {
        Logger.log('DIFF  id=' + id + '  field=' + field +
                   '  prev=' + JSON.stringify(oldVal) +
                   '  curr=' + JSON.stringify(newVal));
        mismatches++;
      }
    });
  });

  // Also log Status values for the first 3 entities so we can sanity-check
  Logger.log('--- Status sample (first 3 entities by ID) ---');
  Object.keys(currentById).slice(0, 3).forEach(function(id) {
    var currStatus = currentById[id]['Status'] || '(empty)';
    var prevStatus = previousById[id] ? (previousById[id]['Status'] || '(empty)') : '(not in previous)';
    Logger.log('id=' + id + '  curr.Status=' + JSON.stringify(currStatus) +
               '  prev.Status=' + JSON.stringify(prevStatus));
  });

  var changes = diffSheets_(current, previous);
  Logger.log('Total changes found: ' + changes.length);
  changes.slice(0, 20).forEach(function(c) { Logger.log(JSON.stringify(c)); });

  SpreadsheetApp.getUi().alert(
    'Debug Diff',
    'Field mismatches: ' + mismatches + '\nChanges found: ' + changes.length +
    '\nCheck View → Logs for details.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// --------------- FIELD MAP ---------------

/**
 * Fetches /entities/configurations and returns { uuid → 'Display Name' }
 * for every custom field across all tracked entity types.
 */
function fetchFieldMap_(token) {
  var url = CONFIG.API_BASE + '/entities/configurations?'
    + CONFIG.ENTITY_TYPES.map(function(t) { return 'type[]=' + encodeURIComponent(t); }).join('&');

  var response = apiGet_(url, token);
  var body     = JSON.parse(response.getContentText());
  assertOk_(response, body);

  var map = parseFieldMap_(body);
  Logger.log('Discovered ' + Object.keys(map).length + ' custom field(s).');
  return map;
}

/**
 * Extracts the { uuid → 'Display Name' } map from a parsed configurations response body.
 * Separated from fetchFieldMap_ so setup() can reuse the response it already validated.
 */
function parseFieldMap_(body) {
  var map = {};
  (body.data || []).forEach(function(config) {
    var fields = config.fields || {};
    Object.keys(fields).forEach(function(fieldId) {
      if (UUID_RE.test(fieldId) && !map[fieldId]) {
        map[fieldId] = fields[fieldId].name || fieldId;
      }
    });
  });
  return map;
}

/**
 * Returns the full ordered header array: base fields followed by
 * custom fields sorted alphabetically by display name.
 */
function buildHeaders_(fieldMap) {
  var customNames = Object.keys(fieldMap)
    .map(function(uuid) { return { uuid: uuid, name: fieldMap[uuid] }; })
    .sort(function(a, b) { return a.name.localeCompare(b.name); })
    .map(function(x) { return x.name; });

  return BASE_HEADERS.concat(customNames);
}


// --------------- SHEET ROTATION ---------------

/** Returns true if the Previous sheet has no data rows (first ever run). */
function isPreviousEmpty_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_PREVIOUS);
  return !sheet || sheet.getLastRow() <= 1;
}

/**
 * Copies Current → Previous verbatim.
 * Used on first run to establish a baseline so changelog starts clean.
 */
function bootstrapPrevious_(ss) {
  var current  = ss.getSheetByName(CONFIG.SHEET_CURRENT);
  var previous = getOrCreateSheet_(ss, CONFIG.SHEET_PREVIOUS);
  previous.clearContents();
  var data = current.getDataRange().getValues();
  if (data.length > 0) {
    previous.getRange(1, 1, data.length, data[0].length).setValues(data);
  }
}

function rotateSheets_(ss) {
  var currentSheet = ss.getSheetByName(CONFIG.SHEET_CURRENT);

  if (!currentSheet || currentSheet.getLastRow() <= 1) {
    Logger.log('Nothing in Current to rotate — skipping.');
    return;
  }

  var previousSheet = getOrCreateSheet_(ss, CONFIG.SHEET_PREVIOUS);
  previousSheet.clearContents();

  var data = currentSheet.getDataRange().getValues();
  previousSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  Logger.log('Rotated Current (' + (data.length - 1) + ' rows) → Previous.');
}


// --------------- API ---------------

function getToken_() {
  var token = PropertiesService.getScriptProperties().getProperty(CONFIG.TOKEN_KEY);
  if (!token) throw new Error('Script Property "PB_TOKEN" is not set.');
  return token;
}

function fetchAllEntities_(token) {
  var all = [];

  CONFIG.ENTITY_TYPES.forEach(function(type) {
    var url = CONFIG.API_BASE + '/entities?type[]=' + encodeURIComponent(type);
    while (url) {
      var response = apiGet_(url, token);
      var body     = JSON.parse(response.getContentText());
      assertOk_(response, body);
      (body.data || []).forEach(function(e) { all.push(e); });
      url = (body.links && body.links.next) ? body.links.next : null;
    }
    Logger.log('Fetched type=' + type + ': running total=' + all.length);
  });

  return all;
}

function apiGet_(url, token) {
  var options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + token,
      'X-Version': '1',
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  var backoff = 1000;
  for (var attempt = 0; attempt < 4; attempt++) {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 429) return response;
    Utilities.sleep(backoff);
    backoff = Math.min(backoff * 2, 30000);
  }
  throw new Error('Rate-limit retries exhausted for: ' + url);
}

function assertOk_(response, body) {
  var code = response.getResponseCode();
  if (code >= 200 && code < 300) return;
  var detail = (body && body.errors && body.errors[0] && body.errors[0].detail)
    ? body.errors[0].detail
    : JSON.stringify(body);
  throw new Error('Productboard API ' + code + ': ' + detail);
}


// --------------- DATA TRANSFORMATION ---------------

function entitiesToRows_(entities, fieldMap, headers) {
  // Build ID → entity map for parent name resolution
  var byId = {};
  entities.forEach(function(e) { byId[e.id] = e; });

  // Invert fieldMap to displayName → uuid for header-order lookup
  var nameToUuid = {};
  Object.keys(fieldMap).forEach(function(uuid) {
    nameToUuid[fieldMap[uuid]] = uuid;
  });

  return entities.map(function(e) {
    var f = e.fields || {};

    // ---- built-in fields ----
    var name       = typeof f.name === 'string' ? f.name : '';
    var status     = (f.status && f.status.name) ? f.status.name : '';
    var ownerEmail = (f.owner  && f.owner.email)  ? f.owner.email  : '';
    var archived   = f.archived === true ? 'true' : 'false';

    var tags = Array.isArray(f.tags)
      ? f.tags.map(function(t) { return t.name; }).join(', ')
      : '';

    var tfStart = (f.timeframe && f.timeframe.startDate) ? f.timeframe.startDate : '';
    var tfEnd   = (f.timeframe && f.timeframe.endDate)   ? f.timeframe.endDate   : '';

    var parentId = '', parentType = '', parentName = '';
    var rels = (e.relationships && e.relationships.data) ? e.relationships.data : [];
    for (var i = 0; i < rels.length; i++) {
      if (rels[i].type === 'parent' && rels[i].target) {
        parentId   = rels[i].target.id   || '';
        parentType = rels[i].target.type || '';
        var parentEntity = byId[parentId];
        if (parentEntity && typeof parentEntity.fields.name === 'string') {
          parentName = parentEntity.fields.name;
        }
        break;
      }
    }

    var url = (e.links && e.links.html) ? e.links.html : '';

    // Base values in BASE_HEADERS order
    var baseValues = [
      e.id || '', e.type || '', name, status, ownerEmail,
      tags, tfStart, tfEnd,
      parentId, parentType, parentName,
      archived, e.createdAt || '', e.updatedAt || '', url
    ];

    // Custom field values — one per header beyond the base headers
    var customValues = headers.slice(BASE_HEADERS.length).map(function(displayName) {
      var uuid = nameToUuid[displayName];
      return uuid ? serializeFieldValue_(f[uuid]) : '';
    });

    return baseValues.concat(customValues);
  });
}

/**
 * Converts any Productboard field value to a plain string for sheet storage.
 */
function serializeFieldValue_(val) {
  if (val === null || val === undefined) return '';
  if (typeof val === 'string')  return val;
  if (typeof val === 'number')  return String(val);
  if (typeof val === 'boolean') return String(val);

  // { name: '...', id: '...' }  — single-select, status
  if (val.name !== undefined) return val.name;

  // { value: ... }
  if (val.value !== undefined) return String(val.value);

  // { text: '...' }  — rich text
  if (val.text !== undefined) return val.text;

  // { url: '...' }
  if (val.url !== undefined) return val.url;

  // { startDate, endDate }  — timeframe
  if (val.startDate !== undefined) {
    return [val.startDate, val.endDate].filter(Boolean).join(' → ');
  }

  // Array — multi-select, teams, tags
  if (Array.isArray(val)) {
    return val.map(function(item) {
      return item && item.name ? item.name : String(item);
    }).join(', ');
  }

  return JSON.stringify(val);
}


// --------------- SHEET READ / WRITE ---------------

function writeSheet_(ss, sheetName, headers, rows) {
  var sheet = getOrCreateSheet_(ss, sheetName);
  sheet.clearContents();

  var allRows = [headers].concat(rows);
  sheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  Logger.log('Wrote ' + rows.length + ' rows to "' + sheetName + '".');
}

/**
 * Returns an array of plain objects keyed by column header.
 * Returns [] if the sheet doesn't exist or has no data rows.
 */
function readSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  // getDisplayValues returns every cell as the string the user sees —
  // avoids type-coercion surprises where getValues() returns Date objects
  // or numbers that String() formats differently across Current vs Previous.
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  var rows = dataRange.getDisplayValues();

  return rows.map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      obj[h.trim()] = row[i] ? row[i].trim() : '';
    });
    return obj;
  });
}

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    Logger.log('Created sheet "' + name + '".');
  }
  return sheet;
}


// --------------- DIFF ---------------

function diffSheets_(currentRows, previousRows) {
  var changes = [];
  var now = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssZ"
  );

  var currentById  = indexById_(currentRows);
  var previousById = indexById_(previousRows);

  // Derive which fields to compare from the Current rows
  // (includes custom fields automatically — excludes noise columns)
  var diffFields = currentRows.length > 0
    ? Object.keys(currentRows[0]).filter(function(h) {
        return DIFF_EXCLUDE.indexOf(h) === -1;
      })
    : [];

  // Added
  Object.keys(currentById).forEach(function(id) {
    if (!previousById[id]) {
      var r = currentById[id];
      changes.push([now, id, r['Name'] || '', r['Type'] || '', 'Added', '', '', '']);
    }
  });

  // Removed
  Object.keys(previousById).forEach(function(id) {
    if (!currentById[id]) {
      var r = previousById[id];
      changes.push([now, id, r['Name'] || '', r['Type'] || '', 'Removed', '', '', '']);
    }
  });

  // Modified — compare all diffFields for entities present in both
  Object.keys(currentById).forEach(function(id) {
    if (!previousById[id]) return;
    var curr = currentById[id];
    var prev = previousById[id];

    diffFields.forEach(function(field) {
      var oldVal = prev[field] !== undefined ? prev[field] : '';
      var newVal = curr[field] !== undefined ? curr[field] : '';
      if (oldVal !== newVal) {
        changes.push([now, id, curr['Name'] || '', curr['Type'] || '',
                      'Modified', field, oldVal, newVal]);
      }
    });
  });

  return changes;
}

function indexById_(rows) {
  var map = {};
  rows.forEach(function(r) { if (r['ID']) map[r['ID']] = r; });
  return map;
}


// --------------- CHANGELOG APPEND ---------------

function appendChangelog_(ss, changes) {
  var sheet = getOrCreateSheet_(ss, CONFIG.SHEET_CHANGELOG);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CHANGELOG_HEADERS.length).setValues([CHANGELOG_HEADERS]);
    sheet.getRange(1, 1, 1, CHANGELOG_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, changes.length, CHANGELOG_HEADERS.length).setValues(changes);
}
