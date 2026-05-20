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
    .addItem('🔔 Slack Notifications...', 'openSlackConfigDialog')
    .addItem('🔍 Debug Diff',            'debugDiff')
    .addToUi();
}


// --------------- CONFIGURATION ---------------

var CONFIG = {
  TOKEN_KEY:        'PB_TOKEN',
  SHEET_CURRENT:    'Current',
  SHEET_PREVIOUS:   'Previous',
  SHEET_CHANGELOG:  'Changelog',
  API_BASE:         'https://api.productboard.com/v2',
  ENTITY_TYPES:     ['product', 'component', 'feature', 'subfeature'],
  SLACK_TOKEN_KEY:  'SLACK_BOT_TOKEN',
  SLACK_CHANNEL_KEY:'SLACK_CHANNEL_ID',
  SLACK_FIELDS_KEY: 'SLACK_NOTIFY_FIELDS',
  SLACK_MAX_BLOCKS: 46,
  SLACK_API_POST:   'https://slack.com/api/chat.postMessage',
};

// Built-in columns always written first, in this order
var BASE_HEADERS = [
  'ID', 'Type', 'Name', 'Status', 'Owner Email',
  'Tags', 'Timeframe Start', 'Timeframe End',
  'Parent ID', 'Parent Type', 'Parent Name',
  'Releases', 'Objectives', 'Initiatives',
  'Archived', 'Created At', 'Updated At', 'PB URL'
];

// Entity types fetched only for name lookup (not written as rows to the sheet)
var LOOKUP_TYPES = ['release', 'objective', 'initiative'];

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

  var entities   = fetchAllEntities_(token);
  resolveRelationshipPagination_(entities, token);
  var nameLookup = fetchNameLookup_(token);
  var rows       = entitiesToRows_(entities, fieldMap, headers, nameLookup);
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
    sendSlackNotifications_(changes);
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

  // Check for duplicate column headers — a collision causes readSheet_ to
  // overwrite built-in field values with custom field values of the same name
  var rawHeaders = ss.getSheetByName(CONFIG.SHEET_CURRENT)
    .getRange(1, 1, 1, ss.getSheetByName(CONFIG.SHEET_CURRENT).getLastColumn())
    .getValues()[0].map(String);
  var headerCounts = {};
  rawHeaders.forEach(function(h) { headerCounts[h] = (headerCounts[h] || 0) + 1; });
  var dupes = Object.keys(headerCounts).filter(function(h) { return headerCounts[h] > 1; });
  if (dupes.length > 0) {
    Logger.log('DUPLICATE HEADERS (causing silent overwrites): ' + dupes.join(', '));
  } else {
    Logger.log('No duplicate headers found.');
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

  // Log Status values for the first 3 entities that have a non-empty Status in either sheet
  Logger.log('--- Status sample (first 3 entities with a Status value) ---');
  var statusSampleCount = 0;
  Object.keys(currentById).forEach(function(id) {
    if (statusSampleCount >= 3) return;
    var currStatus = currentById[id]['Status'] || '';
    var prevStatus = previousById[id] ? (previousById[id]['Status'] || '') : '';
    if (currStatus || prevStatus) {
      Logger.log('id=' + id + '  type=' + (currentById[id]['Type'] || '') +
                 '  curr.Status=' + JSON.stringify(currStatus) +
                 '  prev.Status=' + JSON.stringify(prevStatus));
      statusSampleCount++;
    }
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
 *
 * If a custom field's display name collides with a built-in column name,
 * it is prefixed with "[Custom] " to prevent readSheet_() from overwriting
 * the built-in value when building row objects keyed by header name.
 * fieldMap is updated in-place so entitiesToRows_() uses the same names.
 */
function buildHeaders_(fieldMap) {
  var customEntries = Object.keys(fieldMap)
    .map(function(uuid) { return { uuid: uuid, name: fieldMap[uuid] }; })
    .sort(function(a, b) { return a.name.localeCompare(b.name); });

  var customNames = customEntries.map(function(x) {
    var name = x.name;
    if (BASE_HEADERS.indexOf(name) !== -1) {
      name = '[Custom] ' + name;
      fieldMap[x.uuid] = name;
    }
    return name;
  });

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

/**
 * Fetches releases, objectives, and initiatives and returns { id → name }.
 * Used to resolve relationship target names without adding those entity
 * types as rows in the hierarchy sheet.
 */
function fetchNameLookup_(token) {
  var lookup = {};

  LOOKUP_TYPES.forEach(function(type) {
    var url = CONFIG.API_BASE + '/entities?type[]=' + encodeURIComponent(type);
    while (url) {
      var response = apiGet_(url, token);
      var code     = response.getResponseCode();
      var body     = JSON.parse(response.getContentText());

      // Some workspaces don't have all entity types enabled — skip gracefully
      if (code === 400 || code === 404) {
        Logger.log('Skipping lookup type "' + type + '" — not available in this workspace (HTTP ' + code + ').');
        break;
      }

      assertOk_(response, body);
      (body.data || []).forEach(function(e) {
        lookup[e.id] = (e.fields && typeof e.fields.name === 'string') ? e.fields.name : e.id;
      });
      url = (body.links && body.links.next) ? body.links.next : null;
    }
  });

  Logger.log('Fetched ' + Object.keys(lookup).length + ' lookup entities (releases/objectives/initiatives).');
  return lookup;
}

/**
 * For any entity whose inline relationships were truncated (links.next is set),
 * fetches remaining pages from GET /entities/{id}/relationships and appends
 * them to e.relationships.data in-place so entitiesToRows_ sees a complete list.
 */
function resolveRelationshipPagination_(entities, token) {
  var extraCalls = 0;

  entities.forEach(function(e) {
    var nextUrl = (e.relationships && e.relationships.links && e.relationships.links.next)
      ? e.relationships.links.next
      : null;

    while (nextUrl) {
      var response = apiGet_(nextUrl, token);
      var body     = JSON.parse(response.getContentText());
      assertOk_(response, body);

      (body.data || []).forEach(function(rel) {
        e.relationships.data.push(rel);
      });

      nextUrl = (body.links && body.links.next) ? body.links.next : null;
      extraCalls++;
    }
  });

  if (extraCalls > 0) {
    Logger.log('Resolved relationship pagination: ' + extraCalls + ' extra API call(s).');
  }
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

function entitiesToRows_(entities, fieldMap, headers, nameLookup) {
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
    var releases = [], objectives = [], initiatives = [];
    var rels = (e.relationships && e.relationships.data) ? e.relationships.data : [];

    rels.forEach(function(rel) {
      if (!rel.target) return;
      var targetId   = rel.target.id   || '';
      var targetType = rel.target.type || '';

      if (rel.type === 'parent') {
        parentId   = targetId;
        parentType = targetType;
        var parentEntity = byId[targetId];
        if (parentEntity && typeof parentEntity.fields.name === 'string') {
          parentName = parentEntity.fields.name;
        }
      } else if (rel.type === 'link') {
        var linkedName = nameLookup[targetId] || targetId;
        if (targetType === 'release')    releases.push(linkedName);
        else if (targetType === 'objective')  objectives.push(linkedName);
        else if (targetType === 'initiative') initiatives.push(linkedName);
      }
    });

    var url = (e.links && e.links.html) ? e.links.html : '';

    // Base values in BASE_HEADERS order
    var baseValues = [
      e.id || '', e.type || '', name, status, ownerEmail,
      tags, tfStart, tfEnd,
      parentId, parentType, parentName,
      releases.join(', '), objectives.join(', '), initiatives.join(', '),
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

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];

  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      obj[String(h)] = row[i] !== undefined && row[i] !== null ? String(row[i]) : '';
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
      diffFields.forEach(function(field) {
        var val = r[field] !== undefined ? r[field] : '';
        if (val !== '') {
          changes.push([now, id, r['Name'] || '', r['Type'] || '', 'Added', field, '', val]);
        }
      });
    }
  });

  // Removed
  Object.keys(previousById).forEach(function(id) {
    if (!currentById[id]) {
      var r = previousById[id];
      changes.push([now, id, r['Name'] || '', r['Type'] || '', 'Removed', '', '', '']);
      diffFields.forEach(function(field) {
        var val = r[field] !== undefined ? r[field] : '';
        if (val !== '') {
          changes.push([now, id, r['Name'] || '', r['Type'] || '', 'Removed', field, val, '']);
        }
      });
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


// --------------- SLACK NOTIFICATIONS ---------------

/**
 * Opens the Slack configuration dialog from the menu.
 */
function openSlackConfigDialog() {
  var html = HtmlService.createHtmlOutputFromFile('SlackConfig')
    .setWidth(520)
    .setHeight(620)
    .setTitle('Slack Notifications');
  SpreadsheetApp.getUi().showModalDialog(html, 'Slack Notifications');
}

/**
 * Called from SlackConfig.html on load.
 * Returns current saved config + the list of fields available to watch.
 */
function getSlackConfigData() {
  var props = PropertiesService.getScriptProperties();
  var rawFields = props.getProperty(CONFIG.SLACK_FIELDS_KEY);
  var notifyFields = [];
  try { notifyFields = rawFields ? JSON.parse(rawFields) : []; } catch (e) {}

  // Build available fields from live sheet headers when possible
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getSheetByName(CONFIG.SHEET_CURRENT);
  var baseFields;
  if (currentSheet && currentSheet.getLastRow() >= 1) {
    baseFields = currentSheet.getRange(1, 1, 1, currentSheet.getLastColumn())
      .getValues()[0]
      .map(String)
      .filter(function(h) { return h && DIFF_EXCLUDE.indexOf(h) === -1; });
  } else {
    baseFields = BASE_HEADERS.filter(function(h) { return DIFF_EXCLUDE.indexOf(h) === -1; });
  }

  return {
    botToken:        props.getProperty(CONFIG.SLACK_TOKEN_KEY)   || '',
    channelId:       props.getProperty(CONFIG.SLACK_CHANNEL_KEY) || '',
    notifyFields:    notifyFields,
    availableFields: ['Added', 'Removed'].concat(baseFields),
  };
}

/**
 * Called from SlackConfig.html on save.
 */
function saveSlackConfig(botToken, channelId, selectedFields) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty(CONFIG.SLACK_TOKEN_KEY,   botToken   || '');
    props.setProperty(CONFIG.SLACK_CHANNEL_KEY, channelId  || '');
    props.setProperty(CONFIG.SLACK_FIELDS_KEY,  JSON.stringify(selectedFields || []));
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/**
 * Called from SlackConfig.html to verify the connection without saving.
 */
function testSlackMessage(botToken, channelId) {
  if (!botToken || !channelId) {
    return { ok: false, error: 'Bot token and channel are required.' };
  }
  try {
    var payload = JSON.stringify({
      channel: channelId,
      text: 'Test from Productboard Tracker — Slack is connected.',
    });
    var response = UrlFetchApp.fetch(CONFIG.SLACK_API_POST, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + botToken },
      payload: payload,
      muteHttpExceptions: true,
    });
    var body = JSON.parse(response.getContentText());
    if (body.ok) return { ok: true };
    return { ok: false, error: body.error || 'Slack returned ok:false' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/**
 * Returns true if a change row should generate a Slack notification
 * given the user's saved field preferences.
 *
 * change row: [timestamp, entityId, entityName, entityType, changeType, fieldName, oldVal, newVal]
 */
function isNotifiable_(changeRow, notifyFields) {
  var changeType = changeRow[4];
  var fieldName  = changeRow[5];

  if (changeType === 'Added' || changeType === 'Removed') {
    // Only the entity-level summary row (empty fieldName) triggers a notification.
    // Field-detail rows for Added/Removed are suppressed to avoid flooding.
    if (fieldName !== '') return false;
    return notifyFields.indexOf(changeType) !== -1;
  }

  if (changeType === 'Modified') {
    return notifyFields.indexOf(fieldName) !== -1;
  }

  return false;
}

/**
 * Escapes Slack mrkdwn special characters in a string and truncates if needed.
 */
function slackEscape_(str) {
  var s = String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  if (s.length > 300) s = s.slice(0, 297) + '...';
  return s;
}

/**
 * Builds a Slack Block Kit blocks array from the filtered changes.
 * Groups by entity; each entity gets a section + optional field list + divider.
 */
function buildSlackBlocks_(filteredChanges) {
  // Group rows by entityId, preserving first-appearance order
  var order = [];
  var groups = {};
  filteredChanges.forEach(function(row) {
    var id = row[1];
    if (!groups[id]) {
      groups[id] = [];
      order.push(id);
    }
    groups[id].push(row);
  });

  var blocks = [];

  order.forEach(function(id, idx) {
    if (blocks.length >= CONFIG.SLACK_MAX_BLOCKS - 2) return; // leave room for truncation notice

    var rows      = groups[id];
    var firstRow  = rows[0];
    var name      = slackEscape_(firstRow[2]);
    var type      = slackEscape_(firstRow[3]);
    var changeType = firstRow[4];

    var headline = '*' + name + '* · ' + type;
    if (changeType === 'Added')   headline += '  _(Added)_';
    if (changeType === 'Removed') headline += '  _(Removed)_';

    var modifiedRows = rows.filter(function(r) { return r[4] === 'Modified'; });

    var sectionBlock = { type: 'section', text: { type: 'mrkdwn', text: headline } };

    if (modifiedRows.length > 0) {
      // Slack section fields max is 10 items
      var fieldItems = modifiedRows.slice(0, 10).map(function(r) {
        return {
          type: 'mrkdwn',
          text: '*' + slackEscape_(r[5]) + '*\n' + slackEscape_(r[6]) + ' → ' + slackEscape_(r[7]),
        };
      });
      sectionBlock.fields = fieldItems;

      // If more than 10 modified fields, emit overflow sections
      var overflow = modifiedRows.slice(10);
      while (overflow.length > 0 && blocks.length < CONFIG.SLACK_MAX_BLOCKS - 2) {
        blocks.push(sectionBlock);
        if (idx < order.length - 1) blocks.push({ type: 'divider' });
        var chunk = overflow.splice(0, 10);
        sectionBlock = {
          type: 'section',
          fields: chunk.map(function(r) {
            return {
              type: 'mrkdwn',
              text: '*' + slackEscape_(r[5]) + '*\n' + slackEscape_(r[6]) + ' → ' + slackEscape_(r[7]),
            };
          }),
        };
      }
    }

    blocks.push(sectionBlock);
    if (idx < order.length - 1) blocks.push({ type: 'divider' });
  });

  // Truncation notice
  var rendered = order.slice(0, blocks.filter(function(b) { return b.type === 'section'; }).length);
  if (rendered.length < order.length) {
    var remaining = order.length - rendered.length;
    blocks.push({
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: '_...and ' + remaining + ' more change(s) not shown. See the Changelog sheet for the full list._',
      },
    });
  }

  return blocks;
}

/**
 * Filters the change rows by the user's saved notify-fields config and posts
 * a Block Kit message to Slack. Silent-skips if Slack is not configured.
 */
function sendSlackNotifications_(changes) {
  var props      = PropertiesService.getScriptProperties();
  var botToken   = props.getProperty(CONFIG.SLACK_TOKEN_KEY);
  var channelId  = props.getProperty(CONFIG.SLACK_CHANNEL_KEY);
  var rawFields  = props.getProperty(CONFIG.SLACK_FIELDS_KEY);

  if (!botToken || !channelId || rawFields === null) return;

  var notifyFields;
  try { notifyFields = JSON.parse(rawFields); } catch (e) { return; }
  if (!notifyFields.length) return;

  var filtered = changes.filter(function(row) {
    return isNotifiable_(row, notifyFields);
  });
  if (!filtered.length) return;

  var blocks = buildSlackBlocks_(filtered);
  if (!blocks.length) return;

  var payload = JSON.stringify({
    channel: channelId,
    blocks:  blocks,
    text:    'Productboard change alert (' + filtered.length + ' update(s))',
  });

  try {
    var response = UrlFetchApp.fetch(CONFIG.SLACK_API_POST, {
      method:           'post',
      contentType:      'application/json',
      headers:          { 'Authorization': 'Bearer ' + botToken },
      payload:          payload,
      muteHttpExceptions: true,
    });
    var body = JSON.parse(response.getContentText());
    if (body.ok) {
      Logger.log('Slack notification sent (' + filtered.length + ' update(s)).');
    } else {
      Logger.log('Slack notification failed: ' + body.error);
    }
  } catch (e) {
    Logger.log('Slack notification error: ' + e.message);
  }
}
