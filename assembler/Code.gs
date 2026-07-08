/* =========================================================================
   Assembly Instructions API - Google Apps Script
   =========================================================================
   Sheet layout:
   - "Products" tab: Product | TabName | ModelURL | Counter
   - "Ledger" tab:   Timestamp | Product | StepIDs | Action | RelatedStepIDs
                     | BeforeJSON | AfterJSON | Author | Notes
   - One tab per product (TabName), columns = HEADERS below.

   Step identity:
   - StepID is permanent (e.g. "Demo Product-000014"), never reused.
   - Order is a float that controls display sequence; display numbers are
     computed client-side from sorted Order and never stored.
   - Rows are never hard-deleted: Status flips to "deleted" so the Ledger's
     references stay resolvable forever.

   Setup:
   1. Create a Google Sheet, open Extensions -> Apps Script, paste this file.
   2. Change TOKEN below.
   3. Run setup() once (authorize when prompted).
   4. Deploy -> New deployment -> Web app:
      Execute as: Me | Who has access: Anyone
   5. Copy the /exec URL into the viewer's data-script-url attribute.
   ========================================================================= */

var TOKEN = 'CHANGE-ME-to-a-long-random-string';

var PRODUCTS_TAB = 'Products';
var LEDGER_TAB = 'Ledger';

var HEADERS = ['StepID','Order','Subassembly','Directions','PartsNeeded','ToolsNeeded',
               'VisibleParts','HighlightedParts','VisibleArrows','HighlightedArrows',
               'Camera','BackgroundColor','Transition','Status','LastModified'];

var LEDGER_HEADERS = ['Timestamp','Product','StepIDs','Action','RelatedStepIDs',
                      'BeforeJSON','AfterJSON','Author','Notes'];

/* ========================= entry points ========================= */

function doGet(e) {
  try {
    if (!auth_(e.parameter.token)) return json_({ ok: false, error: 'bad token' });
    var action = e.parameter.action || '';

    if (action === 'products') {
      return json_({ ok: true, products: getProducts_().map(function (p) {
        return { name: p.name, modelUrl: p.modelUrl };
      })});
    }

    if (action === 'bundle') {
      var p = findProduct_(e.parameter.product);
      if (!p) return json_({ ok: false, error: 'unknown product: ' + e.parameter.product });
      return json_({ ok: true,
        product: { name: p.name, modelUrl: p.modelUrl },
        steps: getSteps_(p) });
    }

    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'bad json body' }); }

  if (!auth_(body.token)) return json_({ ok: false, error: 'bad token' });

  var p = findProduct_(body.product);
  if (!p) return json_({ ok: false, error: 'unknown product: ' + body.product });

  var author = String(body.author || 'unknown');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (body.action) {
      case 'create': return json_(createStep_(p, body, author));
      case 'edit':   return json_(editStep_(p, body, author));
      case 'merge':  return json_(mergeSteps_(p, body, author));
      case 'split':  return json_(splitStep_(p, body, author));
      case 'delete': return json_(deleteStep_(p, body, author));
      default:       return json_({ ok: false, error: 'unknown action: ' + body.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ========================= write operations ========================= */

function createStep_(p, body, author) {
  var sheet = tab_(p.tab);
  var fields = body.fields || {};
  var stepId = nextStepId_(p);
  var order = Number(fields.order);
  if (!isFinite(order)) order = maxOrder_(sheet) + 1;

  var row = fieldsToRow_(stepId, order, fields);
  sheet.appendRow(row);

  var after = rowToObj_(row);
  logLedger_(p.name, [stepId], 'created', [], null, after, author, body.notes || '');
  return { ok: true, stepId: stepId };
}

function editStep_(p, body, author) {
  var sheet = tab_(p.tab);
  var found = findRow_(sheet, body.stepId);
  if (!found) return { ok: false, error: 'step not found: ' + body.stepId };

  var before = rowToObj_(found.values);
  var fields = body.fields || {};
  var order = Number(fields.order);
  if (!isFinite(order)) order = Number(found.values[1]);

  var row = fieldsToRow_(body.stepId, order, fields);
  row[13] = found.values[13] || 'active';  // preserve Status on edit
  sheet.getRange(found.rowIndex, 1, 1, HEADERS.length).setValues([row]);

  var after = rowToObj_(row);
  logLedger_(p.name, [body.stepId], 'edited', [], before, after, author, body.notes || '');
  return { ok: true, stepId: body.stepId };
}

function mergeSteps_(p, body, author) {
  var sheet = tab_(p.tab);
  var ids = body.stepIds || [];
  if (ids.length < 2) return { ok: false, error: 'merge needs 2+ stepIds' };

  var sources = [], minOrder = Infinity;
  for (var i = 0; i < ids.length; i++) {
    var f = findRow_(sheet, ids[i]);
    if (!f) return { ok: false, error: 'step not found: ' + ids[i] };
    sources.push(f);
    var o = Number(f.values[1]);
    if (o < minOrder) minOrder = o;
  }

  // New merged step takes the earliest position among its sources.
  var fields = body.fields || {};
  fields.order = minOrder;
  var newId = nextStepId_(p);
  sheet.appendRow(fieldsToRow_(newId, minOrder, fields));

  // Sources become deleted (never removed).
  var beforeObjs = [];
  sources.forEach(function (f) {
    beforeObjs.push(rowToObj_(f.values));
    sheet.getRange(f.rowIndex, 14).setValue('deleted');
    sheet.getRange(f.rowIndex, 15).setValue(new Date().toISOString());
  });

  var after = rowToObj_(fieldsToRow_(newId, minOrder, fields));
  logLedger_(p.name, [newId], 'merged', ids, beforeObjs, after, author, body.notes || '');
  return { ok: true, stepId: newId };
}

function splitStep_(p, body, author) {
  var sheet = tab_(p.tab);
  var found = findRow_(sheet, body.stepId);
  if (!found) return { ok: false, error: 'step not found: ' + body.stepId };

  var copies = Math.max(2, Math.min(6, Number(body.copies) || 2));
  var source = rowToObj_(found.values);
  var srcOrder = Number(found.values[1]);

  // Space the copies between the source's order and the next active order.
  var next = nextOrderAfter_(sheet, srcOrder);
  var span = (next === null) ? copies : (next - srcOrder);
  var newIds = [], afterObjs = [];

  for (var i = 0; i < copies; i++) {
    var id = nextStepId_(p);
    var order = srcOrder + (span * i) / copies;
    var fields = objToFields_(source);
    fields.order = order;
    var row = fieldsToRow_(id, order, fields);
    sheet.appendRow(row);
    newIds.push(id);
    afterObjs.push(rowToObj_(row));
  }

  sheet.getRange(found.rowIndex, 14).setValue('deleted');
  sheet.getRange(found.rowIndex, 15).setValue(new Date().toISOString());

  logLedger_(p.name, newIds, 'split', [body.stepId], source, afterObjs, author, body.notes || '');
  return { ok: true, stepIds: newIds };
}

function deleteStep_(p, body, author) {
  var sheet = tab_(p.tab);
  var found = findRow_(sheet, body.stepId);
  if (!found) return { ok: false, error: 'step not found: ' + body.stepId };

  var before = rowToObj_(found.values);
  sheet.getRange(found.rowIndex, 14).setValue('deleted');
  sheet.getRange(found.rowIndex, 15).setValue(new Date().toISOString());

  logLedger_(p.name, [body.stepId], 'deleted', [], before, null, author, body.notes || '');
  return { ok: true };
}

/* ========================= reads ========================= */

function getProducts_() {
  var sheet = tab_(PRODUCTS_TAB);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    out.push({ name: String(values[i][0]), tab: String(values[i][1]),
               modelUrl: String(values[i][2]), rowIndex: i + 1 });
  }
  return out;
}

function findProduct_(name) {
  var list = getProducts_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].name === String(name)) return list[i];
  }
  return null;
}

function getSteps_(p) {
  var sheet = tab_(p.tab);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    if (String(row[13]) !== 'active') continue;
    out.push(rowToObj_(row));
  }
  out.sort(function (a, b) { return a.order - b.order; });
  return out;
}

/* ========================= row <-> object mapping ========================= */

function rowToObj_(row) {
  return {
    stepId: String(row[0]),
    order: Number(row[1]),
    subassembly: String(row[2] || ''),
    directions: String(row[3] || ''),
    parts: splitLines_(row[4]),
    tools: splitLines_(row[5]),
    visibleParts: parseMaybeJson_(row[6], 'all'),
    highlightedParts: parseMaybeJson_(row[7], []),
    visibleArrows: parseMaybeJson_(row[8], []),
    highlightedArrows: parseMaybeJson_(row[9], []),
    camera: parseMaybeJson_(row[10], { auto: { azimuth: 30, elevation: 20, distance: 2.2 }, controlsEnabled: true, fov: 50 }),
    backgroundColor: String(row[11] || '#0e1013'),
    transition: String(row[12] || 'smooth'),
    status: String(row[13] || 'active'),
    lastModified: String(row[14] || '')
  };
}

function fieldsToRow_(stepId, order, f) {
  return [
    stepId,
    order,
    String(f.subassembly || ''),
    String(f.directions || ''),
    String(f.partsNeeded || ''),                       // newline-separated text
    String(f.toolsNeeded || ''),
    jsonOrAll_(f.visibleParts),
    JSON.stringify(f.highlightedParts || []),
    JSON.stringify(f.visibleArrows || []),
    JSON.stringify(f.highlightedArrows || []),
    JSON.stringify(f.camera || {}),
    String(f.backgroundColor || '#0e1013'),
    String(f.transition || 'smooth'),
    'active',
    new Date().toISOString()
  ];
}

// Turn a read object back into a fields payload (used by split to clone).
function objToFields_(o) {
  return {
    subassembly: o.subassembly,
    directions: o.directions,
    partsNeeded: (o.parts || []).join('\n'),
    toolsNeeded: (o.tools || []).join('\n'),
    visibleParts: o.visibleParts,
    highlightedParts: o.highlightedParts,
    visibleArrows: o.visibleArrows,
    highlightedArrows: o.highlightedArrows,
    camera: o.camera,
    backgroundColor: o.backgroundColor,
    transition: o.transition
  };
}

/* ========================= helpers ========================= */

function auth_(token) { return String(token || '') === TOKEN; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function tab_(name) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('missing tab: ' + name);
  return s;
}

function findRow_(sheet, stepId) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(stepId)) {
      return { rowIndex: i + 1, values: values[i] };
    }
  }
  return null;
}

function maxOrder_(sheet) {
  var values = sheet.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < values.length; i++) {
    var o = Number(values[i][1]);
    if (isFinite(o) && o > max) max = o;
  }
  return max;
}

function nextOrderAfter_(sheet, order) {
  var values = sheet.getDataRange().getValues();
  var next = null;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][13]) !== 'active') continue;
    var o = Number(values[i][1]);
    if (o > order && (next === null || o < next)) next = o;
  }
  return next;
}

function nextStepId_(p) {
  var sheet = tab_(PRODUCTS_TAB);
  var n = Number(sheet.getRange(p.rowIndex, 4).getValue()) || 0;
  n += 1;
  sheet.getRange(p.rowIndex, 4).setValue(n);
  var pad = ('000000' + n).slice(-6);
  return p.name + '-' + pad;
}

function splitLines_(v) {
  var s = String(v || '').trim();
  if (!s) return [];
  return s.split('\n').map(function (x) { return x.trim(); }).filter(Boolean);
}

function parseMaybeJson_(v, fallback) {
  var s = String(v || '').trim();
  if (!s) return fallback;
  if (s === 'all') return 'all';
  try { return JSON.parse(s); } catch (e) { return fallback; }
}

function jsonOrAll_(v) {
  if (v === 'all' || v === undefined || v === null) return 'all';
  return JSON.stringify(v);
}

function logLedger_(product, stepIds, action, relatedIds, before, after, author, notes) {
  tab_(LEDGER_TAB).appendRow([
    new Date().toISOString(),
    product,
    (stepIds || []).join(', '),
    action,
    (relatedIds || []).join(', '),
    before === null ? '' : JSON.stringify(before),
    after === null ? '' : JSON.stringify(after),
    author,
    notes || ''
  ]);
}

/* ========================= one-time setup ========================= */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var products = ss.getSheetByName(PRODUCTS_TAB) || ss.insertSheet(PRODUCTS_TAB);
  if (products.getLastRow() === 0) {
    products.appendRow(['Product', 'TabName', 'ModelURL', 'Counter']);
    products.appendRow(['Demo Product', 'Demo Product',
      'https://cdn.jsdelivr.net/gh/davidthefuturist/plz@main/Body1.glb', 2]);
  }

  var ledger = ss.getSheetByName(LEDGER_TAB) || ss.insertSheet(LEDGER_TAB);
  if (ledger.getLastRow() === 0) ledger.appendRow(LEDGER_HEADERS);

  var demo = ss.getSheetByName('Demo Product') || ss.insertSheet('Demo Product');
  if (demo.getLastRow() === 0) {
    demo.appendRow(HEADERS);
    demo.appendRow(['Demo Product-000001', 1, 'Base frame',
      'Orient the base as shown. Confirm the mounting face is up.',
      'Base frame x1', 'None',
      'all', '[]', '[]', '[]',
      JSON.stringify({ auto: { azimuth: 35, elevation: 22, distance: 2.2 }, controlsEnabled: true, fov: 50 }),
      '#0e1013', 'smooth', 'active', new Date().toISOString()]);
    demo.appendRow(['Demo Product-000002', 2, 'Fastening',
      'Locate the four M5 bosses on the top face. View is locked to the inspection angle.',
      'M5x12 SHCS x4\nM5 washer x4', '4mm hex driver\nTorque wrench (4 N-m)',
      'all', '[]', '[]', '[]',
      JSON.stringify({ auto: { azimuth: 90, elevation: 65, distance: 1.8 }, controlsEnabled: false, fov: 45 }),
      '#2a1608', 'instant', 'active', new Date().toISOString()]);
  }
}