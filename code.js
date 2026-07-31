// ═══════════════════════════════════════════════════════
//  NCM B2B TRACKER — v4 INSTANT BATCH (Ultra-Fast)
//  Deploy as Web App: Execute as Me, Access: Anyone
//  Timezone: Asia/Kathmandu  |  Hours: 08:00 – 16:00
// ═══════════════════════════════════════════════════════

const CONFIG = {
  SHEET_NAME: "SHIPMENT GPS",
  START_HOUR: 8,      // 8:00 AM
  END_HOUR: 16,       // 4:00 PM (16:00)
  TIMEZONE: "Asia/Kathmandu"
};

const PASSCODES = {
  "1111": { branch: "TINKUNE",      role: "branch" },
  "2222": { branch: "CHABAHIL",     role: "branch" },
  "3333": { branch: "NAYA BUSPARK", role: "branch" },
  "4444": { branch: "KALANKI",      role: "branch" },
  "5555": { branch: "SATDOBATO",    role: "branch" },
  "6666": { branch: "NEWROAD",      role: "branch" },
  "7777": { role: "driver", branches: ["TINKUNE","CHABAHIL","NAYA BUSPARK","KALANKI","SATDOBATO","NEWROAD"] }
};

const DESTINATIONS = {
  "TINKUNE":      ["NAYA THIMI","SURYABINAYAK","LUBHU","CHABAHIL","NAYA BUSPARK","KALANKI","SATDOBATO","NEWROAD"],
  "CHABAHIL":     ["KAPAN","BUDHANILKANTHA","SANKHU","SUNDARIJAL","NAYA BUSPARK","KALANKI","SATDOBATO","TINKUNE","NEWROAD"],
  "NAYA BUSPARK": ["NEWROAD","KALANKI","SATDOBATO","TINKUNE","CHABAHIL"],
  "KALANKI":      ["THANKOT","SATDOBATO","TINKUNE","CHABAHIL","NAYA BUSPARK","NEWROAD"],
  "SATDOBATO":    ["TINKUNE","CHABAHIL","NAYA BUSPARK","KALANKI","NEWROAD","CHAPAGAU","GODAWARI"],
  "NEWROAD":      ["CHABAHIL","NAYA BUSPARK","KALANKI","SATDOBATO","TINKUNE"]
};

const SUB_BRANCHES = {
  "KAPAN":"CHABAHIL","BUDHANILKANTHA":"CHABAHIL","SANKHU":"CHABAHIL","SUNDARIJAL":"CHABAHIL",
  "NAYA THIMI":"TINKUNE","SURYABINAYAK":"TINKUNE","LUBHU":"TINKUNE",
  "GODAWARI":"SATDOBATO","CHAPAGAU":"SATDOBATO",
  "SWOYAMBHU":"NAYA BUSPARK","BASUNDHARA":"NAYA BUSPARK",
  "THANKOT":"KALANKI"
};

const COL = { DATE:1, ORIGIN:2, DEST:3, SEND_TIME:4, ID:5, VAN:6, STATUS:7, RECV_BY:8, RECV_TIME:9 };

function resolveMainBranch(name) {
  if (!name) return name;
  var upper = name.toString().toUpperCase().trim();
  return SUB_BRANCHES[upper] || upper;
}

function isBusinessHours() {
  var now = new Date();
  var hour = parseInt(Utilities.formatDate(now, CONFIG.TIMEZONE, "H"));
  var minute = parseInt(Utilities.formatDate(now, CONFIG.TIMEZONE, "m"));
  var time = hour * 60 + minute;
  return time >= CONFIG.START_HOUR * 60 && time < CONFIG.END_HOUR * 60;
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(["Date","Origin","Destination","Send Time","Shipment ID","Van No","Status","Received By","Received Time"]);
    sheet.setFrozenRows(1);
    sheet.getRange(2, COL.VAN, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  return sheet;
}

function ensureRows(sheet, minRows) {
  if (sheet.getMaxRows() < minRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows() + 5);
  }
}

/* ===================== WEB APP ===================== */

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (!action) return jsonResponse({ success: false, error: "No action" });

  if (action === "login") {
    var pc = e.parameter.passcode;
    var cfg = PASSCODES[pc];
    if (!cfg) return jsonResponse({ success: false, error: "Invalid passcode" });
    var res = { success: true, role: cfg.role };
    if (cfg.role === "branch") {
      res.branch = cfg.branch;
      res.destinations = DESTINATIONS[cfg.branch] || [];
    } else if (cfg.role === "driver") {
      res.branches = cfg.branches;
    }
    return jsonResponse(res);
  }

  if (action === "getPendingByVan") {
    if (!isBusinessHours()) return jsonResponse({ success: false, error: "Outside business hours (8 AM - 4 PM)" });
    var data = getPendingByVan(e.parameter.branch, e.parameter.vanNo);
    return jsonResponse({ success: true, count: data.length, data: data });
  }

  return jsonResponse({ success: false, error: "Unknown action" });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonResponse({ success: false, error: "Bad JSON" }); }

  if (!isBusinessHours()) {
    return jsonResponse({ success: false, error: "Outside business hours (8 AM - 4 PM)" });
  }

  switch (body.action) {
    case 'sendBatch':    return jsonResponse(handleSendBatch(body));
    case 'receiveBatch': return jsonResponse(handleReceiveBatch(body));
    case 'receiveOne':   return jsonResponse(handleReceiveOne(body));
    case 'receiveAll':   return jsonResponse(handleReceiveAll(body));
    default:             return jsonResponse({ success: false, error: 'Unknown action' });
  }
}

/* ===================== SEND BATCH (INSTANT) ===================== */

function handleSendBatch(data) {
  try {
    var sheet = getSheet();
    var values = sheet.getDataRange().getValues(); // ONE READ
    var now = new Date();
    var tz = CONFIG.TIMEZONE;
    var dateStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    var timeStr = Utilities.formatDate(now, tz, "HH:mm:ss");

    // In-memory duplicate check
    var inTransitIds = {};
    for (var i = 1; i < values.length; i++) {
      var sid = String(values[i][COL.ID - 1] || "").trim();
      var status = values[i][COL.STATUS - 1];
      if (sid && status === "In Transit") inTransitIds[sid] = values[i][COL.ORIGIN - 1];
    }

    var newRows = [];
    var results = [];
    var vanNo = String(data.vanNo || "N/A").trim();
    var origin = data.origin;
    var dest = data.destination;
    var shipments = data.shipments || [];

    for (var j = 0; j < shipments.length; j++) {
      var item = shipments[j];
      var sid = String(item.shipmentId || "").trim();
      if (!sid) continue;

      if (inTransitIds[sid]) {
        results.push({ shipmentId: sid, status: "duplicate", error: sid + " already In Transit from " + inTransitIds[sid] });
      } else {
        newRows.push([dateStr, origin, dest, timeStr, sid, vanNo, "In Transit", "", ""]);
        results.push({ shipmentId: sid, status: "sent" });
        inTransitIds[sid] = origin; // prevent duplicate within same batch
      }
    }

    if (newRows.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      ensureRows(sheet, startRow + newRows.length);
      sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
      sheet.getRange(startRow, COL.VAN, newRows.length, 1).setNumberFormat('@');
    }

    return { success: true, sent: newRows.length, results: results };

  } catch (err) { return { success: false, error: err.toString() }; }
}

/* ===================== RECEIVE BATCH (INSTANT) ===================== */

function handleReceiveBatch(data) {
  try {
    var sheet = getSheet();
    var range = sheet.getDataRange();
    var values = range.getValues(); // ONE READ
    var now = new Date();
    var tz = CONFIG.TIMEZONE;
    var timeStr = Utilities.formatDate(now, tz, "HH:mm:ss");

    var targetIds = {};
    var rawIds = data.shipmentIds || [];
    for (var k = 0; k < rawIds.length; k++) targetIds[String(rawIds[k]).trim()] = true;

    var targetVan = String(data.vanNo).trim();
    var branch = data.branch;
    var receivedCount = 0;
    var foundIds = {};
    var notFoundIds = [];

    for (var i = 1; i < values.length; i++) {
      var sid = String(values[i][COL.ID - 1] || "").trim();
      var dest = values[i][COL.DEST - 1];
      var van = String(values[i][COL.VAN - 1] || "").trim();
      var status = values[i][COL.STATUS - 1];

      if (targetIds[sid] &&
          resolveMainBranch(dest) === branch &&
          van === targetVan &&
          status === "In Transit" &&
          !foundIds[sid]) {

        values[i][COL.STATUS - 1] = "Received";
        values[i][COL.RECV_BY - 1] = branch;
        values[i][COL.RECV_TIME - 1] = timeStr;
        receivedCount++;
        foundIds[sid] = true;
      }
    }

    range.setValues(values); // ONE WRITE BACK

    for (var j = 0; j < rawIds.length; j++) {
      var id = String(rawIds[j]).trim();
      if (!foundIds[id]) notFoundIds.push(id);
    }

    return {
      success: true,
      message: receivedCount + " received" + (notFoundIds.length ? ", " + notFoundIds.length + " not found" : ""),
      receivedCount: receivedCount,
      notFoundIds: notFoundIds
    };

  } catch (err) { return { success: false, error: err.toString() }; }
}

function handleReceiveOne(data) {
  try {
    var sheet = getSheet();
    var range = sheet.getDataRange();
    var values = range.getValues();
    var now = new Date();
    var tz = CONFIG.TIMEZONE;
    var timeStr = Utilities.formatDate(now, tz, "HH:mm:ss");

    for (var i = 1; i < values.length; i++) {
      var sid = String(values[i][COL.ID - 1] || "").trim();
      var dest = values[i][COL.DEST - 1];
      var van = String(values[i][COL.VAN - 1] || "").trim();
      var status = values[i][COL.STATUS - 1];

      if (sid === String(data.shipmentId || "").trim() &&
          resolveMainBranch(dest) === data.branch &&
          van === String(data.vanNo).trim() &&
          status === "In Transit") {

        values[i][COL.STATUS - 1] = "Received";
        values[i][COL.RECV_BY - 1] = data.branch;
        values[i][COL.RECV_TIME - 1] = timeStr;
        range.setValues(values);
        return { success: true, message: data.shipmentId + " received" };
      }
    }
    return { success: false, error: "Not found or already received" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function handleReceiveAll(data) {
  try {
    var sheet = getSheet();
    var range = sheet.getDataRange();
    var values = range.getValues();
    var count = 0;
    var now = new Date();
    var tz = CONFIG.TIMEZONE;
    var timeStr = Utilities.formatDate(now, tz, "HH:mm:ss");

    for (var i = 1; i < values.length; i++) {
      var dest = values[i][COL.DEST - 1];
      var van = String(values[i][COL.VAN - 1] || "").trim();
      var status = values[i][COL.STATUS - 1];

      if (resolveMainBranch(dest) === data.branch &&
          van === String(data.vanNo).trim() &&
          status === "In Transit") {

        values[i][COL.STATUS - 1] = "Received";
        values[i][COL.RECV_BY - 1] = data.branch;
        values[i][COL.RECV_TIME - 1] = timeStr;
        count++;
      }
    }
    range.setValues(values);
    return { success: true, message: count + " shipment(s) received", count: count };
  } catch (err) { return { success: false, error: err.toString() }; }
}

/* ===================== PENDING ===================== */

function getPendingByVan(branch, vanNo) {
  var sheet = getSheet();
  var values = sheet.getDataRange().getValues();
  var result = [];
  var targetVan = String(vanNo).trim();

  for (var i = 1; i < values.length; i++) {
    var rowDest = values[i][COL.DEST - 1];
    var rowVan = String(values[i][COL.VAN - 1] || "").trim();
    var rowStatus = values[i][COL.STATUS - 1];

    if (resolveMainBranch(rowDest) === branch && rowVan === targetVan && rowStatus === "In Transit") {
      result.push({ shipmentId: values[i][COL.ID - 1], origin: values[i][COL.ORIGIN - 1], destination: rowDest });
    }
  }
  return result;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}