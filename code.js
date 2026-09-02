// ═══════════════════════════════════════════════════════
//  NCM B2B TRACKER — BACKEND (Fixed Receive + Timer)
//  Open from: Extensions → Apps Script inside your sheet
// ═══════════════════════════════════════════════════════

const CONFIG = {
  SHEET_NAME: "SHIPMENT GPS",
  TIMEZONE: "Asia/Kathmandu",
  TOTAL_COLS: 9
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

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(["Date","Origin","Destination","Send Time","Shipment ID","Van No","Status","Received By","Received Time"]);
    sheet.setFrozenRows(1);
  }
  // Ensure we always have 9 columns so getDataRange() never shrinks
  var lastCol = sheet.getLastColumn();
  if (lastCol < CONFIG.TOTAL_COLS) {
    sheet.insertColumnsAfter(lastCol, CONFIG.TOTAL_COLS - lastCol);
  }
  return sheet;
}

function getAllValues(sheet) {
  var lr = sheet.getLastRow();
  if (lr < 1) lr = 1;
  return sheet.getRange(1, 1, lr, CONFIG.TOTAL_COLS).getValues();
}

function setAllValues(sheet, values) {
  var lr = values.length;
  if (lr < 1) lr = 1;
  sheet.getRange(1, 1, lr, CONFIG.TOTAL_COLS).setValues(values);
}

/* ===================== SAFE DATE PARSER ===================== */

function parseDateTime(dateVal, timeVal) {
  var d;
  try {
    if (dateVal instanceof Date) {
      d = new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate());
    } else {
      var ds = String(dateVal).trim();
      // Handle yyyy-MM-dd
      var m = ds.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (m) {
        d = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
      } else {
        d = new Date(ds);
      }
    }
  } catch(e) { d = new Date(); }
  
  if (!d || isNaN(d.getTime())) d = new Date();
  
  var h = 0, mn = 0, s = 0;
  try {
    if (timeVal instanceof Date) {
      h = timeVal.getHours();
      mn = timeVal.getMinutes();
      s = timeVal.getSeconds();
    } else {
      var ts = String(timeVal || "00:00:00");
      var p = ts.match(/(\d{1,2}):(\d{2}):(\d{2})/);
      if (p) {
        h = parseInt(p[1]);
        mn = parseInt(p[2]);
        s = parseInt(p[3]);
      }
    }
  } catch(e) {}
  
  d.setHours(h, mn, s, 0);
  return d;
}

/* ===================== INCOMING VANS (LIVE TIMER) ===================== */

function getIncomingVans(branch) {
  try {
    var sheet = getSheet();
    var values = getAllValues(sheet);
    var now = new Date();
    var vans = {};

    for (var i = 1; i < values.length; i++) {
      var dest = values[i][COL.DEST - 1];
      var van = String(values[i][COL.VAN - 1] || "").trim();
      var status = String(values[i][COL.STATUS - 1] || "").trim();
      var origin = values[i][COL.ORIGIN - 1];
      var dateStr = values[i][COL.DATE - 1];
      var sendTime = values[i][COL.SEND_TIME - 1];

      if (van && resolveMainBranch(dest) === branch && status === "In Transit") {
        var key = van + "|" + origin;
        if (!vans[key]) {
          vans[key] = { vanNo: van, origin: origin, dateStr: dateStr, sendTime: sendTime, count: 0 };
        }
        vans[key].count++;
      }
    }

    var result = [];
    for (var key in vans) {
      var v = vans[key];
      var sentAt = parseDateTime(v.dateStr, v.sendTime);
      var ms = now.getTime() - sentAt.getTime();
      if (isNaN(ms) || ms < 0) ms = 0;
      var min = Math.floor(ms / 60000);
      var sec = Math.floor((ms % 60000) / 1000);
      result.push({
        vanNo: v.vanNo,
        origin: v.origin,
        count: v.count,
        elapsedMinutes: min,
        elapsedSeconds: sec,
        isLate: min >= 20
      });
    }

    result.sort(function(a, b) {
      return (a.elapsedMinutes * 60 + a.elapsedSeconds) - (b.elapsedMinutes * 60 + b.elapsedSeconds);
    });

    return result;
  } catch(e) {
    return [];
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
    var data = getPendingByVan(e.parameter.branch, e.parameter.vanNo);
    return jsonResponse({ success: true, count: data.length, data: data });
  }

  if (action === "getIncomingVans") {
    var data = getIncomingVans(e.parameter.branch);
    return jsonResponse({ success: true, data: data });
  }

  return jsonResponse({ success: false, error: "Unknown action" });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonResponse({ success: false, error: "Bad JSON" }); }

  switch (body.action) {
    case 'sendBatch':    return jsonResponse(handleSendBatch(body));
    case 'receiveBatch': return jsonResponse(handleReceiveBatch(body));
    case 'receiveOne':   return jsonResponse(handleReceiveOne(body));
    case 'receiveAll':   return jsonResponse(handleReceiveAll(body));
    default:             return jsonResponse({ success: false, error: 'Unknown action' });
  }
}

/* ===================== SEND BATCH ===================== */

function handleSendBatch(data) {
  try {
    var sheet = getSheet();
    var values = getAllValues(sheet);
    var now = new Date();
    var tz = CONFIG.TIMEZONE;
    var dateStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    var timeStr = Utilities.formatDate(now, tz, "HH:mm:ss");

    var inTransitIds = {};
    for (var i = 1; i < values.length; i++) {
      var sid = String(values[i][COL.ID - 1] || "").trim();
      var status = String(values[i][COL.STATUS - 1] || "").trim();
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
        inTransitIds[sid] = origin;
      }
    }

    if (newRows.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      var needed = startRow + newRows.length - 1;
      if (sheet.getMaxRows() < needed) {
        sheet.insertRowsAfter(sheet.getMaxRows(), needed - sheet.getMaxRows() + 5);
      }
      sheet.getRange(startRow, 1, newRows.length, CONFIG.TOTAL_COLS).setValues(newRows);
    }

    return { success: true, sent: newRows.length, results: results };

  } catch (err) { return { success: false, error: err.toString() }; }
}

/* ===================== RECEIVE BATCH (FIXED) ===================== */

function handleReceiveBatch(data) {
  try {
    var sheet = getSheet();
    var values = getAllValues(sheet);
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
    var modified = false;

    for (var i = 1; i < values.length; i++) {
      var sid = String(values[i][COL.ID - 1] || "").trim();
      var dest = values[i][COL.DEST - 1];
      var van = String(values[i][COL.VAN - 1] || "").trim();
      var status = String(values[i][COL.STATUS - 1] || "").trim();

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
        modified = true;
      }
    }

    if (modified) {
      setAllValues(sheet, values);
    }

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

/* ===================== RECEIVE ONE (FIXED) ===================== */

function handleReceiveOne(data) {
  try {
    var sheet = getSheet();
    var values = getAllValues(sheet);
    var now = new Date();
    var tz = CONFIG.TIMEZONE;
    var timeStr = Utilities.formatDate(now, tz, "HH:mm:ss");

    for (var i = 1; i < values.length; i++) {
      var sid = String(values[i][COL.ID - 1] || "").trim();
      var dest = values[i][COL.DEST - 1];
      var van = String(values[i][COL.VAN - 1] || "").trim();
      var status = String(values[i][COL.STATUS - 1] || "").trim();

      if (sid === String(data.shipmentId || "").trim() &&
          resolveMainBranch(dest) === data.branch &&
          van === String(data.vanNo).trim() &&
          status === "In Transit") {

        values[i][COL.STATUS - 1] = "Received";
        values[i][COL.RECV_BY - 1] = data.branch;
        values[i][COL.RECV_TIME - 1] = timeStr;
        setAllValues(sheet, values);
        return { success: true, message: data.shipmentId + " received" };
      }
    }
    return { success: false, error: "Not found or already received" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

/* ===================== RECEIVE ALL (FIXED) ===================== */

function handleReceiveAll(data) {
  try {
    var sheet = getSheet();
    var values = getAllValues(sheet);
    var count = 0;
    var now = new Date();
    var tz = CONFIG.TIMEZONE;
    var timeStr = Utilities.formatDate(now, tz, "HH:mm:ss");
    var modified = false;

    for (var i = 1; i < values.length; i++) {
      var dest = values[i][COL.DEST - 1];
      var van = String(values[i][COL.VAN - 1] || "").trim();
      var status = String(values[i][COL.STATUS - 1] || "").trim();

      if (resolveMainBranch(dest) === data.branch &&
          van === String(data.vanNo).trim() &&
          status === "In Transit") {

        values[i][COL.STATUS - 1] = "Received";
        values[i][COL.RECV_BY - 1] = data.branch;
        values[i][COL.RECV_TIME - 1] = timeStr;
        count++;
        modified = true;
      }
    }

    if (modified) {
      setAllValues(sheet, values);
    }

    return { success: true, message: count + " shipment(s) received", count: count };
  } catch (err) { return { success: false, error: err.toString() }; }
}

/* ===================== PENDING ===================== */

function getPendingByVan(branch, vanNo) {
  var sheet = getSheet();
  var values = getAllValues(sheet);
  var result = [];
  var targetVan = String(vanNo).trim();

  for (var i = 1; i < values.length; i++) {
    var rowDest = values[i][COL.DEST - 1];
    var rowVan = String(values[i][COL.VAN - 1] || "").trim();
    var rowStatus = String(values[i][COL.STATUS - 1] || "").trim();

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
