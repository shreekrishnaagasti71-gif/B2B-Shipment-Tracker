// ═══════════════════════════════════════════════════════
//  NCM B2B TRACKER — BACKEND v4 (lean + fast)
//  Changes in this version:
//   • Kept the real checkout-reset fix from v3.1: the VAN MOVEMENTS
//     "Date" column was getting silently auto-converted by Sheets from a
//     "yyyy-MM-dd" string into a real Date object, so a row check-in had
//     JUST written could no longer be found as "today's row" on the very
//     next read — normalizeDateStr() + forcing the column to plain text
//     fixes this everywhere dates are compared.
//   • Removed the entire legacy shipment-scan / VAN STATUS / VAN VISITS
//     system (SHIPMENT GPS sheet, send/receive batch, incoming-vans-by-
//     branch, old round counter). None of it is used by the current
//     driver check-in/check-out screens — it was dead weight being
//     written to on every single check-in and check-out for no reason.
//   • Rewrote routeDashboard() to read VAN MOVEMENTS ONCE and group by
//     van in memory, instead of re-reading the whole sheet once PER VAN.
//     With 100+ staff and a growing sheet, that was the main thing that
//     would have gotten slower every single day.
//  Open from: Extensions → Apps Script inside your sheet
//  After pasting: Deploy → Manage deployments → Edit → New version → Deploy
// ═══════════════════════════════════════════════════════

const CONFIG = {
  TIMEZONE: "Asia/Kathmandu"
};

const PASSCODES = {
  // Branch access
  "1111": { branch: "TINKUNE",      role: "branch" },
  "2222": { branch: "CHABAHIL",     role: "branch" },
  "3333": { branch: "BASUNDHARA",   role: "branch" },
  "4444": { branch: "NAYA BUSPARK", role: "branch" },
  "5555": { branch: "SWOYAMBHU",    role: "branch" },
  "6666": { branch: "KALANKI",      role: "branch" },
  "7777": { branch: "SATDOBATO",    role: "branch" },
  "9999": { role: "admin" },
  // Driver passcodes are intentionally the same as the van number.
  "van 1404": { role: "driver", vanNo: "1404" },
  "van 843":  { role: "driver", vanNo: "843" },
  "van 2266": { role: "driver", vanNo: "2266" },
  "van 1836": { role: "driver", vanNo: "1836" }
};

const ROUTES = {
  "ROUTE 1": ["TINKUNE","CHABAHIL","BASUNDHARA","NAYA BUSPARK","SWOYAMBHU","KALANKI","SATDOBATO","TINKUNE"],
  "ROUTE 2": ["TINKUNE","SATDOBATO","KALANKI","SWOYAMBHU","NAYA BUSPARK","BASUNDHARA","CHABAHIL","TINKUNE"]
};

const ROUTE_LABELS = {
  "ROUTE 1": "Tinkune → Chabahil → Basundhara → Naya Buspark → Swoyambhu → Kalanki → Satdobato → Tinkune",
  "ROUTE 2": "Tinkune → Satdobato → Kalanki → Swoyambhu → Naya Buspark → Basundhara → Chabahil → Tinkune"
};

const MAX_ROUNDS_PER_DAY = 3;
const EXPECTED_LEG_MINUTES = 10;
const MIN_TRAVEL_SECONDS = 120;

/* ─── Helpers ────────────────────────────────────────── */

function today() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd");
}

function normalizeDateStr(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, CONFIG.TIMEZONE, "yyyy-MM-dd");
  }
  var str = String(value || "").trim();
  var match = str.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : str;
}

function secondsBetween(a, b) {
  var first = a instanceof Date ? a : new Date(a);
  var second = b instanceof Date ? b : new Date(b);
  if (isNaN(first.getTime()) || isNaN(second.getTime())) return 0;
  return Math.max(0, Math.floor((second.getTime() - first.getTime()) / 1000));
}

function holdTimeLabel(seconds) {
  var total = Math.max(0, Math.floor(Number(seconds) || 0));
  var hours = Math.floor(total / 3600);
  var minutes = Math.floor((total % 3600) / 60);
  if (hours) return hours + " hr" + (minutes ? " " + minutes + " min" : "");
  return minutes + " min";
}

/* ─── VAN MOVEMENTS sheet ─── single source of truth ─── */

const MOVE_COL = {
  DATE:1, VAN:2, ROUTE:3, ROUND:4, STOP:5, STATION:6,
  ARRIVAL:7, DEPARTURE:8, HOLD_SECONDS:9, TRAVEL_SECONDS:10,
  NEXT_STATION:11, STATUS:12, UPDATED:13, HOLD_TIME:14
};

function getMovementSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VAN MOVEMENTS");
  if (!sheet) {
    sheet = ss.insertSheet("VAN MOVEMENTS");
    sheet.appendRow([
      "Date","Van No","Route","Round","Stop No","Station",
      "Check In","Check Out","Hold Seconds","Travel Seconds",
      "Next Station","Status","Last Updated","Hold Time"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(2, MOVE_COL.VAN, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
    sheet.getRange(2, MOVE_COL.ARRIVAL, sheet.getMaxRows() - 1, 2).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    sheet.getRange(2, MOVE_COL.UPDATED, sheet.getMaxRows() - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  } else if (sheet.getLastColumn() < MOVE_COL.HOLD_TIME) {
    sheet.getRange(1, MOVE_COL.HOLD_TIME).setValue("Hold Time");
  }
  sheet.getRange(2, MOVE_COL.DATE, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  return sheet;
}

function getIssueSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ISSUES");
  if (!sheet) {
    sheet = ss.insertSheet("ISSUES");
    sheet.appendRow(["Date","Time","Reported By","Role","Branch","Van No","Issue","Status"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getVanRegistrySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VAN REGISTRY");
  if (!sheet) {
    sheet = ss.insertSheet("VAN REGISTRY");
    sheet.appendRow(["Van No","Passcode Format","Registered Date","Last Activity Time","Status"]);
    sheet.setFrozenRows(1);
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  return sheet;
}

function registerVan(vanNo) {
  var target = String(vanNo || "").trim().toUpperCase();
  if (!target || !/^[A-Z0-9][A-Z0-9-]*$/.test(target)) {
    return { success:false, error:"Van number may contain letters, numbers, and hyphens only" };
  }
  var sheet = getVanRegistrySheet();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === target) {
      sheet.getRange(i + 1, 5).setValue("ACTIVE");
      return { success:true, vanNo:target, existing:true };
    }
  }
  var now = new Date();
  sheet.appendRow([target, "van " + target, Utilities.formatDate(now, CONFIG.TIMEZONE, "yyyy-MM-dd"), now, "ACTIVE"]);
  return { success:true, vanNo:target, existing:false };
}

function getRouteConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ROUTE CONFIG");
  if (!sheet) {
    sheet = ss.insertSheet("ROUTE CONFIG");
    sheet.appendRow(["Route Name","Stations","Status","Created At"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getConfiguredRoutes() {
  var configured = {};
  Object.keys(ROUTES).forEach(function(name) { configured[name] = ROUTES[name].slice(); });
  var values = getRouteConfigSheet().getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var name = String(values[i][0] || "").trim().toUpperCase();
    var active = String(values[i][2] || "ACTIVE").toUpperCase() !== "INACTIVE";
    var stops = String(values[i][1] || "").split(",").map(function(s) {
      return s.trim().toUpperCase();
    }).filter(Boolean);
    if (name && active && stops.length >= 2) configured[name] = stops;
  }
  return configured;
}

function getRouteLabels() {
  var routes = getConfiguredRoutes();
  var labels = {};
  Object.keys(routes).forEach(function(name) {
    labels[name] = routes[name].join(" → ");
  });
  return labels;
}

function routeFor(routeName) {
  return getConfiguredRoutes()[routeName] || null;
}

function createRoute(data) {
  if (String(data.role || "").toLowerCase() !== "admin") {
    return { success:false, error:"Admin access required" };
  }
  var name = String(data.routeName || "").trim().toUpperCase();
  var stops = String(data.stations || "").split(",").map(function(s) {
    return s.trim().toUpperCase();
  }).filter(Boolean);
  if (!name || !/^[A-Z0-9][A-Z0-9 _-]*$/.test(name)) {
    return { success:false, error:"Use a simple route name, for example ROUTE 3" };
  }
  if (stops.length < 2 || stops[0] !== stops[stops.length - 1]) {
    return { success:false, error:"The route must have at least 2 stations and end at its starting station" };
  }
  var routes = getConfiguredRoutes();
  if (routes[name]) return { success:false, error:"That route already exists" };
  getRouteConfigSheet().appendRow([name, stops.join(","), "ACTIVE", new Date()]);
  return { success:true, message:name + " created", route:name, stations:stops };
}

function movementRowsForToday(vanNo) {
  var values = getMovementSheet().getDataRange().getValues();
  var target = String(vanNo || "").trim();
  var dateStr = today();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (normalizeDateStr(values[i][MOVE_COL.DATE - 1]) === dateStr &&
        String(values[i][MOVE_COL.VAN - 1]).trim() === target) {
      rows.push({ rowIndex: i + 1, values: values[i] });
    }
  }
  return rows;
}

// Shared by the single-van path (getRouteState) and the dashboard path
// (routeDashboard), so both agree on exactly the same logic.
function stateFromRow(vanNo, item) {
  if (!item) {
    return {
      vanNo: String(vanNo || "").trim(), started: false, route: null,
      round: 0, stopIndex: -1, station: null, nextStation: null,
      status: "NOT_STARTED"
    };
  }
  var v = item.values;
  if (String(v[MOVE_COL.STATUS - 1]) === "ADMIN_RESET") {
    return {
      vanNo: String(v[MOVE_COL.VAN - 1]).trim(), started: false,
      route: null, round: 0, stopIndex: -1, station: null,
      nextStation: null, status: "NOT_STARTED"
    };
  }
  var route = routeFor(String(v[MOVE_COL.ROUTE - 1]));
  var state = {
    vanNo: String(v[MOVE_COL.VAN - 1]).trim(),
    started: true,
    route: String(v[MOVE_COL.ROUTE - 1]),
    routeLabel: getRouteLabels()[String(v[MOVE_COL.ROUTE - 1])] || "",
    round: Number(v[MOVE_COL.ROUND - 1]) || 1,
    stopIndex: Number(v[MOVE_COL.STOP - 1]) || 0,
    station: v[MOVE_COL.STATION - 1],
    nextStation: v[MOVE_COL.NEXT_STATION - 1],
    arrivalTime: v[MOVE_COL.ARRIVAL - 1] || null,
    departureTime: v[MOVE_COL.DEPARTURE - 1] || null,
    holdSeconds: Number(v[MOVE_COL.HOLD_SECONDS - 1]) || 0,
    travelSeconds: Number(v[MOVE_COL.TRAVEL_SECONDS - 1]) || 0,
    status: String(v[MOVE_COL.STATUS - 1] || "AT_STATION"),
    rowIndex: item.rowIndex
  };
  if (route && state.status === "MOVING") {
    state.nextStation = route[state.stopIndex + 1] || route[0];
  }
  return state;
}

function getRouteState(vanNo) {
  var rows = movementRowsForToday(vanNo);
  return stateFromRow(vanNo, rows.length ? rows[rows.length - 1] : null);
}

function routeStateResponse(state) {
  var now = new Date();
  var elapsed = 0;
  if (state.status === "MOVING") elapsed = secondsBetween(state.departureTime, now);
  if (state.status === "AT_STATION" || state.status === "OFF_DUTY") {
    elapsed = secondsBetween(state.arrivalTime, now);
  }
  return {
    vanNo: state.vanNo,
    started: state.started,
    route: state.route,
    routeLabel: state.routeLabel || "",
    round: state.round || 0,
    stopIndex: state.stopIndex,
    station: state.station,
    nextStation: state.nextStation,
    status: state.status,
    elapsedSeconds: elapsed,
    holdSeconds: state.holdSeconds || 0,
    travelSeconds: state.travelSeconds || 0
  };
}

// FAST PATH: reads VAN MOVEMENTS and VAN REGISTRY ONCE, groups movement
// rows by van in memory, then builds every van's state from that —
// instead of re-reading the whole movement sheet once PER VAN (which
// would have gotten slower every single day as the sheet grew).
function routeDashboard() {
  var dateStr = today();
  var moveValues = getMovementSheet().getDataRange().getValues();
  var latestTodayByVan = {};
  var allVans = {};

  Object.keys(PASSCODES).forEach(function(pc) {
    if (PASSCODES[pc].role === "driver") allVans[PASSCODES[pc].vanNo] = true;
  });
  var registryValues = getVanRegistrySheet().getDataRange().getValues();
  for (var r = 1; r < registryValues.length; r++) {
    var registeredVan = String(registryValues[r][0] || "").trim();
    if (registeredVan) allVans[registeredVan] = true;
  }

  for (var i = 1; i < moveValues.length; i++) {
    var van = String(moveValues[i][MOVE_COL.VAN - 1] || "").trim();
    if (!van) continue;
    allVans[van] = true;
    if (normalizeDateStr(moveValues[i][MOVE_COL.DATE - 1]) !== dateStr) continue;
    latestTodayByVan[van] = { rowIndex: i + 1, values: moveValues[i] };
  }

  return Object.keys(allVans).map(function(vanNo) {
    var state = stateFromRow(vanNo, latestTodayByVan[vanNo]);
    var item = routeStateResponse(state);
    if (!state.started) {
      item.status = "idle";
      item.lastLocation = "Not started";
    } else if (state.status === "MOVING") {
      item.status = "moving";
      item.fromBranch = state.station;
      item.toBranch = state.nextStation;
      item.departTime = state.departureTime;
      item.elapsedMinutes = Math.floor(item.elapsedSeconds / 60);
      item.elapsedSeconds = item.elapsedSeconds % 60;
      item.isLate = item.elapsedMinutes >= EXPECTED_LEG_MINUTES;
    } else if (state.status === "OFF_DUTY") {
      item.status = "off_duty";
      item.lastLocation = state.station;
    } else {
      item.status = "idle";
      item.lastLocation = state.station;
      item.holdElapsedSeconds = item.elapsedSeconds;
    }
    item.roundLabel = state.started ? ("Round " + state.round + "/" + MAX_ROUNDS_PER_DAY) : null;
    return item;
  });
}

function handleRouteCheckIn(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var vanNo = String(data.vanNo || "").trim();
    var station = String(data.station || "").toUpperCase().trim();
    var routeName = String(data.route || "").toUpperCase().trim();
    if (!vanNo || !station) return { success:false, error:"Van and station are required" };

    var sheet = getMovementSheet();
    var state = getRouteState(vanNo);
    if (state.started && routeName && routeName !== state.route) {
      return { success:false, error:"Route is locked for today as " + state.route };
    }
    var route = routeFor(routeName || state.route);
    if (!route) return { success:false, error:"Choose Route 1 or Route 2 first" };

    var now = new Date();

    if (!state.started) {
      if (station !== route[0]) return { success:false, error:"First check-in must be at " + route[0] };
      sheet.appendRow([today(), vanNo, routeName, 1, 0, station, now, "", 0, 0, route[1], "AT_STATION", now]);
      return { success:true, message:"Checked in at " + station, state:routeStateResponse(getRouteState(vanNo)) };
    }

    if (state.status === "OFF_DUTY") return { success:false, error:"This van completed 3 rounds today" };
    if (state.status !== "MOVING") return { success:false, error:"Already checked in at " + state.station + ". Check out before moving." };
    if (station !== state.nextStation) {
      return { success:false, error:"Expected check-in at " + state.nextStation };
    }
    var travelSeconds = secondsBetween(state.departureTime, new Date());
    if (travelSeconds < MIN_TRAVEL_SECONDS) {
      return { success:false, error:"Check-in opens after 2 minutes of travel" };
    }

    var rows = movementRowsForToday(vanNo);
    var previous = rows[rows.length - 1];
    var previousValues = previous.values.slice();
    previousValues[MOVE_COL.TRAVEL_SECONDS - 1] = travelSeconds;
    previousValues[MOVE_COL.STATUS - 1] = "COMPLETE";
    previousValues[MOVE_COL.UPDATED - 1] = now;
    sheet.getRange(previous.rowIndex, 1, 1, previousValues.length).setValues([previousValues]);

    var nextIndex = state.stopIndex + 1;
    var nextRound = state.round;
    if (nextIndex >= route.length - 1) {
      nextRound = state.round + 1;
      nextIndex = 0;
    }
    var nextStation = route[nextIndex + 1] || route[1];
    var finalStop = nextRound > MAX_ROUNDS_PER_DAY;
    if (finalStop) nextRound = MAX_ROUNDS_PER_DAY;
    var currentStatus = finalStop ? "OFF_DUTY" : "AT_STATION";
    sheet.appendRow([
      today(), vanNo, state.route, nextRound, nextIndex, station, now, "",
      0, secondsBetween(state.departureTime, now), nextStation, currentStatus, now
    ]);
    return {
      success:true,
      message: finalStop ? "Arrived at " + station + " — 3 rounds complete" : "Checked in at " + station,
      state:routeStateResponse(getRouteState(vanNo))
    };
  } finally {
    lock.releaseLock();
  }
}

function handleRouteCheckOut(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var vanNo = String(data.vanNo || "").trim();
    var state = getRouteState(vanNo);
    if (!state.started) return { success:false, error:"Check in at TINKUNE first" };
    if (state.status === "MOVING") return { success:false, error:"Already moving to " + state.nextStation };
    if (state.status === "OFF_DUTY") return { success:false, error:"This van completed 3 rounds today" };
    if (state.status !== "AT_STATION") return { success:false, error:"Van is not at a station" };

    var now = new Date();
    var sheet = getMovementSheet();
    var values = sheet.getRange(state.rowIndex, 1, 1, MOVE_COL.UPDATED).getValues()[0];
    var holdSeconds = secondsBetween(state.arrivalTime, now);
    values[MOVE_COL.DEPARTURE - 1] = now;
    values[MOVE_COL.HOLD_SECONDS - 1] = holdSeconds;
    values[MOVE_COL.UPDATED - 1] = now;
    values[MOVE_COL.STATUS - 1] = "MOVING";
    sheet.getRange(state.rowIndex, 1, 1, values.length).setValues([values]);
    sheet.getRange(state.rowIndex, MOVE_COL.HOLD_TIME).setValue(holdTimeLabel(holdSeconds));
    return { success:true, message:"Checked out — heading to " + state.nextStation, state:routeStateResponse(getRouteState(vanNo)) };
  } finally {
    lock.releaseLock();
  }
}

function getMovementHistory(vanNo, dateStr) {
  var values = getMovementSheet().getDataRange().getValues();
  var target = String(vanNo || "").trim();
  var day = dateStr || today();
  var result = [];
  for (var i = 1; i < values.length; i++) {
    if (normalizeDateStr(values[i][MOVE_COL.DATE - 1]) === day &&
        (!target || String(values[i][MOVE_COL.VAN - 1]).trim() === target)) {
      result.push({
        vanNo: values[i][MOVE_COL.VAN - 1], route: values[i][MOVE_COL.ROUTE - 1],
        round: values[i][MOVE_COL.ROUND - 1], stopNo: values[i][MOVE_COL.STOP - 1],
        station: values[i][MOVE_COL.STATION - 1], checkIn: values[i][MOVE_COL.ARRIVAL - 1],
        checkOut: values[i][MOVE_COL.DEPARTURE - 1],
        holdSeconds: Number(values[i][MOVE_COL.HOLD_SECONDS - 1]) || 0,
        holdTime: values[i][MOVE_COL.HOLD_TIME - 1] || holdTimeLabel(values[i][MOVE_COL.HOLD_SECONDS - 1]),
        travelSeconds: Number(values[i][MOVE_COL.TRAVEL_SECONDS - 1]) || 0,
        nextStation: values[i][MOVE_COL.NEXT_STATION - 1],
        status: values[i][MOVE_COL.STATUS - 1]
      });
    }
  }
  return result;
}

function submitIssue(data) {
  var message = String(data.issue || "").trim();
  if (!message) return { success:false, error:"Write the issue before submitting" };
  getIssueSheet().appendRow([
    today(), Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "HH:mm:ss"),
    data.reportedBy || "", data.role || "", data.branch || "",
    data.vanNo || "", message, "OPEN"
  ]);
  return { success:true, message:"Issue sent to admin" };
}

function getIssues(includeClosed) {
  var values = getIssueSheet().getDataRange().getValues();
  var result = [];
  for (var i = 1; i < values.length; i++) {
    if (!includeClosed && String(values[i][7]).toUpperCase() === "CLOSED") continue;
    result.push({
      row: i + 1, date: values[i][0], time: values[i][1],
      reportedBy: values[i][2], role: values[i][3], branch: values[i][4],
      vanNo: values[i][5], issue: values[i][6], status: values[i][7]
    });
  }
  return result.reverse();
}

function closeIssue(row) {
  var sheet = getIssueSheet();
  if (!row || row < 2 || row > sheet.getLastRow()) return { success:false, error:"Invalid issue" };
  sheet.getRange(Number(row), 8).setValue("CLOSED");
  return { success:true, message:"Issue closed" };
}

function adminResetVan(vanNo) {
  var state = getRouteState(vanNo);
  if (!state.started) return { success:false, error:"Van has no active route today" };
  var sheet = getMovementSheet();
  var values = sheet.getRange(state.rowIndex, 1, 1, MOVE_COL.UPDATED).getValues()[0];
  values[MOVE_COL.STATUS - 1] = "ADMIN_RESET";
  values[MOVE_COL.UPDATED - 1] = new Date();
  sheet.getRange(state.rowIndex, 1, 1, values.length).setValues([values]);
  return { success:true, message:"Van " + vanNo + " reset for a new route" };
}

/* ═══════════════════════════════════════════════════════
   WEB APP ENTRY POINTS
═══════════════════════════════════════════════════════ */

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (!action) return jsonResponse({ success: false, error: "No action" });

  if (action === "login") {
    var enteredPasscode = String(e.parameter.passcode || "").trim().toLowerCase();
    var cfg = PASSCODES[enteredPasscode];
    if (!cfg) {
      var vanMatch = enteredPasscode.match(/^van\s+([a-z0-9][a-z0-9-]*)$/i);
      if (vanMatch) {
        var registration = registerVan(vanMatch[1]);
        if (!registration.success) return jsonResponse(registration);
        cfg = { role:"driver", vanNo:registration.vanNo };
      }
    }
    if (!cfg) return jsonResponse({ success: false, error: "Invalid passcode" });
    var res = { success: true, role: cfg.role, routes: getRouteLabels() };
    if (cfg.role === "branch")  { res.branch = cfg.branch; }
    if (cfg.role === "driver")  { res.vanNo = cfg.vanNo; }
    return jsonResponse(res);
  }

  if (action === "getRoutes") {
    return jsonResponse({ success:true, routes:getConfiguredRoutes(), labels:getRouteLabels(), maxRounds:MAX_ROUNDS_PER_DAY });
  }

  if (action === "getRouteState") {
    return jsonResponse({ success:true, state:routeStateResponse(getRouteState(e.parameter.vanNo)) });
  }

  if (action === "getRouteDashboard" || action === "getAllVans") {
    return jsonResponse({ success:true, data:routeDashboard(), date:today() });
  }

  if (action === "getMovementHistory") {
    return jsonResponse({ success:true, data:getMovementHistory(e.parameter.vanNo, e.parameter.date) });
  }

  if (action === "getIssues") {
    return jsonResponse({ success:true, data:getIssues(false) });
  }

  return jsonResponse({ success: false, error: "Unknown action" });
}

function doPost(e) {
  try {
    var body;
    try { body = JSON.parse(e.postData.contents); }
    catch (err) { return jsonResponse({ success: false, error: "Bad JSON" }); }

    switch (body.action) {
      case 'routeCheckIn':  return jsonResponse(handleRouteCheckIn(body));
      case 'routeCheckOut': return jsonResponse(handleRouteCheckOut(body));
      case 'createRoute':   return jsonResponse(createRoute(body));
      case 'submitIssue':   return jsonResponse(submitIssue(body));
      case 'closeIssue':    return jsonResponse(closeIssue(body.row));
      case 'adminResetVan': return jsonResponse(adminResetVan(body.vanNo));
      default:               return jsonResponse({ success: false, error: 'Unknown action' });
    }
  } catch (err) {
    return jsonResponse({ success:false, error:"Server error: " + (err.message || err.toString()) });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
