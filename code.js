// ═══════════════════════════════════════════════════════
//  NCM B2B TRACKER — BACKEND v3
//  Changes from v1:
//   • Route-first van movement tracking — no shipment scan required
//   • Two fixed routes, three rounds per day, with station check-in/out
//   • Van-specific driver passcodes and an admin dashboard
//   • Travel time, station hold time, issues, and live status in Sheets
//   • Legacy shipment endpoints remain available for existing data
//  Open from: Extensions → Apps Script inside your sheet
//  After pasting: Deploy → Manage deployments → Edit → New version → Deploy
// ═══════════════════════════════════════════════════════

const CONFIG = {
  SHEET_NAME: "SHIPMENT GPS",
  TIMEZONE:   "Asia/Kathmandu"
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
  // Driver passcodes are intentionally the same as the van number for the
  // four vans supplied by the operator.
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

// Kept for the legacy shipment round label functions below.
const ROUND_BRANCHES = ["TINKUNE", "SATDOBATO", "KALANKI", "NAYA BUSPARK", "CHABAHIL"];

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

const MAX_ROUNDS_PER_DAY = 3;
const EXPECTED_LEG_MINUTES = 10;
const MIN_TRAVEL_SECONDS = 120;

const COL = { DATE:1, ORIGIN:2, DEST:3, SEND_TIME:4, ID:5, VAN:6, STATUS:7, RECV_BY:8, RECV_TIME:9 };

/* ─── Helpers ────────────────────────────────────────── */

function resolveMainBranch(name) {
  if (!name) return name;
  var upper = name.toString().toUpperCase().trim();
  return SUB_BRANCHES[upper] || upper;
}

function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
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
  if (sheet.getMaxRows() < minRows)
    sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows() + 5);
}

function today() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd");
}

/* ─── ROUTE MOVEMENT TRACKING ──────────────────────────
   One row represents one station visit. Shipment scanning is optional; the
   route board is driven only by driver check-in/check-out events.
*/

const MOVE_COL = {
  DATE:1, VAN:2, ROUTE:3, ROUND:4, STOP:5, STATION:6,
  ARRIVAL:7, DEPARTURE:8, HOLD_SECONDS:9, TRAVEL_SECONDS:10,
  NEXT_STATION:11, STATUS:12, UPDATED:13
};

function getMovementSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VAN MOVEMENTS");
  if (!sheet) {
    sheet = ss.insertSheet("VAN MOVEMENTS");
    sheet.appendRow([
      "Date","Van No","Route","Round","Stop No","Station",
      "Check In","Check Out","Hold Seconds","Travel Seconds",
      "Next Station","Status","Last Updated"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(2, MOVE_COL.VAN, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
    sheet.getRange(2, MOVE_COL.ARRIVAL, sheet.getMaxRows() - 1, 2).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    sheet.getRange(2, MOVE_COL.UPDATED, sheet.getMaxRows() - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  }
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

function touchVanRegistry(vanNo, status) {
  var target = String(vanNo || "").trim().toUpperCase();
  if (!target) return;
  var sheet = getVanRegistrySheet();
  var values = sheet.getDataRange().getValues();
  var now = new Date();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === target) {
      sheet.getRange(i + 1, 4, 1, 2).setValues([[now, status || "ACTIVE"]]);
      return;
    }
  }
  sheet.appendRow([target, "van " + target, Utilities.formatDate(now, CONFIG.TIMEZONE, "yyyy-MM-dd"), now, status || "ACTIVE"]);
}

function routeFor(routeName) {
  return ROUTES[routeName] || null;
}

function secondsBetween(a, b) {
  var first = a instanceof Date ? a : new Date(a);
  var second = b instanceof Date ? b : new Date(b);
  if (isNaN(first.getTime()) || isNaN(second.getTime())) return 0;
  return Math.max(0, Math.floor((second.getTime() - first.getTime()) / 1000));
}

function movementRowsForToday(vanNo) {
  var values = getMovementSheet().getDataRange().getValues();
  var target = String(vanNo || "").trim();
  var dateStr = today();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][MOVE_COL.DATE - 1]).trim() === dateStr &&
        String(values[i][MOVE_COL.VAN - 1]).trim() === target) {
      rows.push({ rowIndex: i + 1, values: values[i] });
    }
  }
  return rows;
}

function getRouteState(vanNo) {
  var rows = movementRowsForToday(vanNo);
  if (!rows.length) {
    return {
      vanNo: String(vanNo || "").trim(), started: false, route: null,
      round: 0, stopIndex: -1, station: null, nextStation: null,
      status: "NOT_STARTED"
    };
  }
  var item = rows[rows.length - 1];
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
    routeLabel: ROUTE_LABELS[String(v[MOVE_COL.ROUTE - 1])] || "",
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

function knownVans() {
  var found = {};
  Object.keys(PASSCODES).forEach(function(pc) {
    if (PASSCODES[pc].role === "driver") found[PASSCODES[pc].vanNo] = true;
  });
  var registryValues = getVanRegistrySheet().getDataRange().getValues();
  for (var r = 1; r < registryValues.length; r++) {
    var registeredVan = String(registryValues[r][0] || "").trim();
    if (registeredVan) found[registeredVan] = true;
  }
  var values = getMovementSheet().getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var van = String(values[i][MOVE_COL.VAN - 1] || "").trim();
    if (van) found[van] = true;
  }
  return Object.keys(found);
}

function routeDashboard() {
  return knownVans().map(function(vanNo) {
    var state = getRouteState(vanNo);
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
      updateVanStatus(vanNo, "AT_STATION", station, route[1], "", 0);
      touchVanRegistry(vanNo, "AT_STATION");
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
    updateVanStatus(vanNo, finalStop ? "OFF_DUTY" : "AT_STATION", station, nextStation, "", nextRound);
    touchVanRegistry(vanNo, finalStop ? "OFF_DUTY" : "AT_STATION");
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
    updateVanStatus(vanNo, "MOVING", state.station, state.nextStation, now, state.round);
    touchVanRegistry(vanNo, "MOVING");
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
    if (String(values[i][MOVE_COL.DATE - 1]).trim() === day &&
        (!target || String(values[i][MOVE_COL.VAN - 1]).trim() === target)) {
      result.push({
        vanNo: values[i][MOVE_COL.VAN - 1], route: values[i][MOVE_COL.ROUTE - 1],
        round: values[i][MOVE_COL.ROUND - 1], stopNo: values[i][MOVE_COL.STOP - 1],
        station: values[i][MOVE_COL.STATION - 1], checkIn: values[i][MOVE_COL.ARRIVAL - 1],
        checkOut: values[i][MOVE_COL.DEPARTURE - 1],
        holdSeconds: Number(values[i][MOVE_COL.HOLD_SECONDS - 1]) || 0,
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
  updateVanStatus(String(vanNo).trim(), "IDLE", state.station, "", "", state.round);
  touchVanRegistry(vanNo, "RESET");
  return { success:true, message:"Van " + vanNo + " reset for a new route" };
}

/* ─── VAN STATUS sheet ──────────────────────────────── */
// Columns: Van No | Status | From Branch | To Branch | Depart Time | Round Count | Last Activity Time

function getVanStatusSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VAN STATUS");
  if (!sheet) {
    sheet = ss.insertSheet("VAN STATUS");
    sheet.appendRow(["Van No","Status","From Branch","To Branch","Depart Time","Round Count","Last Activity Time"]);
    sheet.setFrozenRows(1);
  }
  sheet.getRange(1, 7).setValue("Last Activity Time");
  return sheet;
}

function getVanStatusRow(vanNo) {
  var sheet  = getVanStatusSheet();
  var values = sheet.getDataRange().getValues();
  var target = String(vanNo).trim();
  var todayStr = today();

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === target) {
      var lastDate   = values[i][6];
      var roundCount = Number(values[i][5]) || 0;
      // Reset round count at start of new day
      var lastDateStr = lastDate instanceof Date
        ? Utilities.formatDate(lastDate, CONFIG.TIMEZONE, "yyyy-MM-dd")
        : String(lastDate).substring(0, 10);
      if (lastDate && lastDateStr !== todayStr) {
        roundCount = 0;
        sheet.getRange(i + 1, 6).setValue(0);
      }
      return {
        rowIndex:   i,
        status:     values[i][1],
        fromBranch: values[i][2],
        toBranch:   values[i][3],
        departTime: values[i][4],
        roundCount: roundCount
      };
    }
  }
  return { rowIndex: -1, status: null, fromBranch: null, toBranch: null, departTime: null, roundCount: 0 };
}

function updateVanStatus(vanNo, status, fromBranch, toBranch, departTime, roundCount) {
  var sheet    = getVanStatusSheet();
  var values   = sheet.getDataRange().getValues();
  var target   = String(vanNo).trim();
  var todayStr = today();
  var activityTime = new Date();

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === target) {
      var existingRound = (roundCount !== undefined) ? roundCount : (Number(values[i][5]) || 0);
      sheet.getRange(i + 1, 2, 1, 6).setValues([[status, fromBranch, toBranch, departTime, existingRound, activityTime]]);
      return;
    }
  }
  sheet.appendRow([target, status, fromBranch, toBranch, departTime, roundCount || 0, activityTime]);
}

/* ─── VAN VISITS sheet — dynamic round tracking ─────────
   Columns: Van No | Date | Visited Branches (comma-separated) | Round Count
   Every time a van departs from a branch, we add it to the visited list.
   When the list contains all ROUND_BRANCHES, we close the round:
     increment Round Count, clear Visited Branches list.
   This works for ANY van number with no pre-registration needed.
*/

function getVanVisitsSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VAN VISITS");
  if (!sheet) {
    sheet = ss.insertSheet("VAN VISITS");
    sheet.appendRow(["Van No","Date","Visited Branches","Round Count"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Records that vanNo departed from branch, checks if a round is now complete.
// Returns the current round count for the van (after any increment).
function trackBranchVisit(vanNo, branch) {
  var mainBranch = resolveMainBranch(branch);
  // Only count visits to main ROUND_BRANCHES
  if (ROUND_BRANCHES.indexOf(mainBranch) === -1) return getRoundCount(vanNo);

  var sheet    = getVanVisitsSheet();
  var values   = sheet.getDataRange().getValues();
  var target   = String(vanNo).trim();
  var todayStr = today();

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === target && String(values[i][1]) === todayStr) {
      // Found today's row for this van
      var visited    = values[i][2] ? String(values[i][2]).split(",") : [];
      var roundCount = Number(values[i][3]) || 0;

      // Add this branch if not already visited this leg
      if (visited.indexOf(mainBranch) === -1) visited.push(mainBranch);

      // Check if all ROUND_BRANCHES have been visited → round complete
      var roundComplete = ROUND_BRANCHES.every(function(b) { return visited.indexOf(b) !== -1; });
      if (roundComplete) {
        roundCount++;
        visited = []; // reset for next round
      }

      sheet.getRange(i + 1, 3, 1, 2).setValues([[visited.join(","), roundCount]]);
      return roundCount;
    }
  }

  // No row yet for today — create it
  var newVisited = [mainBranch];
  var newRoundCount = 0;
  sheet.appendRow([target, todayStr, newVisited.join(","), newRoundCount]);
  return newRoundCount;
}

// Returns round count for a van today (0 if not started yet).
function getRoundCount(vanNo) {
  var sheet    = getVanVisitsSheet();
  var values   = sheet.getDataRange().getValues();
  var target   = String(vanNo).trim();
  var todayStr = today();

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === target && String(values[i][1]) === todayStr) {
      return Number(values[i][3]) || 0;
    }
  }
  return 0;
}

// How many distinct ROUND_BRANCHES has this van visited today (for label).
function getVisitedBranches(vanNo) {
  var sheet    = getVanVisitsSheet();
  var values   = sheet.getDataRange().getValues();
  var target   = String(vanNo).trim();
  var todayStr = today();

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === target && String(values[i][1]) === todayStr) {
      var visited = values[i][2] ? String(values[i][2]).split(",").filter(Boolean) : [];
      return visited;
    }
  }
  return [];
}

// Builds a human-readable round label, e.g. "2/3 · 4 stops"
function buildRoundLabel(vanNo) {
  var roundCount = getRoundCount(vanNo);
  var visited    = getVisitedBranches(vanNo);
  if (roundCount === 0 && visited.length === 0) return null;
  var label = "Round " + (roundCount + 1) + "/" + MAX_ROUNDS_PER_DAY;
  if (visited.length > 0) label += " · " + visited.length + "/" + ROUND_BRANCHES.length + " stops";
  return label;
}

/* ─── Pending count helper ─────────────────────────── */

function countPendingForBranchOnVan(branch, vanNo) {
  var sheet  = getSheet();
  var values = sheet.getDataRange().getValues();
  var count  = 0;
  var target = String(vanNo).trim();
  for (var i = 1; i < values.length; i++) {
    var dest   = values[i][COL.DEST - 1];
    var van    = String(values[i][COL.VAN - 1] || "").trim();
    var status = values[i][COL.STATUS - 1];
    if (resolveMainBranch(dest) === branch && van === target && status === "In Transit") count++;
  }
  return count;
}

/* ─── getIncomingVans (Receive tab) ─────────────────── */

function getIncomingVans(branch) {
  var sheet  = getVanStatusSheet();
  var values = sheet.getDataRange().getValues();
  var now    = new Date();
  var result = [];

  for (var i = 1; i < values.length; i++) {
    if (values[i][1] === "MOVING" && values[i][3] === branch) {
      var vanNo      = String(values[i][0]).trim();
      var fromBranch = values[i][2];
      var departCell = values[i][4];
      var departDate = (departCell instanceof Date) ? departCell : new Date(departCell);
      var elapsedMs  = Math.max(0, now.getTime() - departDate.getTime());
      var elapsedMin = Math.floor(elapsedMs / 60000);
      var elapsedSec = Math.floor((elapsedMs % 60000) / 1000);
      var count      = countPendingForBranchOnVan(branch, vanNo);

      if (count > 0) {
        var entry = {
          vanNo:          vanNo,
          origin:         fromBranch,
          elapsedMinutes: elapsedMin,
          elapsedSeconds: elapsedSec,
          count:          count,
          isLate:         elapsedMin >= EXPECTED_LEG_MINUTES
        };
        var rl = buildRoundLabel(vanNo);
        if (rl) entry.roundLabel = rl;
        result.push(entry);
      }
    }
  }
  return result;
}

/* ─── getAllVanMovements (company-wide, used by frontend getAllVans) ─ */

function getAllVanMovements() {
  var sheet  = getVanStatusSheet();
  var values = sheet.getDataRange().getValues();
  var now    = new Date();
  var result = [];

  for (var i = 1; i < values.length; i++) {
    var vanNo  = String(values[i][0]).trim();
    if (!vanNo) continue;
    var status     = values[i][1];
    var fromBranch = values[i][2];
    var toBranch   = values[i][3];

    if (status === "MOVING") {
      var departCell = values[i][4];
      var departDate = (departCell instanceof Date) ? departCell : new Date(departCell);
      var elapsedMs  = Math.max(0, now.getTime() - departDate.getTime());
      var elapsedMin = Math.floor(elapsedMs / 60000);
      var elapsedSec = Math.floor((elapsedMs % 60000) / 1000);

      var entry = {
        vanNo:          vanNo,
        status:         "moving",
        fromBranch:     fromBranch,
        toBranch:       toBranch,
        elapsedMinutes: elapsedMin,
        elapsedSeconds: elapsedSec,
        isLate:         elapsedMin >= EXPECTED_LEG_MINUTES
      };
      var rl = buildRoundLabel(vanNo);
      if (rl) entry.roundLabel = rl;
      result.push(entry);

    } else if (status === "IDLE" || status === "OFF_DUTY") {
      var idleEntry = {
        vanNo:        vanNo,
        status:       status === "OFF_DUTY" ? "off_duty" : "idle",
        lastLocation: fromBranch
      };
      var rCount = getRoundCount(vanNo);
      if (rCount > 0) idleEntry.roundLabel = rCount + "/" + MAX_ROUNDS_PER_DAY + " rounds today";
      result.push(idleEntry);
    }
  }

  result.sort(function(a, b) {
    if (a.status === "moving" && b.status !== "moving") return -1;
    if (a.status !== "moving" && b.status === "moving") return  1;
    return 0;
  });

  return result;
}

// Alias for backward compat — same data, different field name convention
function getAllVanStatus() {
  var raw = getAllVanMovements();
  return raw.map(function(v) {
    if (v.status === "moving") {
      return {
        vanNo:          v.vanNo,
        status:         "MOVING",
        origin:         v.fromBranch,
        destination:    v.toBranch,
        elapsedMinutes: v.elapsedMinutes,
        elapsedSeconds: v.elapsedSeconds,
        isLate:         v.isLate,
        roundLabel:     v.roundLabel || null
      };
    }
    return { vanNo: v.vanNo, status: v.status.toUpperCase(), parkedAt: v.lastLocation, roundLabel: v.roundLabel || null };
  });
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
    if (cfg.role === "driver") touchVanRegistry(cfg.vanNo, "ACTIVE");
    var res = { success: true, role: cfg.role, routes: ROUTE_LABELS };
    if (cfg.role === "branch")  { res.branch = cfg.branch; res.destinations = DESTINATIONS[cfg.branch] || []; }
    if (cfg.role === "driver")  {
      res.vanNo = cfg.vanNo;
      res.routes = ROUTE_LABELS;
    }
    return jsonResponse(res);
  }

  if (action === "getRoutes") {
    return jsonResponse({ success:true, routes:ROUTES, labels:ROUTE_LABELS, maxRounds:MAX_ROUNDS_PER_DAY });
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

  if (action === "getPendingByVan") {
    return jsonResponse({ success: true, data: getPendingByVan(e.parameter.branch, e.parameter.vanNo) });
  }

  if (action === "getIncomingVans") {
    return jsonResponse({ success: true, data: getIncomingVans(e.parameter.branch) });
  }

  // Both action names supported — getAllVans is used by the frontend
  if (action === "getAllVanStatus") {
    return jsonResponse({ success: true, data: getAllVanMovements() });
  }

  return jsonResponse({ success: false, error: "Unknown action" });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonResponse({ success: false, error: "Bad JSON" }); }

  switch (body.action) {
    case 'routeCheckIn':  return jsonResponse(handleRouteCheckIn(body));
    case 'routeCheckOut': return jsonResponse(handleRouteCheckOut(body));
    case 'submitIssue':   return jsonResponse(submitIssue(body));
    case 'closeIssue':    return jsonResponse(closeIssue(body.row));
    case 'adminResetVan': return jsonResponse(adminResetVan(body.vanNo));
    case 'sendBatch':    return jsonResponse(handleSendBatch(body));
    case 'receiveBatch': return jsonResponse(handleReceiveBatch(body));
    case 'receiveOne':   return jsonResponse(handleReceiveOne(body));
    case 'receiveAll':   return jsonResponse(handleReceiveAll(body));
    default:             return jsonResponse({ success: false, error: 'Unknown action' });
  }
}

/* ─── Send Batch ────────────────────────────────────── */

function handleSendBatch(data) {
  try {
    var sheet    = getSheet();
    var values   = sheet.getDataRange().getValues();
    var now      = new Date();
    var tz       = CONFIG.TIMEZONE;
    var dateStr  = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    var timeStr  = Utilities.formatDate(now, tz, "HH:mm:ss");

    var vanNo  = String(data.vanNo || "N/A").trim();
    var origin = data.origin;
    var dest   = data.destination;

    // Track branch visit for round counting
    var roundCount = trackBranchVisit(vanNo, origin);

    var inTransitIds = {};
    for (var i = 1; i < values.length; i++) {
      var sid    = String(values[i][COL.ID - 1] || "").trim();
      var status = values[i][COL.STATUS - 1];
      if (sid && status === "In Transit") inTransitIds[sid] = values[i][COL.ORIGIN - 1];
    }

    var newRows = [], results = [];
    var shipments = data.shipments || [];

    for (var j = 0; j < shipments.length; j++) {
      var item = shipments[j];
      var sid2 = String(item.shipmentId || "").trim();
      if (!sid2) continue;
      if (inTransitIds[sid2]) {
        results.push({ shipmentId: sid2, status: "duplicate", error: sid2 + " already In Transit from " + inTransitIds[sid2] });
      } else {
        newRows.push([dateStr, origin, dest, timeStr, sid2, vanNo, "In Transit", "", ""]);
        results.push({ shipmentId: sid2, status: "sent" });
        inTransitIds[sid2] = origin;
      }
    }

    if (newRows.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      ensureRows(sheet, startRow + newRows.length);
      sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
      sheet.getRange(startRow, COL.VAN, newRows.length, 1).setNumberFormat('@');
      updateVanStatus(vanNo, "MOVING", origin, dest, now, roundCount);
    }

    var rl = buildRoundLabel(vanNo);
    var response = { success: true, sent: newRows.length, results: results, destination: dest };
    if (rl) response.routeLabel = rl;
    return response;

  } catch (err) { return { success: false, error: err.toString() }; }
}

/* ─── Receive Batch ─────────────────────────────────── */

function handleReceiveBatch(data) {
  try {
    var sheet    = getSheet();
    var range    = sheet.getDataRange();
    var values   = range.getValues();
    var now      = new Date();
    var timeStr  = Utilities.formatDate(now, CONFIG.TIMEZONE, "HH:mm:ss");

    var targetIds = {};
    (data.shipmentIds || []).forEach(function(id) { targetIds[String(id).trim()] = true; });

    var targetVan    = String(data.vanNo).trim();
    var branch       = data.branch;
    var receivedCount= 0;
    var foundIds     = {};
    var notFoundIds  = [];

    for (var i = 1; i < values.length; i++) {
      var sid    = String(values[i][COL.ID - 1] || "").trim();
      var dest   = values[i][COL.DEST - 1];
      var van    = String(values[i][COL.VAN - 1] || "").trim();
      var status = values[i][COL.STATUS - 1];

      if (targetIds[sid] && resolveMainBranch(dest) === branch &&
          van === targetVan && status === "In Transit" && !foundIds[sid]) {
        values[i][COL.STATUS - 1]   = "Received";
        values[i][COL.RECV_BY - 1]  = branch;
        values[i][COL.RECV_TIME - 1]= timeStr;
        receivedCount++;
        foundIds[sid] = true;
      }
    }

    range.setValues(values);

    if (receivedCount > 0) {
      updateVanStatus(targetVan, "IDLE", branch, "", "");
    }

    (data.shipmentIds || []).forEach(function(id) {
      if (!foundIds[String(id).trim()]) notFoundIds.push(id);
    });

    return {
      success:       true,
      message:       receivedCount + " received" + (notFoundIds.length ? ", " + notFoundIds.length + " not found" : ""),
      receivedCount: receivedCount,
      notFoundIds:   notFoundIds
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

/* ─── Receive One ───────────────────────────────────── */

function handleReceiveOne(data) {
  try {
    var sheet  = getSheet();
    var range  = sheet.getDataRange();
    var values = range.getValues();
    var now    = new Date();
    var timeStr= Utilities.formatDate(now, CONFIG.TIMEZONE, "HH:mm:ss");

    for (var i = 1; i < values.length; i++) {
      var sid    = String(values[i][COL.ID - 1] || "").trim();
      var dest   = values[i][COL.DEST - 1];
      var van    = String(values[i][COL.VAN - 1] || "").trim();
      var status = values[i][COL.STATUS - 1];

      if (sid === String(data.shipmentId || "").trim() &&
          resolveMainBranch(dest) === data.branch &&
          van === String(data.vanNo).trim() &&
          status === "In Transit") {
        values[i][COL.STATUS - 1]    = "Received";
        values[i][COL.RECV_BY - 1]   = data.branch;
        values[i][COL.RECV_TIME - 1] = timeStr;
        range.setValues(values);
        updateVanStatus(String(data.vanNo).trim(), "IDLE", data.branch, "", "");
        return { success: true, message: data.shipmentId + " received" };
      }
    }
    return { success: false, error: "Not found or already received" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

/* ─── Receive All ───────────────────────────────────── */

function handleReceiveAll(data) {
  try {
    var sheet  = getSheet();
    var range  = sheet.getDataRange();
    var values = range.getValues();
    var count  = 0;
    var timeStr= Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "HH:mm:ss");

    for (var i = 1; i < values.length; i++) {
      var dest   = values[i][COL.DEST - 1];
      var van    = String(values[i][COL.VAN - 1] || "").trim();
      var status = values[i][COL.STATUS - 1];

      if (resolveMainBranch(dest) === data.branch &&
          van === String(data.vanNo).trim() &&
          status === "In Transit") {
        values[i][COL.STATUS - 1]    = "Received";
        values[i][COL.RECV_BY - 1]   = data.branch;
        values[i][COL.RECV_TIME - 1] = timeStr;
        count++;
      }
    }
    range.setValues(values);
    if (count > 0) updateVanStatus(String(data.vanNo).trim(), "IDLE", data.branch, "", "");
    return { success: true, message: count + " shipment(s) received", count: count };
  } catch (err) { return { success: false, error: err.toString() }; }
}

/* ─── Pending ───────────────────────────────────────── */

function getPendingByVan(branch, vanNo) {
  var sheet  = getSheet();
  var values = sheet.getDataRange().getValues();
  var result = [];
  var target = String(vanNo).trim();

  for (var i = 1; i < values.length; i++) {
    var rowDest   = values[i][COL.DEST - 1];
    var rowVan    = String(values[i][COL.VAN - 1] || "").trim();
    var rowStatus = values[i][COL.STATUS - 1];

    if (resolveMainBranch(rowDest) === branch && rowVan === target && rowStatus === "In Transit") {
      result.push({ shipmentId: values[i][COL.ID - 1], origin: values[i][COL.ORIGIN - 1], destination: rowDest });
    }
  }
  return result;
}

/* ─── Response helper ───────────────────────────────── */

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
