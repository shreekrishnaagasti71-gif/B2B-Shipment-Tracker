/*
 * NCM B2B Tracker — Google Sheets backend
 *
 * Bind this script to the Google Sheet that should hold the data.
 * Run setup() once, then deploy as a Web app:
 *   Execute as: Me
 *   Who has access: Anyone
 *
 * The React app can use the deployment URL as VITE_API_BASE_URL.
 * This adapter intentionally mirrors the /api responses used by the app.
 */

const TZ = "Asia/Kathmandu";
const SHEETS = {
  shipments: {
    name: "Shipments",
    headers: ["id", "shipmentId", "origin", "destination", "vanNo", "status", "sentAt", "receivedAt", "receivedBy"],
  },
  movements: {
    name: "Movements",
    headers: ["id", "movementDate", "vanNo", "branch", "arrivalAt", "departureAt", "holdSeconds", "nextBranch", "travelSeconds", "status", "round", "lastUpdatedAt"],
  },
  issues: {
    name: "Issues",
    headers: ["id", "reportedBy", "role", "branch", "vanNo", "message", "status", "createdAt"],
  },
};

const BRANCHES = [
  { name: "TINKUNE", code: "TINKUNE", driverOnly: false },
  { name: "CHABAHIL", code: "CHABAHIL", driverOnly: false },
  { name: "NAYA BUSPARK", code: "NAYA-BUSPARK", driverOnly: false },
  { name: "KALANKI", code: "KALANKI", driverOnly: false },
  { name: "SATDOBATO", code: "SATDOBATO", driverOnly: false },
  { name: "NEWROAD", code: "NEWROAD", driverOnly: false },
  { name: "BASUNDHARA", code: "BASUNDHARA", driverOnly: true },
  { name: "SWOYAMBHU", code: "SWOYAMBHU", driverOnly: true },
];
const BRANCH_NAMES = BRANCHES.map((branch) => branch.name);
const ROUTE_A = ["TINKUNE", "CHABAHIL", "BASUNDHARA", "NAYA BUSPARK", "SWOYAMBHU", "KALANKI", "SATDOBATO"];
const ROUTE_B = ["TINKUNE", "SATDOBATO", "KALANKI", "SWOYAMBHU", "NAYA BUSPARK", "BASUNDHARA", "CHABAHIL"];
const PASSCODES = {
  "1111": { role: "BRANCH", branch: "TINKUNE" },
  "2222": { role: "BRANCH", branch: "CHABAHIL" },
  "3333": { role: "BRANCH", branch: "NAYA BUSPARK" },
  "4444": { role: "BRANCH", branch: "KALANKI" },
  "5555": { role: "BRANCH", branch: "SATDOBATO" },
  "6666": { role: "BRANCH", branch: "NEWROAD" },
  "7777": { role: "DRIVER" },
  "9999": { role: "ADMIN" },
};

function setup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach((key) => {
    const config = SHEETS[key];
    const sheet = spreadsheet.getSheetByName(config.name) || spreadsheet.insertSheet(config.name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
      sheet.setFrozenRows(1);
    }
  });
  spreadsheet.toast("NCM tracker sheets are ready.");
}

function doGet(event) {
  return route_(event, "GET");
}

function doPost(event) {
  return route_(event, "POST");
}

function route_(event, method) {
  try {
    const path = path_(event);
    const body = method === "POST" ? body_(event) : {};
    if (method === "POST" && path === "session") return json_(session_(body));
    if (method === "GET" && path === "movement/branches") return json_(BRANCHES);
    if (method === "GET" && path === "dashboard") return json_(dashboard_());
    if (method === "GET" && path === "shipments") return json_(listShipments_(event.parameter || {}));
    if (method === "GET" && path === "movement/vans") return json_(listVans_());
    if (method === "GET" && path.indexOf("movement/vans/") === 0) return json_(movement_(decodeURIComponent(path.split("/").pop())));
    if (method === "GET" && path === "issues") return json_(listIssues_(event.parameter || {}));
    if (method === "POST" && path === "shipments/send") return json_(sendShipments_(body));
    if (method === "POST" && path === "shipments/receive") return json_(receiveShipments_(body));
    if (method === "POST" && path === "movement/check-in") return json_(checkIn_(body));
    if (method === "POST" && path === "movement/check-out") return json_(checkOut_(body));
    if (method === "POST" && path === "issues") return json_(createIssue_(body));
    if (method === "POST" && path.indexOf("issues/") === 0 && path.endsWith("/close")) {
      return json_(closeIssue_(path.split("/")[1]));
    }
    return json_({ error: "Not found", path: path }, 404);
  } catch (error) {
    return json_({ error: error.message || String(error) }, 400);
  }
}

function path_(event) {
  let value = String((event && event.pathInfo) || (event && event.parameter && event.parameter.path) || "").replace(/^\/+/, "");
  if (value.indexOf("api/") === 0) value = value.slice(4);
  return value || "dashboard";
}

function body_(event) {
  const raw = event && event.postData && event.postData.contents;
  if (!raw) return {};
  return typeof raw === "string" ? JSON.parse(raw) : raw;
}

function json_(value, status) {
  // Apps Script ContentService does not expose custom response headers. The
  // web app is called with a simple text/plain POST by the React client, so
  // browsers do not issue a preflight request.
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function spreadsheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error("Bind this Apps Script project to a Google Sheet first.");
  return spreadsheet;
}

function sheet_(key) {
  const config = SHEETS[key];
  const sheet = spreadsheet_().getSheetByName(config.name);
  if (!sheet) throw new Error("Run setup() before using the tracker.");
  return sheet;
}

function records_(key) {
  const config = SHEETS[key];
  const sheet = sheet_(key);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, config.headers.length).getValues();
  return values.map((row, index) => {
    const record = { _row: index + 2 };
    config.headers.forEach((header, column) => { record[header] = row[column]; });
    return record;
  });
}

function nextId_(key) {
  const ids = records_(key).map((record) => Number(record.id) || 0);
  return (ids.length ? Math.max.apply(null, ids) : 0) + 1;
}

function append_(key, record) {
  const config = SHEETS[key];
  const values = config.headers.map((header) => record[header] === undefined ? "" : record[header]);
  sheet_(key).appendRow(values);
  return record;
}

function update_(key, record) {
  const config = SHEETS[key];
  const values = config.headers.map((header) => record[header] === undefined ? "" : record[header]);
  sheet_(key).getRange(record._row, 1, 1, config.headers.length).setValues([values]);
  return record;
}

function clean_(value) {
  return String(value == null ? "" : value).trim().toUpperCase();
}

function iso_(value) {
  if (!value) return null;
  const date = value instanceof Date ? value : new Date(value);
  return isNaN(date.getTime()) ? null : date.toISOString();
}

function seconds_(value, end) {
  if (!value) return 0;
  const start = value instanceof Date ? value.getTime() : new Date(value).getTime();
  const finish = (end || new Date()).getTime();
  return isNaN(start) ? 0 : Math.max(0, Math.floor((finish - start) / 1000));
}

function today_() {
  const date = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");
  const start = new Date(date + "T00:00:00+05:45");
  return { date: date, start: start, end: new Date(start.getTime() + 86400000) };
}

function inToday_(value) {
  const bounds = today_();
  const date = value instanceof Date ? value : new Date(value);
  return !isNaN(date.getTime()) && date >= bounds.start && date < bounds.end;
}

function session_(body) {
  const config = PASSCODES[String(body.passcode || "").trim()];
  if (!config) throw new Error("Invalid passcode");
  return {
    role: config.role,
    branch: config.branch || null,
    destinations: [],
    driverBranches: BRANCH_NAMES,
  };
}

function listShipments_(params) {
  const status = params.status || "ALL";
  const branch = clean_(params.branch);
  const limit = Math.min(Number(params.limit) || 100, 500);
  return records_("shipments")
    .filter((row) => (status === "ALL" || row.status === status))
    .filter((row) => !branch || clean_(row.origin) === branch || clean_(row.destination) === branch)
    .sort((a, b) => new Date(b.sentAt).getTime() - new Date(a.sentAt).getTime())
    .slice(0, limit)
    .map(mapShipment_);
}

function mapShipment_(row) {
  return {
    id: Number(row.id),
    shipmentId: String(row.shipmentId),
    origin: String(row.origin),
    destination: String(row.destination),
    vanNo: String(row.vanNo),
    status: row.status === "RECEIVED" ? "RECEIVED" : "IN_TRANSIT",
    sentAt: iso_(row.sentAt),
    receivedAt: iso_(row.receivedAt),
    receivedBy: row.receivedBy || null,
    elapsedSeconds: seconds_(row.sentAt, row.receivedAt ? new Date(row.receivedAt) : new Date()),
  };
}

function sendShipments_(body) {
  const origin = clean_(body.origin);
  const destination = clean_(body.destination);
  const vanNo = clean_(body.vanNo);
  const ids = Array.from(new Set((body.shipmentIds || []).map(clean_).filter(Boolean)));
  if (!BRANCH_NAMES.includes(origin) || !BRANCH_NAMES.includes(destination) || !vanNo || !ids.length) {
    throw new Error("Choose valid branches, van number, and shipment IDs");
  }
  const existing = records_("shipments");
  const duplicateSet = new Set(existing.filter((row) => row.status === "IN_TRANSIT").map((row) => String(row.shipmentId)));
  const duplicates = ids.filter((id) => duplicateSet.has(id));
  const accepted = ids.filter((id) => !duplicateSet.has(id));
  const inserted = accepted.map((shipmentId) => append_("shipments", {
    id: nextId_("shipments"),
    shipmentId: shipmentId,
    origin: origin,
    destination: destination,
    vanNo: vanNo,
    status: "IN_TRANSIT",
    sentAt: new Date(),
    receivedAt: "",
    receivedBy: "",
  }));
  return { processed: ids.length, received: inserted.length, duplicates: duplicates, notFound: [], shipments: inserted.map(mapShipment_) };
}

function receiveShipments_(body) {
  const branch = clean_(body.branch);
  const vanNo = clean_(body.vanNo);
  const ids = Array.from(new Set((body.shipmentIds || []).map(clean_).filter(Boolean)));
  const rows = records_("shipments");
  const eligible = rows.filter((row) =>
    ids.indexOf(String(row.shipmentId)) >= 0 &&
    row.status === "IN_TRANSIT" &&
    clean_(row.destination) === branch &&
    clean_(row.vanNo) === vanNo
  );
  const eligibleIds = new Set(eligible.map((row) => String(row.shipmentId)));
  const notFound = ids.filter((id) => !eligibleIds.has(id));
  const now = new Date();
  eligible.forEach((row) => {
    row.status = "RECEIVED";
    row.receivedAt = now;
    row.receivedBy = branch;
    update_("shipments", row);
  });
  return { processed: ids.length, received: eligible.length, duplicates: [], notFound: notFound, shipments: eligible.map(mapShipment_) };
}

function movementRows_(vanNo) {
  return records_("movements")
    .filter((row) => clean_(row.vanNo) === clean_(vanNo) && inToday_(row.arrivalAt))
    .sort((a, b) => new Date(a.arrivalAt).getTime() - new Date(b.arrivalAt).getTime());
}

function route_(rows) {
  const first = rows.find((row) => row.nextBranch);
  return first && clean_(first.nextBranch) === "SATDOBATO" ? ROUTE_B.slice() : ROUTE_A.slice();
}

function summary_(rows, currentRound, currentBranch) {
  const route = route_(rows);
  const roundRows = rows.filter((row) => Number(row.round) === Number(currentRound));
  const visited = Array.from(new Set(roundRows.map((row) => clean_(row.branch)).filter((branch) => route.indexOf(branch) >= 0)));
  const missing = route.filter((branch) => visited.indexOf(branch) < 0);
  return {
    round: Number(currentRound),
    route: route,
    visited: visited,
    missing: missing,
    completed: clean_(currentBranch) === "TINKUNE" && visited.length >= route.length,
  };
}

function movement_(vanNo) {
  const rows = movementRows_(vanNo);
  if (!rows.length) {
    return {
      vanNo: clean_(vanNo), status: "NOT_STARTED", branch: null, nextBranch: null,
      arrivalAt: null, departureAt: null, elapsedSeconds: 0, holdSeconds: 0, travelSeconds: 0,
      round: 1, roundSummary: { round: 1, route: ROUTE_A.slice(), visited: [], missing: ROUTE_A.slice(), completed: false },
      lastUpdatedAt: new Date().toISOString(),
    };
  }
  const row = rows[rows.length - 1];
  const moving = row.status === "MOVING";
  return {
    vanNo: String(row.vanNo), status: moving ? "MOVING" : "AT_BRANCH",
    branch: String(row.branch), nextBranch: row.nextBranch || null,
    arrivalAt: iso_(row.arrivalAt), departureAt: iso_(row.departureAt),
    elapsedSeconds: moving ? seconds_(row.departureAt) : seconds_(row.arrivalAt),
    holdSeconds: Number(row.holdSeconds) || 0,
    travelSeconds: Number(row.travelSeconds) || (moving ? seconds_(row.departureAt) : 0),
    round: Number(row.round) || 1,
    roundSummary: summary_(rows, Number(row.round) || 1, row.branch),
    lastUpdatedAt: iso_(row.lastUpdatedAt) || new Date().toISOString(),
  };
}

function listVans_() {
  const vans = Array.from(new Set(records_("movements").filter((row) => inToday_(row.arrivalAt)).map((row) => clean_(row.vanNo))));
  return vans.map(movement_);
}

function checkIn_(body) {
  const vanNo = clean_(body.vanNo);
  const branch = clean_(body.branch);
  if (!vanNo || BRANCH_NAMES.indexOf(branch) < 0) throw new Error("Choose a valid van and branch");
  const rows = movementRows_(vanNo);
  const previous = rows[rows.length - 1];
  if (!previous) {
    append_("movements", {
      id: nextId_("movements"), movementDate: today_().date, vanNo: vanNo, branch: branch,
      arrivalAt: new Date(), departureAt: "", holdSeconds: 0, nextBranch: "",
      travelSeconds: 0, status: "AT_BRANCH", round: 1, lastUpdatedAt: new Date(),
    });
    return movement_(vanNo);
  }
  if (previous.status !== "MOVING") throw new Error("Van is already at " + previous.branch);
  const now = new Date();
  previous.status = "COMPLETE";
  previous.travelSeconds = seconds_(previous.departureAt, now);
  previous.lastUpdatedAt = now;
  update_("movements", previous);
  append_("movements", {
    id: nextId_("movements"), movementDate: today_().date, vanNo: vanNo, branch: branch,
    arrivalAt: now, departureAt: "", holdSeconds: 0, nextBranch: "",
    travelSeconds: 0, status: "AT_BRANCH",
    round: branch === "TINKUNE" && clean_(previous.branch) !== "TINKUNE" ? Number(previous.round) + 1 : Number(previous.round),
    lastUpdatedAt: now,
  });
  return movement_(vanNo);
}

function checkOut_(body) {
  const vanNo = clean_(body.vanNo);
  const nextBranch = clean_(body.nextBranch);
  if (!vanNo || BRANCH_NAMES.indexOf(nextBranch) < 0) throw new Error("Choose a valid van and next branch");
  const rows = movementRows_(vanNo);
  const current = rows[rows.length - 1];
  if (!current) throw new Error("Check in at a branch first");
  if (current.status === "MOVING") throw new Error("Already heading to " + current.nextBranch);
  const now = new Date();
  current.departureAt = now;
  current.holdSeconds = seconds_(current.arrivalAt, now);
  current.nextBranch = nextBranch;
  current.status = "MOVING";
  current.lastUpdatedAt = now;
  update_("movements", current);
  return movement_(vanNo);
}

function listIssues_(params) {
  const status = params.status || "OPEN";
  return records_("issues")
    .filter((row) => status === "ALL" || row.status === status)
    .sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime())
    .slice(0, 100)
    .map(mapIssue_);
}

function mapIssue_(row) {
  return {
    id: Number(row.id), reportedBy: String(row.reportedBy), role: String(row.role),
    branch: row.branch || null, vanNo: row.vanNo || null, message: String(row.message),
    status: row.status === "CLOSED" ? "CLOSED" : "OPEN", createdAt: iso_(row.createdAt),
  };
}

function createIssue_(body) {
  const reportedBy = String(body.reportedBy || "").trim();
  const message = String(body.message || "").trim();
  if (!reportedBy || !message) throw new Error("Name and message are required");
  const row = append_("issues", {
    id: nextId_("issues"), reportedBy: reportedBy,
    role: String(body.role || "BRANCH"), branch: body.branch ? clean_(body.branch) : "",
    vanNo: body.vanNo ? clean_(body.vanNo) : "", message: message,
    status: "OPEN", createdAt: new Date(),
  });
  return mapIssue_(row);
}

function closeIssue_(id) {
  const rows = records_("issues");
  const row = rows.find((item) => String(item.id) === String(id));
  if (!row) throw new Error("Issue not found");
  row.status = "CLOSED";
  return mapIssue_(update_("issues", row));
}

function dashboard_() {
  const shipments = records_("shipments").sort((a, b) => new Date(b.sentAt).getTime() - new Date(a.sentAt).getTime()).slice(0, 100);
  const vans = listVans_();
  const issues = records_("issues").filter((row) => row.status === "OPEN").sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()).slice(0, 10);
  return {
    generatedAt: new Date().toISOString(),
    inTransitShipments: shipments.filter((row) => row.status === "IN_TRANSIT").length,
    receivedToday: shipments.filter((row) => row.status === "RECEIVED" && inToday_(row.receivedAt)).length,
    movingVans: vans.filter((van) => van.status === "MOVING").length,
    atBranchVans: vans.filter((van) => van.status === "AT_BRANCH").length,
    openIssues: issues.length,
    vans: vans,
    recentShipments: shipments.slice(0, 8).map(mapShipment_),
    recentIssues: issues.map(mapIssue_),
  };
}
