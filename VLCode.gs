// ══════════════════════════════════════════════════════
//  Project Nazuna — Leave Application Command Center
//  Code.gs  v8.0 — Indexed, Stateless, Rate-Limited
// ══════════════════════════════════════════════════════

var SHEET_NAME  = "Leave Results V2.0";
var CONFIG_NAME = "FormConfig";
var SUP_NAME    = "Supervisors";
var LOG_NAME    = "System Log";

// ── STRUCTURED LOGGER ────────────────────────────────
function _log(level, fn, msg, extra) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(LOG_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_NAME);
      sheet.getRange("A1:F1").setValues([["Timestamp","Level","Function","Actor","Message","Extra"]])
        .setBackground("#202124").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160); sheet.setColumnWidth(3, 140);
      sheet.setColumnWidth(4, 180); sheet.setColumnWidth(5, 300); sheet.setColumnWidth(6, 200);
    }
    var actor = "";
    try { actor = Session.getActiveUser().getEmail(); } catch(e) {}
    var tz = Session.getScriptTimeZone();
    sheet.appendRow([
      Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss"),
      level, fn, actor, msg,
      extra ? JSON.stringify(extra) : ""
    ]);
    Logger.log("[" + level + "] " + fn + " — " + msg);
  } catch(e) { Logger.log("LOG_FAIL: " + e.message); }
}

// ── RATE LIMITER (uses CacheService) ─────────────────
// key: string identifier, maxCalls: per window, windowSec: window in seconds
function _checkRateLimit(key, maxCalls, windowSec) {
  var cache   = CacheService.getScriptCache();
  var cacheKey = "rl_" + key;
  var raw     = cache.get(cacheKey);
  var count   = raw ? parseInt(raw) : 0;
  if (count >= maxCalls) {
    _log("WARN", "_checkRateLimit", "Rate limit hit", { key: key, count: count });
    return false;
  }
  cache.put(cacheKey, String(count + 1), windowSec);
  return true;
}

// ── ROW INDEX: stable UUID-based row lookup ───────────
// Ensures column "Row UUID" exists and backfills missing UUIDs.
// Returns a map of { uuid -> 1-based row number } for O(1) lookup.
function _getRowIndex(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var uuidColIdx = headers.findIndex(function(h) {
    return h.toString().replace(/[^a-z0-9]/gi,'').toLowerCase() === "rowuuid";
  });

  // Create the column if missing
  if (uuidColIdx === -1) {
    uuidColIdx = headers.length;
    sheet.getRange(1, uuidColIdx + 1).setValue("Row UUID")
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    _log("INFO", "_getRowIndex", "Created Row UUID column", { col: uuidColIdx + 1 });
  }

  var lastRow = sheet.getLastRow();
  var index   = {};
  if (lastRow < 2) return { index: index, uuidCol: uuidColIdx + 1 };

  var uuidRange  = sheet.getRange(2, uuidColIdx + 1, lastRow - 1, 1);
  var uuidValues = uuidRange.getValues();
  var toWrite    = [];
  var hasGaps    = false;

  for (var i = 0; i < uuidValues.length; i++) {
    var existing = (uuidValues[i][0] || "").toString().trim();
    if (!existing) {
      existing = Utilities.getUuid();
      uuidValues[i][0] = existing;
      hasGaps = true;
    }
    index[existing] = i + 2; // 1-based row number
  }

  if (hasGaps) {
    uuidRange.setValues(uuidValues);
    _log("INFO", "_getRowIndex", "Backfilled missing UUIDs", { count: Object.keys(index).length });
  }

  return { index: index, uuidCol: uuidColIdx + 1 };
}

// ── SHEET HELPERS ─────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('VLIndex')
    .setTitle('Google Play VL Calendar')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSessionEmail() {
  try { return Session.getActiveUser().getEmail(); } catch(e) { return ""; }
}

function getOrCreateConfigSheet() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName(CONFIG_NAME);
  if (!cfg) {
    cfg = ss.insertSheet(CONFIG_NAME);
    cfg.getRange("A1:B1").setValues([["Setting","Value"]])
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    cfg.getRange("A2:B4").setValues([
      ["IsOpen",       "TRUE"],
      ["ActiveMonths", "May 2026, June 2026"],
      ["ClosedMsg",    "The VL filing period is currently closed."]
    ]);
    cfg.setColumnWidth(1, 160); cfg.setColumnWidth(2, 420); cfg.setFrozenRows(1);
    _log("INFO", "getOrCreateConfigSheet", "Created FormConfig sheet");
  }
  return cfg;
}

function getOrCreateSupervisorSheet() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sup = ss.getSheetByName(SUP_NAME);
  if (!sup) {
    sup = ss.insertSheet(SUP_NAME);
    sup.getRange("A1:C1").setValues([["Supervisor Email","Team Name","Role"]])
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    sup.getRange("A2:C2").setValues([["stevenjosephc@google.com","Team Steven","Admin"]]);
    sup.setColumnWidth(1, 250); sup.setColumnWidth(2, 150); sup.setColumnWidth(3, 100);
    sup.setFrozenRows(1);
    _log("INFO", "getOrCreateSupervisorSheet", "Created Supervisors sheet");
  }
  return sup;
}

function getFormConfig() {
  var res = { isOpen: true, activeMonths: ["May 2026", "June 2026"], closedMsg: "Form is closed.", isSupervisor: false, isAdmin: false, teamName: "" };
  try {
    var data = getOrCreateConfigSheet().getRange("A2:B10").getValues();
    data.forEach(function(r) {
      var k = (r[0]||"").toString().trim();
      var v = (r[1]||"").toString().trim();
      if (!k) return;
      if (k === "IsOpen")       res.isOpen       = (v.toUpperCase() === "TRUE");
      if (k === "ActiveMonths") res.activeMonths = v.split(",").map(function(m){ return m.trim(); }).filter(Boolean);
      if (k === "ClosedMsg")    res.closedMsg    = v;
    });
  } catch(e) { _log("ERROR", "getFormConfig", "Config read failed", { err: e.message }); }

  try {
    var email = getSessionEmail().toLowerCase();
    if (email) {
      var supData = getOrCreateSupervisorSheet().getDataRange().getValues();
      for (var i = 1; i < supData.length; i++) {
        if ((supData[i][0]||"").toString().trim().toLowerCase() === email) {
          res.isSupervisor = true;
          res.teamName     = (supData[i][1]||"").toString().trim();
          res.isAdmin      = (supData[i][2]||"").toString().trim().toLowerCase() === "admin";
          break;
        }
      }
    }
  } catch(e) { _log("ERROR", "getFormConfig", "Supervisor lookup failed", { err: e.message }); }

  return res;
}

function getVLSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var h = ["Timestamp","Email Address","LDAP","Channel","VL Date","Team Lead",
             "Reason for VL","Work Group","Site","Accruals Snip-it","Accruals",
             "Month","Date of Birthday","Proof / Artifacts","Status","Comments",
             "Attendance","Confirmation on Status","Email Sent","Row UUID"];
    sheet.getRange(1,1,1,h.length).setValues([h])
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
    _log("INFO", "getVLSheet", "Created " + SHEET_NAME + " sheet");
  }
  return sheet;
}

function ensureColumn(sheet, headerName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idx = headers.findIndex(function(h) {
    return h.toString().toLowerCase().replace(/[^a-z0-9]/g,'') === headerName.toLowerCase().replace(/[^a-z0-9]/g,'');
  });
  if (idx === -1) {
    sheet.getRange(1, headers.length + 1).setValue(headerName)
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    _log("INFO", "ensureColumn", "Added column: " + headerName);
    return headers.length;
  }
  return idx;
}

// ── CALENDAR DATA (fixed closure bug + UUID IDs) ──────
function getCalendarData() {
  try {
    var sheet   = getVLSheet();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return [];

    // Ensure UUID index exists and is backfilled
    var rowIdx  = _getRowIndex(sheet);
    lastCol     = sheet.getLastColumn(); // refresh after possible UUID column add

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx;
    });

    // Re-read lastRow in case UUID backfill changed row count
    lastRow = sheet.getLastRow();
    var data        = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var displayData = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    var tz          = Session.getScriptTimeZone();
    var events      = [];

    for (var i = 0; i < data.length; i++) {
      // ── FIX: capture row-specific references, not closure over loop var ──
      var row        = data[i];
      var displayRow = displayData[i];

      // ── FIX: inline column lookup to avoid stale closure bug ──
      var getVal  = (function(r, cm) {
        return function(colName) {
          var idx = cm[colName.replace(/[^a-zA-Z0-9]/g,'').toLowerCase()];
          return (idx !== undefined && r[idx] != null) ? r[idx] : "";
        };
      })(row, colMap);

      var getDisp = (function(dr, cm) {
        return function(colName) {
          var idx = cm[colName.replace(/[^a-zA-Z0-9]/g,'').toLowerCase()];
          return (idx !== undefined && dr[idx] != null) ? dr[idx] : "";
        };
      })(displayRow, colMap);

      var rawTs  = getVal("Timestamp");
      var rawLdap = getVal("LDAP").toString().trim();
      if (!rawTs || !rawLdap) continue;

      // Format VL date
      var formattedDate = "";
      var rawDate = getVal("VL Date");
      if (Object.prototype.toString.call(rawDate) === '[object Date]' && !isNaN(rawDate.getTime())) {
        formattedDate = Utilities.formatDate(rawDate, tz, "yyyy-MM-dd");
      } else if (rawDate) {
        var parsed = new Date(rawDate);
        formattedDate = !isNaN(parsed.getTime())
          ? Utilities.formatDate(parsed, tz, "yyyy-MM-dd")
          : rawDate.toString();
      }

      // Format timestamp
      var tsStr = "Unknown";
      if (Object.prototype.toString.call(rawTs) === '[object Date]' && !isNaN(rawTs.getTime())) {
        tsStr = Utilities.formatDate(rawTs, tz, "MMM d, yyyy h:mm a");
      } else {
        tsStr = rawTs.toString();
      }

      // Use stable UUID as event ID
      var uuid = getVal("Row UUID").toString().trim();
      if (!uuid) uuid = "ev_" + i; // fallback for rows not yet indexed

      events.push({
        id           : uuid,
        rowNum       : i + 2,
        timestamp    : tsStr,
        email        : getVal("Email Address").toString().trim().toLowerCase(),
        ldap         : getVal("LDAP").toString().trim(),
        channel      : getVal("Channel").toString().trim(),
        date         : formattedDate,
        teamLead     : getVal("Team Lead").toString().trim(),
        reason       : getVal("Reason for VL").toString().trim(),
        workGroup    : getVal("Work Group").toString().trim(),
        site         : getVal("Site").toString().trim(),
        accruals     : getDisp("Accruals").toString().trim(),
        accrualsProof: getVal("Accruals Snip-it").toString().trim(),
        month        : getVal("Month").toString().trim(),
        status       : getVal("Status").toString().trim() || "Pending",
        remarks      : getVal("Comments").toString().trim() || getVal("Remarks").toString().trim(),
        attendance   : getDisp("Attendance").toString().trim(),
        confirmation : getVal("Confirmation on Status").toString().trim(),
        emailSent    : getVal("Email Sent").toString().trim(),
        proof        : getVal("Proof / Artifacts").toString().trim()
      });
    }

    _log("INFO", "getCalendarData", "Loaded events", { count: events.length });
    return events;

  } catch(e) {
    _log("ERROR", "getCalendarData", e.message);
    return [];
  }
}

// ── STATUS UPDATE (UUID-indexed, rate-limited, logged) ─
function updateRequestStatus(eventUuid, oldStatus, newStatus, adminComment) {
  var actor = getSessionEmail();

  // Rate limit: max 30 status changes per minute per script instance
  if (!_checkRateLimit("statusUpdate_" + actor, 30, 60)) {
    return { success: false, message: "Rate limit exceeded. Please wait before making more changes." };
  }

  try {
    var sheet   = getVLSheet();
    var rowIdx  = _getRowIndex(sheet);
    var rowNum  = rowIdx.index[eventUuid];

    if (!rowNum) {
      _log("WARN", "updateRequestStatus", "UUID not found in index", { uuid: eventUuid });
      return { success: false, message: "Row not found. The calendar may be out of sync — please refresh." };
    }

    var statusColIdx = ensureColumn(sheet, "Status") + 1;
    var confColIdx   = ensureColumn(sheet, "Confirmation on Status") + 1;
    var commentColIdx= ensureColumn(sheet, "Comments") + 1;

    // Concurrency check: read current status from the actual row
    var currentStatus    = sheet.getRange(rowNum, statusColIdx).getValue().toString().trim();
    var normalizedCurrent = currentStatus === "" ? "Pending" : currentStatus;
    var normalizedOld     = oldStatus === "" ? "Pending" : oldStatus;

    if (normalizedCurrent.toLowerCase() !== normalizedOld.toLowerCase()) {
      _log("WARN", "updateRequestStatus", "Concurrency conflict", {
        uuid: eventUuid, expected: normalizedOld, found: normalizedCurrent
      });
      return {
        success: false,
        message: "Another admin already updated this to '" + currentStatus + "'.",
        newStatus: currentStatus
      };
    }

    var tz    = Session.getScriptTimeZone();
    var stamp = actor.split('@')[0] + " — " + Utilities.formatDate(new Date(), tz, "MM/dd/yy HH:mm");

    sheet.getRange(rowNum, statusColIdx).setValue(newStatus);
    sheet.getRange(rowNum, confColIdx).setValue(stamp);
    if (adminComment && adminComment.trim()) {
      sheet.getRange(rowNum, commentColIdx).setValue(adminComment.trim());
    }

    _log("INFO", "updateRequestStatus", "Status updated", {
      uuid: eventUuid, row: rowNum, from: oldStatus, to: newStatus,
      comment: adminComment || ""
    });

    return { success: true, newStatus: newStatus };

  } catch(e) {
    _log("ERROR", "updateRequestStatus", e.message, { uuid: eventUuid });
    return { success: false, message: e.message };
  }
}

// ── FORM SUBMISSION (rate-limited, logged) ────────────
function processVLForm(data) {
  var actor = getSessionEmail();

  // Rate limit: max 3 submissions per 5 minutes per user
  if (!_checkRateLimit("submit_" + actor, 3, 300)) {
    _log("WARN", "processVLForm", "Submission rate limit hit", { actor: actor });
    return { status: "error", message: "Too many submissions. Please wait a few minutes before trying again." };
  }

  try {
    var config = getFormConfig();
    if (!config.isOpen) return { status: "closed", message: config.closedMsg };

    var sheet          = getVLSheet();
    var submittedMonth = normalizeMonthFull(data.month || "");
    var normalizedActive = config.activeMonths.map(function(m){ return normalizeMonthFull(m); });

    // Validation
    if (!data.ldap || !data.ldap.trim()) return { status:"error", message:"LDAP is required." };
    if (!data.channel)  return { status:"error", message:"Channel is required." };
    if (!data.vlDate)   return { status:"error", message:"VL Date is required." };
    if (data.channel === "Phone") {
      var day = new Date(data.vlDate + "T00:00:00").getDay();
      if (day === 0 || day === 6) return { status:"error", message:"Phone agents cannot file leave on weekends." };
    }
    if (normalizedActive.indexOf(submittedMonth) < 0) {
      return { status:"error", message:"Selected month (" + submittedMonth + ") is no longer active." };
    }
    if (data.reasonForVL === "Birthday" && !data.bdayDate) {
      return { status:"error", message:"Birthday Leave requires a Date of Birthday." };
    }
    // ── Validate VL date actually falls within the declared active month ──
    var vlDateParsed = new Date(data.vlDate + "T00:00:00");
    var vlMonthFull  = MO_NAMES_GS[vlDateParsed.getMonth()] + " " + vlDateParsed.getFullYear();
    if (vlMonthFull !== submittedMonth) {
      _log("WARN", "processVLForm", "Date/month mismatch", { vlDate: data.vlDate, declaredMonth: submittedMonth });
      return { status:"error", message:"Your VL Date (" + data.vlDate + ") does not fall within the declared month (" + submittedMonth + ")." };
    }

    if (checkVLDuplicate(sheet, data.ldap.trim(), data.channel.trim(), data.vlDate)) {
      _log("WARN", "processVLForm", "Duplicate blocked", { ldap: data.ldap, date: data.vlDate });
      return { status:"error", message:"Duplicate Request Blocked." };
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx;
    });

    var newRow = new Array(headers.length).fill("");
    function setVal(colName, val) {
      var idx = colMap[colName.replace(/[^a-zA-Z0-9]/g,'').toLowerCase()];
      if (idx !== undefined) newRow[idx] = val;
    }

    setVal("Timestamp",              new Date());
    setVal("Email Address",          actor);
    setVal("LDAP",                   data.ldap.trim());
    setVal("Channel",                data.channel);
    setVal("VL Date",                data.vlDate);
    setVal("Team Lead",              data.teamLead);
    setVal("Reason for VL",          data.reasonForVL);
    setVal("Work Group",             data.workGroup);
    setVal("Site",                   data.site);
    setVal("Accruals Snip-it",       data.accrualUrl.trim());
    setVal("Accruals",               Math.round(parseFloat(data.accrualNum)));
    setVal("Month",                  "'" + submittedMonth);
    setVal("Date of Birthday",       data.bdayDate || "N/A");
    setVal("Proof / Artifacts",      data.proofUrl || "N/A");
    setVal("Status",                 "");
    setVal("Confirmation on Status", "");
    setVal("Email Sent",             "");
    setVal("Row UUID",               Utilities.getUuid()); // ← stable ID on insert

    // Find true last row to avoid blank row gaps
    var colA = sheet.getRange("A:A").getValues();
    var trueLastRow = 0;
    for (var r = colA.length - 1; r >= 0; r--) {
      if (colA[r][0] !== "") { trueLastRow = r + 1; break; }
    }
    sheet.getRange(trueLastRow + 1, 1, 1, newRow.length).setValues([newRow]);

    _log("INFO", "processVLForm", "Submission accepted", {
      ldap: data.ldap, date: data.vlDate, channel: data.channel
    });
    return { status: "success" };

  } catch(e) {
    _log("ERROR", "processVLForm", e.message, { ldap: data.ldap || "unknown" });
    return { status:"error", message:"Server error: " + e.message };
  }
}

// ── DASHBOARD PAYLOAD (rate-limited) ─────────────────
function getDashboardPayload() {
  var actor = getSessionEmail();

  // Rate limit: max 20 full reloads per minute per user
  if (!_checkRateLimit("dashboard_" + actor, 20, 60)) {
    _log("WARN", "getDashboardPayload", "Dashboard rate limit hit", { actor: actor });
    return { error: "Too many requests. Please wait before reloading." };
  }

  var payload = {
    email: '', photoUrl: null,
    config: { isOpen: true, activeMonths: ["May 2026", "June 2026"], closedMsg: "",
              isSupervisor: false, isAdmin: false, teamName: "" },
    events: []
  };

  try { payload.email = actor; } catch(e) {}
  try { var cfg = getFormConfig(); if (cfg) payload.config = cfg; } catch(e) {}
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (sheet && sheet.getLastRow() >= 2) {
      var evts = getCalendarData();
      payload.events = Array.isArray(evts) ? evts : [];
    }
  } catch(e) {
    _log("ERROR", "getDashboardPayload", e.message);
  }

  return payload;
}

// ── DUPLICATE CHECK ───────────────────────────────────
function checkVLDuplicate(sheet, ldap, channel, vlDate) {
  try {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return false;

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx;
    });
    if (colMap["ldap"] === undefined || colMap["channel"] === undefined || colMap["vldate"] === undefined) return false;

    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var tz   = Session.getScriptTimeZone();

    for (var i = 0; i < data.length; i++) {
      var sheetLdap    = (data[i][colMap["ldap"]]||"").toString().trim().toLowerCase();
      var sheetChannel = (data[i][colMap["channel"]]||"").toString().trim().toLowerCase();
      var rawDate      = data[i][colMap["vldate"]];
      if (!sheetLdap || !rawDate) continue;
      var sheetDateStr = "";
      if (Object.prototype.toString.call(rawDate) === '[object Date]' && !isNaN(rawDate.getTime())) {
        sheetDateStr = Utilities.formatDate(rawDate, tz, "yyyy-MM-dd");
      } else {
        var p = new Date(rawDate);
        sheetDateStr = !isNaN(p.getTime()) ? Utilities.formatDate(p, tz, "yyyy-MM-dd") : rawDate.toString().trim();
      }
      if (sheetLdap === ldap.toLowerCase() && sheetChannel === channel.toLowerCase() && sheetDateStr === vlDate) return true;
    }
    return false;
  } catch(e) { return false; }
}

// ── MONTH HELPERS ─────────────────────────────────────
var MO_NAMES_GS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
var MO_SHORT_GS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function normalizeMonthFull(raw) {
  raw = (raw || "").toString().trim();
  var d = new Date(raw);
  if (!isNaN(d.getTime()) && raw.length > 8) return MO_NAMES_GS[d.getMonth()] + " " + d.getFullYear();
  var parts = raw.split(" ");
  if (parts.length >= 2) {
    if (MO_NAMES_GS.indexOf(parts[0]) >= 0) return parts[0] + " " + parts[1];
    var mi = MO_SHORT_GS.indexOf(parts[0]);
    if (mi >= 0) return MO_NAMES_GS[mi] + " " + parts[1];
  }
  return raw;
}

// ── FEEDBACK ─────────────────────────────────────────
function submitFeedback(feedbackText) {
  var actor = getSessionEmail() || "Unknown";
  if (!_checkRateLimit("feedback_" + actor, 5, 300)) {
    return { status: 'error', message: 'Too many feedback submissions. Please wait.' };
  }
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Feedback");
    if (!sheet) {
      sheet = ss.insertSheet("Feedback");
      sheet.getRange("A1:C1").setValues([["Timestamp","Email","Feedback"]])
        .setBackground("#f9ab00").setFontColor("#3e2723").setFontWeight("bold");
    }
    sheet.appendRow([new Date(), actor, feedbackText]);
    _log("INFO", "submitFeedback", "Feedback received", { actor: actor });
    return { status: 'success' };
  } catch(e) {
    _log("ERROR", "submitFeedback", e.message);
    return { status: 'error', message: e.message };
  }
}

// ── EMAIL FUNCTIONS (unchanged — kept for compatibility) ─
function getUnsentSummary() {
  try {
    var sheet = getVLSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { count: 0, html: "No data available." };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colMap = {};
    headers.forEach(function(h, idx) { colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx; });
    var statIdx = colMap["status"];
    var confIdx = colMap["confirmationonstatus"];
    var sentIdx = colMap["emailsent"];
    if (statIdx === undefined || confIdx === undefined) return { count: 0, html: "Required columns missing." };
    var data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var appr = 0, den = 0, noal = 0;
    for (var i = 0; i < data.length; i++) {
      var status = (data[i][statIdx] || "").toString().trim();
      var conf   = (data[i][confIdx] || "").toString().trim();
      var sent   = sentIdx !== undefined ? (data[i][sentIdx] || "").toString().trim() : "";
      var lStat  = status.toLowerCase();
      if (status && !lStat.includes("pending") && conf && !sent) {
        if (lStat.includes("approved") || lStat.includes("birthday leave")) appr++;
        else if (lStat.includes("denied")) den++;
        else if (lStat.includes("no alloc")) noal++;
      }
    }
    var total = appr + den + noal;
    var html = total === 0
      ? "<div style='text-align:center;padding:20px;color:var(--text-muted);'>No confirmed, unsent requests.</div>"
      : "<div style='display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;text-align:center;margin-bottom:16px;'>" +
        "<div style='background:var(--green-bg);border:1px solid var(--green-border);border-radius:8px;padding:12px;'><div style='font-size:20px;font-weight:700;color:var(--green-main);'>" + appr + "</div><div style='font-size:10px;font-weight:700;color:var(--green-main);text-transform:uppercase;'>Approved</div></div>" +
        "<div style='background:var(--red-bg);border:1px solid var(--red-border);border-radius:8px;padding:12px;'><div style='font-size:20px;font-weight:700;color:var(--red-main);'>" + den + "</div><div style='font-size:10px;font-weight:700;color:var(--red-main);text-transform:uppercase;'>Denied</div></div>" +
        "<div style='background:var(--pink-bg);border:1px solid var(--pink-border);border-radius:8px;padding:12px;'><div style='font-size:20px;font-weight:700;color:var(--pink-main);'>" + noal + "</div><div style='font-size:10px;font-weight:700;color:var(--pink-main);text-transform:uppercase;'>No Alloc</div></div>" +
        "</div>";
    return { count: total, html: html };
  } catch(e) { return { count: 0, html: "Error." }; }
}

function sendNazunaNotifications() {
  var actor = getSessionEmail();
  if (!_checkRateLimit("sendNotif_" + actor, 3, 300)) {
    return { success: false, message: "Rate limit: wait before sending again." };
  }
  try {
    var sheet = getVLSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "No data." };
    var statusColIdx = ensureColumn(sheet, "Status");
    var confColIdx   = ensureColumn(sheet, "Confirmation on Status");
    var sentColIdx   = ensureColumn(sheet, "Email Sent");
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colMap = {};
    headers.forEach(function(h, idx) { colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx; });
    var data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var displayData = sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
    var tz = Session.getScriptTimeZone();
    var sentCount = 0;
    for (var i = 0; i < data.length; i++) {
      var row    = data[i];
      var status = (row[statusColIdx] || "").toString().trim();
      var conf   = (row[confColIdx]   || "").toString().trim();
      var sent   = (row[sentColIdx]   || "").toString().trim();
      var lStat  = status.toLowerCase();
      if (!status || lStat.includes("pending") || !conf || sent) continue;
      var getVal  = (function(r, cm) { return function(n) { var x = cm[n.replace(/[^a-zA-Z0-9]/g,'').toLowerCase()]; return x !== undefined ? r[x] : ""; }; })(row, colMap);
      var getDisp = (function(dr, cm) { return function(n) { var x = cm[n.replace(/[^a-zA-Z0-9]/g,'').toLowerCase()]; return x !== undefined ? dr[x] : ""; }; })(displayData[i], colMap);
      var emailAddress = getVal("Email Address");
      if (!emailAddress) continue;
      var ldap = getVal("LDAP");
      var vlDateRaw = getVal("VL Date");
      var formattedDate = Object.prototype.toString.call(vlDateRaw) === '[object Date]'
        ? Utilities.formatDate(vlDateRaw, tz, "MMMM dd, yyyy") : vlDateRaw.toString();
      var emoji = "⚠️";
      if (lStat.includes("birthday")) emoji = "🎂";
      else if (lStat.includes("approved")) emoji = "✅";
      else if (lStat.includes("denied"))   emoji = "⛔";
      else if (lStat.includes("no alloc")) emoji = "🚫";
      var htmlBody = createEmailTemplate(ldap, formattedDate, status, "-",
        getVal("Team Lead"), getVal("Reason for VL"), getVal("Work Group"),
        getVal("Comments"), getVal("Timestamp"), emoji,
        getDisp("Accruals"), getDisp("Attendance"), getVal("Site"),
        getVal("Confirmation on Status"));
      try {
        MailApp.sendEmail({ to: emailAddress, subject: emoji + " Google Play VL Calendar Update - " + status, htmlBody: htmlBody, name: "Google Play VL Calendar" });
        var stamp = Utilities.formatDate(new Date(), tz, "MM/dd/yy HH:mm") + " by " + actor.split('@')[0];
        sheet.getRange(i + 2, sentColIdx + 1).setValue(stamp);
        sentCount++;
        _log("INFO", "sendNazunaNotifications", "Email sent", { ldap: ldap, status: status });
      } catch(mailErr) {
        _log("ERROR", "sendNazunaNotifications", "Mail failed for " + ldap, { err: mailErr.message });
      }
    }
    return { success: true, count: sentCount };
  } catch(e) {
    _log("ERROR", "sendNazunaNotifications", e.message);
    return { success: false, message: e.message };
  }
}

function createEmailTemplate(ldap, date, status, queue, team, reason, workgroup, comments, timestamp, emoji, accruals, attendance, site, confirmation) {
  var statusColor = "#5f6368";
  var lowerStatus = status.toString().toLowerCase();
  if      (lowerStatus.includes("birthday")) statusColor = "#1a73e8";
  else if (lowerStatus.includes("approved")) statusColor = "#188038";
  else if (lowerStatus.includes("denied"))   statusColor = "#d93025";
  else if (lowerStatus.includes("no alloc")) statusColor = "#ea4335";
  return '<div style="font-family:\'Google Sans\',Roboto,Arial,sans-serif;max-width:600px;border:1px solid #dadce0;border-radius:8px;overflow:hidden;margin:0 auto;background:#ffffff;">' +
    '<table style="width:100%;border-collapse:collapse;height:6px;"><tr>' +
    '<td style="background:#4285F4;width:25%;height:6px;"></td><td style="background:#EA4335;width:25%;height:6px;"></td>' +
    '<td style="background:#FBBC05;width:25%;height:6px;"></td><td style="background:#34A853;width:25%;height:6px;"></td>' +
    '</tr></table>' +
    '<div style="padding:30px 20px 10px;text-align:center;"><h2 style="color:' + statusColor + ';margin:0;font-size:32px;font-weight:400;">' + emoji + ' ' + status + '</h2></div>' +
    '<div style="padding:20px 30px;">' +
    '<p style="color:#202124;font-size:16px;margin-bottom:20px;">Hi <strong>' + ldap + '</strong>,</p>' +
    '<p style="color:#5f6368;font-size:14px;">Here is the latest update regarding your leave request.</p>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:25px;font-size:14px;">' +
    '<tr style="border-bottom:1px solid #f1f3f4;"><td style="padding:14px 0;color:#5f6368;">VL Date</td><td style="padding:14px 0;text-align:right;color:#202124;font-weight:500;font-size:16px;">' + date + '</td></tr>' +
    '<tr style="border-bottom:1px solid #f1f3f4;"><td style="padding:14px 0;color:#5f6368;">Status</td><td style="padding:14px 0;text-align:right;color:' + statusColor + ';font-weight:bold;">' + status + '</td></tr>' +
    '<tr style="border-bottom:1px solid #f1f3f4;"><td style="padding:14px 0;color:#5f6368;">Site</td><td style="padding:14px 0;text-align:right;color:#202124;">' + (site||"-") + '</td></tr>' +
    '<tr style="border-bottom:1px solid #f1f3f4;"><td style="padding:14px 0;color:#5f6368;">Attendance</td><td style="padding:14px 0;text-align:right;color:#188038;font-weight:600;">' + (attendance||"-") + '</td></tr>' +
    '<tr style="border-bottom:1px solid #f1f3f4;"><td style="padding:14px 0;color:#5f6368;">Accruals</td><td style="padding:14px 0;text-align:right;color:#202124;">' + (accruals||"-") + '</td></tr>' +
    '<tr style="border-bottom:1px solid #f1f3f4;"><td style="padding:14px 0;color:#5f6368;">Workgroup</td><td style="padding:14px 0;text-align:right;color:#1a73e8;font-weight:600;">' + workgroup + '</td></tr>' +
    '<tr style="border-bottom:1px solid #f1f3f4;"><td style="padding:14px 0;color:#5f6368;">Comments</td><td style="padding:14px 0;text-align:right;color:#202124;">' + (comments||"-") + '</td></tr>' +
    '<tr style="border-bottom:1px solid #f1f3f4;"><td style="padding:14px 0;color:#5f6368;">Confirmation</td><td style="padding:14px 0;text-align:right;color:#202124;font-size:11px;">' + (confirmation||"-") + '</td></tr>' +
    '</table>' +
    '<div style="background:#f8f9fa;padding:15px;border-radius:8px;font-size:12px;color:#5f6368;border:1px solid #f1f3f4;margin-top:16px;">' +
    '<strong>Request Details:</strong><br>Submitted: ' + timestamp + '<br>Reason: ' + reason + '<br>Team: ' + team + '</div>' +
    '<hr style="border:0;height:1px;background:#dadce0;margin:25px 0;">' +
    '<p style="font-size:11px;color:#d93025;text-align:center;font-weight:500;">Automated message from Google Play VL Calendar.</p>' +
    '</div></div>';
}

function getUserAvatarUrl() { return null; }

// ── LIGHTWEIGHT EVENTS-ONLY ENDPOINT for live sync polling ──
// Returns only the events array — no config, no supervisor lookup.
// Much cheaper on quota than a full getDashboardPayload call.
function getEventsOnly() {
  var actor = getSessionEmail();

  // Separate, more generous rate limit for polling: 60 calls/minute
  if (!_checkRateLimit("eventsOnly_" + actor, 60, 60)) {
    return { events: null }; // silently skip, don't error
  }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return { events: [] };
    var evts = getCalendarData();
    return { events: Array.isArray(evts) ? evts : [] };
  } catch(e) {
    _log("ERROR", "getEventsOnly", e.message);
    return { events: null };
  }
}

// ── DELETE / ARCHIVE REQUEST ──────────────────────────
function deleteVLRequest(eventUuid) {
  var actor = getSessionEmail();
  if (!actor) return { success: false, message: "Not authenticated." };

  if (!_checkRateLimit("delete_" + actor, 10, 60)) {
    return { success: false, message: "Rate limit exceeded. Please wait before trying again." };
  }

  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var sheet     = getVLSheet();
    var rowIdx    = _getRowIndex(sheet);
    var rowNum    = rowIdx.index[eventUuid];

    if (!rowNum) {
      return { success: false, message: "Request not found. Please refresh the calendar." };
    }

    var lastCol  = sheet.getLastColumn();
    var headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap   = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx;
    });

    var rowData  = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var rowEmail = (rowData[colMap["emailaddress"]] || "").toString().trim().toLowerCase();
    var status   = (rowData[colMap["status"]] || "").toString().trim().toLowerCase();

    // Only the owner can delete
    if (rowEmail !== actor.toLowerCase()) {
      _log("WARN", "deleteVLRequest", "Unauthorized delete attempt", { actor: actor, owner: rowEmail, uuid: eventUuid });
      return { success: false, message: "You can only remove your own requests." };
    }

    // Admins can delete anything; regular users cannot delete approved requests
    var config = getFormConfig();
    var isActorAdmin = config.isAdmin;
    if (!isActorAdmin && (status === "approved" || status.includes("birthday leave"))) {
      return { success: false, message: "Approved requests cannot be removed. Please contact your supervisor." };
    }

    // ── Get or create the Deleted Requests archive sheet ──
    var archiveName  = "Deleted Requests";
    var archiveSheet = ss.getSheetByName(archiveName);
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet(archiveName);
      var archiveHeaders = headers.concat(["Deleted By", "Deleted At"]);
      archiveSheet.getRange(1, 1, 1, archiveHeaders.length).setValues([archiveHeaders])
        .setBackground("#ea4335").setFontColor("#ffffff").setFontWeight("bold");
      archiveSheet.setFrozenRows(1);
      _log("INFO", "deleteVLRequest", "Created Deleted Requests archive sheet");
    }

    // ── Copy row to archive with metadata ──
    var tz         = Session.getScriptTimeZone();
    var deletedAt  = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    var archiveRow = rowData.concat([actor, deletedAt]);
    archiveSheet.appendRow(archiveRow);

    // ── Delete row from main sheet ──
    sheet.deleteRow(rowNum);

    _log("INFO", "deleteVLRequest", "Request archived and deleted", { uuid: eventUuid, actor: actor, row: rowNum });
    return { success: true };

  } catch(e) {
    _log("ERROR", "deleteVLRequest", e.message, { uuid: eventUuid });
    return { success: false, message: "Server error: " + e.message };
  }
}
