/**
 * APOLLO IT PORTAL - Phase 8 FINAL
 */

// ═══════════════════════════════════════════════════════════════
// CONFIG
// ═══════════════════════════════════════════════════════════════
var props = PropertiesService.getScriptProperties();
var WEB_APP_URL = props.getProperty('WEB_APP_URL') || "https://script.google.com/macros/s/AKfycbyVqi9UJ7L5fekICPr_8GdDpsqIM1lTKRDY9butj3HANo-GPhcvyMvS5QMA7uoGy0drcA/exec";

var LOGO_FILE_ID       = "1pUrTtGb_J9lq_wsShppzu5Y_8k4GNg5-";
var TICKET_DB_SHEET_ID = "12q4j3qep_35B8iR9BHGEIRbLtAEoP2ytskL9KQmIazA";
var TICKET_SHEET_GID   = "2091678180";
var DB_SHEET_ID        = "1crAnFCr5gMgyQoX82JAmmurKhn2KMxm55kYEHrT42PY";
var DB_SHEET_GID       = "648987686";

// ═══════════════════════════════════════════════════════════════
// BASIC HELPERS
// ═══════════════════════════════════════════════════════════════
function cleanEmailAddress(emailStr) {
  if (!emailStr) return "";
  var str = emailStr.toString().trim().toLowerCase();
  var matches = str.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/gi);
  if (matches && matches.length > 0) return matches[0];
  return "";
}

function formatSafeDate(dateVal) {
  if (!dateVal) return "-";
  try {
    if (dateVal instanceof Date) return Utilities.formatDate(dateVal, "GMT+8", "dd/MM/yyyy HH:mm:ss");
    var s = dateVal.toString().trim();
    if (s === "" || s === "-" || s === "Pending") return s || "-";
    return s;
  } catch (e) { return "-"; }
}

function sendSafeEmail(emailConfig) {
  if (!emailConfig) emailConfig = {};
  if (emailConfig.to) emailConfig.to = cleanEmailAddress(emailConfig.to);
  var sysConfig = getSystemEmailConfig();
  if (!emailConfig.to || emailConfig.to.toString().trim() === "" || emailConfig.to === "null") {
    emailConfig.to = sysConfig.itGroupEmail || sysConfig.itCheckEmail || sysConfig.itVerifiedEmail;
  }
  if (emailConfig.to) {
    try { MailApp.sendEmail(emailConfig); } catch (e) { Logger.log("sendSafeEmail error: " + e.toString()); }
  } else {
    Logger.log("sendSafeEmail: No valid recipient.");
  }
}

function getScriptExecUrl() {
  var url = "";
  try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  if (!url || typeof url !== "string" || url.trim() === "" || url.indexOf("http") !== 0) {
    url = WEB_APP_URL;
  }
  return url;
}

function extractDriveFileId(url) {
  if (!url) return "";
  var m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  var m2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) return m2[1];
  return "";
}

// ═══════════════════════════════════════════════════════════════
// HELPER: Clean Stage Name
// ═══════════════════════════════════════════════════════════════
function cleanStageName(stage) {
  if (!stage) return "Unknown";
  var s = stage.toString().toUpperCase();
  if (s.indexOf("MANAGER") !== -1) return "Manager";
  if (s.indexOf("HOD") !== -1 && s.indexOf("IT") === -1) return "HOD";
  if (s.indexOf("C-SUITE") !== -1 || s.indexOf("CSUITE") !== -1) return "C-Suite";
  if (s.indexOf("MD") !== -1 || s.indexOf("MANAGING DIRECTOR") !== -1) return "MD";
  if (s.indexOf("IT CHECK") !== -1) return "IT Check";
  if (s.indexOf("IT VERIF") !== -1) return "IT Verification";
  return "Approval Process";
}

function getRejectStageTagFromStatus(status) {
  if (!status) return "";
  var s = status.toString().toUpperCase();
  if (s.indexOf("REJECT") === -1) return "";
  if (s.indexOf("IT CHECK") !== -1) return "IT Check";
  if (s.indexOf("IT VERIF") !== -1) return "IT Verification";
  if (s.indexOf("C-SUITE") !== -1 || s.indexOf("CSUITE") !== -1) return "C-Suite";
  if (s.indexOf("MANAGING DIRECTOR") !== -1 || s.indexOf(" AT MD") !== -1) return "MD";
  if (s.indexOf("HOD") !== -1) return "HOD";
  if (s.indexOf("MANAGER") !== -1) return "Manager";
  return "";
}

// ═══════════════════════════════════════════════════════════════
// CACHE + SPREADSHEET
// ═══════════════════════════════════════════════════════════════
var _memCache = {
  dbSpreadsheet: null, targetSheet: null, initialized: false,
  emailToName: null, nameToEmail: null, routeMap: null, routeRows: null,
  dashboardAccessList: null, systemEmails: {}, approverProfiles: {}, hierarchyCache: {},
  quotationSheet: null, assessmentSheet: null, discussionSheet: null, signatureSheet: null
};

function getDbSpreadsheet() {
  if (_memCache.dbSpreadsheet) return _memCache.dbSpreadsheet;
  var ss = null;
  if (DB_SHEET_ID && DB_SHEET_ID.trim() !== "") {
    try { ss = SpreadsheetApp.openById(DB_SHEET_ID); } catch (e) {}
  }
  if (!ss) { try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {} }
  _memCache.dbSpreadsheet = ss;
  return _memCache.dbSpreadsheet;
}

function getTicketSpreadsheet() {
  var ss = null;
  if (TICKET_DB_SHEET_ID && TICKET_DB_SHEET_ID.trim() !== "") {
    try { ss = SpreadsheetApp.openById(TICKET_DB_SHEET_ID); } catch (e) {}
  }
  if (!ss) { try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {} }
  if (!ss && DB_SHEET_ID && DB_SHEET_ID.trim() !== "") {
    try { ss = SpreadsheetApp.openById(DB_SHEET_ID); } catch (e) {}
  }
  return ss;
}

function getTargetSheet() {
  if (_memCache.targetSheet) return _memCache.targetSheet;
  var ss = getTicketSpreadsheet();
  if (!ss) return null;
  if (TICKET_SHEET_GID) {
    var allSheets = ss.getSheets();
    for (var i = 0; i < allSheets.length; i++) {
      if (allSheets[i].getSheetId().toString() === TICKET_SHEET_GID.toString()) {
        _memCache.targetSheet = allSheets[i];
        return _memCache.targetSheet;
      }
    }
  }
  _memCache.targetSheet = ss.getSheetByName("Form Responses 1") ||
                          ss.getSheetByName("IT Asset Request Form(Responses)") ||
                          ss.getSheetByName("IT Asset Request Form (Responses)") ||
                          ss.getSheetByName("Form_Responses") ||
                          getSheetFlexible(ss, "Responses") ||
                          getSheetFlexible(ss, "Form") ||
                          ss.getSheets()[0];
  return _memCache.targetSheet;
}

function getSheetFlexible(db, targetName) {
  if (!db || !targetName) return null;
  var cleanTarget = targetName.toString().replace(/[\s+=_\-\(\)]/g, "").toLowerCase();
  var sheets = db.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().toString().replace(/[\s+=_\-\(\)]/g, "").toLowerCase();
    if (sName.indexOf(cleanTarget) !== -1 || cleanTarget.indexOf(sName) !== -1) return sheets[i];
  }
  return null;
}

// ═══════════════════════════════════════════════════════════════
// DIRECTORY CACHE
// ═══════════════════════════════════════════════════════════════
function initDirectoryCache() {
  if (_memCache.initialized) return;
  _memCache.emailToName = {};
  _memCache.nameToEmail = {};
  _memCache.routeMap = {};
  _memCache.routeRows = [];
  _memCache.dashboardAccessList = [];
  _memCache.systemEmails = {};
  _memCache.hierarchyCache = {};

  try {
    var db = getDbSpreadsheet();
    if (db) {
      var accessSheet = getSheetFlexible(db, "DashboardAcess") || getSheetFlexible(db, "DashboardAccess") || getSheetFlexible(db, "Dashboard Access");
      if (accessSheet) {
        var aData = accessSheet.getDataRange().getValues();
        if (aData && aData.length > 1) {
          var aHeaders = aData[0] || [];
          var nameCol = 0, roleCol = 1, emailCol = 2;
          for (var c = 0; c < aHeaders.length; c++) {
            var colH = (aHeaders[c] || "").toString().toLowerCase().trim();
            if (colH.indexOf("name") !== -1 || colH.indexOf("nama") !== -1) nameCol = c;
            else if (colH.indexOf("role") !== -1 || colH.indexOf("jawatan") !== -1) roleCol = c;
            else if (colH.indexOf("email") !== -1 || colH.indexOf("emel") !== -1) emailCol = c;
          }
          for (var a = 1; a < aData.length; a++) {
            var aName = (aData[a][nameCol] || "").toString().replace(/\s+/g, " ").trim();
            var aRole = (aData[a][roleCol] || "").toString().trim().toUpperCase();
            var aEmail = cleanEmailAddress(aData[a][emailCol]);
            if (aEmail) {
              _memCache.dashboardAccessList.push({ name: aName, role: aRole, email: aEmail });
              if (aName) {
                _memCache.emailToName[aEmail] = aName.toUpperCase();
                _memCache.nameToEmail[aName.toLowerCase()] = aEmail;
              }
              if (aRole === "IT_CHECK") { _memCache.systemEmails.IT_CHECK_EMAIL = aEmail; _memCache.systemEmails.IT_CHECK_NAME = aName; }
              else if (aRole === "IT_HOD") { _memCache.systemEmails.IT_HOD_EMAIL = aEmail; _memCache.systemEmails.IT_HOD_NAME = aName; }
              else if (aRole === "MD") { _memCache.systemEmails.MD_EMAIL = aEmail; _memCache.systemEmails.MD_NAME = aName; }
              else if (aRole === "IT_VERIFIED") { _memCache.systemEmails.IT_VERIFIED_EMAIL = aEmail; _memCache.systemEmails.IT_VERIFIED_NAME = aName; }
              else if (aRole === "IT_GROUP_EMAIL") { _memCache.systemEmails.IT_GROUP_EMAIL = aEmail; }
            }
          }
        }
      }

      var enSheet = getSheetFlexible(db, "Email=Name") || getSheetFlexible(db, "Email+Name") || getSheetFlexible(db, "EmailName") || getSheetFlexible(db, "Name=Email");
      if (enSheet) {
        var enData = enSheet.getDataRange().getValues();
        if (enData && enData.length > 1) {
          var enHeaders = enData[0] || [];
          var enNameIdx = 0, enEmailIdx = 1;
          for (var h = 0; h < enHeaders.length; h++) {
            var colTitle = (enHeaders[h] || "").toString().toLowerCase().trim();
            if (colTitle.indexOf("email") !== -1 || colTitle.indexOf("emel") !== -1) enEmailIdx = h;
            else if (colTitle.indexOf("name") !== -1 || colTitle.indexOf("nama") !== -1) enNameIdx = h;
          }
          for (var e = 1; e < enData.length; e++) {
            var val0 = (enData[e][enNameIdx] || "").toString().replace(/\s+/g, " ").trim();
            var val1 = cleanEmailAddress(enData[e][enEmailIdx]);
            if (val0.indexOf("@") !== -1 && !val1) {
              val1 = cleanEmailAddress(val0);
              val0 = (enData[e][enEmailIdx] || "").toString().replace(/\s+/g, " ").trim();
            }
            if (val1 && val0) {
              if (!_memCache.emailToName[val1]) _memCache.emailToName[val1] = val0.toUpperCase();
              if (!_memCache.nameToEmail[val0.toLowerCase()]) _memCache.nameToEmail[val0.toLowerCase()] = val1;
            }
          }
        }
      }

      var routeSheet = getSheetFlexible(db, "Route") || db.getSheets()[0];
      if (routeSheet) {
        var rData = routeSheet.getDataRange().getValues();
        if (rData && rData.length > 1) {
          var rHeaders = rData[0] || [];
          var emailIdx = 0, mgrIdx = 3, hodIdx = 4, csIdx = 5;
          for (var h2 = 0; h2 < rHeaders.length; h2++) {
            var hText = (rHeaders[h2] || "").toString().toLowerCase().trim();
            if (hText === "email" || hText.indexOf("email") !== -1 || hText.indexOf("emel") !== -1) emailIdx = h2;
            else if (hText.indexOf("manager") !== -1 || hText.indexOf("mgr") !== -1) mgrIdx = h2;
            else if (hText.indexOf("hod") !== -1 || hText.indexOf("head") !== -1) hodIdx = h2;
            else if (hText.indexOf("csuite") !== -1 || hText.indexOf("c-suite") !== -1) csIdx = h2;
          }
          for (var r = 1; r < rData.length; r++) {
            var rowEmail = cleanEmailAddress(rData[r][emailIdx]);
            var mgrNameRaw = (rData[r][mgrIdx] || "").toString().replace(/\s+/g, " ").trim();
            var hodNameRaw = (rData[r][hodIdx] || "").toString().replace(/\s+/g, " ").trim();
            var csNameRaw  = (rData[r][csIdx]  || "").toString().replace(/\s+/g, " ").trim();
            var routeObj = {
              email: rowEmail,
              mgrRaw: (mgrNameRaw && mgrNameRaw !== "-") ? mgrNameRaw : null,
              hodRaw: (hodNameRaw && hodNameRaw !== "-") ? hodNameRaw : null,
              csRaw:  (csNameRaw && csNameRaw !== "-")  ? csNameRaw  : null
            };
            _memCache.routeRows.push(routeObj);
            if (rowEmail) _memCache.routeMap[rowEmail] = routeObj;
          }
        }
      }
    }
  } catch (err) { Logger.log("initDirectoryCache error: " + err.toString()); }
  _memCache.initialized = true;
}

function getSystemEmailConfig() {
  initDirectoryCache();
  var sys = _memCache.systemEmails || {};
  return {
    itGroupEmail:    sys.IT_GROUP_EMAIL || sys.IT_CHECK_EMAIL || "",
    itCheckEmail:    sys.IT_CHECK_EMAIL || "",
    itCheckName:     sys.IT_CHECK_NAME  || "IT CHECK OFFICER",
    itHodEmail:      sys.IT_HOD_EMAIL   || "",
    itHodName:       sys.IT_HOD_NAME    || "IT HOD",
    mdEmail:         sys.MD_EMAIL       || "",
    mdName:          sys.MD_NAME        || "MANAGING DIRECTOR",
    itVerifiedEmail: sys.IT_VERIFIED_EMAIL || "",
    itVerifiedName:  sys.IT_VERIFIED_NAME  || "IT VERIFIER"
  };
}

function extractUserEmailFromRow(rowValues) {
  if (!rowValues || !rowValues.length) return "";
  for (var i = 0; i < rowValues.length; i++) {
    var rawStr = (rowValues[i] || "").toString().replace(/\s+/g, "").toLowerCase();
    var cleaned = cleanEmailAddress(rawStr);
    if (cleaned) return cleaned;
  }
  return "";
}

function extractUserEmail(rowData) {
  for (var k in rowData) {
    var rawStr = (rowData[k] || "").toString().replace(/\s+/g, "").toLowerCase();
    var cleaned = cleanEmailAddress(rawStr);
    if (cleaned) return cleaned;
  }
  return "";
}

function getNameByEmail(email) {
  if (!email) return "-";
  var cleanEmail = cleanEmailAddress(email);
  var sys = getSystemEmailConfig();
  if (sys.mdEmail && cleanEmail === cleanEmailAddress(sys.mdEmail)) return sys.mdName;
  if (sys.itCheckEmail && cleanEmail === cleanEmailAddress(sys.itCheckEmail)) return sys.itCheckName;
  if (sys.itHodEmail && cleanEmail === cleanEmailAddress(sys.itHodEmail)) return sys.itHodName;
  if (sys.itVerifiedEmail && cleanEmail === cleanEmailAddress(sys.itVerifiedEmail)) return sys.itVerifiedName;
  initDirectoryCache();
  if (_memCache.emailToName && _memCache.emailToName[cleanEmail]) return _memCache.emailToName[cleanEmail];
  return cleanEmail;
}

function getEmailByName(name) {
  if (!name || name === "-" || name.toUpperCase() === "N/A") return null;
  var cleanName = name.toString().replace(/\s+/g, " ").trim().toLowerCase();
  if (cleanName.indexOf("@") !== -1) {
    var asEmail = cleanEmailAddress(cleanName);
    if (asEmail && asEmail.indexOf("@") !== -1) return asEmail;
  }
  initDirectoryCache();
  if (_memCache.nameToEmail && _memCache.nameToEmail[cleanName]) return _memCache.nameToEmail[cleanName];
  if (_memCache.nameToEmail) {
    for (var k in _memCache.nameToEmail) {
      if (k && (k === cleanName || k.indexOf(cleanName) !== -1 || cleanName.indexOf(k) !== -1)) {
        return _memCache.nameToEmail[k];
      }
    }
  }
  return null;
}

function getApprovalHierarchy(userEmail) {
  var email = cleanEmailAddress(userEmail);
  var sys = getSystemEmailConfig();
  initDirectoryCache();

  if (_memCache.hierarchyCache && _memCache.hierarchyCache[email]) return _memCache.hierarchyCache[email];

  var result = {
    managerEmail: null, managerName: null,
    hodEmail: null, hodName: null,
    csuiteEmail: null, csuiteName: null,
    mdEmail: sys.mdEmail, mdName: sys.mdName
  };

  var routeEntry = _memCache.routeMap ? _memCache.routeMap[email] : null;
  if (routeEntry) {
    if (routeEntry.mgrRaw) { result.managerName = routeEntry.mgrRaw; result.managerEmail = getEmailByName(routeEntry.mgrRaw); }
    if (routeEntry.hodRaw) { result.hodName = routeEntry.hodRaw; result.hodEmail = getEmailByName(routeEntry.hodRaw); }
    if (routeEntry.csRaw)  { result.csuiteName = routeEntry.csRaw; result.csuiteEmail = getEmailByName(routeEntry.csRaw); }
  }

  var mgrE = (result.managerEmail || "").toLowerCase().trim();
  var hodE = (result.hodEmail || "").toLowerCase().trim();
  var csE  = (result.csuiteEmail || "").toLowerCase().trim();

  if (csE && hodE && csE === hodE) { result.hodEmail = null; result.hodName = null; hodE = ""; }
  if (csE && mgrE && csE === mgrE) { result.managerEmail = null; result.managerName = null; mgrE = ""; }
  if (hodE && mgrE && hodE === mgrE) { result.managerEmail = null; result.managerName = null; mgrE = ""; }

  if (mgrE && mgrE === email) { result.managerEmail = null; result.managerName = null; }
  if (hodE && hodE === email) { result.hodEmail = null; result.hodName = null; }
  if (csE && csE === email) { result.csuiteEmail = null; result.csuiteName = null; }

  if (result.managerEmail && (!result.managerName || result.managerName === "-")) result.managerName = getNameByEmail(result.managerEmail);
  if (result.hodEmail && (!result.hodName || result.hodName === "-")) result.hodName = getNameByEmail(result.hodEmail);
  if (result.csuiteEmail && (!result.csuiteName || result.csuiteName === "-")) result.csuiteName = getNameByEmail(result.csuiteEmail);

  if (_memCache.hierarchyCache) _memCache.hierarchyCache[email] = result;
  return result;
}

// ═══════════════════════════════════════════════════════════════
// QUOTATION SHEET
// ═══════════════════════════════════════════════════════════════
function getQuotationSheet() {
  if (_memCache.quotationSheet) return _memCache.quotationSheet;
  var ss = getTicketSpreadsheet();
  if (!ss) return null;
  var sheet = ss.getSheetByName("IT_Quotation");
  if (!sheet) {
    sheet = ss.insertSheet("IT_Quotation");
    sheet.appendRow([
      "Ticket ID", "Quotation ID", "Vendor Name", "Price", "Validity Date",
      "Specification", "Quotation File URL", "Remarks",
      "Added By", "Added By Email", "Added Date",
      "Is Selected", "Selected By", "Selected Date", "MD Remarks"
    ]);
    sheet.getRange(1, 1, 1, 15).setFontWeight("bold").setBackground("#7c3aed").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 15, [120, 110, 180, 100, 110, 320, 280, 200, 150, 200, 160, 100, 150, 160, 200]);
  }
  _memCache.quotationSheet = sheet;
  return sheet;
}

function generateQuotationID() {
  var sheet = getQuotationSheet();
  if (!sheet) return "QT-0001";
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "QT-0001";
  var data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var maxNum = 0;
  for (var i = 0; i < data.length; i++) {
    var id = data[i][0];
    if (id && typeof id === 'string') {
      var m = id.match(/(\d+)$/);
      if (m) { var num = parseInt(m[1], 10); if (num > maxNum) maxNum = num; }
    }
  }
  return "QT-" + String(maxNum + 1).padStart(4, '0');
}

function getQuotationsByTicket(ticketId) {
  try {
    if (!ticketId) return [];
    var sheet = getQuotationSheet();
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var results = [];
    var cleanTicketId = ticketId.toString().trim().toUpperCase();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if ((row[0] || "").toString().trim().toUpperCase() === cleanTicketId) {
        results.push({
          quotationId: String(row[1] || ""),
          ticketId: String(row[0] || ""),
          vendorName: String(row[2] || ""),
          price: String(row[3] || ""),
          validityDate: row[4] instanceof Date ? Utilities.formatDate(row[4], "GMT+8", "dd/MM/yyyy") : String(row[4] || ""),
          specification: String(row[5] || ""),
          quotationFileUrl: String(row[6] || ""),
          remarks: String(row[7] || ""),
          addedBy: String(row[8] || ""),
          addedByEmail: String(row[9] || ""),
          addedDate: formatSafeDate(row[10]),
          isSelected: String(row[11] || "").toUpperCase() === "TRUE",
          selectedBy: String(row[12] || ""),
          selectedDate: formatSafeDate(row[13]),
          mdRemarks: String(row[14] || ""),
          rowIndex: i + 1
        });
      }
    }
    return results;
  } catch (err) {
    Logger.log("getQuotationsByTicket error: " + err.toString());
    return [];
  }
}

function saveQuotation(payload) {
  try {
    if (typeof payload === "string") payload = JSON.parse(payload);
    if (!payload || !payload.ticketId || !payload.vendorName) {
      return JSON.stringify({ success: false, message: "Ticket ID & Vendor Name required." });
    }
    if (!payload.price || payload.price.toString().trim() === "") {
      return JSON.stringify({ success: false, message: "Price required." });
    }
    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });

    var myRoles = profile.allRoles || [profile.role];
    var sys = getSystemEmailConfig();
    var isITCheck = myRoles.indexOf("IT_CHECK") !== -1 || (sys.itCheckEmail && cleanEmailAddress(authEmail) === cleanEmailAddress(sys.itCheckEmail));
    if (!isITCheck) return JSON.stringify({ success: false, message: "Only IT Check can add quotations." });

    var existingQuotes = getQuotationsByTicket(payload.ticketId);
    var isFirstQuote = existingQuotes.length === 0 && !payload.quotationId;
    var fileUrl = (payload.quotationFileUrl || "").toString().trim();
    if (isFirstQuote && fileUrl === "") {
      return JSON.stringify({ success: false, message: "Quotation pertama mesti ada PDF. Sila upload fail." });
    }

    var sheet = getQuotationSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet unavailable." });

    var quotationId = payload.quotationId || generateQuotationID();
    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");

    var rowData = [
      payload.ticketId.toString().trim().toUpperCase(),
      quotationId,
      payload.vendorName.toString().trim(),
      payload.price.toString().trim(),
      "",
      "",
      fileUrl,
      (payload.remarks || "").toString(),
      profile.name || authEmail,
      authEmail,
      nowStr,
      "FALSE", "", "", ""
    ];

    if (payload.quotationId) {
      var allData = sheet.getDataRange().getValues();
      for (var i = 1; i < allData.length; i++) {
        if ((allData[i][1] || "").toString() === payload.quotationId) {
          if ((allData[i][11] || "").toString().toUpperCase() === "TRUE") {
            return JSON.stringify({ success: false, message: "Tidak boleh edit — MD sudah pilih quotation ini." });
          }
          sheet.getRange(i + 1, 1, 1, 11).setValues([rowData.slice(0, 11)]);
          return JSON.stringify({ success: true, message: "Quotation updated.", quotationId: quotationId, mode: "update" });
        }
      }
    }

    sheet.appendRow(rowData);
    return JSON.stringify({ success: true, message: "Quotation added.", quotationId: quotationId, mode: "create" });
  } catch (err) {
    Logger.log("saveQuotation error: " + err.toString());
    return JSON.stringify({ success: false, message: err.toString() });
  }
}

function deleteQuotation(payload) {
  try {
    if (typeof payload === "string") payload = JSON.parse(payload);
    if (!payload || !payload.quotationId) return JSON.stringify({ success: false, message: "Quotation ID required." });
    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });

    var myRoles = profile.allRoles || [profile.role];
    var sys = getSystemEmailConfig();
    var isITCheck = myRoles.indexOf("IT_CHECK") !== -1 || (sys.itCheckEmail && cleanEmailAddress(authEmail) === cleanEmailAddress(sys.itCheckEmail));
    if (!isITCheck) return JSON.stringify({ success: false, message: "Only IT Check can delete." });

    var sheet = getQuotationSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet unavailable." });

    var allData = sheet.getDataRange().getValues();
    for (var i = 1; i < allData.length; i++) {
      if ((allData[i][1] || "").toString() === payload.quotationId) {
        if ((allData[i][11] || "").toString().toUpperCase() === "TRUE") {
          return JSON.stringify({ success: false, message: "Cannot delete — MD already selected." });
        }
        sheet.deleteRow(i + 1);
        return JSON.stringify({ success: true, message: "Quotation deleted." });
      }
    }
    return JSON.stringify({ success: false, message: "Quotation not found." });
  } catch (err) {
    Logger.log("deleteQuotation error: " + err.toString());
    return JSON.stringify({ success: false, message: err.toString() });
  }
}

function getQuotationsForPortal(ticketId) {
  try {
    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });
    var quotations = getQuotationsByTicket(ticketId);
    return JSON.stringify({ success: true, data: quotations });
  } catch (err) {
    return JSON.stringify({ success: false, message: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════════
// SIGNATURE SHEET (MD)
// ═══════════════════════════════════════════════════════════════
function getSignatureSheet() {
  if (_memCache.signatureSheet) return _memCache.signatureSheet;
  var ss = getTicketSpreadsheet();
  if (!ss) return null;
  var sheet = ss.getSheetByName("MD_Signature");
  if (!sheet) {
    sheet = ss.insertSheet("MD_Signature");
    sheet.appendRow(["Email", "Signature File URL", "Updated Date"]);
    sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#7c3aed").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 3, [200, 300, 160]);
  }
  _memCache.signatureSheet = sheet;
  return sheet;
}

function getMDSignature() {
  try {
    var authEmail = getAuthenticatedUserEmail();
    var sheet = getSignatureSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet unavailable." });
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || "").toString().toLowerCase() === authEmail) {
        return JSON.stringify({
          success: true,
          data: {
            signatureUrl: String(data[i][1] || ""),
            updatedDate: formatSafeDate(data[i][2])
          }
        });
      }
    }
    return JSON.stringify({ success: true, data: null });
  } catch (err) {
    return JSON.stringify({ success: false, message: err.toString() });
  }
}

function saveMDSignature(payload) {
  try {
    if (typeof payload === "string") payload = JSON.parse(payload);
    if (!payload || !payload.base64Data) return JSON.stringify({ success: false, message: "Signature data required." });

    var authEmail = getAuthenticatedUserEmail();
    if (!authEmail) return JSON.stringify({ success: false, message: "Not authorized." });

    var FOLDER_NAME = "Apollo IT - Signatures";
    var folder;
    var folders = DriveApp.getFoldersByName(FOLDER_NAME);
    if (folders.hasNext()) folder = folders.next();
    else folder = DriveApp.createFolder(FOLDER_NAME);

    var cleaned = payload.base64Data.replace(/^data:image\/[a-z]+;base64,/, "");
    var bytes = Utilities.base64Decode(cleaned);
    var blob = Utilities.newBlob(bytes, payload.mimeType || "image/png", "Signature_" + authEmail.replace(/@.*/, "") + ".png");
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

    var fileUrl = file.getUrl();
    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");

    var sheet = getSignatureSheet();
    var data = sheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || "").toString().toLowerCase() === authEmail) {
        sheet.getRange(i + 1, 2).setValue(fileUrl);
        sheet.getRange(i + 1, 3).setValue(nowStr);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([authEmail, fileUrl, nowStr]);

    return JSON.stringify({ success: true, data: { signatureUrl: fileUrl } });
  } catch (err) {
    Logger.log("saveMDSignature error: " + err.toString());
    return JSON.stringify({ success: false, message: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════════
// ATTACHMENT UPLOAD
// ═══════════════════════════════════════════════════════════════
function saveAssessmentAttachment(base64Data, fileName, mimeType, ticketId) {
  try {
    var FOLDER_NAME = "Apollo IT - Attachments";
    var folder;
    var folders = DriveApp.getFoldersByName(FOLDER_NAME);
    if (folders.hasNext()) folder = folders.next();
    else folder = DriveApp.createFolder(FOLDER_NAME);

    var cleaned = base64Data.replace(/^data:[^;]+;base64,/, "");
    var bytes = Utilities.base64Decode(cleaned);
    var safeName = ticketId + "_" + fileName.replace(/[^\w\-\.]/g, "_");
    var blob = Utilities.newBlob(bytes, mimeType || "application/octet-stream", safeName);
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return { success: true, url: file.getUrl(), fileId: file.getId(), fileName: safeName };
  } catch (err) {
    Logger.log("saveAssessmentAttachment error: " + err.toString());
    return { success: false, message: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════
// ASSESSMENT SHEET
// ═══════════════════════════════════════════════════════════════
function getAssessmentSheet() {
  if (_memCache.assessmentSheet) return _memCache.assessmentSheet;
  var ss = getTicketSpreadsheet();
  if (!ss) return null;
  var sheet = ss.getSheetByName("IT_Assessment");
  if (!sheet) {
    sheet = ss.insertSheet("IT_Assessment");
    sheet.appendRow([
      "Ticket ID", "New Asset Provisioning", "IT Assessment Justification",
      "User Operational Issue", "Recommended Action Plan",
      "Assessed By", "Assessed By Email", "Assessed Date", "Attachment Links JSON"
    ]);
    sheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  _memCache.assessmentSheet = sheet;
  return sheet;
}

function getAssessmentData(ticketId) {
  try {
    var sheet = getAssessmentSheet();
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return null;
    var clean = ticketId.toString().trim().toUpperCase();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || "").toString().trim().toUpperCase() === clean) {
        var attLinks = [];
        try { attLinks = JSON.parse(data[i][8] || "[]"); } catch (e) {}
        return {
          newAssetProvisioning: String(data[i][1] || ""),
          assessmentJustification: String(data[i][2] || ""),
          userOperationalIssue: String(data[i][3] || ""),
          recommendedActionPlan: String(data[i][4] || ""),
          assessedBy: String(data[i][5] || ""),
          assessedByEmail: String(data[i][6] || ""),
          assessedDate: formatSafeDate(data[i][7]),
          attachmentLinks: attLinks
        };
      }
    }
    return null;
  } catch (e) { return null; }
}

function saveAssessmentData(ticketId, payload) {
  try {
    var sheet = getAssessmentSheet();
    if (!sheet) return false;
    var data = sheet.getDataRange().getValues();
    var clean = ticketId.toString().trim().toUpperCase();
    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");
    var attJson = JSON.stringify(payload.attachmentLinks || []);
    var rowData = [
      clean,
      payload.newAssetProvisioning || "",
      payload.assessmentJustification || "",
      payload.userOperationalIssue || "",
      payload.recommendedActionPlan || "",
      payload.assessedBy || "",
      payload.assessedByEmail || "",
      nowStr,
      attJson
    ];
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || "").toString().trim().toUpperCase() === clean) {
        sheet.getRange(i + 1, 1, 1, 9).setValues([rowData]);
        return true;
      }
    }
    sheet.appendRow(rowData);
    return true;
  } catch (err) {
    Logger.log("saveAssessmentData error: " + err.toString());
    return false;
  }
}

// ═══════════════════════════════════════════════════════════════
// DISCUSSION
// ═══════════════════════════════════════════════════════════════
function getAllDiscussionRows() {
  var rows = [];
  var ss = getTicketSpreadsheet();
  if (!ss) return rows;

  try {
    var s1 = ss.getSheetByName("Discussion");
    if (s1) {
      var d1 = s1.getDataRange().getValues();
      for (var i = 1; i < d1.length; i++) {
        var tid1 = (d1[i][0] || "").toString().trim().toUpperCase();
        if (!tid1) continue;
        rows.push({
          ticketId: tid1,
          timestamp: formatSafeDate(d1[i][1]),
          commenterName: String(d1[i][2] || ""),
          commenterEmail: String(d1[i][3] || ""),
          role: String(d1[i][4] || ""),
          comment: String(d1[i][5] || "")
        });
      }
    }
  } catch (e) { Logger.log("Read Discussion error: " + e.toString()); }

  try {
    var s2 = ss.getSheetByName("IT_Discussion");
    if (s2) {
      var d2 = s2.getDataRange().getValues();
      for (var j = 1; j < d2.length; j++) {
        var tid2 = (d2[j][1] || "").toString().trim().toUpperCase();
        if (!tid2) continue;
        rows.push({
          ticketId: tid2,
          timestamp: formatSafeDate(d2[j][6]),
          commenterName: String(d2[j][2] || ""),
          commenterEmail: String(d2[j][3] || ""),
          role: String(d2[j][4] || ""),
          comment: String(d2[j][5] || "")
        });
      }
    }
  } catch (e) { Logger.log("Read IT_Discussion error: " + e.toString()); }

  return rows;
}

function getDiscussionSheet() {
  if (_memCache.discussionSheet) return _memCache.discussionSheet;
  var ss = getTicketSpreadsheet();
  if (!ss) return null;
  var sheet = ss.getSheetByName("Discussion");
  if (!sheet) sheet = ss.getSheetByName("IT_Discussion");
  if (!sheet) {
    sheet = ss.insertSheet("Discussion");
    sheet.appendRow([
      "Ticket ID", "Timestamp", "Commenter Name", "Commenter Email", "Role", "Comment"
    ]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#059669").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  _memCache.discussionSheet = sheet;
  return sheet;
}

function canUserAccessTicketDiscussion(state, authEmail, profile) {
  if (!state || !authEmail || !profile) return false;

  var myRoles = profile.allRoles || [profile.role];
  var myEmail = (authEmail || "").toString().toLowerCase().trim();

  if (myRoles.indexOf("IT_CHECK") !== -1 || myRoles.indexOf("IT_HOD") !== -1 || myRoles.indexOf("IT_VERIFIED") !== -1) return true;

  var h = state.hierarchy || {};

  if (state.emailUser && state.emailUser.toLowerCase() === myEmail) return true;

  if (h.managerEmail && h.managerEmail.toLowerCase() === myEmail) return true;
  if (h.hodEmail && h.hodEmail.toLowerCase() === myEmail) return true;
  if (h.csuiteEmail && h.csuiteEmail.toLowerCase() === myEmail) return true;
  if (h.mdEmail && h.mdEmail.toLowerCase() === myEmail) return true;

  return false;
}

function canUserChatOnTicket(state, authEmail, profile) {
  if (!state || !authEmail || !profile) return false;
  if (state.isTerminated || state.isCompleted) return false;
  if (!canUserAccessTicketDiscussion(state, authEmail, profile)) return false;

  var myRoles = profile.allRoles || [profile.role];
  var myEmail = (authEmail || "").toString().toLowerCase().trim();
  var rd = state.rowData || {};
  var h = state.hierarchy || {};

  if (myRoles.indexOf("IT_VERIFIED") !== -1) return true;
  if (myRoles.indexOf("MD") !== -1) return false;

  if (myRoles.indexOf("IT_CHECK") !== -1 || myRoles.indexOf("IT_HOD") !== -1) {
    if (state.stage === "MD" || state.isReadyProc) return false;
    return true;
  }

  if (h.managerEmail && h.managerEmail.toLowerCase() === myEmail) {
    var mgrDate = String(rd["Manager Approval Date"] || "").trim();
    return !(mgrDate && mgrDate !== "Pending" && mgrDate !== "-" && mgrDate.indexOf("Bypassed") === -1);
  }

  if (h.hodEmail && h.hodEmail.toLowerCase() === myEmail) {
    var hodDate = String(rd["HOD Approval Date"] || "").trim();
    return !(hodDate && hodDate !== "Pending" && hodDate !== "-" && hodDate.indexOf("Bypassed") === -1);
  }

  if (h.csuiteEmail && h.csuiteEmail.toLowerCase() === myEmail) {
    var csDate = String(rd["C-Suite Approval Date"] || "").trim();
    return !(csDate && csDate !== "Pending" && csDate !== "-" && csDate.indexOf("Bypassed") === -1);
  }

  if (state.emailUser && state.emailUser.toLowerCase() === myEmail) return true;

  return false;
}

function getDiscussionMetrics() {
  try {
    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });

    var allRows = getAllDiscussionRows();
    if (!allRows.length) {
      return JSON.stringify({ success: true, data: { ticketsWithDiscussion: 0, totalDiscussions: 0 } });
    }

    var ticketsWithChat = {};
    for (var i = 0; i < allRows.length; i++) {
      if (allRows[i].ticketId) ticketsWithChat[allRows[i].ticketId] = true;
    }

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: true, data: { ticketsWithDiscussion: 0, totalDiscussions: 0 } });

    var rawData = sheet.getDataRange().getValues();
    if (rawData.length < 2) return JSON.stringify({ success: true, data: { ticketsWithDiscussion: 0, totalDiscussions: 0 } });

    var headers = rawData[0];
    var allData = rawData.slice(1);

    var visibleTicketSet = {};
    for (var r = 0; r < allData.length; r++) {
      var state = determineTicketWorkflowState(allData[r], headers, r + 2);
      if (!state.ticketId) continue;
      var tid = state.ticketId.toString().trim().toUpperCase();
      if (!ticketsWithChat[tid]) continue;
      if (!canUserAccessTicketDiscussion(state, authEmail, profile)) continue;
      visibleTicketSet[tid] = true;
    }

    var visibleCommentCount = 0;
    for (var j = 0; j < allRows.length; j++) {
      if (visibleTicketSet[allRows[j].ticketId]) visibleCommentCount++;
    }

    return JSON.stringify({
      success: true,
      data: {
        ticketsWithDiscussion: Object.keys(visibleTicketSet).length,
        totalDiscussions: visibleCommentCount
      }
    });
  } catch (err) {
    Logger.log("getDiscussionMetrics error: " + err.toString());
    return JSON.stringify({ success: false, message: err.toString() });
  }
}

function getDiscussionComments(ticketId) {
  try {
    if (!ticketId) return JSON.stringify({ success: false, message: "Ticket ID required." });

    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet unavailable." });
    var rawData = sheet.getDataRange().getValues();
    var headers = rawData[0];
    var allData = rawData.slice(1);
    var state = null;
    for (var i = 0; i < allData.length; i++) {
      var rData = {};
      for (var h = 0; h < headers.length; h++) {
        var cleanH = headers[h].toString().trim();
        if (!rData.hasOwnProperty(cleanH)) rData[cleanH] = allData[i][h];
      }
      if ((rData["Ticket ID"] || "").toString().trim().toUpperCase() === ticketId.toString().trim().toUpperCase()) {
        state = determineTicketWorkflowState(allData[i], headers, i + 2);
        break;
      }
    }
    if (!state) return JSON.stringify({ success: false, message: "Ticket not found." });
    if (!canUserAccessTicketDiscussion(state, authEmail, profile)) {
      return JSON.stringify({ success: false, message: "Access denied for this ticket." });
    }

    var clean = ticketId.toString().trim().toUpperCase();
    var allRows = getAllDiscussionRows();
    var results = [];

    for (var j = 0; j < allRows.length; j++) {
      if (allRows[j].ticketId === clean) {
        results.push({
          commentId: "C-" + j,
          ticketId: allRows[j].ticketId,
          commenterName: allRows[j].commenterName,
          commenterEmail: allRows[j].commenterEmail,
          role: allRows[j].role,
          comment: allRows[j].comment,
          timestamp: allRows[j].timestamp
        });
      }
    }
    return JSON.stringify({ success: true, data: results });
  } catch (err) {
    Logger.log("getDiscussionComments error: " + err.toString());
    return JSON.stringify({ success: false, message: err.toString() });
  }
}

function addDiscussionComment(payload) {
  try {
    if (typeof payload === "string") payload = JSON.parse(payload);
    if (!payload || !payload.ticketId || !payload.comment) {
      return JSON.stringify({ success: false, message: "Ticket ID & comment required." });
    }

    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet unavailable." });
    var rawData = sheet.getDataRange().getValues();
    var headers = rawData[0];
    var allData = rawData.slice(1);
    var state = null;
    for (var i = 0; i < allData.length; i++) {
      var rData = {};
      for (var h = 0; h < headers.length; h++) {
        var cleanH = headers[h].toString().trim();
        if (!rData.hasOwnProperty(cleanH)) rData[cleanH] = allData[i][h];
      }
      if ((rData["Ticket ID"] || "").toString().trim().toUpperCase() === payload.ticketId.toString().trim().toUpperCase()) {
        state = determineTicketWorkflowState(allData[i], headers, i + 2);
        break;
      }
    }
    if (!state) return JSON.stringify({ success: false, message: "Ticket not found." });

    if (!canUserChatOnTicket(state, authEmail, profile)) {
      return JSON.stringify({ success: false, message: "You are not allowed to comment on this ticket at the current stage." });
    }

    var discSheet = getDiscussionSheet();
    if (!discSheet) return JSON.stringify({ success: false, message: "Sheet unavailable." });

    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");
    discSheet.appendRow([
      payload.ticketId.toString().trim().toUpperCase(),
      nowStr,
      payload.commenterName || profile.name || authEmail,
      payload.commenterEmail || authEmail,
      payload.role || profile.role || "",
      payload.comment.toString()
    ]);

    return JSON.stringify({ success: true, data: { message: "Comment posted." } });
  } catch (err) {
    Logger.log("addDiscussionComment error: " + err.toString());
    return JSON.stringify({ success: false, message: err.toString() });
  }
}

function getDiscussionMessages(ticketId) {
  try {
    if (!ticketId) return JSON.stringify({ success: false, message: "Ticket ID required." });
    var authEmail = getAuthenticatedUserEmail();
    var clean = ticketId.toString().trim().toUpperCase();
    var allRows = getAllDiscussionRows();
    var results = [];

    for (var i = 0; i < allRows.length; i++) {
      if (allRows[i].ticketId === clean) {
        results.push({
          messageId: "MSG-" + i,
          ticketId: allRows[i].ticketId,
          senderName: allRows[i].commenterName,
          senderEmail: allRows[i].commenterEmail,
          senderRole: allRows[i].role,
          message: allRows[i].comment,
          timestamp: allRows[i].timestamp,
          isRead: true,
          isMine: allRows[i].commenterEmail.toLowerCase() === authEmail.toLowerCase()
        });
      }
    }
    return JSON.stringify({ success: true, data: results });
  } catch (err) {
    return JSON.stringify({ success: false, message: err.toString() });
  }
}

function postDiscussionMessage(payload) {
  try {
    if (typeof payload === "string") payload = JSON.parse(payload);
    if (!payload || !payload.ticketId || !payload.message) return JSON.stringify({ success: false, message: "Invalid." });
    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });
    var sheet = getDiscussionSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet unavailable." });

    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");
    sheet.appendRow([
      payload.ticketId.toString().trim().toUpperCase(),
      nowStr,
      profile.name || authEmail,
      authEmail,
      profile.role || "",
      payload.message.toString()
    ]);
    return JSON.stringify({ success: true, data: { message: "Comment posted." } });
  } catch (err) {
    return JSON.stringify({ success: false, message: err.toString() });
  }
}

function markDiscussionRead(payload) {
  return JSON.stringify({ success: true });
}

// ═══════════════════════════════════════════════════════════════
// PHASE 4: IT CHECK SUBMIT TO MD
// ═══════════════════════════════════════════════════════════════
function itCheckSubmitToMD(payload) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "System busy." });
    if (typeof payload === "string") payload = JSON.parse(payload);
    if (!payload || !payload.ticketId) return JSON.stringify({ success: false, message: "Invalid payload." });

    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });

    var myRoles = profile.allRoles || [profile.role];
    var sys = getSystemEmailConfig();
    var isITCheck = myRoles.indexOf("IT_CHECK") !== -1 || (sys.itCheckEmail && cleanEmailAddress(authEmail) === cleanEmailAddress(sys.itCheckEmail));
    if (!isITCheck) return JSON.stringify({ success: false, message: "Only IT Check can submit to MD." });

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet not found." });

    var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var targetRowValues = null, targetRowIndex = -1;

    for (var i = 0; i < allData.length; i++) {
      var rData = {};
      for (var h = 0; h < headers.length; h++) {
        var cleanH = headers[h].toString().trim();
        if (!rData.hasOwnProperty(cleanH)) rData[cleanH] = allData[i][h];
      }
      if ((rData["Ticket ID"] || "").toString().trim().toUpperCase() === payload.ticketId.toString().trim().toUpperCase()) {
        targetRowValues = allData[i]; targetRowIndex = i + 2; break;
      }
    }

    if (!targetRowValues) return JSON.stringify({ success: false, message: "Ticket not found." });

    var state = determineTicketWorkflowState(targetRowValues, headers, targetRowIndex);
    if (state.isTerminated || state.isCompleted) return JSON.stringify({ success: false, message: "Already finalized." });
    if (state.stage !== "IT_CHECK") return JSON.stringify({ success: false, message: "Ticket is not at IT Check stage." });

    if (state.isSimpleAsset) {
      return JSON.stringify({
        success: false,
        message: "Simple asset (Monitor/Keyboard/Mouse) does not need MD approval. Please use 'Mark as Completed' instead."
      });
    }

    var quotations = getQuotationsByTicket(payload.ticketId);
    if (!quotations || quotations.length === 0) {
      return JSON.stringify({ success: false, message: "Please add at least 1 quotation before submitting to MD." });
    }

    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");
    updateSingleColStatus(sheet, targetRowIndex, headers, "IT Check (Execution)", profile.name || authEmail);
    updateSingleColStatus(sheet, targetRowIndex, headers, "IT Check Date", nowStr);
    updateSingleColStatus(sheet, targetRowIndex, headers, "Submitted to MD Date", nowStr);
    updateSingleColStatus(sheet, targetRowIndex, headers, "Status", "Pending MD");

    var rowData = {};
    for (var i2 = 0; i2 < headers.length; i2++) {
      rowData[headers[i2].toString().trim()] = targetRowValues[i2];
    }
    var emailUser = extractUserEmailFromRow(targetRowValues);
    var namaUser = rowData["Full Name"] || rowData["Nama Penuh"] || "Staff";
    var assetName = rowData["Asset Type"] || rowData["Jenis Aset"] || "N/A";
    var assetQty = rowData["Quantity"] || rowData["Kuantiti"] || "1";
    var hierarchy = getApprovalHierarchy(emailUser);

    var chainData = buildChainDataStructure(hierarchy, emailUser, state.isSimpleAsset, rowData);
    syncStepStatus(chainData, "IT CHECK", "Verified via Email", nowStr);
    var pdfBlob = buildPdfForm(rowData, payload.ticketId, chainData);

    var portalUrl = WEB_APP_URL;
    sendSafeEmail({
      to: sys.mdEmail, name: "Apollo IT Portal",
      subject: "[MD REVIEW] Quotation Ready: " + payload.ticketId + " (" + namaUser + ")",
      htmlBody: buildUniformEmailHtml(
        "MANAGING DIRECTOR (MD) - QUOTATION REVIEW",
        "Dear <b>" + sys.mdName + " (MD)</b>,",
        namaUser, payload.ticketId, assetName, assetQty,
        "Standard Asset - " + quotations.length + " Quotation(s) Submitted",
        portalUrl, null, rowData,
        {
          introText: "IT Check has prepared <b>" + quotations.length + " quotation(s)</b> for your review. Please log in to the MD Portal to select the preferred quotation.",
          headerBg: "#7c3aed"
        }
      ),
      attachments: [pdfBlob]
    });

    return JSON.stringify({ success: true, data: { message: "Submitted to MD. " + quotations.length + " quotation(s) sent for review." } });
  } catch (err) {
    Logger.log("itCheckSubmitToMD error: " + err.toString());
    return JSON.stringify({ success: false, message: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ═══════════════════════════════════════════════════════════════
// PHASE 4: MD SELECT QUOTATION (With Signature)
// ═══════════════════════════════════════════════════════════════
function mdSelectQuotation(payload) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "System busy." });
    if (typeof payload === "string") payload = JSON.parse(payload);
    if (!payload || !payload.ticketId || !payload.quotationId) {
      return JSON.stringify({ success: false, message: "Ticket ID & Quotation ID required." });
    }

    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });

    var myRoles = profile.allRoles || [profile.role];
    var sys = getSystemEmailConfig();
    var isMD = myRoles.indexOf("MD") !== -1 || (sys.mdEmail && cleanEmailAddress(authEmail) === cleanEmailAddress(sys.mdEmail));
    if (!isMD) return JSON.stringify({ success: false, message: "Only MD can select quotation." });

    var signatureUrl = "";
    if (payload.signatureBase64) {
      var sigUploadRes = saveMDSignature({ base64Data: payload.signatureBase64, mimeType: payload.signatureMimeType });
      if (!sigUploadRes.success) return JSON.stringify({ success: false, message: "Signature upload failed." });
      var sigParsed = JSON.parse(sigUploadRes);
      signatureUrl = sigParsed.data.signatureUrl;
    } else if (payload.signatureUrl) {
      signatureUrl = payload.signatureUrl;
    } else {
      var sheet = getSignatureSheet();
      var sData = sheet.getDataRange().getValues();
      for (var si = 1; si < sData.length; si++) {
        if ((sData[si][0] || "").toString().toLowerCase() === authEmail) {
          signatureUrl = String(sData[si][1] || "");
          break;
        }
      }
    }

    if (!signatureUrl) {
      return JSON.stringify({ success: false, message: "MD signature is required. Please save a signature first." });
    }

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet not found." });

    var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var targetRowValues = null, targetRowIndex = -1;

    for (var i = 0; i < allData.length; i++) {
      var rData = {};
      for (var h = 0; h < headers.length; h++) {
        var cleanH = headers[h].toString().trim();
        if (!rData.hasOwnProperty(cleanH)) rData[cleanH] = allData[i][h];
      }
      if ((rData["Ticket ID"] || "").toString().trim().toUpperCase() === payload.ticketId.toString().trim().toUpperCase()) {
        targetRowValues = allData[i]; targetRowIndex = i + 2; break;
      }
    }

    if (!targetRowValues) return JSON.stringify({ success: false, message: "Ticket not found." });

    var state = determineTicketWorkflowState(targetRowValues, headers, targetRowIndex);
    if (state.isTerminated || state.isCompleted) {
      return JSON.stringify({ success: false, message: "Ticket sudah finalized." });
    }
    if (state.stage !== "MD" && !state.isReadyProc) {
      return JSON.stringify({ success: false, message: "Ticket tidak dalam stage MD." });
    }

    var qSheet = getQuotationSheet();
    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");
    var qData = qSheet.getDataRange().getValues();
    var found = false;
    var selectedVendor = "";

    for (var q = 1; q < qData.length; q++) {
      var qid = (qData[q][1] || "").toString().trim();
      var qticket = (qData[q][0] || "").toString().trim().toUpperCase();
      if (qticket === payload.ticketId.toString().trim().toUpperCase()) {
        if (qid === payload.quotationId) {
          qSheet.getRange(q + 1, 12).setValue("TRUE");
          qSheet.getRange(q + 1, 13).setValue(profile.name || authEmail);
          qSheet.getRange(q + 1, 14).setValue(nowStr);
          qSheet.getRange(q + 1, 15).setValue(payload.mdRemarks || "");
          selectedVendor = (qData[q][2] || "").toString();
          found = true;
        } else {
          qSheet.getRange(q + 1, 12).setValue("FALSE");
        }
      }
    }

    if (!found) return JSON.stringify({ success: false, message: "Quotation not found." });

    updateSingleColStatus(sheet, targetRowIndex, headers, "MD Approver", profile.name || authEmail);
    updateSingleColStatus(sheet, targetRowIndex, headers, "MD Approval Date", nowStr);
    updateSingleColStatus(sheet, targetRowIndex, headers, "Status", "Approved - Ready for Procurement");

    var rowData = {};
    for (var i2 = 0; i2 < headers.length; i2++) {
      rowData[headers[i2].toString().trim()] = targetRowValues[i2];
    }
    var emailUser = extractUserEmailFromRow(targetRowValues);
    var namaUser = rowData["Full Name"] || rowData["Nama Penuh"] || "Staff";
    var assetName = rowData["Asset Type"] || rowData["Jenis Aset"] || "N/A";
    var assetQty = rowData["Quantity"] || rowData["Kuantiti"] || "1";
    var hierarchy = getApprovalHierarchy(emailUser);

    var chainData = buildChainDataStructure(hierarchy, emailUser, false, rowData);
    syncStepStatus(chainData, "MANAGING DIRECTOR (MD)", "Approved via Email", nowStr);
    var pdfBlob = buildPdfForm(rowData, payload.ticketId, chainData, signatureUrl);

    var recipients = [sys.itCheckEmail, sys.itHodEmail].filter(Boolean).join(",");
    sendSafeEmail({
      to: recipients, name: "Apollo IT Portal",
      subject: "[MD SELECTED] Quotation Approved: " + payload.ticketId + " (" + selectedVendor + ")",
      htmlBody: buildUniformEmailHtml(
        "QUOTATION SELECTED - READY FOR PROCUREMENT",
        "Dear <b>IT Team</b>,",
        namaUser, payload.ticketId, assetName, assetQty,
        "Standard Asset (MD Selected: " + selectedVendor + ")",
        null, null, rowData,
        {
          introText: "MD has selected quotation from <b>" + selectedVendor + "</b>. Please proceed with procurement." +
                     (payload.mdRemarks ? "<br><br><b>MD Remarks:</b> " + payload.mdRemarks : ""),
          headerBg: "#2e7d32"
        }
      ),
      attachments: [pdfBlob]
    });

    return JSON.stringify({ success: true, data: { message: "Quotation selected.", vendor: selectedVendor } });
  } catch (err) {
    Logger.log("mdSelectQuotation error: " + err.toString());
    return JSON.stringify({ success: false, message: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ═══════════════════════════════════════════════════════════════
// PHASE 4: MD REJECT TICKET
// ═══════════════════════════════════════════════════════════════
function mdRejectTicket(payload) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "System busy." });
    if (typeof payload === "string") payload = JSON.parse(payload);
    if (!payload || !payload.ticketId || !payload.remarks) {
      return JSON.stringify({ success: false, message: "Ticket ID & remarks required." });
    }

    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });

    var myRoles = profile.allRoles || [profile.role];
    var sys = getSystemEmailConfig();
    var isMD = myRoles.indexOf("MD") !== -1 || (sys.mdEmail && cleanEmailAddress(authEmail) === cleanEmailAddress(sys.mdEmail));
    if (!isMD) return JSON.stringify({ success: false, message: "Only MD can reject at this stage." });

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet not found." });

    var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var targetRowValues = null, targetRowIndex = -1;

    for (var i = 0; i < allData.length; i++) {
      var rData = {};
      for (var h = 0; h < headers.length; h++) {
        var cleanH = headers[h].toString().trim();
        if (!rData.hasOwnProperty(cleanH)) rData[cleanH] = allData[i][h];
      }
      if ((rData["Ticket ID"] || "").toString().trim().toUpperCase() === payload.ticketId.toString().trim().toUpperCase()) {
        targetRowValues = allData[i]; targetRowIndex = i + 2; break;
      }
    }

    if (!targetRowValues) return JSON.stringify({ success: false, message: "Ticket not found." });

    var state = determineTicketWorkflowState(targetRowValues, headers, targetRowIndex);
    if (state.stage !== "MD") {
      return JSON.stringify({ success: false, message: "Ticket is not pending MD review." });
    }

    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");
    updateSingleColStatus(sheet, targetRowIndex, headers, "MD Approver", profile.name || authEmail);
    updateSingleColStatus(sheet, targetRowIndex, headers, "MD Approval Date", nowStr);
    updateSingleColStatus(sheet, targetRowIndex, headers, "Status", "Rejected at MD");

    var rowData = {};
    for (var i2 = 0; i2 < headers.length; i2++) {
      rowData[headers[i2].toString().trim()] = targetRowValues[i2];
    }
    var emailUser = extractUserEmailFromRow(targetRowValues);
    var namaUser = rowData["Full Name"] || rowData["Nama Penuh"] || "Staff";
    var assetName = rowData["Asset Type"] || rowData["Jenis Aset"] || "N/A";
    var assetQty = rowData["Quantity"] || rowData["Kuantiti"] || "1";
    var hierarchy = getApprovalHierarchy(emailUser);

    var chainData = buildChainDataStructure(hierarchy, emailUser, false, rowData);
    markChainRejected(chainData, "MANAGING DIRECTOR (MD)", nowStr);
    var pdfBlob = buildPdfForm(rowData, payload.ticketId, chainData);

    var itEmail = sys.itCheckEmail || "it@apollofood.com.my";
    sendSafeEmail({
      to: itEmail, name: "Apollo IT Portal",
      subject: "[MD REJECTED] Quotation Not Selected: " + payload.ticketId + " (" + namaUser + ")",
      htmlBody: buildUniformEmailHtml(
        "MANAGING DIRECTOR (MD) - REJECTED",
        "Dear <b>IT Team</b>,",
        namaUser, payload.ticketId, assetName, assetQty,
        "Standard Asset (Rejected by MD)",
        null, null, rowData,
        {
          introText: "MD has reviewed the submitted quotation(s) and rejected the request. Please review and take action accordingly.",
          headerBg: "#c62828",
          remarks: payload.remarks
        }
      ),
      attachments: [pdfBlob]
    });

    return JSON.stringify({ success: true, data: { message: "Ticket rejected by MD. IT team notified." } });
  } catch (err) {
    Logger.log("mdRejectTicket error: " + err.toString());
    return JSON.stringify({ success: false, message: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ═══════════════════════════════════════════════════════════════
// PHASE 4: IT CHECK MARK AS COMPLETED
// ═══════════════════════════════════════════════════════════════
function itCheckMarkComplete(payload) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "System busy." });
    if (typeof payload === "string") payload = JSON.parse(payload);
    if (!payload || !payload.ticketId) return JSON.stringify({ success: false, message: "Ticket ID required." });

    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });

    var myRoles = profile.allRoles || [profile.role];
    var sys = getSystemEmailConfig();
    var isITCheck = myRoles.indexOf("IT_CHECK") !== -1 || (sys.itCheckEmail && cleanEmailAddress(authEmail) === cleanEmailAddress(sys.itCheckEmail));
    if (!isITCheck) return JSON.stringify({ success: false, message: "Only IT Check can mark complete." });

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet not found." });

    var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var targetRowValues = null, targetRowIndex = -1;

    for (var i = 0; i < allData.length; i++) {
      var rData = {};
      for (var h = 0; h < headers.length; h++) {
        var cleanH = headers[h].toString().trim();
        if (!rData.hasOwnProperty(cleanH)) rData[cleanH] = allData[i][h];
      }
      if ((rData["Ticket ID"] || "").toString().trim().toUpperCase() === payload.ticketId.toString().trim().toUpperCase()) {
        targetRowValues = allData[i]; targetRowIndex = i + 2; break;
      }
    }

    if (!targetRowValues) return JSON.stringify({ success: false, message: "Ticket not found." });

    var state = determineTicketWorkflowState(targetRowValues, headers, targetRowIndex);
    if (state.isTerminated) return JSON.stringify({ success: false, message: "Ticket is rejected." });

    var statusUpper = (state.status || "").toUpperCase();
    var canComplete = false;
    if (statusUpper.indexOf("READY FOR PROCUREMENT") !== -1) canComplete = true;
    if (state.stage === "IT_CHECK" && state.isSimpleAsset) canComplete = true;

    if (!canComplete) {
      return JSON.stringify({ success: false, message: "Ticket not ready. Status: " + state.status });
    }

    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");
    updateSingleColStatus(sheet, targetRowIndex, headers, "Status", "Approved & Completed");
    updateSingleColStatus(sheet, targetRowIndex, headers, "IT HOD Verified", profile.name || authEmail);
    updateSingleColStatus(sheet, targetRowIndex, headers, "IT HOD Date", nowStr);

    var rowData = {};
    for (var i2 = 0; i2 < headers.length; i2++) {
      rowData[headers[i2].toString().trim()] = targetRowValues[i2];
    }
    var emailUser = extractUserEmailFromRow(targetRowValues);
    var namaUser = rowData["Full Name"] || rowData["Nama Penuh"] || "Staff";
    var assetName = rowData["Asset Type"] || rowData["Jenis Aset"] || "N/A";
    var assetQty = rowData["Quantity"] || rowData["Kuantiti"] || "1";

    var recipients = [emailUser, sys.itCheckEmail, sys.itHodEmail].filter(Boolean).join(",");
    sendSafeEmail({
      to: recipients, name: "Apollo IT Portal",
      subject: "[COMPLETED] Ticket " + payload.ticketId + " Fully Completed",
      htmlBody: buildUniformEmailHtml(
        "TICKET FULLY COMPLETED",
        "Dear <b>" + namaUser + " & IT Team</b>,",
        namaUser, payload.ticketId, assetName, assetQty,
        state.isSimpleAsset ? "Simple Asset (Completed)" : "Standard Asset (Completed)",
        null, null, rowData,
        {
          introText: "The IT Asset Request <b>[" + payload.ticketId + "]</b> has been fully completed and delivered.",
          headerBg: "#2e7d32"
        }
      )
    });

    return JSON.stringify({ success: true, data: { message: "Ticket marked as completed." } });
  } catch (err) {
    Logger.log("itCheckMarkComplete error: " + err.toString());
    return JSON.stringify({ success: false, message: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ═══════════════════════════════════════════════════════════════
// PHASE 7: IT CHECK MARK READY FOR PROCUREMENT (Simple Asset only)
// ═══════════════════════════════════════════════════════════════
function itCheckMarkReady(payload) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "System busy." });
    if (typeof payload === "string") payload = JSON.parse(payload);
    if (!payload || !payload.ticketId) return JSON.stringify({ success: false, message: "Ticket ID required." });

    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });

    var myRoles = profile.allRoles || [profile.role];
    var sys = getSystemEmailConfig();
    var isITCheck = myRoles.indexOf("IT_CHECK") !== -1 || (sys.itCheckEmail && cleanEmailAddress(authEmail) === cleanEmailAddress(sys.itCheckEmail));
    if (!isITCheck) return JSON.stringify({ success: false, message: "Only IT Check can mark ready." });

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet not found." });

    var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var targetRowValues = null, targetRowIndex = -1;

    for (var i = 0; i < allData.length; i++) {
      var rData = {};
      for (var h = 0; h < headers.length; h++) {
        var cleanH = headers[h].toString().trim();
        if (!rData.hasOwnProperty(cleanH)) rData[cleanH] = allData[i][h];
      }
      if ((rData["Ticket ID"] || "").toString().trim().toUpperCase() === payload.ticketId.toString().trim().toUpperCase()) {
        targetRowValues = allData[i]; targetRowIndex = i + 2; break;
      }
    }

    if (!targetRowValues) return JSON.stringify({ success: false, message: "Ticket not found." });

    var state = determineTicketWorkflowState(targetRowValues, headers, targetRowIndex);
    if (state.isTerminated || state.isCompleted) return JSON.stringify({ success: false, message: "Already finalized." });
    if (state.stage !== "IT_CHECK") return JSON.stringify({ success: false, message: "Ticket is not at IT Check stage." });
    if (!state.isSimpleAsset) return JSON.stringify({ success: false, message: "Only simple asset can use this action." });

    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");
    updateSingleColStatus(sheet, targetRowIndex, headers, "IT Check (Execution)", profile.name || authEmail);
    updateSingleColStatus(sheet, targetRowIndex, headers, "IT Check Date", nowStr);
    updateSingleColStatus(sheet, targetRowIndex, headers, "Status", "Approved - Ready for Procurement");

    return JSON.stringify({ success: true, data: { message: "Ticket marked ready for procurement." } });
  } catch (err) {
    Logger.log("itCheckMarkReady error: " + err.toString());
    return JSON.stringify({ success: false, message: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ═══════════════════════════════════════════════════════════════
// PHASE 4: IT CHECK REJECT TICKET
// ═══════════════════════════════════════════════════════════════
function itCheckRejectTicket(payload) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "System busy." });
    if (typeof payload === "string") payload = JSON.parse(payload);
    if (!payload || !payload.ticketId || !payload.remarks) {
      return JSON.stringify({ success: false, message: "Ticket ID & remarks required." });
    }

    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });

    var myRoles = profile.allRoles || [profile.role];
    var sys = getSystemEmailConfig();
    var isITCheck = myRoles.indexOf("IT_CHECK") !== -1 || (sys.itCheckEmail && cleanEmailAddress(authEmail) === cleanEmailAddress(sys.itCheckEmail));
    if (!isITCheck) return JSON.stringify({ success: false, message: "Only IT Check can reject at this stage." });

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet not found." });

    var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var targetRowValues = null, targetRowIndex = -1;

    for (var i = 0; i < allData.length; i++) {
      var rData = {};
      for (var h = 0; h < headers.length; h++) {
        var cleanH = headers[h].toString().trim();
        if (!rData.hasOwnProperty(cleanH)) rData[cleanH] = allData[i][h];
      }
      if ((rData["Ticket ID"] || "").toString().trim().toUpperCase() === payload.ticketId.toString().trim().toUpperCase()) {
        targetRowValues = allData[i]; targetRowIndex = i + 2; break;
      }
    }

    if (!targetRowValues) return JSON.stringify({ success: false, message: "Ticket not found." });

    var state = determineTicketWorkflowState(targetRowValues, headers, targetRowIndex);
    if (state.stage !== "IT_CHECK") return JSON.stringify({ success: false, message: "Not at IT Check stage." });

    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");
    updateSingleColStatus(sheet, targetRowIndex, headers, "Status", "Rejected at IT Check");
    updateSingleColStatus(sheet, targetRowIndex, headers, "IT Check (Execution)", profile.name || authEmail);
    updateSingleColStatus(sheet, targetRowIndex, headers, "IT Check Date", nowStr);

    var rowData = {};
    for (var i2 = 0; i2 < headers.length; i2++) {
      rowData[headers[i2].toString().trim()] = targetRowValues[i2];
    }
    var emailUser = extractUserEmailFromRow(targetRowValues);
    var namaUser = rowData["Full Name"] || rowData["Nama Penuh"] || "Staff";
    var assetName = rowData["Asset Type"] || rowData["Jenis Aset"] || "N/A";
    var assetQty = rowData["Quantity"] || rowData["Kuantiti"] || "1";
    var hierarchy = getApprovalHierarchy(emailUser);

    var chainData = buildChainDataStructure(hierarchy, emailUser, state.isSimpleAsset, rowData);
    markChainRejected(chainData, "IT CHECK", nowStr);
    var pdfBlob = buildPdfForm(rowData, payload.ticketId, chainData);
    var recipients = [emailUser, sys.itCheckEmail].filter(Boolean).join(",");

    sendSafeEmail({
      to: recipients, name: "Apollo IT Portal",
      subject: "[REJECTED] IT Asset Request " + payload.ticketId + " (" + namaUser + ")",
      htmlBody: buildUniformEmailHtml(
        "REQUEST REJECTED BY IT CHECK",
        "Dear <b>" + namaUser + "</b>,",
        namaUser, payload.ticketId, assetName, assetQty,
        "Standard Asset (Rejected)", null, null, rowData,
        { introText: "Your IT Asset Request was rejected by IT Check.", headerBg: "#c62828", remarks: payload.remarks }
      ),
      attachments: [pdfBlob]
    });

    return JSON.stringify({ success: true, data: { message: "Ticket rejected." } });
  } catch (err) {
    Logger.log("itCheckRejectTicket error: " + err.toString());
    return JSON.stringify({ success: false, message: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ═══════════════════════════════════════════════════════════════
// ON FORM SUBMIT
// ═══════════════════════════════════════════════════════════════
function onFormSubmit(e) {
  var sheet = getTargetSheet();
  var lastRow = (e && e.range) ? e.range.getRow() : sheet.getLastRow();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowValues = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var sys = getSystemEmailConfig();

  var rowData = {};
  for (var i = 0; i < headers.length; i++) {
    var cleanHeader = headers[i].toString().trim();
    if (rowData.hasOwnProperty(cleanHeader)) rowData[cleanHeader + "_" + i] = rowValues[i];
    else rowData[cleanHeader] = rowValues[i];
  }

  var ticketID = generateTicketID(sheet);
  var ticketCol = getOrCreateCol(sheet, headers, "Ticket ID");
  sheet.getRange(lastRow, ticketCol).setValue(ticketID);
  rowData["Ticket ID"] = ticketID;

  var assetType = (rowData["Asset Type"] || rowData["Jenis Aset"] || "").toString().toLowerCase();
  var isSimpleAsset = (assetType.indexOf("mouse") !== -1 || assetType.indexOf("keyboard") !== -1 || assetType.indexOf("monitor") !== -1);

  var emailUser = extractUserEmailFromRow(rowValues);
  var namaUser = rowData["Full Name"] || rowData["Nama Penuh"] || "Staff";
  var hierarchy = getApprovalHierarchy(emailUser);

  initSheetStatus(sheet, lastRow, headers, isSimpleAsset, hierarchy);

  var firstTargetEmail = sys.itVerifiedEmail || sys.itCheckEmail || sys.itGroupEmail;
  var firstStepTitle = "IT STEP: INITIAL IT VERIFICATION REQUIRED";

  updateSingleColStatus(sheet, lastRow, headers, "Status", "Pending IT Verification");

  var chainData = buildChainDataStructure(hierarchy, emailUser, isSimpleAsset, rowData);
  var initialPdf = buildPdfForm(rowData, ticketID, chainData);
  var firstApproverName = getNameByEmail(firstTargetEmail);

  sendSafeEmail({
    to: firstTargetEmail,
    name: "Apollo IT Portal",
    subject: firstStepTitle + ": " + ticketID + " (" + namaUser + ")",
    htmlBody: buildUniformEmailHtml(
      firstStepTitle,
      "Dear <b>" + (firstApproverName !== "-" ? firstApproverName : "IT Verifier") + "</b>,",
      namaUser, ticketID,
      rowData["Asset Type"] || rowData["Jenis Aset"] || "N/A",
      rowData["Quantity"] || rowData["Kuantiti"] || "1",
      isSimpleAsset ? "Simple Asset (Short Route)" : "Standard Asset Request",
      WEB_APP_URL, null, rowData
    ),
    attachments: [initialPdf]
  });

  if (emailUser && emailUser.toString().trim() !== "") {
    sendSafeEmail({
      to: emailUser,
      name: "Apollo IT Portal",
      subject: "Request Received [" + ticketID + "] - Pending IT Verification",
      htmlBody: buildUniformEmailHtml(
        "REQUEST RECEIVED & UNDER INITIAL REVIEW",
        "Dear <b>" + namaUser + "</b>,", namaUser, ticketID,
        rowData["Asset Type"] || rowData["Jenis Aset"] || "N/A",
        rowData["Quantity"] || rowData["Kuantiti"] || "1",
        isSimpleAsset ? "Simple Asset Request" : "Standard Asset Request",
        null, null, rowData,
        { introText: "Your IT Asset Request has been received. No action is required from you at this time.", headerBg: "#0284c7" }
      )
    });
  }
}

// ═══════════════════════════════════════════════════════════════
// WEB APP ROUTING
// ═══════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    var targetPage = (e && e.parameter) ? (e.parameter.page || e.parameter.view) : null;

    if (!targetPage || targetPage === 'login') {
      return HtmlService.createHtmlOutputFromFile('Login_Template')
        .setTitle('Apollo IT Portal - Login')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    if (e && e.parameter && e.parameter.action) {
      var action = e.parameter.action;
      if (action === "md_approve") {
        return HtmlService.createHtmlOutput(
          "<div style='font-family:Arial;text-align:center;padding:60px 20px;background:#f5f5f5;'>" +
          "<h2 style='color:#7c3aed;'>MD Review Required</h2>" +
          "<p style='color:#555;font-size:14px;'>Please log in to the MD Portal to review quotations and select.</p>" +
          "<a href='" + WEB_APP_URL + "?page=md' style='display:inline-block;margin-top:20px;padding:12px 24px;background:#7c3aed;color:#fff;text-decoration:none;border-radius:6px;font-weight:bold;'>Open MD Portal</a>" +
          "</div>"
        );
      }
      var ticketId = e.parameter.ticketId;
      var row = parseInt(e.parameter.row, 10);
      var isSimple = (e.parameter.isSimple === "true");
      var stage = e.parameter.stage || "Approval Process";
      var remarks = e.parameter.remarks || "";
      var result = executeApprovalAction(action, ticketId, row, isSimple, stage, remarks);
      if (result && result.pdfBlob) {
        return renderAutoDownloadHtml(result.pdfBlob, result.title, result.approverName, ticketId, result.message);
      } else {
        return HtmlService.createHtmlOutput("<div style='font-family:Arial;text-align:center;padding:40px;'><h3>" + (result.message || "Action processed.") + "</h3></div>");
      }
    }

    var pageMap = {
      'approval': { file: 'Portal_Approval', title: 'Apollo IT Asset Approval Portal' },
      'verify':   { file: 'IT_Verification_Portal_Template', title: 'Apollo IT Verification Portal' },
      'it':       { file: 'IT_Execution_Portal_Template', title: 'Apollo IT Execution Portal' },
      'md':       { file: 'Executive_MD_Portal_Template', title: 'Apollo IT Executive MD Portal' },
      'admin':    { file: 'IT_AdminPortal_Template', title: 'Apollo IT Admin Portal' }
    };

    if (pageMap[targetPage]) {
      var entry = pageMap[targetPage];
      try {
        return HtmlService.createHtmlOutputFromFile(entry.file)
          .setTitle(entry.title)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      } catch (fileErr) {
        return HtmlService.createHtmlOutput(
          "<div style='font-family:Arial;padding:40px;text-align:center;color:#c62828;'><h2>Template Not Found</h2><p>File <b>" + entry.file + "</b> does not exist.</p></div>"
        );
      }
    }

    return HtmlService.createHtmlOutputFromFile('Login_Template')
      .setTitle('Apollo IT Portal - Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput("<div style='font-family:Arial;padding:40px;color:#c62828;text-align:center;'><h3>Error</h3><pre>" + err.toString() + "</pre></div>");
  }
}

function getAuthenticatedUserEmail() {
  var email = "";
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!email || email.toString().trim() === "") {
    try { email = Session.getEffectiveUser().getEmail(); } catch (e2) {}
  }
  return cleanEmailAddress(email);
}

function checkActiveUserProfileByEmail(userEmail) {
  if (!userEmail) return JSON.stringify({ success: false, message: "Email required." });

  var cleanEmail = cleanEmailAddress(userEmail);
  var profile = getApproverProfile(cleanEmail);

  if (!profile) {
    return JSON.stringify({ success: false, message: "Your account (" + cleanEmail + ") is not registered in the system." });
  }

  var redirectPage = 'approval';
  var availablePages = [];
  
  var isITCheck = (profile.allRoles && profile.allRoles.indexOf("IT_CHECK") !== -1);
  var isITVerified = (profile.allRoles && profile.allRoles.indexOf("IT_VERIFIED") !== -1);
  var isITHod = (profile.role === 'IT_HOD');
  var isMD = (profile.role === 'MD');

  // KAWAL DUAL-ROLE DI SINI
  if (cleanEmail === 'it@apollofood.com.my') {
    redirectPage = 'verify';
    availablePages = [
      { page: 'verify', label: 'IT Verification Portal (Initial Review)' },
      { page: 'it', label: 'IT Execution Portal (Quotation & Issue)' }
    ];
  } else if (cleanEmail === 'amira.ayob@apollofood.com.my') {
    redirectPage = 'approval';
    availablePages = [
      { page: 'approval', label: 'HOD Approval Portal (Department Approver)' },
      { page: 'admin', label: 'IT Admin Portal (Monitoring & Procurement)' }
    ];
  } else {
    if (isMD) redirectPage = 'md';
    else if (isITVerified && !isITCheck) redirectPage = 'verify';
    else if (isITCheck) redirectPage = 'it';
    else if (isITHod) redirectPage = 'admin';
  }

  return JSON.stringify({ 
    success: true, 
    data: { 
      profile: profile, 
      redirectPage: redirectPage,
      availablePages: availablePages 
    } 
  });
}

function getApproverProfile(userEmail) {
  if (!userEmail) return null;
  var cleanEmail = cleanEmailAddress(userEmail);
  initDirectoryCache();

  if (_memCache.approverProfiles && _memCache.approverProfiles[cleanEmail] !== undefined) {
    return _memCache.approverProfiles[cleanEmail];
  }

  var ROLE_RANKS = { "MANAGER": 1, "HOD": 2, "CSUITE": 3, "MD": 4, "IT_CHECK": 5, "IT_HOD": 6, "IT_VERIFIED": 7 };
  var detectedRoles = [];
  var officialName = null;

  var accessList = _memCache.dashboardAccessList || [];
  for (var i = 0; i < accessList.length; i++) {
    if (cleanEmailAddress(accessList[i].email) === cleanEmail) {
      if (accessList[i].name) officialName = accessList[i].name;
      var rawRole = (accessList[i].role || "MANAGER").toUpperCase();
      if (rawRole.indexOf("IT_VERIFIED") !== -1) detectedRoles.push("IT_VERIFIED");
      else if (rawRole.indexOf("IT_CHECK") !== -1) detectedRoles.push("IT_CHECK");
      else if (rawRole.indexOf("IT_HOD") !== -1) detectedRoles.push("IT_HOD");
      else if (rawRole.indexOf("MD") !== -1) detectedRoles.push("MD");
      else if (rawRole.indexOf("CSUITE") !== -1 || rawRole.indexOf("C-SUITE") !== -1) detectedRoles.push("CSUITE");
      else if (rawRole.indexOf("HOD") !== -1) detectedRoles.push("HOD");
      else if (rawRole.indexOf("MANAGER") !== -1) detectedRoles.push("MANAGER");
    }
  }

  if (detectedRoles.length === 0) {
    var routeRows = _memCache.routeRows || [];
    for (var j = 0; j < routeRows.length; j++) {
      var rObj = routeRows[j];
      if (rObj.mgrRaw) { var mE = getEmailByName(rObj.mgrRaw); if (mE && mE.toLowerCase() === cleanEmail) { detectedRoles.push("MANAGER"); if (!officialName) officialName = rObj.mgrRaw; } }
      if (rObj.hodRaw) { var hE = getEmailByName(rObj.hodRaw); if (hE && hE.toLowerCase() === cleanEmail) { detectedRoles.push("HOD"); if (!officialName) officialName = rObj.hodRaw; } }
      if (rObj.csRaw) { var cE = getEmailByName(rObj.csRaw); if (cE && cE.toLowerCase() === cleanEmail) { detectedRoles.push("CSUITE"); if (!officialName) officialName = rObj.csRaw; } }
    }
  }

  if (detectedRoles.length === 0) {
    if (_memCache.approverProfiles) _memCache.approverProfiles[cleanEmail] = null;
    return null;
  }

  var highestRole = "MANAGER";
  var maxRank = 0;
  for (var r = 0; r < detectedRoles.length; r++) {
    var rank = ROLE_RANKS[detectedRoles[r]] || 1;
    if (rank > maxRank) { maxRank = rank; highestRole = detectedRoles[r]; }
  }

  if (!officialName) officialName = getNameByEmail(cleanEmail);

  var profile = {
    isApprover: true, role: highestRole, allRoles: detectedRoles,
    name: (officialName && officialName !== cleanEmail && officialName !== "-") ? officialName : (highestRole + " Officer"),
    email: cleanEmail, routeSummary: highestRole + " Portal Access Verified"
  };

  if (_memCache.approverProfiles) _memCache.approverProfiles[cleanEmail] = profile;
  return profile;
}

// ═══════════════════════════════════════════════════════════════
// EXECUTE APPROVAL ACTION
// ═══════════════════════════════════════════════════════════════
function executeApprovalAction(action, ticketId, row, isSimple, stage, remarks) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) {
      return { success: false, pdfBlob: null, title: "SYSTEM BUSY", approverName: "SYSTEM", ticketId: ticketId, message: "System busy - please try again shortly." };
    }

    var sheet = getTargetSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    var sys = getSystemEmailConfig();

    var rowData = {};
    for (var i = 0; i < headers.length; i++) {
      var cleanHeader = headers[i].toString().trim();
      if (rowData.hasOwnProperty(cleanHeader)) rowData[cleanHeader + "_" + i] = rowValues[i];
      else rowData[cleanHeader] = rowValues[i];
    }

    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");
    var emailUser = extractUserEmailFromRow(rowValues);
    var namaUser = rowData["Full Name"] || rowData["Nama Penuh"] || "Staff";
    var assetName = rowData["Asset Type"] || rowData["Jenis Aset"] || "N/A";
    var assetQty = rowData["Quantity"] || rowData["Kuantiti"] || "1";
    var hierarchy = getApprovalHierarchy(emailUser);

    if (action === "it_verify_approve") {
      var existingITVerify = (rowData["IT Verification Date"] || "").toString().trim();
      if (existingITVerify && existingITVerify !== "Pending" && existingITVerify !== "-") {
        return { success: true, pdfBlob: null, title: "ALREADY VERIFIED", approverName: "SYSTEM", ticketId: ticketId, message: "Already verified. Duplicate click blocked." };
      }
    }
    if (action === "manager_approve") {
      var existingMgr = (rowData["Manager Approval Date"] || "").toString().trim();
      if (existingMgr && existingMgr !== "Pending" && existingMgr !== "-" && existingMgr.indexOf("Bypassed") === -1) {
        return { success: true, pdfBlob: null, title: "ALREADY APPROVED", approverName: "SYSTEM", ticketId: ticketId, message: "Already approved. Duplicate click blocked." };
      }
    }
    if (action === "hod_approve") {
      var existingHod = (rowData["HOD Approval Date"] || "").toString().trim();
      if (existingHod && existingHod !== "Pending" && existingHod !== "-" && existingHod.indexOf("Bypassed") === -1) {
        return { success: true, pdfBlob: null, title: "ALREADY APPROVED", approverName: "SYSTEM", ticketId: ticketId, message: "Already approved. Duplicate click blocked." };
      }
    }
    if (action === "csuite_approve") {
      var existingCs = (rowData["C-Suite Approval Date"] || "").toString().trim();
      if (existingCs && existingCs !== "Pending" && existingCs !== "-" && existingCs.indexOf("Bypassed") === -1) {
        return { success: true, pdfBlob: null, title: "ALREADY APPROVED", approverName: "SYSTEM", ticketId: ticketId, message: "Already approved. Duplicate click blocked." };
      }
    }
    if (action === "it_check_approve") {
      var existingITCheck = (rowData["IT Check Date"] || "").toString().trim();
      if (existingITCheck && existingITCheck !== "Pending" && existingITCheck !== "-") {
        return { success: true, pdfBlob: null, title: "ALREADY CHECKED", approverName: "SYSTEM", ticketId: ticketId, message: "Already verified. Duplicate click blocked." };
      }
    }

    if (action === "reject") {
      var currentStatus = (rowData["Status"] || "").toString().trim();
      if (currentStatus.indexOf("Rejected") !== -1 || currentStatus.indexOf("Cancelled") !== -1 || currentStatus.indexOf("Approved & Completed") !== -1) {
        return { success: true, pdfBlob: null, title: "ALREADY FINALIZED", approverName: "SYSTEM", ticketId: ticketId, message: "Ticket already finalized - reject cannot be repeated." };
      }

      var rejectStage = cleanStageName(stage);
      updateSingleColStatus(sheet, row, headers, "Status", "Rejected at " + rejectStage);
      var rejectChain = buildChainDataStructure(hierarchy, emailUser, isSimple, rowData);
      markChainRejected(rejectChain, stage || "Approval Process", nowStr);
      var rejectedPdf = buildPdfForm(rowData, ticketId, rejectChain);
      var recipients = emailUser ? (emailUser + "," + (sys.itGroupEmail || sys.itCheckEmail)) : (sys.itGroupEmail || sys.itCheckEmail);
      
      sendSafeEmail({
        to: recipients, name: "Apollo IT Portal",
        subject: "[REJECTED] IT Asset Request " + ticketId + " (" + namaUser + ")",
        htmlBody: buildUniformEmailHtml("REQUEST REJECTED: " + ticketId, "Dear <b>" + namaUser + " & IT Team</b>,",
          namaUser, ticketId, assetName, assetQty, isSimple ? "Simple Asset" : "Standard Asset", null, null, rowData,
          { introText: "REJECTED at stage: <b>" + rejectStage + "</b>.", headerBg: "#c62828", remarks: remarks }),
        attachments: [rejectedPdf]
      });
      return { success: true, pdfBlob: rejectedPdf, title: "REQUEST REJECTED", approverName: "SYSTEM", ticketId: ticketId, message: "Rejected." };
    }

    if (action === "it_verify_approve") {
      var verifierName = sys.itVerifiedName || "IT VERIFIER";
      updateSingleColStatus(sheet, row, headers, "IT Verification", verifierName);
      updateSingleColStatus(sheet, row, headers, "IT Verification Date", nowStr);

      var nextTarget = "", nextAction = "", nextTitle = "";
      if (hierarchy.managerEmail && hierarchy.managerEmail !== emailUser) {
        nextTarget = hierarchy.managerEmail; nextAction = "manager_approve"; nextTitle = "STEP 1: MANAGER APPROVAL REQUIRED";
      } else if (hierarchy.hodEmail && hierarchy.hodEmail !== emailUser) {
        nextTarget = hierarchy.hodEmail; nextAction = "hod_approve"; nextTitle = "STEP 1: HOD APPROVAL REQUIRED";
      } else if (hierarchy.csuiteEmail && hierarchy.csuiteEmail !== emailUser) {
        nextTarget = hierarchy.csuiteEmail; nextAction = "csuite_approve"; nextTitle = "STEP 1: C-SUITE APPROVAL REQUIRED";
      } else {
        nextTarget = sys.itCheckEmail; nextAction = "it_check_approve"; nextTitle = "IT STEP: IT CHECK & VERIFICATION REQUIRED";
      }

      updateSingleColStatus(sheet, row, headers, "Status", "Verified by IT ➔ Pending Route Approval");

      var chainData = buildChainDataStructure(hierarchy, emailUser, isSimple, rowData);
      var pdfBlob = buildPdfForm(rowData, ticketId, chainData);
      var nextApproverName = getNameByEmail(nextTarget);

      sendSafeEmail({
        to: nextTarget, name: "Apollo IT Portal",
        subject: nextTitle + ": " + ticketId + " (" + namaUser + ")",
        htmlBody: buildUniformEmailHtml(nextTitle, "Dear <b>" + (nextApproverName !== "-" ? nextApproverName : "Approver") + "</b>,",
          namaUser, ticketId, assetName, assetQty, isSimple ? "Simple Asset Request" : "Standard Asset Request",
          WEB_APP_URL, null, rowData),
        attachments: [pdfBlob]
      });

      return { success: true, pdfBlob: pdfBlob, title: "IT INITIAL VERIFICATION COMPLETED", approverName: verifierName, ticketId: ticketId, message: "Forwarded to " + nextTarget };
    }

    if (action === "manager_approve") {
      var mgrName = hierarchy.managerName || "MANAGER";
      updateSingleColStatus(sheet, row, headers, "Manager Approver", mgrName);
      updateSingleColStatus(sheet, row, headers, "Manager Approval Date", nowStr);
      updateSingleColStatus(sheet, row, headers, "Status", "Approved by Manager");

      var nextAction, nextTarget, nextTitle;
      if (hierarchy.hodEmail && hierarchy.hodEmail !== emailUser) {
        nextAction = "hod_approve"; nextTarget = hierarchy.hodEmail; nextTitle = "STEP 2: HOD APPROVAL REQUIRED";
      } else if (hierarchy.csuiteEmail && hierarchy.csuiteEmail !== emailUser) {
        nextAction = "csuite_approve"; nextTarget = hierarchy.csuiteEmail; nextTitle = "STEP 2: C-SUITE APPROVAL REQUIRED";
      } else {
        nextAction = "it_check_approve"; nextTarget = sys.itCheckEmail; nextTitle = "IT STEP: IT CHECK & VERIFICATION REQUIRED";
      }

      var chainData = buildChainDataStructure(hierarchy, emailUser, isSimple, rowData);
      syncStepStatus(chainData, "MANAGER APPROVAL", "Approved via Email", nowStr);
      var pdfBlob = buildPdfForm(rowData, ticketId, chainData);
      var nextApproverName = getNameByEmail(nextTarget);

      sendSafeEmail({
        to: nextTarget, name: "Apollo IT Portal",
        subject: nextTitle + ": " + ticketId + " (" + namaUser + ")",
        htmlBody: buildUniformEmailHtml(nextTitle, "Dear <b>" + (nextApproverName !== "-" ? nextApproverName : "Approver") + "</b>,",
          namaUser, ticketId, assetName, assetQty, isSimple ? "Simple Asset Request" : "Standard Asset Request",
          WEB_APP_URL, null, rowData),
        attachments: [pdfBlob]
      });

      return { success: true, pdfBlob: pdfBlob, title: "MANAGER APPROVAL COMPLETED", approverName: mgrName, ticketId: ticketId, message: "PDF downloading..." };
    }

    if (action === "hod_approve") {
      var hName = hierarchy.hodName || "HOD";
      updateSingleColStatus(sheet, row, headers, "HOD Approver", hName);
      updateSingleColStatus(sheet, row, headers, "HOD Approval Date", nowStr);
      updateSingleColStatus(sheet, row, headers, "Status", "Approved by HOD");

      var chainDataHod = buildChainDataStructure(hierarchy, emailUser, isSimple, rowData);
      syncStepStatus(chainDataHod, "HOD APPROVAL", "Approved via Email", nowStr);

      var nextActionStd, nextTargetStd, nextTitleStd;
      if (hierarchy.csuiteEmail && hierarchy.csuiteEmail !== emailUser) {
        nextActionStd = "csuite_approve"; nextTargetStd = hierarchy.csuiteEmail; nextTitleStd = "STEP 3: C-SUITE APPROVAL REQUIRED";
      } else {
        nextActionStd = "it_check_approve"; nextTargetStd = sys.itCheckEmail; nextTitleStd = "IT STEP: IT CHECK & VERIFICATION REQUIRED";
      }

      var pdfStdBlob = buildPdfForm(rowData, ticketId, chainDataHod);
      var nextTargetName = getNameByEmail(nextTargetStd);

      sendSafeEmail({
        to: nextTargetStd, name: "Apollo IT Portal",
        subject: nextTitleStd + ": " + ticketId + " (" + namaUser + ")",
        htmlBody: buildUniformEmailHtml(nextTitleStd, "Dear <b>" + (nextTargetName !== "-" ? nextTargetName : "Approver") + "</b>,",
          namaUser, ticketId, assetName, assetQty, isSimple ? "Simple Asset Request" : "Standard Asset Request",
          WEB_APP_URL, null, rowData),
        attachments: [pdfStdBlob]
      });

      return { success: true, pdfBlob: pdfStdBlob, title: "HOD APPROVAL COMPLETED", approverName: hName, ticketId: ticketId, message: "PDF downloading..." };
    }

    if (action === "csuite_approve") {
      var csName = hierarchy.csuiteName || "C-SUITE";
      updateSingleColStatus(sheet, row, headers, "C-Suite Approver", csName);
      updateSingleColStatus(sheet, row, headers, "C-Suite Approval Date", nowStr);
      updateSingleColStatus(sheet, row, headers, "Status", "Pending IT Check");
      updateSingleColStatus(sheet, row, headers, "IT Check (Execution)", "Pending");
      updateSingleColStatus(sheet, row, headers, "IT Check Date", "Pending");

      var chainDataCs = buildChainDataStructure(hierarchy, emailUser, isSimple, rowData);
      syncStepStatus(chainDataCs, "C-SUITE APPROVAL", "Approved via Email", nowStr);

      var pdfCsBlob = buildPdfForm(rowData, ticketId, chainDataCs);

      sendSafeEmail({
        to: sys.itCheckEmail, name: "Apollo IT Portal",
        subject: "[IT STEP] IT Check Required: " + ticketId + " (" + namaUser + ")",
        htmlBody: buildUniformEmailHtml("IT CHECK & VERIFICATION REQUIRED", "Dear <b>ADMIN IT</b>,",
          namaUser, ticketId, assetName, assetQty,
          isSimple ? "Simple Asset (C-Suite Approved)" : "Standard Asset (C-Suite Approved)",
          WEB_APP_URL, null, rowData,
          { introText: "Please log in to the IT Execution Portal to review the ticket, prepare quotation(s), and submit to MD.", headerBg: "#7c3aed" }),
        attachments: [pdfCsBlob]
      });

      return { success: true, pdfBlob: pdfCsBlob, title: "C-SUITE APPROVAL COMPLETED", approverName: csName, ticketId: ticketId, message: "Forwarded to IT Check..." };
    }

    if (action === "it_check_approve") {
      var itCheckName = sys.itCheckName;
      updateSingleColStatus(sheet, row, headers, "IT Check (Execution)", itCheckName);
      updateSingleColStatus(sheet, row, headers, "IT Check Date", nowStr);

      var chainDataItCheck = buildChainDataStructure(hierarchy, emailUser, isSimple, rowData);
      syncStepStatus(chainDataItCheck, "IT CHECK", "Verified via Email", nowStr);

      if (isSimple) {
        updateSingleColStatus(sheet, row, headers, "Status", "Approved & Completed");
        updateSingleColStatus(sheet, row, headers, "IT HOD Verified", sys.itHodName);
        updateSingleColStatus(sheet, row, headers, "IT HOD Date", nowStr);
        var pdfSimpleCompleted = buildPdfForm(rowData, ticketId, chainDataItCheck);
        var recipientsSimple = [sys.itCheckEmail, sys.itGroupEmail].filter(Boolean).join(",");

        sendSafeEmail({
          to: recipientsSimple, name: "Apollo IT Portal",
          subject: "[COMPLETED] Simple Asset Request: " + ticketId + " (" + namaUser + ")",
          htmlBody: buildUniformEmailHtml("REQUEST FULLY APPROVED & COMPLETED", "Dear <b>IT Team</b>,",
            namaUser, ticketId, assetName, assetQty, "Simple Asset (Completed)", null, null, rowData,
            { introText: "Simple IT Asset Request fully verified.", headerBg: "#2e7d32" }),
          attachments: [pdfSimpleCompleted]
        });
        return { success: true, pdfBlob: pdfSimpleCompleted, title: "IT CHECK VERIFIED & COMPLETED", approverName: itCheckName, ticketId: ticketId, message: "Completed" };
      } else {
        return {
          success: true, pdfBlob: null, title: "IT CHECK", approverName: itCheckName, ticketId: ticketId,
          message: "Please log in to the IT Execution Portal to add quotation(s) and submit to MD."
        };
      }
    }

    if (action === "md_approve") {
      return {
        success: true, pdfBlob: null, title: "MD REVIEW", approverName: sys.mdName, ticketId: ticketId,
        message: "Please log in to the MD Portal to review and select quotation."
      };
    }

    throw new Error("Unrecognized action: " + action);

  } catch (err) {
    Logger.log("executeApprovalAction error: " + err.toString());
    return { success: false, pdfBlob: null, title: "ERROR", approverName: "SYSTEM", ticketId: ticketId, message: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ═══════════════════════════════════════════════════════════════
// BUILD CHAIN DATA
// ═══════════════════════════════════════════════════════════════
function buildChainDataStructure(hierarchy, emailUser, isSimpleAsset, rowData) {
  var chain = [];
  rowData = rowData || {};
  var sys = getSystemEmailConfig();

  if (hierarchy.managerName && hierarchy.managerName !== "-" && hierarchy.managerEmail) {
    var isSelfMgr = (hierarchy.managerEmail === emailUser);
    var mgrDate = rowData["Manager Approval Date"];
    var isMgrApproved = (mgrDate && mgrDate !== "Pending" && mgrDate !== "-" && mgrDate.toString().indexOf("Bypassed") === -1);
    chain.push({
      role: "MANAGER APPROVAL", name: hierarchy.managerName, email: hierarchy.managerEmail,
      status: isMgrApproved ? "Approved via Email" : (isSelfMgr ? "Submitted (Self)" : "Pending"),
      date: isMgrApproved ? mgrDate : "-"
    });
  }

  if (hierarchy.hodName && hierarchy.hodName !== "-" && hierarchy.hodEmail) {
    var isSelfHod = (hierarchy.hodEmail === emailUser);
    var hodDate = rowData["HOD Approval Date"];
    var isHodApproved = (hodDate && hodDate !== "Pending" && hodDate !== "-" && hodDate.toString().indexOf("Bypassed") === -1);
    chain.push({
      role: "HOD APPROVAL", name: hierarchy.hodName, email: hierarchy.hodEmail,
      status: isHodApproved ? "Approved via Email" : (isSelfHod ? "Submitted (Self)" : "Pending"),
      date: isHodApproved ? hodDate : "-"
    });
  }

  if (hierarchy.csuiteName && hierarchy.csuiteName !== "-" && hierarchy.csuiteEmail) {
    var isSelfCs = (hierarchy.csuiteEmail === emailUser);
    var csDate = rowData["C-Suite Approval Date"];
    var isCsApproved = (csDate && csDate !== "Pending" && csDate !== "-" && csDate.toString().indexOf("Bypassed") === -1);
    chain.push({
      role: "C-SUITE APPROVAL", name: hierarchy.csuiteName, email: hierarchy.csuiteEmail,
      status: isCsApproved ? "Approved via Email" : (isSelfCs ? "Submitted (Self)" : "Pending"),
      date: isCsApproved ? csDate : "-"
    });
  }

  var itCheckDate = rowData["IT Check Date"];
  var isItCheckApproved = (itCheckDate && itCheckDate !== "Pending" && itCheckDate !== "-");
  chain.push({
    role: "IT CHECK", name: sys.itCheckName, email: sys.itCheckEmail,
    status: isItCheckApproved ? "Verified via Email" : "Pending",
    date: isItCheckApproved ? itCheckDate : "-"
  });

  if (!isSimpleAsset && hierarchy.mdName && hierarchy.mdName !== "-" && hierarchy.mdEmail) {
    var mdDate = rowData["MD Approval Date"];
    var isMdApproved = (mdDate && mdDate !== "Pending" && mdDate !== "-" && mdDate.toString().indexOf("Bypassed") === -1);
    var isSelfMd = (hierarchy.mdEmail === emailUser);
    chain.push({
      role: "MANAGING DIRECTOR (MD)", name: hierarchy.mdName || sys.mdName, email: hierarchy.mdEmail || sys.mdEmail,
      status: isMdApproved ? "Approved via Email" : (isSelfMd ? "Submitted (Self)" : "Pending"),
      date: isMdApproved ? mdDate : "-"
    });
  }

  return { isSimple: isSimpleAsset, steps: chain };
}

function syncStepStatus(chainObj, targetRole, newStatus, dateStr) {
  if (!chainObj || !chainObj.steps) return;
  for (var i = 0; i < chainObj.steps.length; i++) {
    if (chainObj.steps[i].role.toUpperCase() === targetRole.toUpperCase()) {
      chainObj.steps[i].status = newStatus;
      chainObj.steps[i].date = dateStr;
      break;
    }
  }
}

function markChainRejected(chainObj, stageTitle, dateStr) {
  var upperStage = stageTitle.toUpperCase();
  var found = false;
  for (var i = 0; i < chainObj.steps.length; i++) {
    var step = chainObj.steps[i];
    var stepRole = step.role.toUpperCase();
    var isMatch = false;
    if (upperStage.indexOf("IT CHECK") !== -1 && stepRole.indexOf("IT CHECK") !== -1) isMatch = true;
    else if ((upperStage.indexOf("MD") !== -1 || upperStage.indexOf("MANAGING DIRECTOR") !== -1) && stepRole.indexOf("MD") !== -1) isMatch = true;
    else if (upperStage.indexOf("C-SUITE") !== -1 && stepRole.indexOf("C-SUITE") !== -1) isMatch = true;
    else if (upperStage.indexOf("HOD") !== -1 && stepRole === "HOD APPROVAL") isMatch = true;
    else if (upperStage.indexOf("MANAGER") !== -1 && stepRole === "MANAGER APPROVAL") isMatch = true;
    if (isMatch && !found) { step.status = "REJECTED"; step.date = dateStr; found = true; }
  }
  if (!found) {
    for (var j = 0; j < chainObj.steps.length; j++) {
      if (chainObj.steps[j].status === "Pending") { chainObj.steps[j].status = "REJECTED"; chainObj.steps[j].date = dateStr; break; }
    }
  }
}

// ═══════════════════════════════════════════════════════════════
// EMAIL TEMPLATE RENDERER
// ═══════════════════════════════════════════════════════════════
function buildUniformEmailHtml(stepTitle, greetingTitle, staffName, ticketId, assetName, qty, routeText, portalUrl, unusedRejectUrl, extraData, customConfig) {
  var htmlTemplate = HtmlService.createTemplateFromFile('Email_Template');
  extraData = extraData || {};
  customConfig = customConfig || {};
  htmlTemplate.stepTitle = stepTitle || "IT ASSET REQUEST";
  var cleanGreeting = (greetingTitle || "Dear Approver,").toString().replace(/<\/?p>/gi, "");
  htmlTemplate.greetingTitle = cleanGreeting;
  htmlTemplate.staffName = staffName || "-";
  htmlTemplate.ticketId = ticketId || "-";
  htmlTemplate.assetName = assetName || "N/A";
  htmlTemplate.qty = qty || "1";
  htmlTemplate.routeText = routeText || "-";
  htmlTemplate.portalUrl = portalUrl || WEB_APP_URL;
  htmlTemplate.department = String(extraData.department || extraData.Department || extraData.Jabatan || '-');
  var rawJustification = extraData.justification || extraData["Justification for Request"] || extraData["Justification for request"] || '-';
  htmlTemplate.justification = String(rawJustification);
  var rawReason = extraData.reason || extraData["Reason / Justification"] || extraData.Reason || '-';
  htmlTemplate.reason = String(rawReason);
  htmlTemplate.attachment = String(extraData.attachment || extraData.Attachment || '-');
  htmlTemplate.introText = customConfig.introText || 'An IT Asset Request has been submitted:';
  htmlTemplate.headerBg = customConfig.headerBg || '#d32f2f';
  htmlTemplate.remarks = customConfig.remarks || null;
  return htmlTemplate.evaluate().getContent();
}

function renderAutoDownloadHtml(pdfBlob, title, approverName, ticketId, message) {
  var htmlTemplate = HtmlService.createTemplateFromFile('Download_PDF_Template');
  htmlTemplate.title = title;
  htmlTemplate.approverName = approverName;
  htmlTemplate.ticketId = ticketId;
  htmlTemplate.message = message;
  htmlTemplate.pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
  htmlTemplate.fileName = "IT_Asset_Request_" + ticketId + "_UPDATED.pdf";
  return htmlTemplate.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════════════════════════════════════════════════════
// SHEET UTILS
// ═══════════════════════════════════════════════════════════════
function generateTicketID(sheet) {
  if (!sheet) sheet = getTargetSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "ITA-0001";
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var ticketIdCol = getOrCreateCol(sheet, headers, "Ticket ID");
  var data = sheet.getRange(2, ticketIdCol, lastRow - 1, 1).getValues();
  var maxNum = 0;
  for (var i = 0; i < data.length; i++) {
    var id = data[i][0];
    if (id && typeof id === 'string') {
      var match = id.match(/(\d+)$/);
      if (match) { var num = parseInt(match[1], 10); if (num > maxNum) maxNum = num; }
    }
  }
  return "ITA-" + String(maxNum + 1).padStart(4, '0');
}

function initSheetStatus(sheet, row, headers, isSimpleAsset, hierarchy) {
  var hasMgr = hierarchy && hierarchy.managerEmail;
  var hasHod = hierarchy && hierarchy.hodEmail;
  var hasCs  = hierarchy && hierarchy.csuiteEmail;
  var hasMd  = !isSimpleAsset && hierarchy && hierarchy.mdEmail;

  updateSingleColStatus(sheet, row, headers, "Status", "Pending IT Verification");
  updateSingleColStatus(sheet, row, headers, "IT Verification", "Pending");
  updateSingleColStatus(sheet, row, headers, "IT Verification Date", "Pending");
  updateSingleColStatus(sheet, row, headers, "Manager Approver", hasMgr ? "Pending" : "Bypassed / Not Required");
  updateSingleColStatus(sheet, row, headers, "Manager Approval Date", hasMgr ? "Pending" : "-");
  updateSingleColStatus(sheet, row, headers, "HOD Approver", hasHod ? "Pending" : "Bypassed / Not Required");
  updateSingleColStatus(sheet, row, headers, "HOD Approval Date", hasHod ? "Pending" : "-");
  updateSingleColStatus(sheet, row, headers, "C-Suite Approver", hasCs ? "Pending" : "Bypassed / Not Required");
  updateSingleColStatus(sheet, row, headers, "C-Suite Approval Date", hasCs ? "Pending" : "-");
  updateSingleColStatus(sheet, row, headers, "MD Approver", isSimpleAsset ? "Bypassed (Simple Asset)" : (hasMd ? "Pending" : "Bypassed / Not Required"));
  updateSingleColStatus(sheet, row, headers, "MD Approval Date", (isSimpleAsset || !hasMd) ? "-" : "Pending");
  updateSingleColStatus(sheet, row, headers, "IT Check (Execution)", "Pending");
  updateSingleColStatus(sheet, row, headers, "IT Check Date", "Pending");
  updateSingleColStatus(sheet, row, headers, "IT HOD Verified", "Pending");
  updateSingleColStatus(sheet, row, headers, "IT HOD Date", "Pending");
}

function updateSingleColStatus(sheet, row, headers, colName, value) {
  var colIdx = getOrCreateCol(sheet, headers, colName);
  sheet.getRange(row, colIdx).setValue(value || "-");
}

function getOrCreateCol(sheet, headers, colName) {
  var idx = headers.indexOf(colName) + 1;
  if (idx === 0) {
    var lastCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, lastCol).setValue(colName);
    headers.push(colName);
    return lastCol;
  }
  return idx;
}

function getApolloLogoSrc() {
  try {
    var file = DriveApp.getFileById(LOGO_FILE_ID);
    var blob = file.getBlob();
    var bytes = blob.getBytes();
    if (bytes && bytes.length > 0) {
      return "data:" + (blob.getContentType() || "image/png") + ";base64," + Utilities.base64Encode(bytes);
    }
  } catch (e) {}
  return "https://drive.google.com/thumbnail?id=" + LOGO_FILE_ID + "&sz=w500";
}

function buildPdfForm(rowData, ticketID, chainData, signatureUrl) {
  var logoImgSrc = getApolloLogoSrc();
  var timestamp = rowData["Timestamp"] ? Utilities.formatDate(new Date(rowData["Timestamp"]), "GMT+8", "dd/MM/yyyy HH:mm:ss") : "-";

  var steps = chainData.steps || [];
  var numBoxes = steps.length;
  var colWidth = (100 / (numBoxes > 0 ? numBoxes : 1)).toFixed(2) + "%";

  var thHtml = "", tdHtml = "";
  for (var i = 0; i < steps.length; i++) {
    var step = steps[i];
    var statusText = step.status || "Pending";
    var statusClass = "badge-pending";
    if (statusText.indexOf("Approved") !== -1 || statusText.indexOf("Verified") !== -1) statusClass = "badge-approved";
    else if (statusText === "REJECTED") statusClass = "badge-rejected";
    else if (statusText.indexOf("Self") !== -1) statusClass = "badge-self";
    
    var sigHtml = "";
    if (step.role.toUpperCase().indexOf("MD") !== -1 && signatureUrl) {
      var fileId = extractDriveFileId(signatureUrl);
      if (fileId) {
        sigHtml = "<br><img src='https://drive.google.com/thumbnail?id=" + fileId + "&sz=w100' style='max-height:40px; max-width:100px; display:block; margin:4px auto;' alt='MD Signature' />";
      }
    }

    thHtml += "<th style='width: " + colWidth + ";'>" + step.role + "</th>";
    tdHtml += "<td class='sign-space'><b>Name:</b><br>" + (step.name || "-") + sigHtml + "<br><br><b>Status:</b><br><span class='status-badge " + statusClass + "'>" + statusText + "</span><br><br><b>Date:</b><br>" + (step.date || "-") + "</td>";
  }
  var boxesHtml = "<tr>" + thHtml + "</tr><tr>" + tdHtml + "</tr>";

  var justReqText = rowData["Justification for Request"] || rowData["Justification for request"] || "";
  var justReqRowHtml = "";
  if (justReqText && justReqText.toString().trim() !== "" && justReqText.toString().trim() !== "-" && justReqText.toString().trim() !== "N/A") {
    justReqRowHtml = "<tr><td class='label'>Justification for Request</td><td colspan='3'>" + justReqText + "</td></tr>";
  }

  var assessmentData = getAssessmentData(ticketID);
  var assess1 = assessmentData ? assessmentData.newAssetProvisioning : "";
  var assess2 = assessmentData ? assessmentData.assessmentJustification : "";
  var assess3 = assessmentData ? assessmentData.userOperationalIssue : "";
  var assess4 = assessmentData ? assessmentData.recommendedActionPlan : "";

  var pdfHtmlContent =
    "<!DOCTYPE html><html><head><meta charset='UTF-8'><style>" +
    "@page { size: A4 portrait; margin: 8mm; }" +
    "body { font-family: Arial, sans-serif; color: #111; font-size: 10.5px; line-height: 1.35; margin: 0; padding: 0; }" +
    ".header-table { width: 100%; border-bottom: 3.5px solid #d32f2f; padding-bottom: 6px; margin-bottom: 10px; }" +
    ".header-table td { border: none; vertical-align: middle; }" +
    ".logo-img { height: 50px; max-width: 160px; object-fit: contain; }" +
    ".title-area { text-align: right; }" +
    ".title-area h1 { margin: 0; color: #b71c1c; font-size: 17px; text-transform: uppercase; font-weight: bold; }" +
    ".title-area p { margin: 2px 0 0 0; font-size: 10.5px; font-weight: bold; color: #444; text-transform: uppercase; }" +
    ".section-title { background-color: #d32f2f; color: #ffffff; padding: 4px 7px; font-weight: bold; font-size: 10.5px; margin-top: 10px; margin-bottom: 4px; text-transform: uppercase; }" +
    "table.data-table { width: 100%; border-collapse: collapse; margin-bottom: 5px; }" +
    "table.data-table td { border: 1px solid #999; padding: 5px 7px; font-size: 10px; text-align: left; vertical-align: top; }" +
    ".label { background-color: #fdf2f2; font-weight: bold; width: 20%; color: #b71c1c; }" +
    ".assess-box { min-height: 55px; line-height: 1.4; }" +
    ".empty-line { color: #888; font-style: italic; }" +
    ".handwrite-line { border-bottom: 1px solid #666; display: inline-block; min-width: 140px; height: 14px; }" +
    ".policy-box { border: 1px solid #e57373; background-color: #fff9f9; padding: 6px 8px; font-size: 8.5px; color: #222; margin-bottom: 5px; text-align: justify; line-height: 1.3; }" +
    ".policy-box ul { margin: 2px 0 2px 12px; padding: 0; }" +
    ".policy-box li { margin-bottom: 2px; }" +
    "table.approval-grid { width: 100%; border-collapse: collapse; margin-top: 5px; table-layout: fixed; word-wrap: break-word; }" +
    "table.approval-grid td, table.approval-grid th { border: 1px solid #444; padding: 4px; font-size: 8.5px; text-align: left; }" +
    "table.approval-grid th { background-color: #ffebee; color: #b71c1c; text-align: center; font-size: 9px; font-weight: bold; text-transform: uppercase; padding: 5px 3px; border: 1px solid #d32f2f; }" +
    ".sign-space { height: 75px; vertical-align: top; font-size: 8px; color: #222; background-color: #fff; line-height: 1.25; }" +
    ".status-badge { font-weight: bold; }" +
    ".badge-approved { color: #2e7d32; }" +
    ".badge-rejected { color: #c62828; }" +
    ".badge-pending { color: #e65100; }" +
    ".badge-self { color: #1565c0; }" +
    "</style></head><body>" +
    "<table class='header-table'><tr>" +
      "<td style='width: 35%;'><img src='" + logoImgSrc + "' class='logo-img' alt='Apollo Logo' /></td>" +
      "<td class='title-area' style='width: 65%;'><h1>Apollo Food Holdings Berhad</h1><p>IT ASSET & EQUIPMENT REQUEST FORM</p></td>" +
    "</tr></table>" +
    "<div class='section-title'>SECTION A: REQUESTOR DETAILS</div>" +
    "<table class='data-table'>" +
      "<tr><td class='label'>Ticket Request ID</td><td><b style='font-size:11px; color:#d32f2f;'>" + ticketID + "</b></td><td class='label'>Date & Time Requested</td><td>" + timestamp + "</td></tr>" +
      "<tr><td class='label'>Full Name</td><td><b>" + (rowData["Full Name"] || rowData["Nama Penuh"] || "N/A") + "</b></td><td class='label'>Department</td><td><b>" + (rowData["Department"] || rowData["Jabatan"] || "N/A") + "</b></td></tr>" +
      "<tr><td class='label'>Email Address</td><td colspan='3'>" + (extractUserEmail(rowData) || "N/A") + "</td></tr>" +
    "</table>" +
    "<div class='section-title'>SECTION B: REQUEST & EQUIPMENT DETAILS</div>" +
    "<table class='data-table'>" +
      "<tr><td class='label'>Asset Requested</td><td><b>" + (rowData["Asset Type"] || rowData["Jenis Aset"] || "N/A") + "</b></td><td class='label'>Quantity Requested</td><td><b>" + (rowData["Quantity"] || rowData["Kuantiti"] || "1") + "</b></td></tr>" +
      "<tr><td class='label'>Reason / Justification</td><td colspan='3'>" + (rowData["Reason / Justification"] || rowData["Reason"] || "N/A") + "</td></tr>" +
      justReqRowHtml +
    "</table>" +
    "<div class='section-title'>SECTION C: IT TECHNICAL ASSESSMENT & HOD APPROVAL</div>" +
    "<table class='data-table'>" +
      "<tr>" +
        "<td class='label'>New Asset Provisioning</td>" +
        "<td class='assess-box'>" + (assess1 ? assess1 : "<span class='empty-line'>&mdash;</span>") + "</td>" +
        "<td class='label'>IT Assessment & Justification</td>" +
        "<td class='assess-box'>" + (assess2 ? assess2 : "<span class='empty-line'>&mdash;</span>") + "</td>" +
      "</tr>" +
      "<tr>" +
        "<td class='label'>User Operational Issue</td>" +
        "<td class='assess-box'>" + (assess3 ? assess3 : "<span class='empty-line'>&mdash;</span>") + "</td>" +
        "<td class='label'>Recommended Action Plan</td>" +
        "<td class='assess-box'>" + (assess4 ? assess4 : "<span class='empty-line'>&mdash;</span>") + "</td>" +
      "</tr>" +
      "<tr>" +
        "<td class='label'>IT Management Verification</td>" +
        "<td colspan='3' style='padding:8px 10px;'>" +
          "Name: <span class='handwrite-line'></span> &nbsp;&nbsp; Date: <span class='handwrite-line'></span> &nbsp;&nbsp; Sign: <span class='handwrite-line'></span>" +
        "</td>" +
      "</tr>" +
    "</table>" +
    "<div class='section-title'>SECTION D: IT ASSET POLICY ACKNOWLEDGEMENT & TERMS</div>" +
    "<div class='policy-box'>" +
      "By submitting or approving this form, the Requestor and Department acknowledge and agree to abide strictly by Apollo Food Holdings Berhad IT and Information Security Policy (S-ITS-PL-001):" +
      "<ul>" +
        "<li><b>1. Custody & Care (Sec 4.3.4 & 5.2.8):</b> All devices remain the property of the Company and are provided primarily for business use. Users must safeguard all hardware entrusted to them.</li>" +
        "<li><b>2. Software & Device Control (Sec 5.2.5 & 5.2.6):</b> Only IT is authorised to install software. Users must not install unapproved/unlicensed software, nor tamper with device configurations.</li>" +
        "<li><b>3. Incident Reporting (Sec 5.2.9 & 7.3.1):</b> Any lost, stolen, or malfunctioning devices must be reported to IT immediately via the ticketing system.</li>" +
        "<li><b>4. Asset Return (Sec 7.1.4):</b> All Company devices, accessories, and access items must be returned to IT on or before the employee's last working day, or upon role reassignment.</li>" +
      "</ul>" +
    "</div>" +
    "<div class='section-title'>SECTION E: DIGITAL APPROVAL WORKFLOW</div>" +
    "<table class='approval-grid'>" + boxesHtml + "</table>" +
    "<p style='font-size: 8px; color: #555; text-align: center; margin-top: 12px; font-style: italic;'>* Digitally verified via Apollo IT Portal.</p>" +
    "</body></html>";

  var pdfBlob = Utilities.newBlob(pdfHtmlContent, "text/html", "IT_Asset_Request_" + ticketID + ".html").getAs("application/pdf");
  pdfBlob.setName("IT_Asset_Request_" + ticketID + ".pdf");
  return pdfBlob;
}

// ═══════════════════════════════════════════════════════════════
// GET APPROVAL PORTAL DATA (DENGAN CACHE & FORCE REFRESH)
// ═══════════════════════════════════════════════════════════════
function getApprovalPortalData(forceRefresh) {
  try {
    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    var sys = getSystemEmailConfig();

    if (!profile) return JSON.stringify({ success: false, message: "Account (" + (authEmail || "Unknown") + ") not registered." });

    // ═══ CACHE LOGIC ═══
    var cache = CacheService.getScriptCache();
    var cacheKey = "portal_data_" + authEmail;
    
    if (forceRefresh !== true) {
      var cachedData = cache.get(cacheKey);
      if (cachedData) return cachedData;
    }
    // ═══════════════════

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Request sheet not found." });

    var rawData = sheet.getDataRange().getValues();
    if (!rawData || rawData.length < 2) {
      var emptyResponse = JSON.stringify({
        success: true,
        data: {
          approver: profile,
          metrics: { pending: 0, approved: 0, rejected: 0, total: 0, pendingManager: 0, pendingHod: 0, pendingCsuite: 0, pendingMd: 0, pendingIt: 0, pendingQuotation: 0, readyProcurement: 0, discussionCount: 0, discussion: 0 },
          pendingRequests: [], historyRequests: [], allRequests: []
        }
      });
      cache.put(cacheKey, emptyResponse, 60);
      return emptyResponse;
    }

    var headers = rawData[0];
    var allData = rawData.slice(1);

    var pendingRequests = [];
    var historyRequests = [];
    var allRequestsMap = {};

    var pendingMgrCount = 0, pendingHodCount = 0, pendingCsCount = 0, pendingMdCount = 0, pendingItCount = 0;
    var pendingQuotationCount = 0, readyProcurementCount = 0;
    var myApprovedCount = 0, myRejectedCount = 0;

    var myCleanEmail = (authEmail || "").toString().toLowerCase().trim();
    var myAllRoles = profile.allRoles || [profile.role];
    var hasITCheckRole    = myAllRoles.indexOf("IT_CHECK") !== -1;
    var hasITVerifiedRole = myAllRoles.indexOf("IT_VERIFIED") !== -1;
    var hasITHodRole      = myAllRoles.indexOf("IT_HOD") !== -1;
    var hasMDRole         = myAllRoles.indexOf("MD") !== -1;

    var isITCheck    = (hasITCheckRole || (sys.itCheckEmail && myCleanEmail === cleanEmailAddress(sys.itCheckEmail)));
    var isITHod      = (hasITHodRole || (sys.itHodEmail && myCleanEmail === cleanEmailAddress(sys.itHodEmail)));
    var isMD         = (hasMDRole || (sys.mdEmail && myCleanEmail === cleanEmailAddress(sys.mdEmail)));
    var isITVerified = (hasITVerifiedRole || (sys.itVerifiedEmail && myCleanEmail === cleanEmailAddress(sys.itVerifiedEmail)));
    var isITAny      = isITCheck || isITHod || isITVerified;

    var ticketsWithChat = {};
    try {
      var discRows = getAllDiscussionRows();
      for (var di = 0; di < discRows.length; di++) {
        if (discRows[di].ticketId) ticketsWithChat[discRows[di].ticketId] = true;
      }
    } catch (e) { Logger.log("disc fetch error: " + e.toString()); }

    var visibleDiscussionTicketSet = {};

    for (var r = 0; r < allData.length; r++) {
      var rowValues = allData[r];
      var rowIndex = r + 2;
      var state = determineTicketWorkflowState(rowValues, headers, rowIndex);
      if (!state.ticketId) continue;

      if (isMD && state.isSimpleAsset) continue;

      var hierarchy = state.hierarchy || {};
      var rd = state.rowData;

      var mgrEmail = (hierarchy.managerEmail || "").toLowerCase().trim();
      var hodEmail = (hierarchy.hodEmail || "").toLowerCase().trim();
      var csEmail  = (hierarchy.csuiteEmail || "").toLowerCase().trim();

      var isInRoute = false;
      if (isITAny || isMD) isInRoute = true;
      else {
        if (mgrEmail === myCleanEmail) isInRoute = true;
        else if (hodEmail === myCleanEmail) isInRoute = true;
        else if (csEmail === myCleanEmail) isInRoute = true;
      }
      if (!isInRoute) continue;

      var mgrDateRaw = String(rd["Manager Approval Date"] || "").trim();
      var hodDateRaw = String(rd["HOD Approval Date"] || "").trim();
      var csDateRaw  = String(rd["C-Suite Approval Date"] || "").trim();
      var mdDateRaw  = String(rd["MD Approval Date"] || "").trim();
      var itCheckDateRaw = String(rd["IT Check Date"] || "").trim();
      var itVerifyDateRaw = String(rd["IT Verification Date"] || "").trim();

      var isMgrDone = mgrDateRaw && mgrDateRaw !== "Pending" && mgrDateRaw !== "-" && mgrDateRaw.indexOf("Bypassed") === -1;
      var isHodDone = hodDateRaw && hodDateRaw !== "Pending" && hodDateRaw !== "-" && hodDateRaw.indexOf("Bypassed") === -1;
      var isCsDone  = csDateRaw && csDateRaw !== "Pending" && csDateRaw !== "-" && csDateRaw.indexOf("Bypassed") === -1;
      var isMdDone  = mdDateRaw && mdDateRaw !== "Pending" && mdDateRaw !== "-" && mdDateRaw.indexOf("Bypassed") === -1;

      var hasReachedMyStage = false;

      if (state.isTerminated) {
        if (isITAny) hasReachedMyStage = true;
        var rejectTag = getRejectStageTagFromStatus(state.status);

        if (rejectTag === "Manager" && mgrEmail && mgrEmail === myCleanEmail) hasReachedMyStage = true;
        else if (rejectTag === "HOD" && hodEmail && hodEmail === myCleanEmail) hasReachedMyStage = true;
        else if (rejectTag === "C-Suite" && csEmail && csEmail === myCleanEmail) hasReachedMyStage = true;
        else if (rejectTag === "MD" && isMD) hasReachedMyStage = true;
        else if (rejectTag === "IT Check" && isITCheck) hasReachedMyStage = true;
        else if (rejectTag === "IT Verification" && isITVerified) hasReachedMyStage = true;
      } else if (state.isCompleted) {
        hasReachedMyStage = true;
      } else {
        if (isITAny || isMD) hasReachedMyStage = true;
        else {
          if (mgrEmail === myCleanEmail && (isMgrDone || state.stage === "MANAGER")) hasReachedMyStage = true;
          else if (hodEmail === myCleanEmail && (isMgrDone || state.stage === "HOD")) hasReachedMyStage = true;
          else if (csEmail === myCleanEmail && (isHodDone || state.stage === "CSUITE")) hasReachedMyStage = true;
        }
      }
      if (!hasReachedMyStage) continue;

      if (state.stage === "MANAGER") pendingMgrCount++;
      else if (state.stage === "HOD") pendingHodCount++;
      else if (state.stage === "CSUITE") pendingCsCount++;
      else if (state.stage === "MD") pendingMdCount++;
      else if (state.stage === "IT_CHECK") pendingItCount++;
      else if (state.stage === "IT_VERIFIED") pendingQuotationCount++;

      var statusUpper = (state.status || "").toUpperCase();
      if (statusUpper.indexOf("READY FOR PROCUREMENT") !== -1) readyProcurementCount++;

      // FIX METRICS: Count approved based on actual approvals made by this user
      var hasApproved = false;
      if (isITCheck && itCheckDateRaw && itCheckDateRaw !== "Pending" && itCheckDateRaw !== "-") hasApproved = true;
      if (isITVerified && itVerifyDateRaw && itVerifyDateRaw !== "Pending" && itVerifyDateRaw !== "-") hasApproved = true;
      if (isMD && isMdDone) hasApproved = true;
      if (mgrEmail === myCleanEmail && isMgrDone) hasApproved = true;
      if (hodEmail === myCleanEmail && isHodDone) hasApproved = true;
      if (csEmail === myCleanEmail && isCsDone) hasApproved = true;
      
      if (hasApproved) myApprovedCount++;

      if (state.isTerminated) {
        myRejectedCount++;
      }

      var tsRaw = rd["Timestamp"] instanceof Date ? rd["Timestamp"].getTime() : 0;

      var qCount = 0;
      var selectedVendor = "";
      var selectedPrice = "";
      if (isITCheck || isMD || isITHod) {
        try {
          var qs = getQuotationsByTicket(state.ticketId);
          qCount = qs ? qs.length : 0;
          for (var qi = 0; qi < qs.length; qi++) {
            if (qs[qi].isSelected) {
              selectedVendor = qs[qi].vendorName || "";
              selectedPrice = qs[qi].price || "";
              break;
            }
          }
        } catch (e) { qCount = 0; }
      }

      var isMyTurn = false;
      if (state.requiredEmail && state.requiredEmail.toLowerCase().trim() === myCleanEmail) isMyTurn = true;
      if (isMD && state.stage === "MD") isMyTurn = true;
      if (isITCheck && state.stage === "IT_CHECK") isMyTurn = true;
      if (isITVerified && state.stage === "IT_VERIFIED") isMyTurn = true;

      var isActionableForMe = isMyTurn && !state.isTerminated && !state.isCompleted;

      if (isITCheck && statusUpper.indexOf("READY FOR PROCUREMENT") !== -1) {
        isActionableForMe = true;
      }
      if (isITCheck && statusUpper.indexOf("PENDING IT CHECK") !== -1) {
        isActionableForMe = true;
      }

      var card = {
        "Ticket ID": String(state.ticketId || ""),
        "Timestamp": formatSafeDate(rd["Timestamp"]),
        "TimestampRaw": tsRaw,
        "Full Name": String(state.namaUser || "Staff"),
        "Department": String(rd["Department"] || rd["Jabatan"] || "-"),
        "Asset Type": String(rd["Asset Type"] || rd["Jenis Aset"] || "-"),
        "Quantity": String(rd["Quantity"] || rd["Kuantiti"] || "1"),
        "Reason / Justification": String(rd["Reason / Justification"] || rd["Reason"] || "-"),
        "Attachment": String(rd["Attachment"] || ""),
        "Status": String(state.status || "Pending Approval"),
        "Manager Approver": String(rd["Manager Approver"] || "-"),
        "Manager Approval Date": formatSafeDate(rd["Manager Approval Date"]),
        "HOD Approver": String(rd["HOD Approver"] || "-"),
        "HOD Approval Date": formatSafeDate(rd["HOD Approval Date"]),
        "C-Suite Approver": String(rd["C-Suite Approver"] || "-"),
        "C-Suite Approval Date": formatSafeDate(rd["C-Suite Approval Date"]),
        "MD Approver": String(rd["MD Approver"] || "-"),
        "MD Approval Date": formatSafeDate(rd["MD Approval Date"]),
        "requesterEmail": String(state.emailUser || ""),
        "currentStage": String(state.stage || ""),
        "isCompleted": Boolean(state.isCompleted),
        "isTerminated": Boolean(state.isTerminated),
        "isSimple": Boolean(state.isSimpleAsset),
        "isReadyProc": Boolean(state.isReadyProc),
        "hasDiscussion": false,
        "quotationCount": Number(qCount),
        "selectedVendor": String(selectedVendor || ""),
        "selectedPrice": String(selectedPrice || ""),
        "submittedToMDDate": formatSafeDate(rd["Submitted to MD Date"]),
        "rowIndex": Number(rowIndex)
      };

      var tidUpper = String(state.ticketId || "").toUpperCase();
      if (ticketsWithChat[tidUpper] && canUserAccessTicketDiscussion(state, authEmail, profile)) {
        card["hasDiscussion"] = true;
        visibleDiscussionTicketSet[tidUpper] = true;
      }

      allRequestsMap[state.ticketId] = card;

      if (isActionableForMe) pendingRequests.push(card);
      else historyRequests.push(card);
    }

    function sortLatest(a, b) { return (b.TimestampRaw || 0) - (a.TimestampRaw || 0); }
    pendingRequests.sort(sortLatest);
    historyRequests.sort(sortLatest);

    var allRequests = Object.keys(allRequestsMap).map(function(k) { return allRequestsMap[k]; });
    allRequests.sort(sortLatest);

    var discussionCount = Object.keys(visibleDiscussionTicketSet).length;

    var finalResponse = JSON.stringify({
      success: true,
      data: {
        approver: profile,
        metrics: {
          pending: pendingRequests.length, approved: myApprovedCount, rejected: myRejectedCount,
          total: allRequests.length, pendingManager: pendingMgrCount, pendingHod: pendingHodCount,
          pendingCsuite: pendingCsCount, pendingMd: pendingMdCount, pendingIt: pendingItCount,
          pendingQuotation: pendingQuotationCount, readyProcurement: readyProcurementCount,
          discussionCount: discussionCount,
          discussion: discussionCount
        },
        pendingRequests: pendingRequests, historyRequests: historyRequests, allRequests: allRequests
      }
    });

    cache.put(cacheKey, finalResponse, 30); // Simpan 30 saat
    return finalResponse;
  } catch (err) {
    Logger.log("getApprovalPortalData error: " + err.toString());
    return JSON.stringify({ success: false, message: "Error: " + err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════════
// GET REQUEST DETAILS
// ═══════════════════════════════════════════════════════════════
function getRequestDetails(ticketId) {
  try {
    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet not found." });

    var rawData = sheet.getDataRange().getValues();
    if (!rawData || rawData.length < 2) return JSON.stringify({ success: false, message: "No data." });

    var headers = rawData[0];
    var allData = rawData.slice(1);
    var targetRowValues = null, targetRowIndex = -1;

    for (var i = 0; i < allData.length; i++) {
      var rData = {};
      for (var h = 0; h < headers.length; h++) {
        var cleanH = headers[h].toString().trim();
        if (!rData.hasOwnProperty(cleanH)) rData[cleanH] = allData[i][h];
      }
      if ((rData["Ticket ID"] || "").toString().trim().toUpperCase() === ticketId.toString().trim().toUpperCase()) {
        targetRowValues = allData[i]; targetRowIndex = i + 2; break;
      }
    }

    if (!targetRowValues) return JSON.stringify({ success: false, message: "Request not found." });

    var state = determineTicketWorkflowState(targetRowValues, headers, targetRowIndex);
    var sys = getSystemEmailConfig();
    var myAllRoles = profile.allRoles || [profile.role];
    var isMD = (myAllRoles.indexOf("MD") !== -1 || (sys.mdEmail && authEmail === cleanEmailAddress(sys.mdEmail)));
    var isITCheck = (myAllRoles.indexOf("IT_CHECK") !== -1 || (sys.itCheckEmail && authEmail === cleanEmailAddress(sys.itCheckEmail)));
    var isITVerified = (myAllRoles.indexOf("IT_VERIFIED") !== -1 || (sys.itVerifiedEmail && authEmail === cleanEmailAddress(sys.itVerifiedEmail)));

    var isActionableForMe = false;
    if (!state.isTerminated && !state.isCompleted) {
      if (isMD && state.stage === "MD") isActionableForMe = true;
      else if (isITCheck && state.stage === "IT_CHECK") isActionableForMe = true;
      else if (isITVerified && state.stage === "IT_VERIFIED") isActionableForMe = true;
      else if (state.requiredEmail && state.requiredEmail.toLowerCase() === authEmail) isActionableForMe = true;
    }
    var statusUpper = (state.status || "").toUpperCase();
    if (isITCheck && statusUpper.indexOf("READY FOR PROCUREMENT") !== -1) isActionableForMe = true;
    if (isITCheck && statusUpper.indexOf("PENDING IT CHECK") !== -1) isActionableForMe = true;

    var chainData = buildChainDataStructure(state.hierarchy, state.emailUser, state.isSimpleAsset, state.rowData);
    var workflow = (chainData.steps || []).map(function(s) {
      return { role: String(s.role||""), name: String(s.name||"-"), email: String(s.email||"-"), status: String(s.status||"Pending"), date: String(s.date||"-") };
    });

    var assessment = getAssessmentData(ticketId);
    var quotations = getQuotationsByTicket(ticketId);
    var rd = state.rowData;
    var requestObj = {
      "Timestamp": formatSafeDate(rd["Timestamp"]),
      "Ticket ID": String(state.ticketId || ""),
      "Full Name": String(state.namaUser || "Staff"),
      "Department": String(rd["Department"] || rd["Jabatan"] || "-"),
      "Asset Type": String(rd["Asset Type"] || rd["Jenis Aset"] || "-"),
      "Quantity": String(rd["Quantity"] || rd["Kuantiti"] || "1"),
      "Reason / Justification": String(rd["Reason / Justification"] || rd["Reason"] || "-"),
      "Attachment": String(rd["Attachment"] || ""),
      "Status": String(state.status || "Pending Approval")
    };

    var canChat = canUserChatOnTicket(state, authEmail, profile);

    return JSON.stringify({
      success: true,
      data: {
        ticketId: state.ticketId, request: requestObj, requesterEmail: state.emailUser,
        status: state.status, currentStage: state.stage,
        isActionableForUser: isActionableForMe, canChat: canChat,
        workflow: workflow, assessment: assessment, quotations: quotations,
        isSimple: state.isSimpleAsset, isReadyProc: state.isReadyProc
      }
    });
  } catch (err) {
    return JSON.stringify({ success: false, message: "Error: " + err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════════
// ADMIN PORTAL DATA
// ═══════════════════════════════════════════════════════════════
function getAdminPortalData() {
  try {
    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    var sys = getSystemEmailConfig();
    if (!profile) return JSON.stringify({ success: false, message: "Not registered." });

    var isAllowed = (profile.role === "IT_HOD" || profile.role === "ADMIN" || (sys.itHodEmail && authEmail === cleanEmailAddress(sys.itHodEmail)));
    if (!isAllowed) return JSON.stringify({ success: false, message: "Admin Portal for IT HOD only." });

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet not found." });

    var rawData = sheet.getDataRange().getValues();
    if (!rawData || rawData.length < 2) {
      return JSON.stringify({ success: true, data: { approver: profile, metrics: { total: 0, pending: 0, approved: 0, rejected: 0, discussionCount: 0, discussion: 0 }, pendingRequests: [], historyRequests: [], allRequests: [] } });
    }

    var headers = rawData[0];
    var allData = rawData.slice(1);
    var pendingRequests = [], historyRequests = [], allRequestsMap = {};
    var pendingMgrCount = 0, pendingHodCount = 0, pendingCsCount = 0, pendingMdCount = 0, pendingItCount = 0;
    var readyProcCount = 0;
    var approvedCount = 0, rejectedCount = 0;

    var discussionTicketSet = {};
    try {
      var discRows = getAllDiscussionRows();
      for (var di = 0; di < discRows.length; di++) {
        if (discRows[di].ticketId) discussionTicketSet[discRows[di].ticketId] = true;
      }
    } catch (e) {}

    for (var r = 0; r < allData.length; r++) {
      var state = determineTicketWorkflowState(allData[r], headers, r + 2);
      if (!state.ticketId) continue;

      if (state.stage === "MANAGER") pendingMgrCount++;
      else if (state.stage === "HOD") pendingHodCount++;
      else if (state.stage === "CSUITE") pendingCsCount++;
      else if (state.stage === "MD") pendingMdCount++;
      else if (state.stage === "IT_CHECK") pendingItCount++;
      if (state.isReadyProc) readyProcCount++;
      if (state.isTerminated) rejectedCount++;
      else if (state.isCompleted) approvedCount++;

      var rd = state.rowData;
      var tsRaw = rd["Timestamp"] instanceof Date ? rd["Timestamp"].getTime() : 0;

      var qCount = 0;
      var qList = [];
      try {
        var qs = getQuotationsByTicket(state.ticketId);
        qCount = qs ? qs.length : 0;
        qList = qs || [];
      } catch (e) { qCount = 0; }

      var card = {
        "Ticket ID": String(state.ticketId || ""), "Timestamp": formatSafeDate(rd["Timestamp"]),
        "TimestampRaw": tsRaw, "Full Name": String(state.namaUser || "Staff"),
        "Department": String(rd["Department"] || rd["Jabatan"] || "-"),
        "Asset Type": String(rd["Asset Type"] || rd["Jenis Aset"] || "-"),
        "Quantity": String(rd["Quantity"] || rd["Kuantiti"] || "1"),
        "Reason / Justification": String(rd["Reason / Justification"] || rd["Reason"] || "-"),
        "Status": String(state.status || "Pending"),
        "currentStage": String(state.stage || ""),
        "quotationCount": Number(qCount),
        "quotations": qList,
        "hasDiscussion": Boolean(discussionTicketSet[String(state.ticketId || "").toUpperCase()]),
        "isCompleted": Boolean(state.isCompleted), "isTerminated": Boolean(state.isTerminated),
        "isReadyProc": Boolean(state.isReadyProc)
      };

      allRequestsMap[state.ticketId] = card;
      if (!state.isTerminated && !state.isCompleted) pendingRequests.push(card);
      else historyRequests.push(card);
    }

    function sortLatest(a, b) { return (b.TimestampRaw || 0) - (a.TimestampRaw || 0); }
    pendingRequests.sort(sortLatest);
    historyRequests.sort(sortLatest);
    var allRequests = Object.keys(allRequestsMap).map(function(k) { return allRequestsMap[k]; });
    allRequests.sort(sortLatest);

    var discussionCount = Object.keys(discussionTicketSet).length;

    return JSON.stringify({
      success: true,
      data: {
        approver: profile,
        metrics: { total: allRequests.length, pending: pendingRequests.length, approved: approvedCount, rejected: rejectedCount,
                   pendingManager: pendingMgrCount, pendingHod: pendingHodCount, pendingCsuite: pendingCsCount, pendingMd: pendingMdCount, pendingIt: pendingItCount,
                   readyProcurement: readyProcCount, discussionCount: discussionCount, discussion: discussionCount },
        pendingRequests: pendingRequests, historyRequests: historyRequests, allRequests: allRequests
      }
    });
  } catch (err) {
    return JSON.stringify({ success: false, message: "Error: " + err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════════
// SUBMIT APPROVAL DECISION
// ═══════════════════════════════════════════════════════════════
function submitApprovalDecision(payload) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "System busy." });
    if (!payload || !payload.ticketId || !payload.decision) return JSON.stringify({ success: false, message: "Invalid payload." });

    var authEmail = getAuthenticatedUserEmail();
    var profile = getApproverProfile(authEmail);
    if (!profile) return JSON.stringify({ success: false, message: "Not authorized." });
    if (profile.role === "IT_HOD") return JSON.stringify({ success: false, message: "HOD IT is Monitoring only." });

    var sheet = getTargetSheet();
    if (!sheet) return JSON.stringify({ success: false, message: "Sheet not found." });

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify({ success: false, message: "Not found." });

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    var targetRowValues = null, targetRowIndex = -1;

    for (var i = 0; i < allData.length; i++) {
      var rData = {};
      for (var h = 0; h < headers.length; h++) {
        var cleanH = headers[h].toString().trim();
        if (!rData.hasOwnProperty(cleanH)) rData[cleanH] = allData[i][h];
      }
      if ((rData["Ticket ID"] || "").toString().trim().toUpperCase() === payload.ticketId.toString().trim().toUpperCase()) {
        targetRowValues = allData[i]; targetRowIndex = i + 2; break;
      }
    }

    if (!targetRowValues) return JSON.stringify({ success: false, message: "Request not found." });

    var state = determineTicketWorkflowState(targetRowValues, headers, targetRowIndex);
    if (state.isTerminated || state.isCompleted) return JSON.stringify({ success: false, message: "Already finalized." });

    var decision = payload.decision;
    var remarks = payload.remarks || "";

    if (decision === "Rejected" && !remarks.trim()) return JSON.stringify({ success: false, message: "Remarks required." });

    if (decision === "Approved" && (state.stage === "IT_VERIFIED" || state.status.indexOf("Pending IT Verification") !== -1)) {
      var combinedAttachments = [];
      var rawLinks = payload.attachmentLinks || [];
      for (var li = 0; li < rawLinks.length; li++) {
        var lk = rawLinks[li];
        if (lk && lk.url && lk.url.trim() !== "") {
          combinedAttachments.push({ type: "link", url: lk.url.trim(), label: (lk.label || "").trim() || ("Link " + (combinedAttachments.length + 1)) });
        }
      }
      var rawFiles = payload.attachmentFiles || [];
      for (var fi = 0; fi < rawFiles.length; fi++) {
        var f = rawFiles[fi];
        if (f && f.base64 && f.name) {
          var upResult = saveAssessmentAttachment(f.base64, f.name, f.mimeType || "application/octet-stream", state.ticketId);
          if (upResult && upResult.success) {
            combinedAttachments.push({ type: "file", url: upResult.url, fileId: upResult.fileId, label: f.label || upResult.fileName, fileName: upResult.fileName });
          }
        }
      }
      saveAssessmentData(state.ticketId, {
        newAssetProvisioning: payload.newAssetProvisioning || "",
        assessmentJustification: payload.assessmentJustification || "",
        userOperationalIssue: payload.userOperationalIssue || "",
        recommendedActionPlan: payload.recommendedActionPlan || "",
        assessedBy: profile.name || profile.email,
        assessedByEmail: profile.email,
        attachmentLinks: combinedAttachments
      });
    }

    if (decision === "Approved") {
      var targetAction = (state.stage === "IT_VERIFIED" || state.status.indexOf("Pending IT Verification") !== -1) ? "it_verify_approve" : (state.action || "");
      if (!targetAction) return JSON.stringify({ success: false, message: "No action available at this stage." });
      executeApprovalAction(targetAction, state.ticketId, targetRowIndex, state.isSimpleAsset, state.stageTitle, remarks);
    } else if (decision === "Rejected") {
      executeApprovalAction("reject", state.ticketId, targetRowIndex, state.isSimpleAsset, state.stageTitle || (state.stage + " Approval"), remarks);
    }

    return JSON.stringify({ success: true, data: { message: "Request " + state.ticketId + " " + decision.toLowerCase() + " successfully." } });
  } catch (err) {
    Logger.log("submitApprovalDecision error: " + err.toString());
    return JSON.stringify({ success: false, message: "Error: " + err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ═══════════════════════════════════════════════════════════════
// DETERMINE TICKET WORKFLOW STATE
// ═══════════════════════════════════════════════════════════════
function determineTicketWorkflowState(rowValues, headers, rowIndex) {
  var rowData = {};
  for (var i = 0; i < headers.length; i++) {
    var cleanHeader = headers[i].toString().trim();
    if (!rowData.hasOwnProperty(cleanHeader)) rowData[cleanHeader] = rowValues[i];
  }

  var sys = getSystemEmailConfig();
  var ticketId = (rowData["Ticket ID"] || "").toString().trim();
  var emailUser = extractUserEmailFromRow(rowValues);
  var namaUser = rowData["Full Name"] || rowData["Nama Penuh"] || "Staff";
  var assetType = (rowData["Asset Type"] || rowData["Jenis Aset"] || "").toString().toLowerCase();
  var isSimpleAsset = (assetType.indexOf("mouse") !== -1 || assetType.indexOf("keyboard") !== -1 || assetType.indexOf("monitor") !== -1);
  var status = (rowData["Status"] || "Pending Approval").toString().trim();
  var statusUpper = status.toUpperCase();
  var hierarchy = getApprovalHierarchy(emailUser);

  var itVerifyDate = (rowData["IT Verification Date"] || "").toString().trim();
  var mgrDate = (rowData["Manager Approval Date"] || "").toString().trim();
  var hodDate = (rowData["HOD Approval Date"] || "").toString().trim();
  var csDate  = (rowData["C-Suite Approval Date"] || "").toString().trim();
  var mdDate  = (rowData["MD Approval Date"] || "").toString().trim();
  var itCheckDate = (rowData["IT Check Date"] || "").toString().trim();

  var isItVerifyDone = itVerifyDate && itVerifyDate !== "Pending" && itVerifyDate !== "-";
  var isMgrDone = mgrDate && mgrDate !== "Pending" && mgrDate !== "-" && mgrDate.indexOf("Bypassed") === -1;
  var isHodDone = hodDate && hodDate !== "Pending" && hodDate !== "-" && hodDate.indexOf("Bypassed") === -1;
  var isCsDone  = csDate && csDate !== "Pending" && csDate !== "-" && csDate.indexOf("Bypassed") === -1;
  var isMdDone  = mdDate && mdDate !== "Pending" && mdDate !== "-" && mdDate.indexOf("Bypassed") === -1;
  var isItCheckDone = itCheckDate && itCheckDate !== "Pending" && itCheckDate !== "-";

  var isTerminated = statusUpper.indexOf("REJECTED") !== -1 || statusUpper.indexOf("CANCELLED") !== -1;
  var isCompleted  = statusUpper.indexOf("APPROVED & COMPLETED") !== -1;
  var isReadyProc  = statusUpper.indexOf("READY FOR PROCUREMENT") !== -1;
  var isPendingMD  = statusUpper.indexOf("PENDING MD") !== -1;
  var isPendingITCheck = statusUpper.indexOf("PENDING IT CHECK") !== -1;

  var stage = "UNKNOWN";
  var requiredEmail = null;
  var stageTitle = "";
  var action = "";

  if (isTerminated) {
    stage = "TERMINATED";
    stageTitle = "Rejected";
  } else if (isCompleted) {
    stage = "COMPLETED";
    stageTitle = "Completed";
  } else if (isReadyProc) {
    stage = "READY_FOR_PROCUREMENT";
    stageTitle = "Ready for Procurement";
  } else if (isPendingMD) {
    stage = "MD";
    stageTitle = "MD Approval";
    requiredEmail = sys.mdEmail;
    action = "md_approve";
  } else if (isPendingITCheck) {
    stage = "IT_CHECK";
    stageTitle = "IT Check & Quotation";
    requiredEmail = sys.itCheckEmail;
  } else {
    if (!isItVerifyDone) {
      stage = "IT_VERIFIED";
      stageTitle = "IT Verification";
      requiredEmail = sys.itVerifiedEmail || sys.itCheckEmail;
      action = "it_verify_approve";
    } else if (!isMgrDone && hierarchy.managerEmail) {
      stage = "MANAGER";
      stageTitle = "Manager Approval";
      requiredEmail = hierarchy.managerEmail;
      action = "manager_approve";
    } else if (!isHodDone && hierarchy.hodEmail) {
      stage = "HOD";
      stageTitle = "HOD Approval";
      requiredEmail = hierarchy.hodEmail;
      action = "hod_approve";
    } else if (!isCsDone && hierarchy.csuiteEmail) {
      stage = "CSUITE";
      stageTitle = "C-Suite Approval";
      requiredEmail = hierarchy.csuiteEmail;
      action = "csuite_approve";
    } else if (!isItCheckDone) {
      stage = "IT_CHECK";
      stageTitle = "IT Check & Quotation";
      requiredEmail = sys.itCheckEmail;
    } else {
      stage = "UNKNOWN";
      stageTitle = status;
    }
  }

  return {
    ticketId: ticketId,
    emailUser: emailUser,
    namaUser: namaUser,
    status: status,
    stage: stage,
    stageTitle: stageTitle,
    requiredEmail: requiredEmail,
    action: action,
    rowData: rowData,
    hierarchy: hierarchy,
    isSimpleAsset: isSimpleAsset,
    isCompleted: isCompleted,
    isTerminated: isTerminated,
    isReadyProc: isReadyProc,
    rowIndex: rowIndex
  };
}