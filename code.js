function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Nothing ERP - Audit Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function verifyLogin(loginId, password) {
  var userId = String(loginId || '').trim();
  var pass = String(password || '').trim();

  if (!userId || !pass) {
    return { success: false, message: "Please enter User ID and Password." };
  }

  var userList = [
    { userId: 'Admin', password: 'Admin@12345', role: 'admin' },
    { userId: 'QA', password: 'QA@12345', role: 'employee' },
    { userId: 'Network', password: 'Network@12345', role: 'network' },
    { userId: 'NET', password: 'Network@12345', role: 'network' }
  ];

  for (var i = 0; i < userList.length; i++) {
    if (userList[i].userId.toLowerCase() === userId.toLowerCase() &&
        userList[i].password === pass) {
      return {
        success: true,
        role: userList[i].role,
        userId: userList[i].userId
      };
    }
  }

  return {
    success: false,
    message: "Incorrect User ID or Password."
  };
}

function getAppSettings() {
  return {
    appName: 'Nothing ERP',
    appShortName: 'ERP',
    version: '1.0',
    loginTitle: 'Nothing ERP - Audit Form',
    supportEmail: 'support@nothingerp.com'
  };
}

function getRolePermissions(role) {
  var permissions = {
    admin: ['dashboard', 'audit', 'reports', 'mapping', 'locks', 'settings'],
    employee: ['dashboard', 'audit', 'reports'],
    network: ['dashboard', 'audit', 'mapping', 'network']
  };

  return permissions[String(role || '').toLowerCase()] || [];
}

function getAuditData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Check list 1.4 ") || ss.getSheetByName("Check list 1.4");
    if (!sheet) return [{ error: "Check list sheet not found." }];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var data = sheet.getRange(2, 2, lastRow - 1, 6).getDisplayValues();

    return data.map(function (row, index) {
      return {
        id: index + 1,
        category: row[0] || "",
        checkpoint: row[1] || "",
        applicability: row[2] || "",
        maxMarks: row[3] || 0,
        immediateFix: row[4] || "",
        capa: row[5] || ""
      };
    });
  } catch (e) {
    return [{ error: e.toString() }];
  }
}

function getHistoricalNCData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Repeated NC data");
    if (!sheet) return [{ error: "Repeated NC data sheet not found." }];

    var data = sheet.getDataRange().getDisplayValues();
    var records = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][3]) {
        records.push({
          auditDate: data[i][0] || "",
          auditorName: data[i][1] || "",
          scName: data[i][2] || "",
          scCode: data[i][3] || "",
          region: data[i][4] || "",
          state: data[i][5] || "",
          city: data[i][6] || "",
          observation: data[i][7] || "",
          score: data[i][8] || ""
        });
      }
    }

    return records;
  } catch (e) {
    return [{ error: e.toString() }];
  }
}

function getMappingData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var sheet = null;

    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().toLowerCase().includes("mapping")) {
        sheet = sheets[i];
        break;
      }
    }

    if (!sheet) return [{ error: "Mapping sheet not found." }];

    var data = sheet.getDataRange().getDisplayValues();
    var records = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] || data[i][4]) {
        var rawEmails = [
          data[i][8], data[i][9], data[i][10], data[i][11], data[i][12], data[i][13]
        ];
        var validEmails = rawEmails.filter(function (e) {
          return e && e.toString().includes('@');
        });

        records.push({
          scCode: data[i][0] || "N/A",
          svcNameColB: data[i][1] || "",
          city: data[i][1] || "",
          state: data[i][2] || "",
          region: data[i][3] || "",
          svcName: data[i][4] || "",
          svcStatus: data[i][5] || "",
          priorityDesk: data[i][6] || "",
          bmName: data[i][8] || "N/A",
          rmName: data[i][9] || "N/A",
          spocEmails: validEmails
        });
      }
    }

    return records;
  } catch (e) {
    return [{ error: e.toString() }];
  }
}

function getNCClosureData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("NC closure") || ss.getSheetByName("NC Closure");
    if (!sheet) return [{ error: "NC closure sheet not found." }];

    var data = sheet.getDataRange().getDisplayValues();
    var records = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][1]) {
        records.push({
          auditDate: data[i][0] || "",
          scName: data[i][1] || "",
          scCode: data[i][2] || "",
          scType: data[i][3] || "",
          scCategory: data[i][4] || "",
          region: data[i][5] || "",
          state: data[i][6] || "",
          city: data[i][7] || "",
          ncPoints: data[i][8] || "",
          score: data[i][9] || "",
          totalObserved: data[i][10] || "",
          closedCount: data[i][11] || "",
          groupAgeing: data[i][12] || "",
          evidenceLink: data[i][13] || "",
          openCount: data[i][14] || "",
          ageing: data[i][15] || ""
        });
      }
    }

    return records;
  } catch (e) {
    return [{ error: e.toString() }];
  }
}

function acquireLock(scCode, auditorId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Audit Locks");
    if (!sheet) {
      sheet = ss.insertSheet("Audit Locks");
      sheet.appendRow(["SC Code", "Locked By", "Locked At", "Expires At"]);
      sheet.setFrozenRows(1);
      sheet.hideSheet();
    }

    var now = new Date();
    var expires = new Date(now.getTime() + 4 * 60 * 60 * 1000);

    var data = sheet.getDataRange().getDisplayValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === scCode) {
        var expiresAt = new Date(data[i][3]);

        if (expiresAt > now) {
          if (data[i][1] === auditorId) {
            sheet.getRange(i + 1, 1, 1, 4).setValues([[scCode, auditorId, now.toISOString(), expires.toISOString()]]);
            return { success: true, lockedBy: auditorId, isOwn: true };
          }
          return { success: false, lockedBy: data[i][1], lockedAt: data[i][2], expiresAt: data[i][3] };
        } else {
          sheet.getRange(i + 1, 1, 1, 4).setValues([[scCode, auditorId, now.toISOString(), expires.toISOString()]]);
          return { success: true, lockedBy: auditorId, isOwn: false };
        }
      }
    }

    sheet.appendRow([scCode, auditorId, now.toISOString(), expires.toISOString()]);
    return { success: true, lockedBy: auditorId, isOwn: false };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function releaseLock(scCode, auditorId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Audit Locks");
    if (!sheet) return { success: true };

    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === scCode && data[i][1] === auditorId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function checkLock(scCode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Audit Locks");
    if (!sheet) return { locked: false };

    var data = sheet.getDataRange().getDisplayValues();
    var now = new Date();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === scCode) {
        var expiresAt = new Date(data[i][3]);
        if (expiresAt > now) {
          return { locked: true, lockedBy: data[i][1], lockedAt: data[i][2], expiresAt: data[i][3] };
        }
      }
    }

    return { locked: false };
  } catch (e) {
    return { locked: false };
  }
}

function getAllActiveLocks() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Audit Locks");
    if (!sheet) return [];

    var data = sheet.getDataRange().getDisplayValues();
    var now = new Date();
    var locks = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var expiresAt = new Date(data[i][3]);
        var isActive = expiresAt > now;
        var minutesLeft = isActive ? Math.round((expiresAt - now) / 60000) : 0;

        locks.push({
          scCode: data[i][0],
          lockedBy: data[i][1],
          lockedAt: data[i][2],
          expiresAt: data[i][3],
          isActive: isActive,
          minutesLeft: minutesLeft
        });
      }
    }

    return locks;
  } catch (e) {
    return [];
  }
}

function adminForceReleaseLock(scCode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Audit Locks");
    if (!sheet) return { success: true };

    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === scCode) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAnomalyAlerts() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Master Audit Logs");
    if (!sheet) return { alerts: [], message: "No audit log data yet." };

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return { alerts: [], message: "Insufficient audit history." };

    var siteMap = {};
    for (var i = 1; i < data.length; i++) {
      var scName = data[i][2];
      var scCode = data[i][3];
      var pct = parseFloat(data[i][9]);
      var dateStr = data[i][0];

      if (!scName || isNaN(pct)) continue;
      var key = scCode || scName;

      if (!siteMap[key]) {
        siteMap[key] = { scName: scName, scCode: scCode, entries: [] };
      }

      siteMap[key].entries.push({ date: dateStr, score: pct });
    }

    var alerts = [];

    Object.keys(siteMap).forEach(function (key) {
      var site = siteMap[key];
      var entries = site.entries;

      if (entries.length < 3) return;

      var scores = entries.map(function (e) { return e.score; });
      var latest = scores[scores.length - 1];
      var baseline = scores.slice(0, scores.length - 1);
      var mean = baseline.reduce(function (s, v) { return s + v; }, 0) / baseline.length;

      var variance = baseline.reduce(function (s, v) {
        return s + Math.pow(v - mean, 2);
      }, 0) / baseline.length;

      var stdDev = Math.sqrt(variance);

      if (stdDev < 1) return;

      var zScore = (latest - mean) / stdDev;

      if (zScore < -2) {
        var severity = zScore < -3 ? "CRITICAL" : "WARNING";
        alerts.push({
          scName: site.scName,
          scCode: site.scCode,
          latestScore: latest,
          baseline: mean.toFixed(1),
          stdDev: stdDev.toFixed(1),
          zScore: zScore.toFixed(2),
          drop: (mean - latest).toFixed(1),
          severity: severity,
          lastAuditDate: entries[entries.length - 1].date,
          auditCount: entries.length
        });
      }
    });

    alerts.sort(function (a, b) {
      if (a.severity !== b.severity) return a.severity === "CRITICAL" ? -1 : 1;
      return parseFloat(a.zScore) - parseFloat(b.zScore);
    });

    return { alerts: alerts, totalSitesAnalysed: Object.keys(siteMap).length };
  } catch (e) {
    return { alerts: [], message: e.message };
  }
}

function saveDraft(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Draft Audits");
    if (!sheet) {
      sheet = ss.insertSheet("Draft Audits");
      sheet.appendRow(["SC Code", "SC Name", "Saved At", "State JSON"]);
      sheet.setFrozenRows(1);
      sheet.hideSheet();
    }

    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === payload.scCode) {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[payload.scCode, payload.scName, payload.savedAt, JSON.stringify(payload.state)]]);
        return { success: true };
      }
    }

    sheet.appendRow([payload.scCode, payload.scName, payload.savedAt, JSON.stringify(payload.state)]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function loadDraft(scCode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Draft Audits");
    if (!sheet) return null;

    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === scCode && data[i][3]) {
        return {
          scCode: data[i][0],
          scName: data[i][1],
          savedAt: data[i][2],
          state: JSON.parse(data[i][3])
        };
      }
    }

    return null;
  } catch (e) {
    return null;
  }
}

function clearDraft(scCode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Draft Audits");
    if (!sheet) return { success: true };

    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === scCode) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveEvidenceImage(payload) {
  try {
    var scCode = String(payload.scCode || '').trim();
    var checkpointId = String(payload.checkpointId || '').trim();
    var dataUrl = payload.dataUrl || "";

    if (!dataUrl) {
      return { success: false, message: "No image data found." };
    }

    var match = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
    if (!match) {
      return { success: false, message: "Invalid image format." };
    }

    var mimeType = match[1];
    var base64 = match[2];
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, "evidence_" + scCode + "_" + checkpointId + "_" + Date.now() + ".png");

    var folderId = "PASTE_YOUR_DRIVE_FOLDER_ID_HERE";
    if (!folderId || folderId === "PASTE_YOUR_DRIVE_FOLDER_ID_HERE") {
      return { success: false, message: "Please set Drive folder ID in saveEvidenceImage()" };
    }

    var folder = DriveApp.getFolderById(folderId);
    var file = folder.createFile(blob);

    return {
      success: true,
      fileId: file.getId(),
      fileUrl: file.getUrl()
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function saveAuditData(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var summarySheet = ss.getSheetByName("Master Audit Logs");
    if (!summarySheet) {
      summarySheet = ss.insertSheet("Master Audit Logs");
      summarySheet.appendRow(["Audit Date", "Auditor ID", "SC Name", "SC Code", "Region", "State", "City", "Total Score", "Max Score", "Percentage", "NC Count"]);
      summarySheet.setFrozenRows(1);
    }

    summarySheet.appendRow([
      payload.auditDate,
      payload.auditorId,
      payload.scName,
      payload.scCode,
      payload.region,
      payload.state,
      payload.city,
      payload.totalScore,
      payload.maxScore,
      payload.percentage,
      payload.ncCount
    ]);

    var ncSheet = ss.getSheetByName("Repeated NC data");
    if (!ncSheet) {
      ncSheet = ss.insertSheet("Repeated NC data");
      ncSheet.appendRow(["Audit Date", "Auditor Name", "SC Name", "SC Code", "Region", "State", "City", "Observation", "Score"]);
    }

    if (payload.ncList && payload.ncList.length > 0) {
      for (var i = 0; i < payload.ncList.length; i++) {
        var nc = payload.ncList[i];
        var obs = nc.checkpoint;
        if (nc.remark && nc.remark !== "N/A") obs += "\nRemark: " + nc.remark;
        if (nc.evidence) obs += "\nEvidence: " + nc.evidence;
        if (nc.evidenceImages && nc.evidenceImages.length > 0) {
          obs += "\nImages: " + nc.evidenceImages.join(" | ");
        }

        ncSheet.appendRow([
          payload.auditDate,
          payload.auditorId,
          payload.scName,
          payload.scCode,
          payload.region,
          payload.state,
          payload.city,
          obs,
          payload.percentage
        ]);
      }
    }

    var highlightsText = payload.highlights || "No highlights recorded.";
    var rcaText = payload.rca || "No RCA recorded.";

    var auditViewObs = "[AUDIT_VIEW]\nHighlights: " + highlightsText + "\nRCA & Corrective Action: " + rcaText +
      "\nTotal NCs: " + (payload.ncCount || 0);

    ncSheet.appendRow([
      payload.auditDate,
      payload.auditorId,
      payload.scName,
      payload.scCode,
      payload.region,
      payload.state,
      payload.city,
      auditViewObs,
      payload.percentage
    ]);

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAuditViewData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Repeated NC data");
    if (!sheet) return [];

    var data = sheet.getDataRange().getDisplayValues();
    var summaryRows = [];
    var rawRows = [];

    for (var i = 1; i < data.length; i++) {
      var obs = data[i][7] || "";
      if (!data[i][3] && !data[i][2]) continue;

      if (obs.indexOf("[AUDIT_VIEW]") === 0) {
        var highlights = "";
        var rca = "";

        var lines = obs.split("\n");
        lines.forEach(function (line) {
          if (line.indexOf("Highlights: ") === 0) highlights = line.replace("Highlights: ", "");
          if (line.indexOf("RCA & Corrective Action: ") === 0) rca = line.replace("RCA & Corrective Action: ", "");
        });

        summaryRows.push({
          auditDate: data[i][0] || "",
          auditorName: data[i][1] || "",
          scName: data[i][2] || "",
          scCode: data[i][3] || "",
          region: data[i][4] || "",
          state: data[i][5] || "",
          city: data[i][6] || "",
          highlights: highlights || "—",
          score: data[i][8] || "",
          rca: rca || "—",
          source: "submitted"
        });
      } else {
        rawRows.push({
          auditDate: data[i][0] || "",
          auditorName: data[i][1] || "",
          scName: data[i][2] || "",
          scCode: data[i][3] || "",
          region: data[i][4] || "",
          state: data[i][5] || "",
          city: data[i][6] || "",
          observation: obs,
          score: data[i][8] || ""
        });
      }
    }

    if (summaryRows.length > 0) {
      summaryRows.reverse();
      return summaryRows;
    }

    var grouped = {};
    rawRows.forEach(function (r) {
      var key = (r.scCode || r.scName) + "|||" + r.auditDate;
      if (!grouped[key]) {
        grouped[key] = {
          auditDate: r.auditDate,
          auditorName: r.auditorName,
          scName: r.scName,
          scCode: r.scCode,
          region: r.region,
          state: r.state,
          city: r.city,
          observations: [],
          score: r.score,
          source: "historical"
        };
      }
      if (r.observation && r.observation.trim()) grouped[key].observations.push(r.observation.trim());
      if (r.score) grouped[key].score = r.score;
    });

    var fallbackRows = Object.values(grouped).map(function (g) {
      var obsText = g.observations.join(" | ");
      return {
        auditDate: g.auditDate,
        auditorName: g.auditorName,
        scName: g.scName,
        scCode: g.scCode,
        region: g.region,
        state: g.state,
        city: g.city,
        highlights: obsText || "—",
        score: g.score,
        rca: "—",
        source: "historical"
      };
    });

    fallbackRows.sort(function (a, b) {
      return b.auditDate.localeCompare(a.auditDate);
    });

    return fallbackRows;
  } catch (e) {
    return [{ error: e.toString() }];
  }
}

function getAdminAuditConsoleData(filters) {
  try {
    var rows = getAuditViewData();

    var region = (filters && filters.region) || "ALL";
    var month = (filters && filters.month) || "ALL";
    var search = ((filters && filters.search) || "").toLowerCase();

    if (region !== "ALL") {
      rows = rows.filter(function (r) {
        return (r.region || "").toLowerCase() === region.toLowerCase();
      });
    }

    if (month !== "ALL") {
      rows = rows.filter(function (r) {
        var d = parseGsDate(r.auditDate);
        if (!d) return false;
        var key = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
        return key === month;
      });
    }

    if (search) {
      rows = rows.filter(function (r) {
        return (r.scName || "").toLowerCase().indexOf(search) !== -1 ||
          (r.scCode || "").toLowerCase().indexOf(search) !== -1 ||
          (r.auditorName || "").toLowerCase().indexOf(search) !== -1 ||
          (r.region || "").toLowerCase().indexOf(search) !== -1;
      });
    }

    return rows;
  } catch (e) {
    return [];
  }
}

function sendQAAlertEmail(emailAddress, scName, htmlBody) {
  try {
    MailApp.sendEmail({
      to: emailAddress,
      subject: 'URGENT: QA Pre-Visit Risk Alert - ' + scName,
      htmlBody: htmlBody
    });
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function sendFinalAuditReport(scCode, scName, emails, htmlBody) {
  if (!emails || emails.length === 0) return { success: false, message: "No valid SPOC email addresses found." };
  try {
    MailApp.sendEmail({
      to: emails.join(','),
      subject: 'Final Audit Report: ' + scName + ' (' + scCode + ')',
      htmlBody: htmlBody
    });
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function parseGsDate(dateStr) {
  if (!dateStr) return null;
  var str = dateStr.toString().trim();

  var parts = str.split('-');
  if (parts.length === 3) {
    var day = parseInt(parts[0], 10);
    var monthNames = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
    var monthIndex = -1;

    for (var k = 0; k < monthNames.length; k++) {
      if (monthNames[k] === String(parts[1]).trim().toLowerCase().substring(0, 3)) {
        monthIndex = k;
        break;
      }
    }

    var year = parseInt(parts[2], 10);
    if (monthIndex !== -1 && !isNaN(day) && !isNaN(year)) {
      return new Date(year, monthIndex, day);
    }
  }

  var d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function standardizeObs(freeText) {
  if (!freeText) return null;

  var text = freeText.toString().toLowerCase().trim();
  var NOISE = ["all ok", "all things are maintained", "no gaps found", "none", "nil", "satisfactory", "good", "ok"];
  for (var n = 0; n < NOISE.length; n++) {
    if (text === NOISE[n] || text.indexOf(NOISE[n]) === 0) return null;
  }

  var mappings = [
    { name: "1. Authorization Certificate", keys: ["certificate","authorization","authorised","authorized","expired certificate","certi","licence","license","renewal"] },
    { name: "2. SOPs & File Maintenance", keys: ["job sheet","jobsheet","sop","rwr","inward register","outward register","signature","signed","sign missing","challan","bulletin","file not","register not","gate pass","work order","documentation","doc not","papers"] },
    { name: "3. CRM Data & Email Hygiene", keys: ["address","gmail","email","classification","call source","customer category","dealer","imei","alternate contact","crm entry","crm not","customer detail","mobile number","contact not","wrong address","incomplete address","gmil","data entry"] },
    { name: "4. ESD Safety & Equipment", keys: ["esd","grounding","apron","wrist band","wristband","slipper","gloves","earthed","mat","voltage","polybag","poly bag","anti-static","static","esd mat","wrist strap","earthing"] },
    { name: "5. Notice Board & Escalation", keys: ["notice board","poster","escalation matrix","price list","timing","operating hours","holiday list","f1 board","display board","rate card","board not","timings not","working hours"] },
    { name: "6. Inventory & Stock", keys: ["inventory","stock","short","shortage","excess","mismatch","part","dispatched","binning","swelling","swollen","adhesive","label","spare","component","stock count","physical count","part missing"] },
    { name: "7. Branding & Signage", keys: ["branding","signage","f1 board","direction","light board","reception wall","logo","brand","fascia","glow sign","hoarding","banner"] },
    { name: "8. Uniform & Greeting", keys: ["t-shirt","tshirt","uniform","greeting","greet","id card","badge","grooming","tone","dress code","apron not","welcome","name tag","identity card"] },
    { name: "9. Customer Data Privacy", keys: ["data not","data removed","data erased","wiped","privacy","reset","data delete","personal data","customer data","data wipe","format","factory reset"] },
    { name: "10. ESD Checklist", keys: ["esd checklist","esd checksheet","esd register","esd log","esd record"] },
    { name: "11. CRM Login / ID Sharing", keys: ["crm id","single id","same id","one crm","manager id","ex-employee","inactive id","login id","shared login","common id","one id","single login","multiple use"] },
    { name: "12. 5S & Cleanliness", keys: ["5s","7s","clean","hygiene","dustbin","organized","organised","trc","dirt","rack","dust","paint peel","sweep","mop","untidy","messy","clutter","arrangement","workbench"] },
    { name: "13. TAT & Delay Issues", keys: ["15 min","15 minute","tat","less than 15","4min","5 min","7 min","delay","closed in","turnaround","time taken","repair time","pendency","pending job","aged","overdue","sla"] },
    { name: "14. First Aid & Fire Safety", keys: ["first aid","medicine","antiseptic","cream","bandage","fire","extinguisher","co2","safety kit","emergency","fire safety","first-aid"] },
    { name: "15. Tools & Lipo Safety", keys: ["lipo","lipo bag","multimeter","suction pen","camera protector","calibration","torque","coaxial","scope meter","screwdriver","tool kit","soldering","heat gun","tweezers","equipment not"] },
    { name: "16. Facility & Infrastructure", keys: ["ac","air conditioning","cctv","surveillance","power backup","electricity","water dispenser","leakage","inverter","ups","generator","toilet","restroom","air conditioner","hvac","seating","lighting","broken","damaged"] },
    { name: "17. Google Rating & Feedback", keys: ["google rating","feedback","calling feedback","customer call","nps","rating","review","satisfaction","survey","follow-up call","google review"] },
    { name: "Repair Quality & TAT", keys: ["repair not done","re-repair","repeat repair","quality of repair","poor repair","wrong diagnosis","callback","come back","return visit","rework"] },
    { name: "Customer Interaction", keys: ["rude","behaviour","behavior","attitude","misbehav","staff behav","not polite","impolite","communication","spoke rudely","customer complaint"] },
    { name: "Display & Demo", keys: ["display","demo unit","demo phone","demo device","display not working","sample unit","product display","demo not"] },
    { name: "Accessories & Spare Availability", keys: ["accessory","accessories","earphone","charger","cable","availability","not available","out of stock for accessory"] }
  ];

  for (var i = 0; i < mappings.length; i++) {
    for (var j = 0; j < mappings[i].keys.length; j++) {
      var k = mappings[i].keys[j];
      var matches = k.length > 5 ? text.indexOf(k) !== -1 : new RegExp("\\b" + k.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b").test(text);
      if (matches) return mappings[i].name;
    }
  }

  var cleaned = freeText.toString().trim().replace(/\s+/g, " ");
  var firstSentence = cleaned.split(/[.;,\n]/)[0].trim();
  return "Other: " + (firstSentence.length > 50 ? firstSentence.substring(0, 50) + "…" : firstSentence);
}

function runLocalPredictiveEngine(payload) {
  try {
    var scCode = payload.scCode || "";
    var scName = payload.scName || "Site";
    var ncs = payload.historicalNCs || [];
    var today = new Date();

    var WEAK_KEYWORDS = ["retrain","retraining","refresher","reminded","counsell","counseling","awareness","instructed","re-explained","warning issued","communication shared","staff informed","human error","supervisor reminded","team informed"];

    function isWeakCapa(text) {
      if (!text) return false;
      var lower = text.toString().toLowerCase();
      return WEAK_KEYWORDS.some(function (k) { return lower.indexOf(k) !== -1; });
    }

    var freqMap = {};
    ncs.forEach(function (nc) {
      var key = standardizeObs(nc.observation);
      if (!freqMap[key]) freqMap[key] = { count: 0, dates: [], capas: [] };
      freqMap[key].count++;
      freqMap[key].dates.push(nc.date ? new Date(nc.date) : null);
      freqMap[key].capas.push(nc.capa || "");
    });

    var totalNCs = ncs.length || 1;
    var entries = Object.keys(freqMap).map(function (k) { return { checkpoint: k, data: freqMap[k] }; }).sort(function (a, b) { return b.data.count - a.data.count; });

    function daysSince(d) {
      if (!d || isNaN(d.getTime())) return 9999;
      return Math.round((today - d) / 86400000);
    }

    var predictions = entries.map(function (entry) {
      var data = entry.data;
      var recurrence = data.count;
      var validDates = data.dates.filter(function (d) { return d && !isNaN(d.getTime()); }).sort(function (a, b) { return b - a; });
      var mostRecent = validDates[0] || null;
      var recencyDays = daysSince(mostRecent);
      var hasWeakCapa = data.capas.some(isWeakCapa);
      var hasNoCapa = data.capas.filter(function (c) { return c && c.trim(); }).length === 0;
      var paretoShare = (recurrence / totalNCs) * 100;

      var repeatAfterCapa = hasWeakCapa && recurrence >= 2;
      var score = 0;

      score += Math.min(recurrence * 15, 45);
      score += recencyDays < 90 ? 25 : recencyDays < 180 ? 15 : recencyDays < 365 ? 8 : 2;
      score += hasWeakCapa ? 20 : 0;
      score += hasNoCapa ? 15 : 0;
      score += paretoShare > 15 ? 10 : 0;
      score += repeatAfterCapa ? 5 : 0;

      var level = recurrence < 2 ? "LOW" : score >= 60 ? "HIGH" : score >= 30 ? "MEDIUM" : "LOW";
      var capaEff = hasNoCapa ? "Unknown" : hasWeakCapa ? "Weak / Superficial" : "Potentially Effective";
      var recentStr = mostRecent ? Utilities.formatDate(mostRecent, Session.getScriptTimeZone(), "dd-MMM-yyyy") : "Unknown";

      return {
        checkpoint: entry.checkpoint,
        level: level,
        recurrence: recurrence,
        historicalContext: "Failed " + recurrence + "x. Last: " + recentStr + ". Pareto: " + paretoShare.toFixed(1) + "%. CAPA: " + capaEff + ".",
        insight: buildInsight(level, recencyDays, recurrence, hasWeakCapa, repeatAfterCapa, paretoShare),
        action: getFieldAction(entry.checkpoint)
      };
    });

    var ranked = predictions.filter(function (p) { return p.level === "HIGH"; })
      .concat(predictions.filter(function (p) { return p.level === "MEDIUM"; }))
      .concat(predictions.filter(function (p) { return p.level === "LOW"; }));

    var highCount = ranked.filter(function (p) { return p.level === "HIGH"; }).length;
    var mediumCount = ranked.filter(function (p) { return p.level === "MEDIUM"; }).length;

    var tableRows = ranked.map(function (p) {
      var color = p.level === "HIGH" ? "#c0392b" : p.level === "MEDIUM" ? "#e67e22" : "#27ae60";
      var badge = '<span style="background:' + color + ';color:#fff;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:bold;">' + p.level + '</span>';
      return '<tr><td style="padding:10px;border:1px solid #333;font-weight:bold;font-size:13px;color:#fff;">' + p.checkpoint + '</td>' +
        '<td style="padding:10px;border:1px solid #333;font-size:12px;color:#ccc;">' + p.historicalContext + '</td>' +
        '<td style="padding:10px;border:1px solid #333;text-align:center;">' + badge + '</td>' +
        '<td style="padding:10px;border:1px solid #333;font-size:12px;color:#ccc;">' + p.insight + '</td>' +
        '<td style="padding:10px;border:1px solid #333;font-size:12px;color:#ccc;">' + p.action + '</td></tr>';
    }).join('');

    var noDataMsg = ncs.length === 0 ? '<p style="color:#22c55e;"><strong>No historical NCs recorded.</strong> Conduct a standard comprehensive audit.</p>' : '';
    var execSummary = '<p style="color:#ccc;">Analysis of <strong style="color:#fff;">' + ncs.length + '</strong> records for <strong style="color:#fff;">' + scName + ' (' + scCode + ')</strong> — <span style="color:#c0392b;font-weight:bold;">' + highCount + ' HIGH</span> · <span style="color:#e67e22;font-weight:bold;">' + mediumCount + ' MEDIUM</span> risk checkpoints.</p>';

    var html = '<h3 style="color:#EA1515;margin-bottom:16px;">QA Pre-Visit Risk Alert</h3>' +
      '<h4 style="color:#fff;border-bottom:1px solid #2A2A2A;padding-bottom:8px;margin-bottom:12px;">Summary</h4>' +
      execSummary + noDataMsg;

    if (ranked.length > 0) {
      html += '<h4 style="color:#fff;border-bottom:1px solid #2A2A2A;padding-bottom:8px;margin:20px 0 12px;">Predictive Risk Table</h4>' +
        '<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;background:#0A0A0A;">' +
        '<thead><tr style="background:#111;"><th style="padding:10px;border:1px solid #2A2A2A;color:#666;font-size:10px;letter-spacing:0.1em;text-transform:uppercase;">Checkpoint</th>' +
        '<th style="padding:10px;border:1px solid #2A2A2A;color:#666;font-size:10px;letter-spacing:0.1em;text-transform:uppercase;">Context</th>' +
        '<th style="padding:10px;border:1px solid #2A2A2A;color:#666;font-size:10px;letter-spacing:0.1em;text-transform:uppercase;text-align:center;">Risk</th>' +
        '<th style="padding:10px;border:1px solid #2A2A2A;color:#666;font-size:10px;letter-spacing:0.1em;text-transform:uppercase;">Insight</th>' +
        '<th style="padding:10px;border:1px solid #2A2A2A;color:#666;font-size:10px;letter-spacing:0.1em;text-transform:uppercase;">Field Action</th>' +
        '</tr></thead><tbody>' + tableRows + '</tbody></table></div>';
    }

    html += '<h4 style="color:#fff;margin-top:20px;margin-bottom:8px;">Tactical Directives</h4>' +
      '<ul style="color:#ccc;line-height:2;padding-left:20px;">' +
      '<li>Dedicate at least <strong style="color:#fff;">40% of audit time</strong> to High-Risk checkpoints.</li>' +
      '<li>For Repeat NCs closed via retraining, <strong style="color:#fff;">physically witness</strong> process execution.</li>' +
      '<li>Collect <strong style="color:#fff;">physical or digital evidence</strong> for all high-risk checkpoints.</li>' +
      '</ul>' +
      '<p style="border-top:1px solid #2A2A2A;padding-top:12px;margin-top:16px;font-style:italic;color:#666;font-size:12px;">Do not miss or ignore remaining audit points. Maintain a comprehensive review of all areas.</p>';

    return html;
  } catch (e) {
    return '<strong style="color:red;">Engine Error:</strong> ' + e.message;
  }
}

function buildInsight(level, recencyDays, recurrence, hasWeakCapa, repeatAfterCapa, paretoShare) {
  var parts = [];
  if (repeatAfterCapa) parts.push("Band-Aid CAPA risk: failed again after retraining-based closure.");
  if (recencyDays < 90) parts.push("Very recent failure (within 90 days) — likely still unresolved.");
  else if (recencyDays < 180) parts.push("Failed within last 6 months.");
  if (recurrence >= 3) parts.push("Chronic recurrence: " + recurrence + " failures — systemic gap.");
  if (paretoShare > 15) parts.push("Top Pareto contributor (" + paretoShare.toFixed(1) + "% of NCs).");
  if (hasWeakCapa && !repeatAfterCapa) parts.push("Previous corrective action was superficial.");
  if (parts.length === 0) parts.push("Monitor closely; pattern not strong enough for high-confidence prediction.");
  return parts.join(" ");
}

function getFieldAction(checkpoint) {
  var actions = {
    "1.": "Physically inspect authorization certificate on-wall. Check expiry date against today.",
    "2.": "Sample 5 random job sheets. Verify signatures, dates, RWR completion, and bulletin acknowledgment.",
    "3.": "Pull 10 CRM records. Check address completeness, email format, call source, and IMEI tagging.",
    "4.": "Test ESD wrist strap with continuity tester. Inspect mat grounding and staff aprons.",
    "5.": "Photograph notice board. Verify escalation matrix version, updated price list, and operating hours.",
    "6.": "Cross-check physical stock count against CRM. Isolate and flag swollen batteries.",
    "7.": "Photograph all external and internal signage. Verify F1 board illumination and logo standards.",
    "8.": "Observe staff during live customer interaction. Check ID cards and uniform compliance.",
    "9.": "Request last 3 device handover records. Verify data wipe completion certificates.",
    "10.": "Ask for ESD checklist register. Verify daily entries for past 7 days are signed.",
    "11.": "Log into CRM. Confirm each active staff member has a unique, personal login ID.",
    "12.": "Conduct 5S walkthrough. Photograph dustbin, workstation, and parts rack condition.",
    "13.": "Pull TAT report from CRM. Flag any job closed in under 15 minutes without justification.",
    "14.": "Open first aid box — verify all items present and within expiry. Check extinguisher seal.",
    "15.": "Power on multimeter. Check calibration sticker on torque screwdriver. Inspect lipo bag.",
    "16.": "Test AC operation, CCTV live feed, and power backup during visit. Document with photos.",
    "17.": "Check live Google rating. Ask staff to demonstrate the post-service feedback calling process."
  };

  for (var prefix in actions) {
    if (checkpoint.indexOf(prefix) === 0) return actions[prefix];
  }

  return "Physically verify compliance with direct observation and documentary evidence.";
}

function getAdminReportViewData() {
  return getAuditViewData();
}
