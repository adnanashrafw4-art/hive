// ============================================================
// HIVE Attendance Tracker — Google Apps Script Backend v2
// ============================================================
// SETUP INSTRUCTIONS:
// 1. Go to script.google.com → New Project → paste this code
// 2. Replace SPREADSHEET_ID below with your Google Sheet ID
// 3. Deploy → New Deployment → Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Copy the Web App URL into the HTML file (via the ⚙ Configure button)
// ============================================================

const SPREADSHEET_ID   = "YOUR_SPREADSHEET_ID_HERE"; // ← REPLACE THIS
const MEMBERS_SHEET    = "Members";
const SESSIONS_SHEET   = "Sessions";
const ATTENDANCE_SHEET = "Attendance";
const GUESTS_SHEET     = "Guests";
const SUMMARY_SHEET    = "MeetingSummaries";

// ── Entry points ──────────────────────────────────────────────
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const p      = e.parameter || {};
  const body   = e.postData  ? JSON.parse(e.postData.contents || "{}") : p;
  const action = p.action    || body.action;
  let result;
  try {
    switch (action) {
      case "getMembers":              result = getMembers();                       break;
      case "getSessions":             result = getSessions();                      break;
      case "getAttendance":           result = getAttendance(p);                  break;
      case "addMember":               result = addMember(p);                      break;
      case "addSession":              result = addSession(p);                      break;
      case "markAttendance":          result = markAttendance(p);                 break;
      case "getStats":                result = getStats();                         break;
      case "saveGuests":              result = saveGuests(p);                      break;
      case "saveSummary":             result = saveSummary(p);                     break;
      case "getAllGuestsAndSummaries": result = getAllGuestsAndSummaries();         break;
      case "setup":                   result = setupSheets();                      break;
      default:                        result = { error: "Unknown action: " + action };
    }
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Sheet helper ──────────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ── Setup ─────────────────────────────────────────────────────
function setupSheets() {
  const sheets = [
    { name: MEMBERS_SHEET,    headers: ["MemberID","Name","Phone","Role","JoinDate","Active"] },
    { name: SESSIONS_SHEET,   headers: ["SessionID","Title","Date","Location","Notes","CreatedAt"] },
    { name: ATTENDANCE_SHEET, headers: ["AttendanceID","SessionID","MemberID","Status","MarkedAt"] },
    { name: GUESTS_SHEET,     headers: ["SessionID","GuestsJSON","UpdatedAt"] },
    { name: SUMMARY_SHEET,    headers: ["SessionID","Highlights","Decisions","ActionItemsJSON","UpdatedAt"] },
  ];
  sheets.forEach(({ name, headers }) => {
    const sh = getSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#F5C518");
    }
  });
  return { success: true, message: "All sheets initialized!" };
}

// ── Members ───────────────────────────────────────────────────
function getMembers() {
  const data = getSheet(MEMBERS_SHEET).getDataRange().getValues();
  if (data.length <= 1) return { members: [] };
  const headers = data[0];
  const members = data.slice(1).map(row => {
    const obj = {}; headers.forEach((h,i) => obj[h] = row[i]); return obj;
  }).filter(m => m.Active !== false && m.Active !== "FALSE" && m.Active !== 0);
  return { members };
}

function addMember(p) {
  const sheet = getSheet(MEMBERS_SHEET);
  const id    = "M" + Date.now();
  sheet.appendRow([id, p.name||"", p.phone||"", p.role||"Member", new Date().toISOString().split("T")[0], true]);
  return { success: true, memberId: id };
}

// ── Sessions ──────────────────────────────────────────────────
function getSessions() {
  const data = getSheet(SESSIONS_SHEET).getDataRange().getValues();
  if (data.length <= 1) return { sessions: [] };
  const headers = data[0];
  const sessions = data.slice(1).map(row => {
    const obj = {}; headers.forEach((h,i) => obj[h] = row[i]); return obj;
  });
  sessions.sort((a,b) => new Date(b.Date) - new Date(a.Date));
  return { sessions };
}

function addSession(p) {
  const sheet = getSheet(SESSIONS_SHEET);
  const id    = "S" + Date.now();
  sheet.appendRow([id, p.title||"HIVE Meeting", p.date||new Date().toISOString().split("T")[0], p.location||"Dubai, UAE", p.notes||"", new Date().toISOString()]);
  return { success: true, sessionId: id };
}

// ── Attendance ────────────────────────────────────────────────
function getAttendance(p) {
  const data = getSheet(ATTENDANCE_SHEET).getDataRange().getValues();
  if (data.length <= 1) return { attendance: [] };
  const headers = data[0];
  let att = data.slice(1).map(row => { const obj={}; headers.forEach((h,i)=>obj[h]=row[i]); return obj; });
  if (p.sessionId) att = att.filter(a => a.SessionID === p.sessionId);
  return { attendance: att };
}

function markAttendance(p) {
  const sheet     = getSheet(ATTENDANCE_SHEET);
  const data      = sheet.getDataRange().getValues();
  const headers   = data[0];
  const sidIdx    = headers.indexOf("SessionID");
  const midIdx    = headers.indexOf("MemberID");
  const statusIdx = headers.indexOf("Status");
  const markedIdx = headers.indexOf("MarkedAt");
  for (let i = 1; i < data.length; i++) {
    if (data[i][sidIdx] === p.sessionId && data[i][midIdx] === p.memberId) {
      sheet.getRange(i+1, statusIdx+1).setValue(p.status);
      sheet.getRange(i+1, markedIdx+1).setValue(new Date().toISOString());
      return { success: true, action: "updated" };
    }
  }
  sheet.appendRow(["A"+Date.now(), p.sessionId, p.memberId, p.status, new Date().toISOString()]);
  return { success: true, action: "created" };
}

// ── Guests ────────────────────────────────────────────────────
function saveGuests(p) {
  const sheet    = getSheet(GUESTS_SHEET);
  const data     = sheet.getDataRange().getValues();
  const headers  = data[0];
  const sidIdx   = headers.indexOf("SessionID");
  for (let i = 1; i < data.length; i++) {
    if (data[i][sidIdx] === p.sessionId) {
      sheet.getRange(i+1, 2).setValue(p.guests);
      sheet.getRange(i+1, 3).setValue(new Date().toISOString());
      return { success: true, action: "updated" };
    }
  }
  sheet.appendRow([p.sessionId, p.guests, new Date().toISOString()]);
  return { success: true, action: "created" };
}

// ── Meeting Summaries ─────────────────────────────────────────
function saveSummary(p) {
  const sheet   = getSheet(SUMMARY_SHEET);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const sidIdx  = headers.indexOf("SessionID");
  for (let i = 1; i < data.length; i++) {
    if (data[i][sidIdx] === p.sessionId) {
      sheet.getRange(i+1, 2).setValue(p.highlights   || "");
      sheet.getRange(i+1, 3).setValue(p.decisions    || "");
      sheet.getRange(i+1, 4).setValue(p.actionItems  || "[]");
      sheet.getRange(i+1, 5).setValue(new Date().toISOString());
      return { success: true, action: "updated" };
    }
  }
  sheet.appendRow([p.sessionId, p.highlights||"", p.decisions||"", p.actionItems||"[]", new Date().toISOString()]);
  return { success: true, action: "created" };
}

// ── Get all guests + summaries (bulk load) ─────────────────────
function getAllGuestsAndSummaries() {
  const gData  = getSheet(GUESTS_SHEET).getDataRange().getValues();
  const sData  = getSheet(SUMMARY_SHEET).getDataRange().getValues();

  const guests  = {};
  if (gData.length > 1) {
    gData.slice(1).forEach(row => {
      const sid = row[0];
      if (sid) {
        try { guests[sid] = JSON.parse(row[1] || "[]"); }
        catch { guests[sid] = []; }
      }
    });
  }

  const summary = {};
  if (sData.length > 1) {
    sData.slice(1).forEach(row => {
      const sid = row[0];
      if (sid) {
        let actionItems = [];
        try { actionItems = JSON.parse(row[3] || "[]"); } catch {}
        summary[sid] = { highlights: row[1]||"", decisions: row[2]||"", actionItems };
      }
    });
  }

  return { guests, summary };
}

// ── Stats ─────────────────────────────────────────────────────
function getStats() {
  const membersData    = getMembers().members;
  const sessionsData   = getSessions().sessions;
  const attendanceData = getAttendance({}).attendance;

  // Count total guests
  const gData  = getSheet(GUESTS_SHEET).getDataRange().getValues();
  let totalGuests = 0;
  if (gData.length > 1) {
    gData.slice(1).forEach(row => {
      try { totalGuests += JSON.parse(row[1] || "[]").length; } catch {}
    });
  }

  const totalMembers  = membersData.length;
  const totalSessions = sessionsData.length;

  const sessionStats = sessionsData.map(s => {
    const records = attendanceData.filter(a => a.SessionID === s.SessionID);
    const present = records.filter(a => a.Status === "present").length;
    const excused = records.filter(a => a.Status === "excused").length;
    const absent  = records.filter(a => a.Status === "absent").length;
    return {
      sessionId: s.SessionID, title: s.Title, date: s.Date,
      present, excused, absent, total: records.length,
      rate: totalMembers > 0 ? Math.round((present / totalMembers) * 100) : 0
    };
  });

  const memberStats = membersData.map(m => {
    const records = attendanceData.filter(a => a.MemberID === m.MemberID);
    const present = records.filter(a => a.Status === "present").length;
    return {
      memberId: m.MemberID, name: m.Name, role: m.Role,
      present, total: totalSessions,
      rate: totalSessions > 0 ? Math.round((present / totalSessions) * 100) : 0
    };
  });

  const overallRate = totalSessions > 0 && totalMembers > 0
    ? Math.round(attendanceData.filter(a => a.Status === "present").length / (totalSessions * totalMembers) * 100)
    : 0;

  return { totalMembers, totalSessions, overallRate, totalGuests, sessionStats, memberStats };
}
