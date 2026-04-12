// ============================================================
// HIVE Attendance Tracker — Google Apps Script Backend v3
// OPTIMIZED: Single SS open per request, bulk load on init
// ============================================================

const SPREADSHEET_ID   = "1ZefudqafwpHT4r6iXcMim1H_j6a8UKwj-IcDIIq9EEU";
const MEMBERS_SHEET    = "Members";
const SESSIONS_SHEET   = "Sessions";
const ATTENDANCE_SHEET = "Attendance";
const GUESTS_SHEET     = "Guests";
const SUMMARY_SHEET    = "MeetingSummaries";

// Cache the spreadsheet object per execution
let _ss = null;
function ss() {
  if (!_ss) _ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ss;
}
function getSheet(name) {
  return ss().getSheetByName(name) || ss().insertSheet(name);
}

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
      case "loadAll":          result = loadAll();            break; // ← NEW: single bulk load
      case "addMember":        result = addMember(p);         break;
      case "addSession":       result = addSession(p);        break;
      case "markAttendance":   result = markAttendance(p);    break;
      case "saveGuests":       result = saveGuests(p);        break;
      case "saveSummary":      result = saveSummary(p);       break;
      case "setup":            result = setupSheets();        break;
      default:                 result = { error: "Unknown action: " + action };
    }
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── BULK LOAD — one call returns everything ───────────────────
// This replaces getMembers + getSessions + getStats + getAllGuestsAndSummaries
function loadAll() {
  // Open spreadsheet once, read all sheets in parallel
  const membersData    = getSheet(MEMBERS_SHEET).getDataRange().getValues();
  const sessionsData   = getSheet(SESSIONS_SHEET).getDataRange().getValues();
  const attendanceData = getSheet(ATTENDANCE_SHEET).getDataRange().getValues();
  const guestsData     = getSheet(GUESTS_SHEET).getDataRange().getValues();
  const summaryData    = getSheet(SUMMARY_SHEET).getDataRange().getValues();

  // Parse members
  const members = parseRows(membersData).filter(m =>
    m.Active !== false && m.Active !== "FALSE" && m.Active !== 0
  );

  // Parse sessions (newest first)
  const sessions = parseRows(sessionsData).sort((a,b) => new Date(b.Date) - new Date(a.Date));

  // Parse attendance
  const attendance = parseRows(attendanceData);

  // Parse guests
  const guests = {};
  if (guestsData.length > 1) {
    guestsData.slice(1).forEach(row => {
      if (row[0]) {
        try { guests[row[0]] = JSON.parse(row[1] || "[]"); }
        catch { guests[row[0]] = []; }
      }
    });
  }

  // Parse summaries
  const summary = {};
  if (summaryData.length > 1) {
    summaryData.slice(1).forEach(row => {
      if (row[0]) {
        let actionItems = [];
        try { actionItems = JSON.parse(row[3] || "[]"); } catch {}
        summary[row[0]] = { highlights: row[1]||"", decisions: row[2]||"", actionItems };
      }
    });
  }

  // Compute stats inline (no extra sheet reads)
  const totalMembers  = members.length;
  const totalSessions = sessions.length;
  let totalGuests = 0;
  Object.values(guests).forEach(g => totalGuests += g.length);

  const sessionStats = sessions.map(s => {
    const records = attendance.filter(a => a.SessionID === s.SessionID);
    const present = records.filter(a => a.Status === "present").length;
    const excused = records.filter(a => a.Status === "excused").length;
    const absent  = records.filter(a => a.Status === "absent").length;
    return {
      sessionId: s.SessionID, title: s.Title, date: s.Date,
      present, excused, absent, total: records.length,
      rate: totalMembers > 0 ? Math.round((present / totalMembers) * 100) : 0
    };
  });

  const memberStats = members.map(m => {
    const records = attendance.filter(a => a.MemberID === m.MemberID);
    const present = records.filter(a => a.Status === "present").length;
    return {
      memberId: m.MemberID, name: m.Name, role: m.Role,
      present, total: totalSessions,
      rate: totalSessions > 0 ? Math.round((present / totalSessions) * 100) : 0
    };
  });

  const overallRate = totalSessions > 0 && totalMembers > 0
    ? Math.round(attendance.filter(a => a.Status === "present").length / (totalSessions * totalMembers) * 100)
    : 0;

  return {
    members, sessions, attendance, guests, summary,
    stats: { totalMembers, totalSessions, overallRate, totalGuests, sessionStats, memberStats }
  };
}

// ── Add Member ────────────────────────────────────────────────
function addMember(p) {
  const sheet = getSheet(MEMBERS_SHEET);
  const id    = "M" + Date.now();
  sheet.appendRow([id, p.name||"", p.phone||"", p.role||"Member", new Date().toISOString().split("T")[0], true]);
  return { success: true, memberId: id };
}

// ── Add Session ───────────────────────────────────────────────
function addSession(p) {
  const sheet = getSheet(SESSIONS_SHEET);
  const id    = "S" + Date.now();
  sheet.appendRow([id, p.title||"HIVE Meeting", p.date||new Date().toISOString().split("T")[0], p.location||"Dubai, UAE", p.notes||"", new Date().toISOString()]);
  return { success: true, sessionId: id };
}

// ── Mark Attendance ───────────────────────────────────────────
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

// ── Save Guests ───────────────────────────────────────────────
function saveGuests(p) {
  const sheet  = getSheet(GUESTS_SHEET);
  const data   = sheet.getDataRange().getValues();
  const sidIdx = data[0] ? data[0].indexOf("SessionID") : 0;
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

// ── Save Summary ──────────────────────────────────────────────
function saveSummary(p) {
  const sheet  = getSheet(SUMMARY_SHEET);
  const data   = sheet.getDataRange().getValues();
  const sidIdx = data[0] ? data[0].indexOf("SessionID") : 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][sidIdx] === p.sessionId) {
      sheet.getRange(i+1, 2).setValue(p.highlights  || "");
      sheet.getRange(i+1, 3).setValue(p.decisions   || "");
      sheet.getRange(i+1, 4).setValue(p.actionItems || "[]");
      sheet.getRange(i+1, 5).setValue(new Date().toISOString());
      return { success: true, action: "updated" };
    }
  }
  sheet.appendRow([p.sessionId, p.highlights||"", p.decisions||"", p.actionItems||"[]", new Date().toISOString()]);
  return { success: true, action: "created" };
}

// ── Setup Sheets ──────────────────────────────────────────────
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
      sh.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#F5C518");
    }
  });
  return { success: true, message: "All sheets initialized!" };
}

// ── Helper: parse sheet rows into objects ─────────────────────
function parseRows(data) {
  if (!data || data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => obj[h] = row[i]);
    return obj;
  });
}
