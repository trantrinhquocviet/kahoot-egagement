// Code.gs — DUA Edu Career Quest
// Deploy as Web App: Execute as Me, Access: Anyone

const WEBHOOK_URL = "";

function getOrCreateSheet(name, headers, headerColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground(headerColor || "#123A34")
      .setFontColor("#ffffff");
  }
  return sheet;
}

// ── Import danh sách email vào sheet Account ──
function importAccounts(emails) {
  const sheet = getOrCreateSheet("Account", ["Email", "Imported At"], "#16A34A");

  // Xoá data cũ (giữ header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  const timestamp = new Date().toLocaleString("vi-VN", { timeZone: "Asia/Ho_Chi_Minh" });
  const rows = emails.map(email => [email, timestamp]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

// ── Ghi điểm vào sheet Score ──
function saveScore(email, scoreData, session) {
  const sheet = getOrCreateSheet(
    "Score",
    ["Timestamp", "Session", "Email", "Data", "Biz", "Eng", "Creative", "Total", "Career"],
    "#0EA5A4"
  );

  const timestamp = new Date().toLocaleString("vi-VN", { timeZone: "Asia/Ho_Chi_Minh" });
  const total = Math.max(scoreData.data, scoreData.biz, scoreData.eng, scoreData.creative);
  const career = Object.entries(scoreData).reduce((a, b) => a[1] > b[1] ? a : b)[0];
  const sess = session || "";

  // Upsert: tìm row cùng session + email
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === sess && String(data[i][2]).toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i + 1, 1, 1, 9).setValues([[
        timestamp, sess, email, scoreData.data, scoreData.biz, scoreData.eng, scoreData.creative, total, career
      ]]);
      return;
    }
  }
  sheet.appendRow([timestamp, sess, email, scoreData.data, scoreData.biz, scoreData.eng, scoreData.creative, total, career]);
}

// ── Lobby: check-in & game state ──
function checkInLobby(session, email) {
  const sheet = getOrCreateSheet(
    "Lobby",
    ["Session", "Email", "Checked In At"],
    "#f59e0b"
  );
  // Xoá check-in cũ của email này trong session
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(session) && String(data[i][1]).toLowerCase() === email.toLowerCase()) {
      sheet.deleteRow(i + 1);
    }
  }
  const timestamp = new Date().toLocaleString("vi-VN", { timeZone: "Asia/Ho_Chi_Minh" });
  sheet.appendRow([session, email, timestamp]);
}

function getLobbyState(session) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Lấy game state từ Config sheet
  let configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    configSheet = ss.insertSheet("Config");
    configSheet.appendRow(["Key", "Value"]);
    configSheet.appendRow(["game_state", "waiting"]);
    configSheet.appendRow(["game_session", ""]);
  }
  const configData = configSheet.getDataRange().getValues();
  let state = "waiting";
  let gameSession = "";
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0] === "game_state") state = configData[i][1];
    if (configData[i][0] === "game_session") gameSession = configData[i][1];
  }

  // Lấy danh sách players trong lobby của session này
  const lobbySheet = ss.getSheetByName("Lobby");
  let players = [];
  if (lobbySheet && lobbySheet.getLastRow() > 1) {
    const rows = lobbySheet.getRange(2, 1, lobbySheet.getLastRow() - 1, 2).getDisplayValues();
    players = [...new Set(rows.filter(r => String(r[0]) === String(session)).map(r => String(r[1])).filter(Boolean))];
  }

  return { state, gameSession, players };
}

function setGameState(state, session) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    configSheet = ss.insertSheet("Config");
    configSheet.appendRow(["Key", "Value"]);
  }
  const data = configSheet.getDataRange().getValues();
  let foundState = false, foundSession = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === "game_state") { configSheet.getRange(i + 1, 2).setValue(state); foundState = true; }
    if (data[i][0] === "game_session") { configSheet.getRange(i + 1, 2).setValue(session || ""); foundSession = true; }
  }
  if (!foundState) configSheet.appendRow(["game_state", state]);
  if (!foundSession) configSheet.appendRow(["game_session", session || ""]);
}

// ── Quản lý Sessions ──
function createSession(sessionId, sessionName) {
  const sheet = getOrCreateSheet("Sessions", ["Session ID", "Session Name", "Created At", "Status"], "#ea580c");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(sessionId)) {
      sheet.getRange(i + 1, 2).setValue(sessionName);
      sheet.getRange(i + 1, 4).setValue("active");
      return;
    }
  }
  const timestamp = new Date().toLocaleString("vi-VN", { timeZone: "Asia/Ho_Chi_Minh" });
  sheet.appendRow([sessionId, sessionName, timestamp, "active"]);
}

function getActiveSessions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sessions");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return rows
    .filter(r => String(r[3]).toLowerCase() === "active" && String(r[0]))
    .map(r => ({ id: String(r[0]), name: String(r[1]), createdAt: String(r[2]) }));
}

function getAllSessions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sessions");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return rows
    .filter(r => String(r[0]))
    .map(r => ({ id: String(r[0]), name: String(r[1]), createdAt: String(r[2]), status: String(r[3]) }))
    .reverse(); // newest first
}

function endSession(sessionId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sessions");
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(sessionId)) {
      sheet.getRange(i + 1, 4).setValue("ended");
    }
  }
}


// ── Web App endpoint ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === "importAccounts") {
      importAccounts(data.emails);
    } else if (data.action === "saveScore") {
      saveScore(data.email, data.score);
    } else if (data.action === "saveAnswer") {
      saveAnswer(data.session, data.questionIdx, data.email, data.answerIdx);
    } else if (data.action === "checkIn") {
      checkInLobby(data.session, data.email);
    } else if (data.action === "setGameState") {
      setGameState(data.state, data.session);
    } else if (data.action === "createSession") {
      createSession(data.sessionId, data.sessionName);
    } else if (data.action === "endSession") {
      endSession(data.sessionId);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Lưu câu trả lời của học viên ──
function saveAnswer(session, questionIdx, email, answerIdx) {
  const sheet = getOrCreateSheet(
    "Answers",
    ["Session", "Question", "Email", "Answer", "Timestamp"],
    "#6366f1"
  );
  const timestamp = new Date().toLocaleString("vi-VN", { timeZone: "Asia/Ho_Chi_Minh" });
  sheet.appendRow([session, questionIdx, email, answerIdx, timestamp]);
}

// ── Lấy thống kê câu trả lời theo session + câu hỏi ──
function getQuestionStats(session, questionIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Answers");
  const votes = [0, 0, 0, 0];
  if (!sheet || sheet.getLastRow() < 2) return votes;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  rows.forEach(r => {
    if (String(r[0]) === String(session) && Number(r[1]) === Number(questionIdx)) {
      const idx = Number(r[3]);
      if (idx >= 0 && idx <= 3) votes[idx]++;
    }
  });
  return votes;
}

// ── Lấy danh sách điểm từ sheet Score (có thể lọc theo session) ──
function getScores(session) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Score");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  return rows
    .filter(r => r[2] && (!session || String(r[1]) === String(session)))
    .map(r => ({
      session: r[1],
      email: r[2],
      data: r[3],
      biz: r[4],
      eng: r[5],
      creative: r[6],
      total: r[7],
      career: r[8]
    }))
    .sort((a, b) => b.total - a.total);
}

function doGet(e) {
  if (e.parameter.action === "getAccounts") {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Account");
    let emails = [];
    if (sheet && sheet.getLastRow() > 1) {
      emails = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
        .map(r => String(r[0]).trim().toLowerCase())
        .filter(e => e.includes("@"));
    }
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, emails }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === "checkIn") {
    checkInLobby(e.parameter.session, e.parameter.email);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === "setGameState") {
    setGameState(e.parameter.state, e.parameter.session);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === "getLobbyState") {
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, ...getLobbyState(e.parameter.session) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === "getActiveSessions") {
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, sessions: getActiveSessions() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === "getScores") {
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, scores: getScores(e.parameter.session || "") }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === "getAllSessions") {
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, sessions: getAllSessions() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === "getQuestionStats") {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        votes: getQuestionStats(e.parameter.session, e.parameter.q)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === "saveAnswer") {
    saveAnswer(e.parameter.session, e.parameter.q, e.parameter.email, e.parameter.ans);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === "saveScore") {
    const scoreObj = {
      data: Number(e.parameter.data)||0,
      biz: Number(e.parameter.biz)||0,
      eng: Number(e.parameter.eng)||0,
      creative: Number(e.parameter.creative)||0
    };
    saveScore(e.parameter.email, scoreObj, e.parameter.session);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === "createSession") {
    createSession(e.parameter.sessionId, e.parameter.sessionName);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}
