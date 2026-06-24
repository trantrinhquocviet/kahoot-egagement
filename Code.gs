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
function saveScore(email, scoreData) {
  const sheet = getOrCreateSheet(
    "Score",
    ["Timestamp", "Email", "Data", "Biz", "Eng", "Creative", "Total", "Career"],
    "#0EA5A4"
  );

  const timestamp = new Date().toLocaleString("vi-VN", { timeZone: "Asia/Ho_Chi_Minh" });
  const total = Math.max(scoreData.data, scoreData.biz, scoreData.eng, scoreData.creative);
  const career = Object.entries(scoreData).reduce((a, b) => a[1] > b[1] ? a : b)[0];

  // Tìm row cũ của email này và cập nhật
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i + 1, 1, 1, 8).setValues([[
        timestamp, email, scoreData.data, scoreData.biz, scoreData.eng, scoreData.creative, total, career
      ]]);
      return;
    }
  }

  // Chưa có → append mới
  sheet.appendRow([timestamp, email, scoreData.data, scoreData.biz, scoreData.eng, scoreData.creative, total, career]);
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
    configSheet.appendRow(["game_state", "waiting"]);
    configSheet.appendRow(["game_session", ""]);
  }
  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === "game_state") configSheet.getRange(i + 1, 2).setValue(state);
    if (data[i][0] === "game_session") configSheet.getRange(i + 1, 2).setValue(session || "");
  }
}

// ── Cờ kích hoạt buổi học ──
function setSessionActive(active) {
  const sheet = getOrCreateSheet("Active", ["Active", "Updated At"], "#7C3AED");
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  const timestamp = new Date().toLocaleString("vi-VN", { timeZone: "Asia/Ho_Chi_Minh" });
  sheet.appendRow([active ? "true" : "false", timestamp]);
}

function getSessionActive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Active");
  if (!sheet || sheet.getLastRow() < 2) return false;
  return String(sheet.getRange(2, 1).getValue()).trim().toLowerCase() === "true";
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
    } else if (data.action === "setSessionActive") {
      setSessionActive(data.active);
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

// ── Lấy danh sách điểm từ sheet Score ──
function getScores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Score");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  return rows
    .filter(r => r[1])
    .map(r => ({
      email: r[1],
      data: r[2],
      biz: r[3],
      eng: r[4],
      creative: r[5],
      total: r[6],
      career: r[7]
    }))
    .sort((a, b) => b.total - a.total);
}

function doGet(e) {
  if (e.parameter.action === "getSessionActive") {
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, active: getSessionActive() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
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
  if (e.parameter.action === "getScores") {
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, scores: getScores() }))
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
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}
