const SHEET_ID        = '1XddJWtMdMzePaGWLekbvFBqgEbvWOVh2qJfsUA8g2SI';
const LOG_SHEET       = 'Registros';
const CONFIRMED_SHEET = 'confirmados';
const GUESTS_SHEET    = 'Invitados';
const ASISTENTES_SHEET= 'asistentes';
const DESAFIO_SHEET   = 'desafio';

function doGet(e) {
  if (!e || !e.parameter) return jsonResponse({ error: 'no params' });
  return handleRequest(e);
}

function doPost(e) {
  if (!e || !e.parameter) return jsonResponse({ error: 'no params' });
  return handleRequest(e);
}

function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

function handleRequest(e) {
  var params = e.parameter;
  var action = params.action;
  try {
    if (action === 'check')          return checkQR(params.qr);
    if (action === 'register')       return registerScan(params.qr, params.username, params.timestamp);
    if (action === 'add_guests')     return addGuests(params);
    if (action === 'check_player')   return checkPlayer(params.qr);
    if (action === 'save_score')     return saveScore(params);
    if (action === 'leaderboard')    return getLeaderboard();
    if (action === 'search_player')  return searchPlayer(params.query);
    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

// ── ACCESS CONTROL ───────────────────────────────────────

function checkQR(qr) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var logSheet = ss.getSheetByName(LOG_SHEET);
  if (logSheet && logSheet.getLastRow() > 1) {
    var logData = logSheet.getDataRange().getValues();
    for (var i = 1; i < logData.length; i++) {
      if (String(logData[i][2]).trim() === String(qr).trim())
        return jsonResponse({ status: 'already_scanned' });
    }
  }
  var confSheet = ss.getSheetByName(CONFIRMED_SHEET);
  if (confSheet) {
    var confData = confSheet.getDataRange().getValues();
    var headers = confData[0].map(function(h) { return String(h).toLowerCase().trim(); });
    var qrCol      = headers.indexOf('qr');
    var nameCol    = headers.indexOf('name');
    var statusCol  = headers.indexOf('status');
    var inviteeCol = headers.indexOf('invitee_name');
    for (var j = 1; j < confData.length; j++) {
      if (String(confData[j][qrCol]).trim() === String(qr).trim()) {
        return jsonResponse({
          status:      'confirmed',
          name:        confData[j][nameCol],
          userStatus:  confData[j][statusCol],
          inviteeName: inviteeCol >= 0 ? String(confData[j][inviteeCol]).trim() : ''
        });
      }
    }
  }
  return jsonResponse({ status: 'not_found' });
}

function registerScan(qr, username, timestamp) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var logSheet = ss.getSheetByName(LOG_SHEET);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET);
    logSheet.appendRow(['Timestamp', 'Username', 'QR Data', 'Name', 'Status', 'Invitee Name']);
  }
  var name = '', userStatus = '', inviteeName = '';
  var confSheet = ss.getSheetByName(CONFIRMED_SHEET);
  if (confSheet) {
    var confData = confSheet.getDataRange().getValues();
    var headers = confData[0].map(function(h) { return String(h).toLowerCase().trim(); });
    var qrCol      = headers.indexOf('qr');
    var nameCol    = headers.indexOf('name');
    var statusCol  = headers.indexOf('status');
    var inviteeCol = headers.indexOf('invitee_name');
    for (var i = 1; i < confData.length; i++) {
      if (String(confData[i][qrCol]).trim() === String(qr).trim()) {
        name        = confData[i][nameCol];
        userStatus  = confData[i][statusCol];
        inviteeName = inviteeCol >= 0 ? String(confData[i][inviteeCol]).trim() : '';
        break;
      }
    }
  }
  logSheet.appendRow([timestamp, username, qr, name, userStatus, inviteeName]);
  return jsonResponse({ status: 'ok' });
}

function addGuests(params) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var gSheet = ss.getSheetByName(GUESTS_SHEET);
  if (!gSheet) {
    gSheet = ss.insertSheet(GUESTS_SHEET);
    gSheet.appendRow(['Timestamp', 'Operador', 'QR Host', 'Host Name', 'Cantidad', 'Nombre Invitado', 'Contacto']);
  }
  gSheet.appendRow([
    params.timestamp || new Date().toISOString(),
    params.username  || '',
    params.qr        || '',
    params.host_name || '',
    parseInt(params.count) || 1,
    params.guest_name || '',
    params.contact    || ''
  ]);
  return jsonResponse({ status: 'ok', count: parseInt(params.count) || 1 });
}

// ── DESAFÍO ──────────────────────────────────────────────

function checkPlayer(qr) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var name = '', email = '';

  var aSheet = ss.getSheetByName(ASISTENTES_SHEET);
  if (aSheet) {
    var aData = aSheet.getDataRange().getValues();
    var ah    = aData[0].map(function(h) { return String(h).toLowerCase().trim(); });
    var aqr   = ah.indexOf('qr');
    var an    = ah.indexOf('name');
    var ae    = ah.indexOf('email');
    for (var i = 1; i < aData.length; i++) {
      if (String(aData[i][aqr]).trim() === String(qr).trim()) {
        name  = an >= 0 ? String(aData[i][an]) : '';
        email = ae >= 0 ? String(aData[i][ae]) : '';
        break;
      }
    }
  }

  if (!name) {
    var cSheet = ss.getSheetByName(CONFIRMED_SHEET);
    if (cSheet) {
      var cData = cSheet.getDataRange().getValues();
      var ch    = cData[0].map(function(h) { return String(h).toLowerCase().trim(); });
      var cqr   = ch.indexOf('qr');
      var cn    = ch.indexOf('name');
      for (var j = 1; j < cData.length; j++) {
        if (String(cData[j][cqr]).trim() === String(qr).trim()) {
          name = String(cData[j][cn]);
          break;
        }
      }
    }
  }

  var attempts = 0, bestScore = 0;
  var dSheet = ss.getSheetByName(DESAFIO_SHEET);
  if (dSheet && dSheet.getLastRow() > 1) {
    var dData = dSheet.getDataRange().getValues();
    var dh    = dData[0].map(function(h) { return String(h).toLowerCase().trim(); });
    var dqr   = dh.indexOf('qr');
    var ds    = dh.indexOf('score');
    for (var k = 1; k < dData.length; k++) {
      if (String(dData[k][dqr]).trim() === String(qr).trim()) {
        attempts++;
        var sc = parseInt(dData[k][ds]) || 0;
        if (sc > bestScore) bestScore = sc;
      }
    }
  }

  if (!name) return jsonResponse({ status: 'not_found' });
  return jsonResponse({ status: 'ok', name: name, email: email, attempts: attempts, bestScore: bestScore });
}

function saveScore(params) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var dSheet = ss.getSheetByName(DESAFIO_SHEET);
  if (!dSheet) {
    dSheet = ss.insertSheet(DESAFIO_SHEET);
    dSheet.appendRow(['Timestamp', 'Operador', 'QR', 'Name', 'Email', 'Score', 'Shots', 'Attempt']);
  }

  var attempt = 1;
  if (dSheet.getLastRow() > 1) {
    var dData = dSheet.getDataRange().getValues();
    var dh    = dData[0].map(function(h) { return String(h).toLowerCase().trim(); });
    var dqr   = dh.indexOf('qr');
    for (var i = 1; i < dData.length; i++) {
      if (String(dData[i][dqr]).trim() === String(params.qr).trim()) attempt++;
    }
  }

  dSheet.appendRow([
    params.timestamp || new Date().toISOString(),
    params.username  || '',
    params.qr        || '',
    params.name      || '',
    params.email     || '',
    parseInt(params.score) || 0,
    params.shots     || '',
    attempt
  ]);

  var rank = getRankForQR(ss, params.qr, parseInt(params.score) || 0);
  return jsonResponse({ status: 'ok', rank: rank, attempt: attempt });
}

function getRankForQR(ss, qr, newScore) {
  var dSheet = ss.getSheetByName(DESAFIO_SHEET);
  if (!dSheet || dSheet.getLastRow() < 2) return 1;
  var dData   = dSheet.getDataRange().getValues();
  var dh      = dData[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var dqr     = dh.indexOf('qr');
  var ds      = dh.indexOf('score');
  var bestMap = {};
  for (var i = 1; i < dData.length; i++) {
    var pqr = String(dData[i][dqr]).trim();
    var sc  = parseInt(dData[i][ds]) || 0;
    if (!bestMap[pqr] || sc > bestMap[pqr]) bestMap[pqr] = sc;
  }
  if (!bestMap[qr] || newScore > bestMap[qr]) bestMap[qr] = newScore;
  var scores = Object.keys(bestMap).map(function(k) { return bestMap[k]; });
  scores.sort(function(a, b) { return b - a; });
  var rank = 1;
  for (var j = 0; j < scores.length; j++) {
    if (scores[j] > newScore) rank++;
  }
  return rank;
}

function getLeaderboard() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var dSheet = ss.getSheetByName(DESAFIO_SHEET);
  if (!dSheet || dSheet.getLastRow() < 2)
    return jsonResponse({ status: 'ok', leaderboard: [], total: 0 });

  var dData   = dSheet.getDataRange().getValues();
  var dh      = dData[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var dqr     = dh.indexOf('qr');
  var ds      = dh.indexOf('score');
  var dn      = dh.indexOf('name');
  var de      = dh.indexOf('email');
  var dat     = dh.indexOf('attempt');
  var bestMap = {};

  for (var i = 1; i < dData.length; i++) {
    var pqr = String(dData[i][dqr]).trim();
    var sc  = parseInt(dData[i][ds])  || 0;
    var att = parseInt(dData[i][dat]) || 1;
    if (!bestMap[pqr]) {
      bestMap[pqr] = { qr: pqr, score: sc, name: String(dData[i][dn]), email: String(dData[i][de]), attempts: att };
    } else {
      if (sc  > bestMap[pqr].score)    bestMap[pqr].score    = sc;
      if (att > bestMap[pqr].attempts) bestMap[pqr].attempts = att;
    }
  }

  var arr = Object.keys(bestMap).map(function(k) { return bestMap[k]; });
  arr.sort(function(a, b) { return b.score - a.score; });
  var top = arr.slice(0, 10).map(function(p, i) {
    return { rank: i + 1, name: p.name, email: p.email, score: p.score, attempts: p.attempts, qr: p.qr };
  });
  return jsonResponse({ status: 'ok', leaderboard: top, total: arr.length });
}

function searchPlayer(query) {
  if (!query || query.length < 2) return jsonResponse({ status: 'ok', results: [] });
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var dSheet = ss.getSheetByName(DESAFIO_SHEET);
  if (!dSheet || dSheet.getLastRow() < 2) return jsonResponse({ status: 'ok', results: [] });

  var dData = dSheet.getDataRange().getValues();
  var dh    = dData[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var dqr   = dh.indexOf('qr');
  var ds    = dh.indexOf('score');
  var dn    = dh.indexOf('name');
  var de    = dh.indexOf('email');
  var q     = String(query).toLowerCase();

  var bestMap = {};
  for (var i = 1; i < dData.length; i++) {
    var pqr    = String(dData[i][dqr]).trim();
    var pname  = String(dData[i][dn]);
    var pemail = String(dData[i][de]);
    var sc     = parseInt(dData[i][ds]) || 0;
    if (pname.toLowerCase().indexOf(q) >= 0 || pemail.toLowerCase().indexOf(q) >= 0) {
      if (!bestMap[pqr] || sc > bestMap[pqr].score)
        bestMap[pqr] = { qr: pqr, name: pname, email: pemail, score: sc };
    }
  }

  var allMap = {};
  for (var j = 1; j < dData.length; j++) {
    var aqr = String(dData[j][dqr]).trim();
    var asc = parseInt(dData[j][ds]) || 0;
    if (!allMap[aqr] || asc > allMap[aqr]) allMap[aqr] = asc;
  }
  var allScores = Object.keys(allMap).map(function(k) { return allMap[k]; });
  allScores.sort(function(a, b) { return b - a; });

  var results = Object.keys(bestMap).map(function(k) {
    var p    = bestMap[k];
    var rank = 1;
    for (var r = 0; r < allScores.length; r++) {
      if (allScores[r] > p.score) rank++; else break;
    }
    return { name: p.name, email: p.email, score: p.score, rank: rank, total: allScores.length };
  });
  results.sort(function(a, b) { return a.rank - b.rank; });
  return jsonResponse({ status: 'ok', results: results });
}

// ── HELPERS ──────────────────────────────────────────────

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
