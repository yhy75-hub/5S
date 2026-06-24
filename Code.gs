// ================================================================
// 事務所 5Sチェックシート - Google Apps Script
// GitHub Pages の index.html から fetch() で呼び出すREST API
// ================================================================

const CFG = {
  members: ['平林','森崎','里見','堀江','細江','林','専務','山本','本城','横塚','四ツ木','有側','伊藤','松谷','毛利'],
  privileged: ['伊藤A','本城A','林A','事務所管理者','5Sリーダー'],
  qMonths: [3, 6, 9, 12],
};

// ================================================================
// エントリーポイント（GETはデータ取得、POSTは書き込み）
// ================================================================
function doGet(e) {
  const p = e.parameter;
  let result;
  try {
    const y = +p.year, m = +p.month;
    // スコアはJSON文字列で渡ってくるのでパース
    const s1 = p.s1 ? JSON.parse(p.s1) : null;
    const s2 = p.s2 ? JSON.parse(p.s2) : null;

    if      (p.action === 'getConfig')     result = CFG;
    else if (p.action === 'getMemberData') result = getMemberData(y, m, p.name);
    else if (p.action === 'getResults')    result = getResults(y, m);
    else if (p.action === 'authenticate')  result = { role: authenticate(p.name, p.pin) };
    else if (p.action === 'saveScores')    result = saveScores(y, m, p.name, s1, s2, p.role);
    else if (p.action === 'submitScores')  result = submitScores(y, m, p.name, s1, s2, p.role);
    else result = { error: 'Unknown action: ' + p.action };
  } catch(err) {
    result = { error: err.message };
  }
  return jsonResponse_(result);
}

function doPost(e) {
  let body, result;
  try {
    body = JSON.parse(e.postData.contents);
    const { action } = body;
    if      (action === 'authenticate')  result = { role: authenticate(body.name, body.pin) };
    else if (action === 'getMemberData') result = getMemberData(body.year, body.month, body.name);
    else if (action === 'saveScores')    result = saveScores(body.year, body.month, body.name, body.s1, body.s2, body.role);
    else if (action === 'submitScores')  result = submitScores(body.year, body.month, body.name, body.s1, body.s2, body.role);
    else if (action === 'getResults')    result = getResults(body.year, body.month);
    else result = { error: 'Unknown action: ' + action };
  } catch(err) {
    result = { error: err.message };
  }
  return jsonResponse_(result);
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
// 認証
// ================================================================
function authenticate(name, pin) {
  if (CFG.privileged.includes(name)) {
    const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || '5s2026';
    return pin === stored ? 'priv' : null;
  }
  if (CFG.members.includes(name)) return 'member';
  return null;
}

// ================================================================
// データ操作
// ================================================================
function getMemberData(year, month, name) {
  const sh = getOrCreateSheet_(year, month);
  const rowIdx = findRow_(sh, name);
  if (!rowIdx) {
    return { s1: Array(10).fill(null), s2: Array(10).fill(null), submitted: false, submittedAt: null };
  }
  const row = sh.getRange(rowIdx, 1, 1, 24).getValues()[0];
  return {
    s1: row.slice(1, 11).map(v => v === '' ? null : Number(v)),
    s2: row.slice(11, 21).map(v => v === '' ? null : Number(v)),
    submitted: row[21] === true || row[21] === 'TRUE',
    submittedAt: row[22] ? String(row[22]) : null,
  };
}

function saveScores(year, month, name, s1, s2, role) {
  const sh = getOrCreateSheet_(year, month);
  if (isSubmitted_(sh, name)) return { success: false, message: '既に提出済みです。' };
  writeRow_(sh, name, s1, s2, false, '', role);
  return { success: true };
}

function submitScores(year, month, name, s1, s2, role) {
  const sh = getOrCreateSheet_(year, month);
  if (isSubmitted_(sh, name)) return { success: false, message: '既に提出済みです。' };
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  const stamp = role === 'priv' ? now + '（管理者入力）' : now;
  writeRow_(sh, name, s1, s2, true, stamp, role);
  return { success: true, submittedAt: stamp };
}

function getResults(year, month) {
  const sh = getOrCreateSheet_(year, month);
  const data = sh.getDataRange().getValues();
  const results = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    results[String(row[0])] = {
      s1: row.slice(1, 11).map(v => v === '' ? null : Number(v)),
      s2: row.slice(11, 21).map(v => v === '' ? null : Number(v)),
      submitted: row[21] === true || row[21] === 'TRUE',
      submittedAt: row[22] ? String(row[22]) : null,
    };
  }
  return results;
}

// ================================================================
// PINコード設定（スクリプトエディタから直接実行）
// ★ PINを変えたいときは下の文字列を書き換えてから実行
// ================================================================
function setAdminPin() {
  const PIN = '5s2026'; // ← ここを変更する
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', PIN);
  Logger.log('PINコードを設定しました: ' + PIN);
}

// ================================================================
// 内部ヘルパー
// ================================================================
function getOrCreateSheet_(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = `${year}-${String(month).padStart(2, '0')}`;
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    const headers = ['名前',
      'S1_1','S1_2','S1_3','S1_4','S1_5','S1_6','S1_7','S1_8','S1_9','S1_10',
      'S2_1','S2_2','S2_3','S2_4','S2_5','S2_6','S2_7','S2_8','S2_9','S2_10',
      '提出済み','提出日時','ロール'];
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setBackground('#1d4ed8').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sh;
}

function findRow_(sh, name) {
  const vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === name) return i + 1;
  }
  return null;
}

function isSubmitted_(sh, name) {
  const rowIdx = findRow_(sh, name);
  if (!rowIdx) return false;
  const val = sh.getRange(rowIdx, 22).getValue();
  return val === true || val === 'TRUE';
}

function writeRow_(sh, name, s1, s2, submitted, submittedAt, role) {
  const rowData = [name,
    ...s1.map(v => v != null ? v : ''),
    ...s2.map(v => v != null ? v : ''),
    submitted, submittedAt, role];
  const rowIdx = findRow_(sh, name);
  if (rowIdx) {
    sh.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sh.appendRow(rowData);
  }
  if (submitted) {
    const ri = findRow_(sh, name);
    sh.getRange(ri, 22).setBackground('#dcfce7');
  }
}
