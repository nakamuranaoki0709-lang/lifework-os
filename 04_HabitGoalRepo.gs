// ==========================================
// 04_HabitGoalRepo.gs : 習慣マスタ ＆ 目標管理 ＆ アプリ設定同期
// ==========================================
function getHabitsFromSheet(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  var habits = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || !row[1]) continue;
    var activeStr = String(row[5] || 'TRUE').toUpperCase();
    if (activeStr === 'FALSE') continue;

    habits.push({
      id: String(row[0]),
      name: String(row[1]),
      normalGoal: String(row[2] || ''),
      minGoal: String(row[3] || ''),
      category: String(row[4] || 'any:daily')
    });
  }
  return habits;
}

function saveHabitMaster(payload) {
  if (Array.isArray(payload)) payload = payload[0];
  setupAllDatabases();
  var isWork = payload.domain === 'work';
  var sheet = isWork
    ? getWorkSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.WORK_HABIT)
    : getPrivateSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.PRIV_HABIT);

  var data = sheet.getDataRange().getDisplayValues();
  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  if (payload.deleteId) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.deleteId)) {
        sheet.getRange(i + 1, 6).setValue('FALSE');
        break;
      }
    }
    return { success: true, habits: getHabitsFromSheet(sheet) };
  }

  var id = payload.id || ((isWork ? 'whab_' : 'hab_') + new Date().getTime());
  var foundRow = -1;
  for (var j = 1; j < data.length; j++) {
    if (String(data[j][0]) === String(id)) {
      foundRow = j + 1;
      break;
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 2, 1, 4).setValues([[payload.name, payload.normalGoal, payload.minGoal, payload.category || 'any:daily']]);
  } else {
    sheet.appendRow([id, payload.name, payload.normalGoal, payload.minGoal, payload.category || 'any:daily', 'TRUE', nowStr]);
  }

  return { success: true, habits: getHabitsFromSheet(sheet) };
}

function reorderHabits(payload) {
  if (Array.isArray(payload)) payload = payload[0];
  if (!payload || !Array.isArray(payload.orderedIds)) return { status: 'skip' };
  setupAllDatabases();
  var isWork = payload.domain === 'work';
  var sheet = isWork
    ? getWorkSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.WORK_HABIT)
    : getPrivateSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.PRIV_HABIT);

  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return { status: 'ok' };
  var rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var rowMap = {};
  rows.forEach(function(r) { rowMap[String(r[0])] = r; });

  var sorted = [];
  var used = {};
  payload.orderedIds.forEach(function(id) {
    var k = String(id);
    if (rowMap[k] && !used[k]) {
      sorted.push(rowMap[k]);
      used[k] = true;
    }
  });
  rows.forEach(function(r) {
    var k = String(r[0]);
    if (!used[k]) sorted.push(r);
  });

  if (sorted.length > 0) {
    sheet.getRange(2, 1, sorted.length, 7).setValues(sorted);
  }
  return { status: 'ok' };
}

function getGoalsFromSheet(sheet) {
  if (!sheet) return { current: { yearly: null, monthly: null, weekly: null }, history: [] };
  var data = sheet.getDataRange().getDisplayValues();
  var rawData = sheet.getDataRange().getValues();
  var current = { yearly: null, monthly: null, weekly: null };
  var history = [];

  if (data.length > 1) {
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var rRow = rawData[i];
      var type = String(row[1] || '').toLowerCase();
      var content = String(row[3] || '');
      if (!content) continue;

      var item = {
        id: String(row[0]),
        type: type,
        period: String(row[2] || ''),
        content: content,
        updatedAt: String(row[4] || ''),
        startDate: cleanDateStr(rRow[5] || row[5] || ''),
        endDate: cleanDateStr(rRow[6] || row[6] || ''),
        progress: (row[7] !== '' && row[7] !== undefined) ? Number(row[7]) : 0,
        nextAction: String(row[8] || '')
      };

      if (current.hasOwnProperty(type) && current[type] === null) {
        current[type] = item;
      }
      history.push(item);
    }
  }
  return { current: current, history: history };
}

function saveGoalItem(payload) {
  if (Array.isArray(payload)) payload = payload[0];
  setupAllDatabases();
  var sheet = payload.domain === 'private'
    ? getPrivateSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.PRIV_GOAL)
    : getWorkSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.WORK_GOAL);

  var now = new Date();
  var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var data = sheet.getDataRange().getDisplayValues();

  if (payload.updateExistingId) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.updateExistingId)) {
        if (payload.content !== undefined) sheet.getRange(i + 1, 4).setValue(payload.content);
        sheet.getRange(i + 1, 5).setValue(timeStr);
        if (payload.startDate !== undefined) sheet.getRange(i + 1, 6).setValue(payload.startDate);
        if (payload.endDate !== undefined) sheet.getRange(i + 1, 7).setValue(payload.endDate);
        if (payload.progress !== undefined) sheet.getRange(i + 1, 8).setValue(Number(payload.progress));
        if (payload.nextAction !== undefined) sheet.getRange(i + 1, 9).setValue(payload.nextAction);
        return getGoalsFromSheet(sheet);
      }
    }
  }

  var labelMap = { yearly: '年間目標', monthly: '月間目標', weekly: '週間目標' };
  var id = 'goal_' + now.getTime() + '_' + payload.type;
  sheet.appendRow([
    id,
    payload.type,
    labelMap[payload.type] || '目標',
    String(payload.content || '').trim(),
    timeStr,
    payload.startDate || '',
    payload.endDate || '',
    Number(payload.progress || 0),
    payload.nextAction || ''
  ]);

  return getGoalsFromSheet(sheet);
}

// アプリ設定シート（夢リスト・クレド・タイムカプセル・カスタム項目）のKey-Value永続化
function saveAppSettingKv(payload) {
  if (Array.isArray(payload)) payload = payload[0];
  if (!payload || !payload.key) return { ok: false };
  setupAllDatabases();
  var ss = getMainSpreadsheet();
  var sheet = ss.getSheetByName(APP_CONFIG.SHEETS.APP_SETTINGS);
  var keyStr = String(payload.key);
  var jsonStr = (typeof payload.value === 'string') ? payload.value : JSON.stringify(payload.value);
  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === keyStr) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[jsonStr, nowStr]]);
      return { ok: true };
    }
  }
  sheet.appendRow([keyStr, jsonStr, nowStr]);
  return { ok: true };
}

function getAllAppSettingsKv() {
  setupAllDatabases();
  var ss = getMainSpreadsheet();
  var sheet = ss.getSheetByName(APP_CONFIG.SHEETS.APP_SETTINGS);
  var res = {};
  if (!sheet || sheet.getLastRow() < 2) return res;
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][0] || '');
    var raw = String(data[i][1] || '');
    if (!k) continue;
    try { res[k] = JSON.parse(raw); } catch (e) { res[k] = raw; }
  }
  return res;
}