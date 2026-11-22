// ==========================================
// 01_Database.gs : 単一スプレッドシート接続・初期化・統合管理
// ==========================================
var _cachedMainSs = null;
var _dbSchemaChecked = false;

function getMainSpreadsheet() {
  if (_cachedMainSs) return _cachedMainSs;
  var id = String(APP_CONFIG.SPREADSHEET_ID || MAIN_SPREADSHEET_ID || '').trim();
  if (id && id !== ScriptApp.getScriptId()) {
    _cachedMainSs = SpreadsheetApp.openById(id);
    return _cachedMainSs;
  }
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    _cachedMainSs = active;
    return _cachedMainSs;
  }
  throw new Error('有効なスプレッドシートIDが設定されていません');
}

function getWorkSpreadsheet() {
  return getMainSpreadsheet();
}

function getPrivateSpreadsheet() {
  return getMainSpreadsheet();
}

function getCanonicalSheetSpecs_() {
  return [
    {
      key: 'WORK_LOG',
      name: APP_CONFIG.SHEETS.WORK_LOG,
      tabHex: '#1e3a8a',
      bgHex: '#e2e8f0',
      fontHex: '#0f172a',
      headers: [
        'ID', '日付', '出勤時間', '退勤時間', '休憩(分)', '実働(時間)', '残業(時間)',
        'タイムスケジュール', '成果(Keep)', '課題(Problem)', '学び(Insight)', '改善(Try)',
        'メモ・本文', '写真URL', '成長スコア', '体調', '睡眠(時間)', '一言メモ', '投稿者', '更新日時',
        '仕事習慣JSON'
      ],
      widths: [130, 110, 85, 85, 80, 85, 85, 260, 200, 200, 200, 200, 240, 180, 85, 80, 85, 180, 100, 140, 180]
    },
    {
      key: 'PRIV_LOG',
      name: APP_CONFIG.SHEETS.PRIV_LOG,
      tabHex: '#4f46e5',
      bgHex: '#e0e7ff',
      fontHex: '#1e1b4b',
      headers: [
        'ID', '日付', '学習・活動スケジュール', '習慣達成JSON', '私用メモ・本文', '写真URL', '充実度スコア', '更新日時',
        '就寝時刻', '起床時刻', '睡眠時間', '体重', '最高血圧', '最低血圧'
      ],
      widths: [130, 110, 260, 180, 240, 180, 90, 140, 85, 85, 85, 80, 85, 85]
    },
    {
      key: 'SOLO_SNS',
      name: APP_CONFIG.SHEETS.SOLO_SNS,
      tabHex: '#0284c7',
      bgHex: '#0f172a',
      fontHex: '#ffffff',
      headers: ['ID', '投稿日時', '更新日時', '本文', 'タグ', 'ピン留め', '画像数', '拡張JSON'],
      widths: [140, 145, 145, 340, 160, 75, 75, 200]
    },
    {
      key: 'TODO_LIST',
      name: APP_CONFIG.SHEETS.TODO_LIST,
      tabHex: '#059669',
      bgHex: '#0f172a',
      fontHex: '#ffffff',
      headers: ['ID', '完了', 'タスク名', 'カテゴリ', '期限', '優先度', '作成日時', '拡張JSON'],
      widths: [140, 70, 280, 110, 110, 85, 140, 220]
    },
    {
      key: 'WORK_HABIT',
      name: APP_CONFIG.SHEETS.WORK_HABIT,
      tabHex: '#2563eb',
      bgHex: '#dbeafe',
      fontHex: '#1e3a8a',
      headers: ['ID', '習慣名', '通常目標', '最小目標', 'カテゴリ', '有効', '作成日時'],
      widths: [120, 190, 240, 220, 100, 70, 140]
    },
    {
      key: 'PRIV_HABIT',
      name: APP_CONFIG.SHEETS.PRIV_HABIT,
      tabHex: '#7c3aed',
      bgHex: '#ede9fe',
      fontHex: '#4c1d95',
      headers: ['ID', '習慣名', '通常目標', '最小目標', 'カテゴリ', '有効', '作成日時'],
      widths: [120, 190, 240, 220, 100, 70, 140]
    },
    {
      key: 'WORK_GOAL',
      name: APP_CONFIG.SHEETS.WORK_GOAL,
      tabHex: '#334155',
      bgHex: '#f1f5f9',
      fontHex: '#0f172a',
      headers: ['ID', '種別', '対象期間', '目標内容', '更新日時', '開始日', '締切日', '進捗率', '次の一手'],
      widths: [130, 100, 110, 280, 140, 105, 105, 80, 220]
    },
    {
      key: 'PRIV_GOAL',
      name: APP_CONFIG.SHEETS.PRIV_GOAL,
      tabHex: '#6d28d9',
      bgHex: '#f3e8ff',
      fontHex: '#3b0764',
      headers: ['ID', '種別', '対象期間', '目標内容', '更新日時', '開始日', '締切日', '進捗率', '次の一手'],
      widths: [130, 100, 110, 280, 140, 105, 105, 80, 220]
    },
    {
      key: 'APP_SETTINGS',
      name: APP_CONFIG.SHEETS.APP_SETTINGS,
      tabHex: '#475569',
      bgHex: '#0f172a',
      fontHex: '#ffffff',
      headers: ['設定キー', '設定値JSON', '最終更新日時'],
      widths: [200, 420, 150]
    }
  ];
}

function ensureSheetSchema_(ss, sheetName, headers, bgHex, tabHex, fontHex) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (tabHex) {
    try { sheet.setTabColor(tabHex); } catch (e) {}
  }
  var lastRow = sheet.getLastRow();
  var firstCell = lastRow > 0 ? String(sheet.getRange(1, 1).getValue() || '') : '';
  if (lastRow === 0 || !firstCell) {
    var hdrRange = sheet.getRange(1, 1, 1, headers.length);
    hdrRange.setValues([headers])
      .setFontWeight('bold')
      .setBackground(bgHex || '#f1f5f9')
      .setFontColor(fontHex || (bgHex === '#0f172a' ? '#ffffff' : '#0f172a'))
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < headers.length) {
    var startCol = sheet.getLastColumn() + 1;
    var addHeaders = headers.slice(startCol - 1);
    sheet.getRange(1, startCol, 1, addHeaders.length)
      .setValues([addHeaders])
      .setFontWeight('bold')
      .setBackground(bgHex || '#f1f5f9')
      .setFontColor(fontHex || (bgHex === '#0f172a' ? '#ffffff' : '#0f172a'))
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
  }
  return sheet;
}

function setupAllDatabases(forceCheck) {
  if (_dbSchemaChecked && !forceCheck) return;
  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var ss = getMainSpreadsheet();
  var specs = getCanonicalSheetSpecs_();

  for (var i = 0; i < specs.length; i++) {
    var sp = specs[i];
    ensureSheetSchema_(ss, sp.name, sp.headers, sp.bgHex, sp.tabHex, sp.fontHex);
  }

  var workHabitSheet = ss.getSheetByName(APP_CONFIG.SHEETS.WORK_HABIT);
  if (workHabitSheet && workHabitSheet.getLastRow() < 2) {
    workHabitSheet.getRange(2, 1, 2, 7).setValues([
      ['whab_1', '朝イチ優先順位付け', '重要タスク3つと時間枠を決める', '最優先タスクを1つだけ決める', '仕事', 'TRUE', nowStr],
      ['whab_2', '退勤前リセット', '机上整理＋明日のTryを記入', '明日の開始タスクを1行書く', '仕事', 'TRUE', nowStr]
    ]);
  }

  var privHabitSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PRIV_HABIT);
  if (privHabitSheet && privHabitSheet.getLastRow() < 2) {
    privHabitSheet.getRange(2, 1, 2, 7).setValues([
      ['hab_1', '読書・インプット', '30分読書して1行メモ', '1ページだけ読む', '学習', 'TRUE', nowStr],
      ['hab_2', '運動・ストレッチ', '20分ウォーキングまたは筋トレ', '深呼吸とストレッチ3分', '健康', 'TRUE', nowStr]
    ]);
  }

  var defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch (e) {}
  }

  if (forceCheck) {
    organizeIntegratedSpreadsheet();
  }

  _dbSchemaChecked = true;
}

function organizeIntegratedSpreadsheet() {
  var ss = getMainSpreadsheet();
  var specs = getCanonicalSheetSpecs_();
  var canonicalNames = {};
  var orderedSheets = [];

  for (var i = 0; i < specs.length; i++) {
    var sp = specs[i];
    canonicalNames[sp.name] = true;
    var sh = ensureSheetSchema_(ss, sp.name, sp.headers, sp.bgHex, sp.tabHex, sp.fontHex);

    // 見出し行の再整形
    var hdrRange = sh.getRange(1, 1, 1, sp.headers.length);
    hdrRange.setValues([sp.headers])
      .setFontWeight('bold')
      .setFontSize(10)
      .setBackground(sp.bgHex)
      .setFontColor(sp.fontHex)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setWrap(true);
    hdrRange.setBorder(null, null, true, null, null, null, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.setRowHeight(1, 30);
    sh.setFrozenRows(1);

    // 列幅最適化
    if (sp.widths) {
      for (var c = 0; c < sp.widths.length; c++) {
        try { sh.setColumnWidth(c + 1, sp.widths[c]); } catch (e) {}
      }
    }

    // 右側の完全な空白余剰列を削除（見通し改善）
    var maxCols = sh.getMaxColumns();
    var keepCols = Math.max(sp.headers.length, sh.getLastColumn());
    if (maxCols > keepCols) {
      try { sh.deleteColumns(keepCols + 1, maxCols - keepCols); } catch (e) {}
    }

    // タブ並び順を確定
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
    orderedSheets.push(sp.name);
  }

  // 不要シート（シート1 / Sheet1 / 完全空の非標準シート）を削除
  var deletedSheets = [];
  var allSheets = ss.getSheets();
  for (var j = allSheets.length - 1; j >= 0; j--) {
    var s = allSheets[j];
    var sName = s.getName();
    if (canonicalNames[sName]) continue;
    var isDefaultName = /^(シート|Sheet)\d+$/i.test(sName);
    var isCompletelyEmpty = (s.getLastRow() === 0 && s.getLastColumn() === 0);
    if ((isDefaultName || isCompletelyEmpty) && ss.getSheets().length > 1) {
      try {
        ss.deleteSheet(s);
        deletedSheets.push(sName);
      } catch (e) {}
    }
  }

  SpreadsheetApp.flush();
  return {
    ok: true,
    spreadsheetName: ss.getName(),
    orderedSheets: orderedSheets,
    deletedSheets: deletedSheets
  };
}

function getInitialBundle() {
  setupAllDatabases();
  var ss = getMainSpreadsheet();
  var ssName = ss.getName();
  var ssUrl = ss.getUrl();

  return {
    workLogs: getWorkLogs(),
    privateLogs: getPrivateLogs(ss),
    workHabits: getHabitsFromSheet(ss.getSheetByName(APP_CONFIG.SHEETS.WORK_HABIT)),
    workGoals: getGoalsFromSheet(ss.getSheetByName(APP_CONFIG.SHEETS.WORK_GOAL)),
    privateHabits: getHabitsFromSheet(ss.getSheetByName(APP_CONFIG.SHEETS.PRIV_HABIT)),
    privateGoals: getGoalsFromSheet(ss.getSheetByName(APP_CONFIG.SHEETS.PRIV_GOAL)),
    notes: loadNormalizedNotes_(),
    todos: loadNormalizedTodos_(),
    meta: {
      unified: true,
      spreadsheetId: APP_CONFIG.SPREADSHEET_ID,
      workSheetName: ssName,
      workSheetUrl: ssUrl,
      privateSheetName: ssName,
      privateSheetUrl: ssUrl,
      serverTime: new Date().toISOString()
    }
  };
}

function getFullAppBundle() {
  return getInitialBundle();
}

function getFolder() {
  var folders = DriveApp.getFoldersByName(APP_CONFIG.PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(APP_CONFIG.PHOTO_FOLDER_NAME);
}