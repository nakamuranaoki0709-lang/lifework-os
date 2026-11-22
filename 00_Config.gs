// ==========================================
// 00_Config.gs : システム全体の設定・スキーマ定義
// ==========================================
var MAIN_SPREADSHEET_ID = '1-9GalbQnpczX7IlbxAu2yyRaERapIxr1OsvEaez3les';

var APP_CONFIG = {
  API_TOKEN: 'Ti3pWxoR6zOjwqleDPufm7sG',
  SPREADSHEET_ID: MAIN_SPREADSHEET_ID,
  PHOTO_FOLDER_NAME: '日報添付写真',
  SHEETS: {
    WORK_LOG: '日報_勤怠ログ',
    WORK_HABIT: '仕事_習慣マスタ',
    WORK_GOAL: '目標履歴',
    PRIV_LOG: '私用_日次ログ',
    PRIV_HABIT: '習慣マスタ',
    PRIV_GOAL: '私用_目標履歴',
    SOLO_SNS: '一人SNS_投稿',
    TODO_LIST: 'タスク_ToDo',
    APP_SETTINGS: 'アプリ設定_マスタ'
  }
};

function cleanDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  var m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  }
  return s.slice(0, 10);
}

function cleanTimeStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  }
  var s = String(val).trim();
  var m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) {
    return ('0' + m[1]).slice(-2) + ':' + m[2];
  }
  return '';
}

function calculateWorkHours(start, end, breakMin) {
  if (!start || !end) return { workHours: 0, overtimeHours: 0 };
  var s = start.split(':');
  var e = end.split(':');
  if (s.length < 2 || e.length < 2) return { workHours: 0, overtimeHours: 0 };
  var startM = parseInt(s[0], 10) * 60 + parseInt(s[1], 10);
  var endM = parseInt(e[0], 10) * 60 + parseInt(e[1], 10);
  if (endM < startM) endM += 24 * 60;
  var netM = Math.max(0, endM - startM - (Number(breakMin) || 0));
  var workH = Math.round((netM / 60) * 100) / 100;
  var overH = Math.max(0, Math.round(((netM - 480) / 60) * 100) / 100);
  return { workHours: workH, overtimeHours: overH };
}

function calculateSleepHours(bedTime, wakeTime) {
  if (!bedTime || !wakeTime) return null;
  var b = bedTime.split(':');
  var w = wakeTime.split(':');
  if (b.length < 2 || w.length < 2) return null;
  var bM = parseInt(b[0], 10) * 60 + parseInt(b[1], 10);
  var wM = parseInt(w[0], 10) * 60 + parseInt(w[1], 10);
  if (wM <= bM) wM += 24 * 60;
  return Math.round(((wM - bM) / 60) * 10) / 10;
}