// ==========================================
// 03_DailyRepo.gs : 仕事・私用の日次ログ＆打刻＆写真メモ
// ==========================================
function getWorkLogs() {
  var sheet = getWorkSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.WORK_LOG);
  if (!sheet) return [];
  var displayData = sheet.getDataRange().getDisplayValues();
  var rawData = sheet.getDataRange().getValues();
  if (displayData.length <= 1) return [];

  var logs = [];
  for (var i = 1; i < displayData.length; i++) {
    var dRow = displayData[i];
    var rRow = rawData[i];
    if (!dRow[0] && !dRow[1]) continue;

    var dateStr = cleanDateStr(rRow[1] || dRow[1]);
    var startTime = cleanTimeStr(rRow[2] || dRow[2]);
    var endTime = cleanTimeStr(rRow[3] || dRow[3]);
    var breakMin = dRow[4] !== '' ? Number(dRow[4]) : 80;
    var hours = calculateWorkHours(startTime, endTime, breakMin);

    var habitMap = {};
    if (dRow[20]) {
      try { habitMap = JSON.parse(dRow[20]); } catch (e) { habitMap = {}; }
    }

    var ratingVal = dRow[14] !== '' ? Number(dRow[14]) : 3;
    var moodVal = (habitMap && habitMap._mood !== undefined) ? Number(habitMap._mood) : ratingVal;
    var qNote = String(dRow[17] || '');

    logs.push({
      id: String(dRow[0] || ('log_' + i)),
      date: dateStr,
      startTime: startTime,
      endTime: endTime,
      breakMinutes: breakMin,
      workHours: hours.workHours,
      overtimeHours: hours.overtimeHours,
      scheduleText: String(dRow[7] || ''),
      reportKeep: String(dRow[8] || ''),
      keep: String(dRow[8] || ''),
      reportProblem: String(dRow[9] || ''),
      problem: String(dRow[9] || ''),
      reportInsight: String(dRow[10] || ''),
      reportTry: String(dRow[11] || ''),
      tryText: String(dRow[11] || ''),
      memo: String(dRow[12] || ''),
      photoUrls: dRow[13] ? String(dRow[13]).split(',') : [],
      growthScore: ratingVal,
      dayRating: ratingVal,
      mood: moodVal,
      condition: dRow[15] !== '' ? Number(dRow[15]) : 3,
      sleepHours: dRow[16] !== '' ? Number(dRow[16]) : 7,
      quickNote: qNote,
      intention: qNote,
      author: String(dRow[18] || '中村'),
      updatedAt: String(dRow[19] || ''),
      habitMap: habitMap
    });
  }
  return logs;
}

function getOrCreateWorkRowByDate(dateStr) {
  var sheet = getWorkSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.WORK_LOG);
  var rawData = sheet.getDataRange().getValues();
  var targetDateStr = cleanDateStr(dateStr);

  for (var i = 1; i < rawData.length; i++) {
    if (cleanDateStr(rawData[i][1]) === targetDateStr) {
      return { rowIndex: i + 1 };
    }
  }

  var newId = 'log_' + new Date().getTime();
  var newRow = [newId, targetDateStr, '', '', 80, 0, 0, '', '', '', '', '', '', '', 3, 3, 7, '', '中村', new Date(), '{}'];
  sheet.appendRow(newRow);
  return { rowIndex: sheet.getLastRow() };
}

function syncWorkDailyLog(payload) {
  if (Array.isArray(payload)) payload = payload[0];
  if (!payload || !payload.date) throw new Error('syncWorkDailyLog: 日付(date)が未指定です');
  setupAllDatabases();

  var rowObj = getOrCreateWorkRowByDate(payload.date);
  var sheet = getWorkSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.WORK_LOG);
  var rowIndex = rowObj.rowIndex;

  var startTime = cleanTimeStr(payload.startTime);
  var endTime = cleanTimeStr(payload.endTime);
  var breakMinutes = (payload.breakMinutes !== undefined && payload.breakMinutes !== '' && payload.breakMinutes !== null)
    ? Number(payload.breakMinutes) : 80;
  var hours = calculateWorkHours(startTime, endTime, breakMinutes);

  var habitObj = {};
  if (typeof payload.habitMap === 'string') {
    try { habitObj = JSON.parse(payload.habitMap); } catch (e) { habitObj = {}; }
  } else if (payload.habitMap && typeof payload.habitMap === 'object') {
    habitObj = JSON.parse(JSON.stringify(payload.habitMap));
  }
  if (payload.mood !== undefined && payload.mood !== null) {
    habitObj._mood = Number(payload.mood);
  }
  var habitJson = JSON.stringify(habitObj);

  var cleanMemo = String(payload.memo || '').replace(/\s*\[local_photo:[^\]]+\]/g, '');
  var ratingVal = (payload.dayRating !== undefined && payload.dayRating !== null)
    ? Number(payload.dayRating)
    : (payload.growthScore || 3);
  var intentionVal = (payload.intention !== undefined && payload.intention !== null)
    ? String(payload.intention)
    : String(payload.quickNote || '');

  sheet.getRange(rowIndex, 3, 1, 11).setValues([[
    startTime,
    endTime,
    breakMinutes,
    hours.workHours,
    hours.overtimeHours,
    payload.scheduleText || '',
    payload.reportKeep !== undefined ? payload.reportKeep : (payload.keep || ''),
    payload.reportProblem !== undefined ? payload.reportProblem : (payload.problem || ''),
    payload.reportInsight || '',
    payload.reportTry !== undefined ? payload.reportTry : (payload.tryText || ''),
    cleanMemo
  ]]);

  sheet.getRange(rowIndex, 15, 1, 7).setValues([[
    ratingVal,
    payload.condition || 3,
    (payload.sleepHours !== undefined && payload.sleepHours !== '' && payload.sleepHours !== null) ? Number(payload.sleepHours) : 7,
    intentionVal,
    sheet.getRange(rowIndex, 19).getValue() || '中村',
    new Date(),
    habitJson
  ]]);

  return { success: true };
}

function getPrivateLogs(privSs) {
  if (!privSs) privSs = getPrivateSpreadsheet();
  var sheet = privSs.getSheetByName(APP_CONFIG.SHEETS.PRIV_LOG);
  if (!sheet) return [];
  var displayData = sheet.getDataRange().getDisplayValues();
  var rawData = sheet.getDataRange().getValues();
  if (displayData.length <= 1) return [];

  var logs = [];
  for (var i = 1; i < displayData.length; i++) {
    var dRow = displayData[i];
    var rRow = rawData[i];
    if (!dRow[0] && !dRow[1]) continue;

    var habitMap = {};
    if (dRow[3]) {
      try { habitMap = JSON.parse(dRow[3]); } catch (e) { habitMap = {}; }
    }

    var bedTime = cleanTimeStr(rRow[8] || dRow[8] || '');
    var wakeTime = cleanTimeStr(rRow[9] || dRow[9] || '');
    var sleepHours = (dRow[10] !== '' && dRow[10] !== undefined) ? Number(dRow[10]) : calculateSleepHours(bedTime, wakeTime);
    var memoStr = String(dRow[4] || '');
    var stepsVal = (habitMap && habitMap._steps !== undefined) ? habitMap._steps : '';
    if (stepsVal === '') {
      var mSteps = memoStr.match(/\[歩数\]\s*(\d+)/);
      if (mSteps) stepsVal = Number(mSteps[1]);
    }

    logs.push({
      id: String(dRow[0] || ('plog_' + i)),
      date: cleanDateStr(rRow[1] || dRow[1]),
      scheduleText: String(dRow[2] || ''),
      habitMap: habitMap,
      memo: memoStr,
      photoUrls: dRow[5] ? String(dRow[5]).split(',') : [],
      score: dRow[6] !== '' ? Number(dRow[6]) : 3,
      sleepQuality: (habitMap && habitMap._sleepQuality !== undefined) ? Number(habitMap._sleepQuality) : 3,
      condition: (habitMap && habitMap._condition !== undefined) ? Number(habitMap._condition) : 3,
      steps: stepsVal,
      updatedAt: String(dRow[7] || ''),
      bedTime: bedTime,
      wakeTime: wakeTime,
      sleepHours: sleepHours,
      weight: (dRow[11] !== '' && dRow[11] !== undefined) ? Number(dRow[11]) : '',
      bpSys: (dRow[12] !== '' && dRow[12] !== undefined) ? Number(dRow[12]) : '',
      bpDia: (dRow[13] !== '' && dRow[13] !== undefined) ? Number(dRow[13]) : ''
    });
  }
  return logs;
}

function getOrCreatePrivateRowByDate(privSs, dateStr) {
  var sheet = privSs.getSheetByName(APP_CONFIG.SHEETS.PRIV_LOG);
  var rawData = sheet.getDataRange().getValues();
  var targetDateStr = cleanDateStr(dateStr);

  for (var i = 1; i < rawData.length; i++) {
    if (cleanDateStr(rawData[i][1]) === targetDateStr) {
      return { rowIndex: i + 1 };
    }
  }

  var newId = 'plog_' + new Date().getTime();
  var newRow = [newId, targetDateStr, '', '{}', '', '', 3, new Date(), '', '', '', '', '', ''];
  sheet.appendRow(newRow);
  return { rowIndex: sheet.getLastRow() };
}

function syncPrivateDailyLog(payload) {
  if (Array.isArray(payload)) payload = payload[0];
  if (!payload || !payload.date) throw new Error('syncPrivateDailyLog: 日付(date)が未指定です');
  setupAllDatabases();

  var privSs = getPrivateSpreadsheet();
  var rowObj = getOrCreatePrivateRowByDate(privSs, payload.date);
  var sheet = privSs.getSheetByName(APP_CONFIG.SHEETS.PRIV_LOG);
  var rowIndex = rowObj.rowIndex;

  var bedTime = cleanTimeStr(payload.bedTime || '');
  var wakeTime = cleanTimeStr(payload.wakeTime || '');
  var sleepHours = (payload.sleepHours !== null && payload.sleepHours !== undefined && payload.sleepHours !== '')
    ? Number(payload.sleepHours)
    : calculateSleepHours(bedTime, wakeTime);

  var habitObj = {};
  if (typeof payload.habitMap === 'string') {
    try { habitObj = JSON.parse(payload.habitMap); } catch (e) { habitObj = {}; }
  } else if (payload.habitMap && typeof payload.habitMap === 'object') {
    habitObj = JSON.parse(JSON.stringify(payload.habitMap));
  }
  if (payload.sleepQuality !== undefined && payload.sleepQuality !== null) {
    habitObj._sleepQuality = Number(payload.sleepQuality);
  }
  if (payload.condition !== undefined && payload.condition !== null) {
    habitObj._condition = Number(payload.condition);
  }
  if (payload.steps !== undefined && payload.steps !== null && payload.steps !== '') {
    habitObj._steps = Number(payload.steps);
  }
  var habitJson = JSON.stringify(habitObj);

  var cleanMemo = String(payload.memo || '').replace(/\s*\[local_photo:[^\]]+\]/g, '');

  var weightVal = (payload.weight !== '' && payload.weight !== null && payload.weight !== undefined && !isNaN(Number(payload.weight))) ? Number(payload.weight) : '';
  var bpSysVal = (payload.bpSys !== '' && payload.bpSys !== null && payload.bpSys !== undefined && !isNaN(Number(payload.bpSys))) ? Number(payload.bpSys) : '';
  var bpDiaVal = (payload.bpDia !== '' && payload.bpDia !== null && payload.bpDia !== undefined && !isNaN(Number(payload.bpDia))) ? Number(payload.bpDia) : '';

  sheet.getRange(rowIndex, 3, 1, 3).setValues([[
    payload.scheduleText || '',
    habitJson,
    cleanMemo
  ]]);

  sheet.getRange(rowIndex, 7, 1, 8).setValues([[
    payload.score || 3,
    new Date(),
    bedTime,
    wakeTime,
    (sleepHours !== null && !isNaN(sleepHours)) ? sleepHours : '',
    weightVal,
    bpSysVal,
    bpDiaVal
  ]]);

  return { success: true };
}

function addMemoWithPhotos(payload) {
  if (Array.isArray(payload)) payload = payload[0];
  setupAllDatabases();
  var folder = getFolder();
  var isPrivate = payload.domain === 'private';
  var sheet, rowIndex, memoCol, photoCol, updatedCol;

  if (isPrivate) {
    var privSs = getPrivateSpreadsheet();
    var pRow = getOrCreatePrivateRowByDate(privSs, payload.date);
    sheet = privSs.getSheetByName(APP_CONFIG.SHEETS.PRIV_LOG);
    rowIndex = pRow.rowIndex;
    memoCol = 5;
    photoCol = 6;
    updatedCol = 8;
  } else {
    var wRow = getOrCreateWorkRowByDate(payload.date);
    sheet = getWorkSpreadsheet().getSheetByName(APP_CONFIG.SHEETS.WORK_LOG);
    rowIndex = wRow.rowIndex;
    memoCol = 13;
    photoCol = 14;
    updatedCol = 20;
  }

  var currentMemo = String(sheet.getRange(rowIndex, memoCol).getValue() || '');
  var photoVal = sheet.getRange(rowIndex, photoCol).getValue();
  var photoUrls = photoVal ? String(photoVal).split(',') : [];

  var uploadedCount = 0;
  if (payload.photoFiles && payload.photoFiles.length > 0) {
    payload.photoFiles.forEach(function(file) {
      var blob = Utilities.newBlob(Utilities.base64Decode(file.base64.split(',')[1]), file.mimeType, file.name);
      var driveFile = folder.createFile(blob);
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrls.push(driveFile.getUrl());
      uploadedCount++;
    });
  }

  var timeStr = payload.timeStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
  var photoTag = uploadedCount > 0 ? ' [photo:' + photoUrls.length + ']' : '';
  var newItem = '[' + timeStr + '] ' + payload.memoText.trim() + photoTag;

  var rawLine = '[' + timeStr + '] ' + payload.memoText.trim();
  var lines = currentMemo ? currentMemo.split('\n') : [];
  var replaced = false;
  for (var i = 0; i < lines.length; i++) {
    if (lines[i] === rawLine) {
      lines[i] = newItem;
      replaced = true;
      break;
    }
  }
  if (!replaced) lines.push(newItem);

  sheet.getRange(rowIndex, memoCol).setValue(lines.join('\n'));
  sheet.getRange(rowIndex, photoCol).setValue(photoUrls.join(','));
  sheet.getRange(rowIndex, updatedCol).setValue(new Date());

  return { success: true, photoUrls: photoUrls, memo: lines.join('\n') };
}