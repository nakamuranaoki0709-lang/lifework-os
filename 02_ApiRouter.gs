// ==========================================
// 02_ApiRouter.gs : HTTP窓口（個別独立処理・診断API）
// ==========================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function jsonApiResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.api === '1') {
    return handleApiGet_(e);
  }
  setupAllDatabases();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('WORK & LIFE OS - 勤怠・習慣・ヘルスケア')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleApiGet_(e) {
  try {
    if (!e || !e.parameter || e.parameter.token !== APP_CONFIG.API_TOKEN) {
      return jsonApiResponse_({ ok: false, error: 'Unauthorized' });
    }
    var action = String(e.parameter.action || '');
    if (action === 'organize') {
      return jsonApiResponse_(organizeIntegratedSpreadsheet());
    }
    if (action === 'migrate') {
      return jsonApiResponse_(migrateToUnifiedSpreadsheet());
    }
    if (action === 'consolidate') {
      return jsonApiResponse_({ ok: true, report: consolidateLegacySheets_() });
    }
    if (e.parameter.fn) {
      var fnName = String(e.parameter.fn);
      var arg = e.parameter.arg ? JSON.parse(e.parameter.arg) : undefined;
      if (typeof globalThis[fnName] === 'function') {
        var result = (arg !== undefined) ? globalThis[fnName](arg) : globalThis[fnName]();
        return jsonApiResponse_({ ok: true, result: result });
      }
      return jsonApiResponse_({ ok: false, error: 'Unknown function: ' + fnName });
    }
    return jsonApiResponse_({ ok: true, bundle: getFullAppBundle() });
  } catch (err) {
    return jsonApiResponse_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function dispatchSingleAction_(act, payload) {
  if (act === 'saveNotesList' || act === 'saveAllNotes' || act === 'saveNotesToSheet') {
    var notesList = Array.isArray(payload) ? payload : (payload && payload.notes ? payload.notes : []);
    saveNormalizedNotes_(notesList);
    return { ok: true };
  }
  if (act === 'saveTodoList' || act === 'saveTodosList' || act === 'saveAllTodos' || act === 'saveTodoListToSheet') {
    var todosList = Array.isArray(payload) ? payload : (payload && payload.todos ? payload.todos : []);
    saveNormalizedTodos_(todosList);
    return { ok: true };
  }
  if (typeof globalThis[act] === 'function') {
    if (Array.isArray(payload)) {
      globalThis[act].apply(null, payload);
    } else {
      globalThis[act](payload);
    }
    return { ok: true };
  }
  throw new Error('未定義のサーバー関数: ' + act);
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) ? e.postData.contents : '{}');

    // 歩数連携Webhook（従来互換）
    if (body.steps !== undefined && !body.batch) {
      var dateStr = body.date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
      var steps = parseInt(body.steps, 10);
      if (!isNaN(steps) && steps > 0) {
        var privSs = getPrivateSpreadsheet();
        var logs = getPrivateLogs(privSs);
        var target = logs.find(function(l) { return l.date === dateStr; }) || {
          date: dateStr, scheduleText: '', habitMap: {}, memo: '', photoUrls: [], score: 4,
          bedTime: '', wakeTime: '', sleepHours: null, weight: '', bpSys: '', bpDia: ''
        };
        var lines = target.memo ? String(target.memo).split('\n').filter(function(l) {
          return l.trim() !== '' && l.indexOf('[歩数] ') !== 0;
        }) : [];
        lines.push('[歩数] ' + steps);
        target.memo = lines.join('\n');
        syncPrivateDailyLog(target);
        SpreadsheetApp.flush();
      }
      return jsonApiResponse_({ ok: true, status: 'ok', date: dateStr, steps: steps });
    }

    if (body.token !== APP_CONFIG.API_TOKEN) {
      return jsonApiResponse_({ ok: false, error: 'Unauthorized' });
    }

    var batch = Array.isArray(body.batch) ? body.batch : [];
    var results = [];
    for (var i = 0; i < batch.length; i++) {
      var item = batch[i];
      if (!item || !item.action) continue;
      try {
        dispatchSingleAction_(String(item.action), item.payload);
        results.push({ id: item.id, ok: true });
      } catch (itemErr) {
        results.push({
          id: item.id,
          ok: false,
          error: String(itemErr && itemErr.message ? itemErr.message : itemErr)
        });
      }
    }
    SpreadsheetApp.flush();
    return jsonApiResponse_({ ok: true, results: results });
  } catch (err) {
    return jsonApiResponse_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}