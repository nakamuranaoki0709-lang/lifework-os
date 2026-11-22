// ==========================================
// 05_NotesTodoRepo.gs : 一人SNS ＆ タスク(ToDo) ＆ 旧シート統合
// ==========================================
function saveNormalizedNotes_(notesArr) {
  if (!Array.isArray(notesArr)) return;
  var privSs = getPrivateSpreadsheet();
  var headers = ['ID', '投稿日時', '更新日時', '本文', 'タグ', 'ピン留め', '画像数', '拡張JSON'];
  var sheet = ensureSheetSchema_(privSs, APP_CONFIG.SHEETS.SOLO_SNS, headers, '#0f172a', '#0284c7');

  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  if (notesArr.length === 0) return;

  var rows = notesArr.map(function(n) {
    var id = String(n.id || ('note_' + Date.now()));
    var created = String(n.createdAt || n.date || n.timestamp || '');
    var updated = String(n.updatedAt || created);
    var text = String(n.content != null ? n.content : (n.text || ''));
    var tags = Array.isArray(n.tags) ? n.tags.join(', ') : String(n.tags || '');
    var pinned = n.pinned ? '★' : '';
    var imgCount = Array.isArray(n.images) ? n.images.length : 0;
    return [id, created, updated, text, tags, pinned, imgCount, JSON.stringify(n)];
  });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function loadNormalizedNotes_() {
  var privSs = getPrivateSpreadsheet();
  var sheet = privSs.getSheetByName(APP_CONFIG.SHEETS.SOLO_SNS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0] && !r[3]) continue;
    var obj = null;
    if (r[7]) { try { obj = JSON.parse(String(r[7])); } catch (e) {} }
    if (!obj || typeof obj !== 'object') {
      obj = {
        id: String(r[0] || ('note_' + i)),
        createdAt: String(r[1] || new Date().toISOString()),
        updatedAt: String(r[2] || r[1] || new Date().toISOString()),
        content: String(r[3] || ''),
        tags: r[4] ? String(r[4]).split(',').map(function(s) { return s.trim(); }).filter(Boolean) : [],
        pinned: Boolean(r[5])
      };
    } else if (r[3]) {
      obj.content = String(r[3]);
    }
    out.push(obj);
  }
  return out;
}

function saveNormalizedTodos_(todosArr) {
  if (!Array.isArray(todosArr)) return;
  var privSs = getPrivateSpreadsheet();
  var headers = ['ID', '完了', 'タスク名', 'カテゴリ', '期限', '優先度', '作成日時', '拡張JSON'];
  var sheet = ensureSheetSchema_(privSs, APP_CONFIG.SHEETS.TODO_LIST, headers, '#0f172a', '#059669');

  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  if (todosArr.length === 0) return;

  var rows = todosArr.map(function(t) {
    var id = String(t.id || ('todo_' + Date.now()));
    var done = (t.completed || t.done) ? '完了' : '未完了';
    var title = String(t.title || t.text || t.content || '');
    var cat = String(t.category || t.project || t.tag || '');
    var due = String(t.dueDate || t.date || '');
    var prio = String(t.priority || t.importance || '');
    var created = String(t.createdAt || '');
    return [id, done, title, cat, due, prio, created, JSON.stringify(t)];
  });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function loadNormalizedTodos_() {
  var privSs = getPrivateSpreadsheet();
  var sheet = privSs.getSheetByName(APP_CONFIG.SHEETS.TODO_LIST);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0] && !r[2]) continue;
    var obj = null;
    if (r[7]) { try { obj = JSON.parse(String(r[7])); } catch (e) {} }
    if (!obj || typeof obj !== 'object') {
      obj = {
        id: String(r[0] || ('todo_' + i)),
        completed: (String(r[1]) === '完了' || r[1] === true),
        title: String(r[2] || ''),
        category: String(r[3] || ''),
        dueDate: String(r[4] || ''),
        priority: String(r[5] || ''),
        createdAt: String(r[6] || '')
      };
    } else {
      if (r[2]) obj.title = String(r[2]);
      obj.completed = (String(r[1]) === '完了' || r[1] === true);
    }
    out.push(obj);
  }
  return out;
}

function getNotesFromSheet() { return loadNormalizedNotes_(); }
function saveNotesList(notes) {
  saveNormalizedNotes_(Array.isArray(notes) ? notes : []);
  SpreadsheetApp.flush();
  return { ok: true };
}
function getTodoListFromSheet() { return loadNormalizedTodos_(); }
function saveTodoList(todos) {
  saveNormalizedTodos_(Array.isArray(todos) ? todos : []);
  SpreadsheetApp.flush();
  return { ok: true };
}

function consolidateLegacySheets_() {
  setupAllDatabases();
  var workSs = getWorkSpreadsheet();
  var privSs = getPrivateSpreadsheet();
  var mergedNotes = 0;
  var mergedTodos = 0;

  var oldNotesSheet = workSs.getSheetByName('Notes_DB');
  if (oldNotesSheet) {
    if (oldNotesSheet.getLastRow() >= 2) {
      var nRows = oldNotesSheet.getRange(2, 1, oldNotesSheet.getLastRow() - 1, 8).getValues();
      var currentNotes = loadNormalizedNotes_();
      var existingIds = {};
      currentNotes.forEach(function(n) { existingIds[String(n.id)] = true; });
      for (var i = 0; i < nRows.length; i++) {
        var r = nRows[i];
        var nid = String(r[0] || '');
        if (!nid || existingIds[nid]) continue;
        var tags = [];
        try { if (r[3]) tags = JSON.parse(String(r[3])); } catch (e) {}
        currentNotes.push({
          id: nid,
          title: String(r[1] || ''),
          content: String(r[2] || ''),
          tags: tags,
          domain: String(r[4] || 'work'),
          pinned: (r[5] === true || String(r[5]).toUpperCase() === 'TRUE'),
          updatedAt: String(r[6] || ''),
          createdAt: String(r[7] || '')
        });
        mergedNotes++;
      }
      if (mergedNotes > 0) saveNormalizedNotes_(currentNotes);
    }
    workSs.deleteSheet(oldNotesSheet);
  }

  var oldTodoSheet = workSs.getSheetByName('Todo_Items');
  if (oldTodoSheet) {
    if (oldTodoSheet.getLastRow() >= 2) {
      var tRows = oldTodoSheet.getRange(2, 1, oldTodoSheet.getLastRow() - 1, Math.min(15, oldTodoSheet.getLastColumn())).getValues();
      var currentTodos = loadNormalizedTodos_();
      var existingTIds = {};
      currentTodos.forEach(function(t) { existingTIds[String(t.id)] = true; });
      var tz = Session.getScriptTimeZone();
      for (var j = 0; j < tRows.length; j++) {
        var tr = tRows[j];
        var tid = String(tr[0] || '');
        if (!tid || existingTIds[tid]) continue;
        var subs = [];
        try { if (tr[11]) subs = JSON.parse(String(tr[11])); } catch (e) {}
        var comp = (tr[9] === true || String(tr[9]).toUpperCase() === 'TRUE');
        currentTodos.push({
          id: tid,
          title: String(tr[1] || ''),
          importance: String(tr[2] || 'mid'),
          urgency: String(tr[3] || 'low'),
          startDate: tr[4] instanceof Date ? Utilities.formatDate(tr[4], tz, 'yyyy-MM-dd') : String(tr[4] || ''),
          dueDate: tr[5] instanceof Date ? Utilities.formatDate(tr[5], tz, 'yyyy-MM-dd') : String(tr[5] || ''),
          project: String(tr[6] || '仕事'),
          status: String(tr[7] || (comp ? 'done' : 'todo')),
          progress: Number(tr[8]) || (comp ? 100 : 0),
          completed: comp,
          completedAt: tr[10] instanceof Date ? Utilities.formatDate(tr[10], tz, 'yyyy-MM-dd') : String(tr[10] || ''),
          subtasks: subs,
          createdAt: tr[12] instanceof Date ? Utilities.formatDate(tr[12], tz, 'yyyy-MM-dd') : String(tr[12] || '')
        });
        mergedTodos++;
      }
      if (mergedTodos > 0) saveNormalizedTodos_(currentTodos);
    }
    workSs.deleteSheet(oldTodoSheet);
  }

  SpreadsheetApp.flush();
  return {
    mergedNotes: mergedNotes,
    mergedTodos: mergedTodos,
    workSheetName: workSs.getName(),
    workSheetUrl: workSs.getUrl(),
    workSheets: workSs.getSheets().map(function(s) { return s.getName(); }),
    privateSheetName: privSs.getName(),
    privateSheetUrl: privSs.getUrl(),
    privateSheets: privSs.getSheets().map(function(s) { return s.getName(); })
  };
}