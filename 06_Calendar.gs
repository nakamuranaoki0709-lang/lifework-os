// ==========================================
// 06_Calendar.gs : Googleカレンダー連携
// ==========================================
function authorizeCalendarPermission() {
  var cals = CalendarApp.getAllCalendars();
  var def = CalendarApp.getDefaultCalendar();
  return {
    ok: true,
    defaultName: def ? def.getName() : 'Primary',
    count: cals ? cals.length : 0
  };
}

function getCalendarEventsForDate(dateStr) {
  var ds = cleanDateStr(dateStr);
  if (!ds) {
    ds = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return getCalendarEventsForRange(ds, ds);
}

function getCalendarEventsForRange(startDateStr, endDateStr) {
  try {
    var sDs = cleanDateStr(startDateStr);
    var eDs = cleanDateStr(endDateStr || startDateStr);
    if (!sDs || !eDs) {
      return { ok: false, error: '日付形式が不正です', events: [] };
    }
    var sp = sDs.split('-');
    var ep = eDs.split('-');
    var start = new Date(parseInt(sp[0], 10), parseInt(sp[1], 10) - 1, parseInt(sp[2], 10), 0, 0, 0);
    var end = new Date(parseInt(ep[0], 10), parseInt(ep[1], 10) - 1, parseInt(ep[2], 10), 23, 59, 59);

    var calendars = [];
    try {
      calendars = CalendarApp.getAllCalendars();
    } catch (authErr) {
      return {
        ok: false,
        needsAuth: true,
        error: 'Googleカレンダーのアクセス権限が未承認です: ' + String(authErr.message || authErr),
        events: []
      };
    }

    if (!calendars || calendars.length === 0) {
      var defCal = CalendarApp.getDefaultCalendar();
      if (defCal) calendars = [defCal];
    }

    var results = [];
    var seenKeys = {};
    var tz = 'Asia/Tokyo';

    for (var i = 0; i < calendars.length; i++) {
      var cal = calendars[i];
      try {
        if (cal.isHidden && cal.isHidden()) continue;
        var calColor = cal.getColor ? cal.getColor() : '#0284c7';
        var calName = cal.getName ? cal.getName() : '';
        var events = cal.getEvents(start, end);

        for (var j = 0; j < events.length; j++) {
          var ev = events[j];
          var title = ev.getTitle() || '(無題の予定)';
          var isAllDay = ev.isAllDayEvent();
          var evDateStr = sDs;
          var sTime = '';
          var eTime = '';
          var sortKey = 0;

          if (isAllDay) {
            var allDayStart = ev.getAllDayStartDate ? ev.getAllDayStartDate() : ev.getStartTime();
            evDateStr = Utilities.formatDate(allDayStart, tz, 'yyyy-MM-dd');
            if (evDateStr < sDs) evDateStr = sDs;
            if (evDateStr > eDs) continue;
            sortKey = 0;
          } else {
            var st = ev.getStartTime();
            var et = ev.getEndTime();
            evDateStr = Utilities.formatDate(st, tz, 'yyyy-MM-dd');
            sTime = Utilities.formatDate(st, tz, 'HH:mm');
            eTime = Utilities.formatDate(et, tz, 'HH:mm');
            var hh = parseInt(sTime.slice(0, 2), 10) || 0;
            var mm = parseInt(sTime.slice(3, 5), 10) || 0;
            sortKey = hh * 60 + mm + 1;
          }

          var uniq = evDateStr + '_' + title + '_' + sTime + '_' + eTime;
          if (seenKeys[uniq]) continue;
          seenKeys[uniq] = true;

          results.push({
            date: evDateStr,
            title: title,
            startTime: sTime,
            endTime: eTime,
            isAllDay: isAllDay,
            color: calColor,
            calendarName: calName,
            sortKey: sortKey
          });
        }
      } catch (calErr) {
        // 個別カレンダー取得失敗時は他カレンダーの取得を継続
      }
    }

    results.sort(function(a, b) {
      if (a.date !== b.date) return a.date < b.date ? -1 : 1;
      return a.sortKey - b.sortKey;
    });
    return { ok: true, events: results };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err), events: [] };
  }
}