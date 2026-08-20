/**
 * 面談枠の自動生成・再生成と、全体ビュー・各クラス予約表の更新。
 */

var CLASS_SHEET_PREFIX = '予約表_';

/**
 * 「面談日」「設定」「クラス」から面談枠を一括生成する。
 * 既存の予約がある場合は枠IDで引き継ぎ、無くなったらエラーにする。
 * @return {{written:number, kept:number}}
 */
function generateSlots() {
  var cfg = getConfig();
  var days = getDays();
  var classes = getClasses();
  if (!days.length) throw new Error('「面談日」シートで「実施する」がTRUEの日付が1つもありません。');
  if (!classes.length) throw new Error('「クラス」シートにクラス名を入力してください。');

  var existing = readSlots_();
  var existingMap = {};
  var lost = [];
  for (var i = 0; i < existing.length; i++) {
    var v = existing[i].v;
    var id = String(v[COL.SLOT_ID - 1]);
    existingMap[id] = v;
  }

  var times = daySlotTimes_(cfg);
  var newSlotIds = {};
  var rows = [];
  var kept = 0;

  for (var d = 0; d < days.length; d++) {
    for (var c = 0; c < classes.length; c++) {
      for (var t = 0; t < times.length; t++) {
        var slotId = makeSlotId_(days[d], classes[c].name, times[t].index);
        newSlotIds[slotId] = true;
        var old = existingMap[slotId];

        if (old) {
          rows.push([
            slotId,
            days[d],
            times[t].start,
            times[t].end,
            classes[c].name,
            classes[c].teacher || '',
            old[COL.STATUS - 1],
            old[COL.NUMBER - 1] || '',
            old[COL.STUDENT - 1] || '',
            old[COL.GUARDIAN - 1] || '',
            old[COL.NOTE - 1] || '',
            old[COL.CODE - 1] || '',
            old[COL.BOOKED_AT - 1] || ''
          ]);
          if (String(old[COL.STATUS - 1]) === STATUS.BOOKED) kept++;
        } else {
          rows.push([
            slotId,
            days[d],
            times[t].start,
            times[t].end,
            classes[c].name,
            classes[c].teacher || '',
            STATUS.OPEN,
            '', '', '', '', '', ''
          ]);
        }
      }
    }
  }

  for (var k = 0; k < existing.length; k++) {
    var ev = existing[k].v;
    var eid = String(ev[COL.SLOT_ID - 1]);
    if (!newSlotIds[eid] && String(ev[COL.STATUS - 1]) === STATUS.BOOKED) {
      lost.push(dateLabel_(ev[COL.DATE - 1]) + ' ' + ev[COL.START - 1] + ' ' +
        ev[COL.CLASS - 1] + ' ' + ev[COL.NUMBER - 1] + '. ' + ev[COL.STUDENT - 1]);
    }
  }

  if (lost.length) {
    throw new Error(
      '次の枠に予約が入っているため再生成を中止しました。先に予約を取り消すか、面談日・クラス・枠数の設定を戻してください。\n' +
      lost.join('\n')
    );
  }

  var sh = sheet_(SH.SLOTS);
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, SLOT_LAST_COL).clearContent();
  }
  sh.getRange(2, 1, rows.length, SLOT_LAST_COL).setValues(rows);
  sh.getRange(2, COL.DATE, rows.length, 1).setNumberFormat('yyyy/mm/dd');
  sh.getRange(2, COL.START, rows.length, 2).setNumberFormat('@');
  sh.getRange(2, COL.BOOKED_AT, rows.length, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  clearSlotCache_();

  logAction_('枠再生成', '', '', '', '', rows.length + '枠 / 引継ぎ ' + kept + '件');
  return { written: rows.length, kept: kept };
}

/** 枠マスタ全行を読む（ヘッダを除く） */
function readSlots_() {
  var sh = sheet_(SH.SLOTS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, SLOT_LAST_COL).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    if (!vals[i][COL.SLOT_ID - 1]) continue;
    out.push({ row: i + 2, v: vals[i] });
  }
  return out;
}

/**
 * 全体ビューを作り直す。
 * 1行目: 学年・各クラス予約進捗サマリーバナー
 * 2行目: 見出しヘッダー
 * 3行目〜: 行 = 日付×時間、列 = クラス。空きが一目で分かる。
 */
function rebuildOverview() {
  var cfg = getConfig();
  var classes = getClasses();
  var days = getDays();
  var times = daySlotTimes_(cfg);
  var slots = readSlots_();
  var roster = getRoster();

  // 進捗率の計算
  var totalStudents = roster.length;
  var bookedCount = 0;
  var classStats = {};
  for (var c = 0; c < classes.length; c++) {
    classStats[classes[c].name] = { total: 0, booked: 0 };
  }
  for (var r = 0; r < roster.length; r++) {
    if (classStats[roster[r].cls]) classStats[roster[r].cls].total++;
  }
  for (var s = 0; s < slots.length; s++) {
    var sv = slots[s].v;
    if (String(sv[COL.STATUS - 1]) === STATUS.BOOKED) {
      bookedCount++;
      var cName = String(sv[COL.CLASS - 1]);
      if (classStats[cName]) classStats[cName].booked++;
    }
  }

  var totalRate = totalStudents ? Math.round((bookedCount / totalStudents) * 100) : 0;
  var summaryText = '📊 学年予約状況: ' + totalStudents + '名中 ' + bookedCount + '名予約済 (' + totalRate + '%) ｜ ';
  var classSummaries = [];
  for (var cc = 0; cc < classes.length; cc++) {
    var cN = classes[cc].name;
    var st = classStats[cN] || { total: 0, booked: 0 };
    var cRate = st.total ? Math.round((st.booked / st.total) * 100) : 0;
    classSummaries.push(cN + ': ' + st.booked + '/' + st.total + '名 (' + cRate + '%)');
  }
  summaryText += classSummaries.join(' ｜ ');

  var byId = {};
  for (var i = 0; i < slots.length; i++) byId[slots[i].v[COL.SLOT_ID - 1]] = slots[i].v;

  var header = ['日付', '時間'].concat(classes.map(function (c) {
    return c.name + (c.teacher ? '\n' + c.teacher : '');
  })).concat(['空き数']);

  var body = [];
  var colors = [];
  for (var d = 0; d < days.length; d++) {
    for (var t = 0; t < times.length; t++) {
      var line = [dateLabel_(days[d]), times[t].start + '–' + times[t].end];
      var lineColor = ['#ffffff', '#ffffff'];
      var free = 0;
      for (var c = 0; c < classes.length; c++) {
        var v = byId[makeSlotId_(days[d], classes[c].name, times[t].index)];
        if (!v) {
          line.push('—'); lineColor.push('#f8f9fa'); continue;
        }
        var statusStr = String(v[COL.STATUS - 1]);
        if (statusStr === STATUS.BOOKED) {
          line.push(v[COL.NUMBER - 1] + '. ' + v[COL.STUDENT - 1]);
          lineColor.push('#ffffff');
        } else if (statusStr === STATUS.BLOCKED) {
          line.push('×');
          lineColor.push('#f1f3f4');
        } else {
          line.push('空き');
          lineColor.push('#e6f4ea');
          free++;
        }
      }
      line.push(free);
      lineColor.push(free === 0 ? '#fce8e6' : '#ffffff');
      body.push(line);
      colors.push(lineColor);
    }
  }

  var sh = ss_().getSheetByName(SH.OVERVIEW) || ss_().insertSheet(SH.OVERVIEW);
  sh.clear();

  // 1行目: 進捗サマリーバナー
  sh.getRange(1, 1, 1, header.length).merge()
    .setValue(summaryText)
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#e8f0fe')
    .setFontColor('#1a73e8')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 28);

  // 2行目: ヘッダー
  sh.getRange(2, 1, 1, header.length).setValues([header]);
  sh.getRange(2, 1, 1, header.length)
    .setFontWeight('bold')
    .setBackground('#f1f3f4')
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(2, 28);

  // 3行目〜: データ行
  if (body.length) {
    sh.getRange(3, 1, body.length, header.length).setValues(body);
    sh.getRange(3, 1, body.length, header.length).setBackgrounds(colors);
    sh.getRange(3, 1, body.length, header.length).setHorizontalAlignment('center').setVerticalAlignment('middle');
    // 日付が変わるところに中太罫線
    for (var r = 0; r < body.length; r += times.length) {
      sh.getRange(r + 3, 1, 1, header.length).setBorder(true, null, null, null, null, null, '#5f6368', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }
  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 120);
  for (var colIdx = 0; colIdx < classes.length; colIdx++) sh.setColumnWidth(3 + colIdx, 150);
  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);
  return body.length;
}

/** 
 * クラスごとの予約表シートを作り直す（各クラスシートを直接マスターとして保持）
 * 左側(A-F): 【生徒別 予約状況（出席番号順）】 - 生徒名簿直接管理
 * 右側(I-O): 【時間枠別 予約表（時間順）】 - O列に4桁の「予約コード」を表示・日付ごとに区切り線
 */
function rebuildClassSheets() {
  var ss = ss_();
  var classes = getClasses();
  var slots = readSlots_();

  var headerLeft = ['出席番号', '生徒氏名', '予約状況', '予約日時', '保護者氏名', '連絡事項'];
  var headerRight = ['日付', '時間', '状態', '出席番号', '生徒氏名', '保護者氏名', '予約コード'];

  for (var c = 0; c < classes.length; c++) {
    var clsName = classes[c].name;
    var name = CLASS_SHEET_PREFIX + clsName;
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);

    var existingStudents = [];
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var rawAB = sh.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var r = 0; r < rawAB.length; r++) {
        var no = Number(rawAB[r][0]);
        var stName = String(rawAB[r][1] || '').trim();
        if (no && stName) {
          existingStudents.push({ no: no, name: stName });
        }
      }
    }

    existingStudents.sort(function (a, b) { return a.no - b.no; });

    var bookedByNo = {};
    for (var i = 0; i < slots.length; i++) {
      var v = slots[i].v;
      if (String(v[COL.CLASS - 1]) === clsName && String(v[COL.STATUS - 1]) === STATUS.BOOKED) {
        bookedByNo[Number(v[COL.NUMBER - 1])] = v;
      }
    }

    // 左側 (A〜F): 生徒別 予約状況
    var bodyLeft = [];
    var colorsLeft = [];
    for (var s = 0; s < existingStudents.length; s++) {
      var st = existingStudents[s];
      var bk = bookedByNo[st.no];
      if (bk) {
        bodyLeft.push([
          st.no,
          st.name,
          '予約済',
          dateLabel_(bk[COL.DATE - 1]) + ' ' + bk[COL.START - 1] + '–' + bk[COL.END - 1],
          bk[COL.GUARDIAN - 1] || '',
          bk[COL.NOTE - 1] || ''
        ]);
        colorsLeft.push(['#ffffff', '#ffffff', '#e6f4ea', '#ffffff', '#ffffff', '#ffffff']);
      } else {
        bodyLeft.push([
          st.no,
          st.name,
          '未予約',
          '—',
          '—',
          '—'
        ]);
        colorsLeft.push(['#fef7e0', '#fef7e0', '#fef7e0', '#fef7e0', '#fef7e0', '#fef7e0']);
      }
    }

    // 右側 (I〜O): 時間枠別 予約表
    var bodyRight = [];
    var colorsRight = [];
    var dateBoundaries = [];
    var prevDate = '';

    for (var j = 0; j < slots.length; j++) {
      var sv = slots[j].v;
      if (String(sv[COL.CLASS - 1]) !== clsName) continue;
      var curDate = dateLabel_(sv[COL.DATE - 1]);
      if (curDate !== prevDate) {
        dateBoundaries.push(bodyRight.length);
        prevDate = curDate;
      }
      var statusStr = String(sv[COL.STATUS - 1]);
      bodyRight.push([
        curDate,
        sv[COL.START - 1] + '–' + sv[COL.END - 1],
        statusStr,
        sv[COL.NUMBER - 1] || '',
        sv[COL.STUDENT - 1] || '',
        sv[COL.GUARDIAN - 1] || '',
        sv[COL.CODE - 1] ? "'" + String(sv[COL.CODE - 1]) : ''
      ]);
      var bg = statusStr === STATUS.OPEN ? '#e6f4ea' : (statusStr === STATUS.BLOCKED ? '#f1f3f4' : '#ffffff');
      colorsRight.push([bg, bg, bg, bg, bg, bg, bg]);
    }

    sh.clear();

    // 左側ヘッダー (A1:F1)
    sh.getRange(1, 1, 1, headerLeft.length).setValues([headerLeft]);
    sh.getRange(1, 1, 1, headerLeft.length)
      .setFontWeight('bold')
      .setBackground('#d9ead3')
      .setVerticalAlignment('middle');

    if (bodyLeft.length) {
      sh.getRange(2, 1, bodyLeft.length, headerLeft.length).setValues(bodyLeft);
      sh.getRange(2, 1, bodyLeft.length, headerLeft.length).setBackgrounds(colorsLeft);
    }

    // 右側ヘッダー (I1:O1)
    sh.getRange(1, 9, 1, headerRight.length).setValues([headerRight]);
    sh.getRange(1, 9, 1, headerRight.length)
      .setFontWeight('bold')
      .setBackground('#e8eaed')
      .setVerticalAlignment('middle');

    if (bodyRight.length) {
      sh.getRange(2, 9, bodyRight.length, headerRight.length).setValues(bodyRight);
      sh.getRange(2, 9, bodyRight.length, headerRight.length).setBackgrounds(colorsRight);

      for (var bIdx = 0; bIdx < dateBoundaries.length; bIdx++) {
        var rNum = dateBoundaries[bIdx] + 2;
        sh.getRange(rNum, 9, 1, headerRight.length)
          .setBorder(true, null, null, null, null, null, '#5f6368', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    }

    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 70);
    sh.setColumnWidth(2, 120);
    sh.setColumnWidth(3, 80);
    sh.setColumnWidth(4, 180);
    sh.setColumnWidth(5, 120);
    sh.setColumnWidth(6, 200);

    sh.setColumnWidth(8, 30);

    sh.setColumnWidth(9, 110);
    sh.setColumnWidth(10, 110);
    sh.setColumnWidth(11, 70);
    sh.setColumnWidth(12, 70);
    sh.setColumnWidth(13, 120);
    sh.setColumnWidth(14, 120);
    sh.setColumnWidth(15, 100);
  }
}

/** キャッシュされた空き枠を読む */
function readSlotsCached_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('all_slots_v1');
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* ignore */ }
  }
  var slots = readSlots_();
  var simple = slots.map(function (s) {
    return {
      v: [
        s.v[COL.SLOT_ID - 1],
        ymd_(s.v[COL.DATE - 1]),
        s.v[COL.START - 1],
        s.v[COL.END - 1],
        s.v[COL.CLASS - 1],
        s.v[COL.TEACHER - 1],
        s.v[COL.STATUS - 1],
        s.v[COL.NUMBER - 1],
        s.v[COL.STUDENT - 1],
        s.v[COL.GUARDIAN - 1],
        s.v[COL.NOTE - 1],
        s.v[COL.CODE - 1],
        s.v[COL.BOOKED_AT - 1]
      ]
    };
  });
  try {
    cache.put('all_slots_v1', JSON.stringify(simple), 30);
  } catch (e) { /* ignore */ }
  return slots;
}

function clearSlotCache_() {
  try { CacheService.getScriptCache().remove('all_slots_v1'); } catch (e) { }
}

function makeSlotId_(date, clsName, idx) {
  return ymdCompact_(date) + '_' + clsName + '_' + idx;
}

function daySlotTimes_(cfg) {
  var times = [];
  var cur = cfg.startMin;
  for (var i = 1; i <= cfg.slotsPerDay; i++) {
    var end = cur + cfg.slotMin;
    times.push({ index: i, start: fromMinutes_(cur), end: fromMinutes_(end) });
    cur = end + cfg.breakMin;
  }
  return times;
}
