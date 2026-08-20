/**
 * 枠マスタの生成と、閲覧用シート(全体ビュー / クラス別予約表)の再構築。
 */

var CLASS_SHEET_PREFIX = '予約表_';

/** 1日分の枠の時刻を計算する。[{index, start, end}] */
function daySlotTimes_(cfg) {
  var out = [];
  var cycle = cfg.slotMin + cfg.breakMin;
  for (var i = 0; i < cfg.slotsPerDay; i++) {
    var s = cfg.startMin + cycle * i;
    out.push({ index: i + 1, start: fromMinutes_(s), end: fromMinutes_(s + cfg.slotMin) });
  }
  return out;
}

function makeSlotId_(date, cls, slotIndex) {
  return ymdCompact_(date) + '_' + String(cls).replace(/[\s　]/g, '') + '_' + slotIndex;
}

/**
 * 設定・面談日・クラスから枠マスタを作り直す。
 * 予約済/ブロックの内容は枠IDが一致する限り引き継ぐ。
 * 引き継げない予約がある場合は何も書き換えずに中止する。
 * @return {{written:number, kept:number}}
 */
function generateSlots() {
  var cfg = getConfig();
  var days = getDays();
  var classes = getClasses();
  if (!days.length) throw new Error('「面談日」シートに、実施するにチェックの入った日付がありません。');
  if (!classes.length) throw new Error('「クラス」シートにクラスが登録されていません。');

  var sh = sheet_(SH.SLOTS);
  var existing = {};
  var last = sh.getLastRow();
  if (last >= 2) {
    var cur = sh.getRange(2, 1, last - 1, SLOT_LAST_COL).getValues();
    for (var i = 0; i < cur.length; i++) {
      var id = String(cur[i][COL.SLOT_ID - 1] || '').trim();
      if (id) existing[id] = cur[i];
    }
  }

  var times = daySlotTimes_(cfg);
  var rows = [];
  var seen = {};
  var kept = 0;

  for (var d = 0; d < days.length; d++) {
    for (var t = 0; t < times.length; t++) {
      for (var c = 0; c < classes.length; c++) {
        var cls = classes[c];
        var id = makeSlotId_(days[d], cls.name, times[t].index);
        seen[id] = true;
        var prev = existing[id];
        var row = [
          id, days[d], times[t].start, times[t].end, cls.name, cls.teacher,
          STATUS.OPEN, '', '', '', '', '', ''
        ];
        if (prev && String(prev[COL.STATUS - 1]) !== STATUS.OPEN) {
          for (var k = COL.STATUS - 1; k < SLOT_LAST_COL; k++) row[k] = prev[k];
          kept++;
        }
        rows.push(row);
      }
    }
  }

  // 消える枠に予約が入っていないか確認
  var lost = [];
  for (var id2 in existing) {
    if (seen[id2]) continue;
    if (String(existing[id2][COL.STATUS - 1]) === STATUS.BOOKED) {
      lost.push(id2 + '（' + existing[id2][COL.STUDENT] + '）');
    }
  }
  if (lost.length) {
    throw new Error(
      '次の枠に予約が入っているため再生成を中止しました。先に予約を取り消すか、面談日・クラス・枠数の設定を戻してください。\n' +
      lost.join('\n')
    );
  }

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
 * 行 = 日付×時間、列 = クラス。空きが一目で分かる。
 */
function rebuildOverview() {
  var cfg = getConfig();
  var classes = getClasses();
  var days = getDays();
  var times = daySlotTimes_(cfg);
  var slots = readSlots_();

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
        var st = String(v[COL.STATUS - 1]);
        if (st === STATUS.BOOKED) {
          line.push(v[COL.NUMBER] + '. ' + v[COL.STUDENT]);
          lineColor.push('#ffffff');
        } else if (st === STATUS.BLOCKED) {
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
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  styleHeader_(sh, header.length);
  sh.getRange(1, 1, 1, header.length).setWrap(true).setHorizontalAlignment('center');
  if (body.length) {
    sh.getRange(2, 1, body.length, header.length).setValues(body);
    sh.getRange(2, 1, body.length, header.length).setBackgrounds(colors);
    sh.getRange(2, 1, body.length, header.length).setHorizontalAlignment('center');
    // 日付が変わるところに罫線
    for (var r = 0; r < body.length; r += times.length) {
      sh.getRange(r + 2, 1, 1, header.length).setBorder(true, null, null, null, null, null, '#5f6368', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }
  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 120);
  for (var cc = 0; cc < classes.length; cc++) sh.setColumnWidth(3 + cc, 150);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  return body.length;
}

/** 
 * クラスごとの予約表シートを作り直す（担任配布・管理用）
 * 左側(A-F): 【生徒別 予約状況（出席番号順）】 - 誰が予約済/未予約か一目でわかる
 * 右側(I-O): 【時間枠別 予約表（時間順）】 - 時間ごとの埋まり具合
 */
function rebuildClassSheets() {
  var ss = ss_();
  var classes = getClasses();
  var slots = readSlots_();
  var roster = getRoster();

  var headerLeft = ['出席番号', '生徒氏名', '予約状況', '予約日時', '保護者氏名', '連絡事項'];
  var headerRight = ['日付', '時間', '状態', '出席番号', '生徒氏名', '保護者氏名', '連絡事項'];

  for (var c = 0; c < classes.length; c++) {
    var clsName = classes[c].name;
    var name = CLASS_SHEET_PREFIX + clsName;
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clear();

    // 該当クラスの生徒を名簿から取得し、出席番号順にソート
    var classStudents = roster.filter(function (s) { return s.cls === clsName; });
    classStudents.sort(function (a, b) { return a.no - b.no; });

    // 該当クラスの予約済み枠を生徒番号でインデックス化
    var bookedByNo = {};
    for (var i = 0; i < slots.length; i++) {
      var v = slots[i].v;
      if (String(v[COL.CLASS - 1]) === clsName && String(v[COL.STATUS - 1]) === STATUS.BOOKED) {
        bookedByNo[Number(v[COL.NUMBER - 1])] = v;
      }
    }

    // 左側: 生徒別 予約状況リスト (出席番号順)
    var bodyLeft = [];
    var colorsLeft = [];
    for (var s = 0; s < classStudents.length; s++) {
      var st = classStudents[s];
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
        // 未予約は薄黄色で強調表示
        colorsLeft.push(['#fef7e0', '#fef7e0', '#fef7e0', '#fef7e0', '#fef7e0', '#fef7e0']);
      }
    }

    // 右側: 時間枠別 予約表 (時間順)
    var bodyRight = [];
    var colorsRight = [];
    for (var j = 0; j < slots.length; j++) {
      var sv = slots[j].v;
      if (String(sv[COL.CLASS - 1]) !== clsName) continue;
      var statusStr = String(sv[COL.STATUS - 1]);
      bodyRight.push([
        dateLabel_(sv[COL.DATE - 1]),
        sv[COL.START - 1] + '–' + sv[COL.END - 1],
        statusStr,
        sv[COL.NUMBER - 1] || '',
        sv[COL.STUDENT - 1] || '',
        sv[COL.GUARDIAN - 1] || '',
        sv[COL.NOTE - 1] || ''
      ]);
      var bg = statusStr === STATUS.OPEN ? '#e6f4ea' : (statusStr === STATUS.BLOCKED ? '#f1f3f4' : '#ffffff');
      colorsRight.push([bg, bg, bg, bg, bg, bg, bg]);
    }

    // --- 書き込み ---
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

    // 右側ヘッダー (I1:O1) - H列は間隔用空列
    sh.getRange(1, 9, 1, headerRight.length).setValues([headerRight]);
    sh.getRange(1, 9, 1, headerRight.length)
      .setFontWeight('bold')
      .setBackground('#e8eaed')
      .setVerticalAlignment('middle');

    if (bodyRight.length) {
      sh.getRange(2, 9, bodyRight.length, headerRight.length).setValues(bodyRight);
      sh.getRange(2, 9, bodyRight.length, headerRight.length).setBackgrounds(colorsRight);
    }

    // 書式設定
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 70);  // 出席番号
    sh.setColumnWidth(2, 120); // 生徒氏名
    sh.setColumnWidth(3, 80);  // 状況
    sh.setColumnWidth(4, 180); // 予約日時
    sh.setColumnWidth(5, 120); // 保護者氏名
    sh.setColumnWidth(6, 200); // 連絡事項

    sh.setColumnWidth(8, 30);  // 区切り列 (H列)

    sh.setColumnWidth(9, 110);  // 日付
    sh.setColumnWidth(10, 110); // 時間
    sh.setColumnWidth(11, 70);  // 状態
    sh.setColumnWidth(12, 70);  // 出席番号
    sh.setColumnWidth(13, 120); // 生徒氏名
    sh.setColumnWidth(14, 120); // 保護者氏名
    sh.setColumnWidth(15, 200); // 連絡事項
  }
  return classes.length;
}

/** まだ予約していない生徒の一覧 [{cls, no, name}] */
function unbookedStudents(clsFilter) {
  var roster = getRoster();
  var slots = readSlots_();
  var booked = {};
  for (var i = 0; i < slots.length; i++) {
    var v = slots[i].v;
    if (String(v[COL.STATUS - 1]) !== STATUS.BOOKED) continue;
    booked[v[COL.CLASS - 1] + '#' + Number(v[COL.NUMBER - 1])] = true;
  }
  return roster.filter(function (s) {
    if (clsFilter && s.cls !== clsFilter) return false;
    return !booked[s.cls + '#' + s.no];
  });
}

var SLOT_CACHE_KEY = 'slots_public';
var SLOT_CACHE_SEC = 15;

/**
 * 空き枠一覧の表示だけが使う短時間キャッシュ。
 * 予約の確定は readSlots_() で読み直したうえでロック内で状態を再確認するので、
 * ここが数秒古くても二重予約にはならない（「ちょうど埋まりました」と案内される）。
 * 日付は JSON 化で1日ずれないよう yyyy-MM-dd 文字列に直して保存する。
 */
function readSlotsCached_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(SLOT_CACHE_KEY);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* 壊れていたら読み直す */ }
  }
  var data = readSlots_();
  var plain = data.map(function (s) {
    var v = s.v.slice();
    v[COL.DATE - 1] = ymd_(v[COL.DATE - 1]);
    return { row: s.row, v: v };
  });
  try {
    cache.put(SLOT_CACHE_KEY, JSON.stringify(plain), SLOT_CACHE_SEC);
  } catch (e) { /* サイズ超過などは無視して素通し */ }
  return plain;
}

function clearSlotCache_() {
  try {
    CacheService.getScriptCache().remove(SLOT_CACHE_KEY);
  } catch (e) { /* noop */ }
}
