/**
 * 担任が「面談を入れられないコマ（だめなコマ）」を指定するためのシート。
 *
 * 行 = 面談日 × コマ、列 = クラス。担任は自分のクラスの列にチェックを入れるだけでよい。
 * チェックした枠は「枠マスタ」で「ブロック」になり、保護者の予約画面から選べなくなる。
 *
 * Web管理画面（?page=admin）での×指定とは常に同期する。
 * どちらから操作しても、このシートが担任にとっての一覧表になる。
 */

/** 1行目=案内バナー、2行目=見出し、3行目以降=データ */
var NG_BANNER_ROW = 1;
var NG_HEADER_ROW = 2;
var NG_FIRST_ROW = 3;

/** だめなコマシートの列番号(1始まり) */
var NG_COL = {
  DATE: 1,
  WEEKDAY: 2,
  INDEX: 3,
  TIME: 4,
  FIRST_CLASS: 5
};

var NG_BANNER_TEXT = '✅ 面談を入れられない枠に、自分のクラスの列でチェックを入れてください。' +
  '　入力後にメニュー「三者面談 ▸ 🚫 だめなコマ ▸ 指定を枠に反映する」を実行すると、保護者の画面から選べなくなります。' +
  '（チェックを外して再度反映すれば元に戻ります）';

/** 予約とぶつかっているセルの色 */
var NG_CONFLICT_BG = '#fce8e6';
var NG_NORMAL_BG = '#ffffff';

/**
 * 「だめなコマ」シートを作り直す。既存のチェックは日付・コマ・クラスで引き継ぐ。
 * @return {number} データ行数
 */
function rebuildNgSheet() {
  var cfg = getConfig();
  var days = getDays();
  var classes = getClasses();
  var times = daySlotTimes_(cfg);
  if (!days.length || !classes.length) return 0;

  var prev = readNgSet_();

  var header = ['日付', '曜日', 'コマ', '時間'].concat(classes.map(function (c) {
    return c.name + (c.teacher ? '\n' + c.teacher : '');
  }));

  var body = [];
  var dateBoundaries = [];
  for (var d = 0; d < days.length; d++) {
    dateBoundaries.push(body.length);
    for (var t = 0; t < times.length; t++) {
      var line = [
        days[d],
        WEEKDAY_JA[days[d].getDay()],
        times[t].index,
        times[t].start + '–' + times[t].end
      ];
      for (var c = 0; c < classes.length; c++) {
        line.push(!!prev[makeSlotId_(days[d], classes[c].name, times[t].index)]);
      }
      body.push(line);
    }
  }

  var ss = ss_();
  var sh = ss.getSheetByName(SH.NG) || ss.insertSheet(SH.NG);
  ensureSheetSize_(sh, body.length + NG_FIRST_ROW + 1, header.length);

  try {
    sh.setFrozenRows(0);
    sh.setFrozenColumns(0);
  } catch (e) { /* 無視 */ }
  // 既存のチェックボックス(入力規則)と結合を解いてから中身をクリアする
  var whole = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  whole.clearDataValidations();
  whole.breakApart();
  sh.clear();

  // 1行目: 案内バナー
  sh.getRange(NG_BANNER_ROW, 1, 1, header.length).merge()
    .setValue(NG_BANNER_TEXT)
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#fef7e0')
    .setFontColor('#b06000')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sh.setRowHeight(NG_BANNER_ROW, 40);

  // 2行目: 見出し
  sh.getRange(NG_HEADER_ROW, 1, 1, header.length).setValues([header]);
  sh.getRange(NG_HEADER_ROW, 1, 1, header.length)
    .setFontWeight('bold')
    .setBackground('#f1f3f4')
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(NG_HEADER_ROW, 30);

  if (body.length) {
    sh.getRange(NG_FIRST_ROW, 1, body.length, header.length).setValues(body);
    sh.getRange(NG_FIRST_ROW, NG_COL.DATE, body.length, 1).setNumberFormat('yyyy/mm/dd');
    sh.getRange(NG_FIRST_ROW, NG_COL.TIME, body.length, 1).setNumberFormat('@');
    sh.getRange(NG_FIRST_ROW, 1, body.length, NG_COL.FIRST_CLASS - 1)
      .setHorizontalAlignment('center')
      .setBackground('#ffffff');

    var chk = sh.getRange(NG_FIRST_ROW, NG_COL.FIRST_CLASS, body.length, classes.length);
    chk.insertCheckboxes();
    chk.setHorizontalAlignment('center');

    // 日付が変わるところに中太罫線
    for (var b = 0; b < dateBoundaries.length; b++) {
      sh.getRange(dateBoundaries[b] + NG_FIRST_ROW, 1, 1, header.length)
        .setBorder(true, null, null, null, null, null, '#5f6368', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }

  sh.setColumnWidth(NG_COL.DATE, 110);
  sh.setColumnWidth(NG_COL.WEEKDAY, 50);
  sh.setColumnWidth(NG_COL.INDEX, 50);
  sh.setColumnWidth(NG_COL.TIME, 120);
  for (var ci = 0; ci < classes.length; ci++) {
    sh.setColumnWidth(NG_COL.FIRST_CLASS + ci, 110);
  }
  // 1行目のバナーは全列を結合しているため、列の固定はできない（行のみ固定する）
  try {
    sh.setFrozenRows(NG_HEADER_ROW);
  } catch (e) { /* 無視 */ }

  return body.length;
}

/**
 * だめなコマシートを読み、{枠ID: true} の集合を返す。シートが無ければ空。
 */
function readNgSet_() {
  var out = {};
  var sh = ss_().getSheetByName(SH.NG);
  if (!sh) return out;

  var last = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (last < NG_FIRST_ROW || lastCol < NG_COL.FIRST_CLASS) return out;

  var header = sh.getRange(NG_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  var vals = sh.getRange(NG_FIRST_ROW, 1, last - NG_FIRST_ROW + 1, lastCol).getValues();

  for (var r = 0; r < vals.length; r++) {
    var dateVal = vals[r][NG_COL.DATE - 1];
    var idx = Number(vals[r][NG_COL.INDEX - 1]);
    if (!dateVal || !idx) continue;
    var date;
    try { date = toDate_(dateVal); } catch (e) { continue; }

    for (var c = NG_COL.FIRST_CLASS - 1; c < lastCol; c++) {
      var clsName = ngClassNameOf_(header[c]);
      if (!clsName) continue;
      if (isNgMark_(vals[r][c])) out[makeSlotId_(date, clsName, idx)] = true;
    }
  }
  return out;
}

/** 見出しは「1組\n山田先生」の形。1行目だけをクラス名として使う。 */
function ngClassNameOf_(headerCell) {
  return String(headerCell == null ? '' : headerCell).split('\n')[0].trim();
}

/** チェックボックスのTRUE、または ×・NG などの手入力を「だめ」とみなす */
function isNgMark_(v) {
  if (v === true) return true;
  var s = norm_(v);
  if (!s) return false;
  return s === 'true' || s === '×' || s === '✕' || s === '✖' || s === 'x' ||
    s === 'ng' || s === '●' || s === '■' || s === '✓' || s === '☑' || s === '1';
}

/** 枠ID "20261029_1組_3" を {ymd, cls, idx} に分解する */
function parseSlotId_(slotId) {
  var m = String(slotId == null ? '' : slotId).match(/^(\d{8})_(.+)_(\d+)$/);
  if (!m) return null;
  return { ymd: m[1], cls: m[2], idx: Number(m[3]) };
}

/**
 * だめなコマシートの内容を「枠マスタ」に反映する。
 * チェック済み → ブロック、チェックなしのブロック → 空きに戻す。
 * 予約が入っている枠は変更せず、conflicts として返す。
 * @return {{blocked:number, unblocked:number, conflicts:Array<string>}}
 */
function applyNgSlots() {
  // 枠マスタの状態列をまとめて書き換えるので、予約処理と同時に走らないようにロックする
  return withLock_(applyNgSlotsInner_);
}

function applyNgSlotsInner_() {
  var ng = readNgSet_();
  var sh = sheet_(SH.SLOTS);
  var last = sh.getLastRow();
  if (last < 2) return { blocked: 0, unblocked: 0, conflicts: [] };

  var slots = readSlots_();
  var statusCol = sh.getRange(2, COL.STATUS, last - 1, 1).getValues();
  var blocked = 0, unblocked = 0;

  for (var i = 0; i < slots.length; i++) {
    var v = slots[i].v;
    var id = String(v[COL.SLOT_ID - 1]);
    var cur = String(v[COL.STATUS - 1]);
    var want = !!ng[id];

    if (want) {
      if (cur === STATUS.BOOKED) continue;   // 予約は消さない。警告は後でまとめて出す
      if (cur !== STATUS.BLOCKED) {
        statusCol[slots[i].row - 2][0] = STATUS.BLOCKED;
        blocked++;
      }
    } else if (cur === STATUS.BLOCKED) {
      statusCol[slots[i].row - 2][0] = STATUS.OPEN;
      unblocked++;
    }
  }

  if (blocked || unblocked) {
    sh.getRange(2, COL.STATUS, last - 1, 1).setValues(statusCol);
    clearSlotCache_();
    markViewsStale_();
  }

  var conflicts = findNgConflicts_(ng, slots);
  try {
    markNgConflicts_(conflicts);
  } catch (e) {
    console.warn('だめなコマシートの警告表示をスキップ:', e);
  }

  if (blocked || unblocked || conflicts.length) {
    logAction_('だめなコマ反映', '', '', '', '',
      'ブロック ' + blocked + '件 / 解除 ' + unblocked + '件' +
      (conflicts.length ? ' / ⚠予約とぶつかり ' + conflicts.length + '件' : ''));
  }

  return { blocked: blocked, unblocked: unblocked, conflicts: conflicts };
}

/**
 * 「だめなコマ」に指定されているのに、すでに予約が入っている枠を洗い出す。
 * 担任がチェックを入れた時点で予約済みだったものが、そのまま残っていないかの確認用。
 * @return {Array<Object>}
 */
function findNgConflicts_(ngSet, slots) {
  var out = [];
  for (var i = 0; i < slots.length; i++) {
    var v = slots[i].v;
    if (String(v[COL.STATUS - 1]) !== STATUS.BOOKED) continue;
    var id = String(v[COL.SLOT_ID - 1]);
    if (!ngSet[id]) continue;
    out.push({
      slotId: id,
      dateLabel: dateLabel_(v[COL.DATE - 1]),
      time: String(v[COL.START - 1]) + '–' + String(v[COL.END - 1]),
      cls: String(v[COL.CLASS - 1]),
      no: v[COL.NUMBER - 1],
      name: String(v[COL.STUDENT - 1] || ''),
      guardian: String(v[COL.GUARDIAN - 1] || ''),
      code: String(v[COL.CODE - 1] || '')
    });
  }
  return out;
}

/** 警告メッセージ1行分 */
function ngConflictLabel_(c) {
  return c.dateLabel + ' ' + c.time + '　' + c.cls + ' ' + c.no + '. ' + c.name +
    (c.guardian ? '（保護者: ' + c.guardian + '）' : '');
}

/**
 * だめなコマシートに、予約とぶつかっているセルの警告（赤色＋メモ＋バナー）を表示する。
 * 解消済みの警告は自動で消える。
 */
function markNgConflicts_(conflicts) {
  var sh = ss_().getSheetByName(SH.NG);
  if (!sh) return;

  var last = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (last < NG_FIRST_ROW || lastCol < NG_COL.FIRST_CLASS) return;

  // 衝突が無く、バナーも通常表示のままなら、消すべき警告も無いので何もしない
  // （予約のたびに呼ばれるので、余計な読み書きを避ける）
  if (!conflicts.length &&
    String(sh.getRange(NG_BANNER_ROW, 1).getValue()) === NG_BANNER_TEXT) return;

  var header = sh.getRange(NG_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  var colOf = {};
  for (var c = NG_COL.FIRST_CLASS - 1; c < lastCol; c++) {
    var clsName = ngClassNameOf_(header[c]);
    if (clsName) colOf[clsName] = c - (NG_COL.FIRST_CLASS - 1);
  }

  var rowCount = last - NG_FIRST_ROW + 1;
  var width = lastCol - NG_COL.FIRST_CLASS + 1;
  var keys = sh.getRange(NG_FIRST_ROW, 1, rowCount, NG_COL.FIRST_CLASS - 1).getValues();
  var rowOf = {};
  for (var r = 0; r < keys.length; r++) {
    var dateVal = keys[r][NG_COL.DATE - 1];
    var idx = Number(keys[r][NG_COL.INDEX - 1]);
    if (!dateVal || !idx) continue;
    try { rowOf[ymdCompact_(dateVal) + '_' + idx] = r; } catch (e) { /* 無視 */ }
  }

  var notes = [], bgs = [];
  for (var r2 = 0; r2 < rowCount; r2++) {
    var noteRow = [], bgRow = [];
    for (var c2 = 0; c2 < width; c2++) { noteRow.push(''); bgRow.push(NG_NORMAL_BG); }
    notes.push(noteRow); bgs.push(bgRow);
  }

  for (var k = 0; k < conflicts.length; k++) {
    var q = parseSlotId_(conflicts[k].slotId);
    if (!q) continue;
    var ri = rowOf[q.ymd + '_' + q.idx];
    var ci = colOf[q.cls];
    if (ri == null || ci == null) continue;
    notes[ri][ci] = '⚠ この枠にはすでに面談の予約が入っています。\n' +
      conflicts[k].no + '. ' + conflicts[k].name +
      (conflicts[k].guardian ? '（保護者: ' + conflicts[k].guardian + '）' : '') + '\n\n' +
      '保護者に連絡して別の時間に移すか、予約を取り消してください。\n' +
      'そのあと「' + MENU.NG_APPLY + '」をもう一度実行すると、この警告は消えます。';
    bgs[ri][ci] = NG_CONFLICT_BG;
  }

  var rng = sh.getRange(NG_FIRST_ROW, NG_COL.FIRST_CLASS, rowCount, width);
  if (JSON.stringify(rng.getNotes()) !== JSON.stringify(notes)) {
    rng.setNotes(notes);
    rng.setBackgrounds(bgs);
  }

  setNgBanner_(sh, lastCol, conflicts);
}

/** 1行目のバナー。予約とぶつかっている指定があるときは赤い警告に切り替える。 */
function setNgBanner_(sh, lastCol, conflicts) {
  var text = NG_BANNER_TEXT;
  var bg = '#fef7e0';
  var fg = '#b06000';

  if (conflicts.length) {
    text = '⚠ すでに予約が入っている枠が ' + conflicts.length + ' 件、だめなコマに指定されています（赤いセル）。' +
      'このままでは、その時間に面談の予約が残ったままです。' +
      '　保護者に連絡して面談時間を移すか予約を取り消し、そのあと「' + MENU.NG_APPLY + '」をもう一度実行してください。';
    bg = NG_CONFLICT_BG;
    fg = '#c5221f';
  }

  var range = sh.getRange(NG_BANNER_ROW, 1, 1, lastCol);
  if (String(range.getValue()) === text) return;
  range.setValue(text).setBackground(bg).setFontColor(fg);
}

/** 枠1つ分のチェックを書き換える（Web管理画面・メニューからの操作を同期する） */
function setNgFlag_(slotId, flag) {
  setNgFlags_([{ slotId: slotId, flag: flag }]);
}

/**
 * 複数枠のチェックをまとめて書き換える。
 * @param {Array<{slotId:string, flag:boolean}>} pairs
 */
function setNgFlags_(pairs) {
  if (!pairs || !pairs.length) return;
  var sh = ss_().getSheetByName(SH.NG);
  if (!sh) return;

  var last = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (last < NG_FIRST_ROW || lastCol < NG_COL.FIRST_CLASS) return;

  var header = sh.getRange(NG_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  var colOf = {};
  for (var c = NG_COL.FIRST_CLASS - 1; c < lastCol; c++) {
    var clsName = ngClassNameOf_(header[c]);
    if (clsName) colOf[clsName] = c - (NG_COL.FIRST_CLASS - 1);
  }

  var rowCount = last - NG_FIRST_ROW + 1;
  var width = lastCol - NG_COL.FIRST_CLASS + 1;
  var keys = sh.getRange(NG_FIRST_ROW, 1, rowCount, NG_COL.FIRST_CLASS - 1).getValues();
  var flagRange = sh.getRange(NG_FIRST_ROW, NG_COL.FIRST_CLASS, rowCount, width);
  var flags = flagRange.getValues();

  var rowOf = {};
  for (var r = 0; r < keys.length; r++) {
    var dateVal = keys[r][NG_COL.DATE - 1];
    var idx = Number(keys[r][NG_COL.INDEX - 1]);
    if (!dateVal || !idx) continue;
    try { rowOf[ymdCompact_(dateVal) + '_' + idx] = r; } catch (e) { /* 無視 */ }
  }

  var dirty = false;
  for (var p = 0; p < pairs.length; p++) {
    var q = parseSlotId_(pairs[p].slotId);
    if (!q) continue;
    var rIdx = rowOf[q.ymd + '_' + q.idx];
    var cIdx = colOf[q.cls];
    if (rIdx == null || cIdx == null) continue;
    var want = !!pairs[p].flag;
    if (isNgMark_(flags[rIdx][cIdx]) !== want) {
      flags[rIdx][cIdx] = want;
      dirty = true;
    }
  }

  if (dirty) flagRange.setValues(flags);
}
