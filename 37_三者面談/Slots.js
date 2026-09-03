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
  // クラスや面談日を直したあとに実行されるので、古い読み取り結果を使わない
  dropRefCaches_();

  var cfg = getConfig();
  var days = getDays();
  var classes = getClasses();
  if (!days.length) throw new Error('「面談日」シートで「実施する」がTRUEの日付が1つもありません。');
  if (!classes.length) throw new Error('「クラス」シートにクラス名を入力してください。');

  // 枠IDは「日付＋クラス名＋コマ番号」で作るため、重複があると枠が二重にできてしまう
  var dupCls = findDuplicates_(classes.map(function (c) { return c.name; }));
  if (dupCls.length) {
    throw new Error('「' + SH.CLASSES + '」シートに同じクラス名が複数あります: ' + dupCls.join('、') +
      '\n重複した行を消してから、もう一度実行してください。');
  }
  var dupDay = findDuplicates_(days.map(function (d) { return dateLabel_(d); }));
  if (dupDay.length) {
    throw new Error('「' + SH.DAYS + '」シートに同じ日付が複数あります: ' + dupDay.join('、') +
      '\n重複した行を消すか、片方の「実施する」のチェックを外してください。');
  }

  if (!(cfg.slotMin >= 1)) {
    throw new Error('「' + SH.CONFIG + '」シートの「面談枠の長さ(分)」に1以上の数を入れてください。');
  }
  if (!(cfg.breakMin >= 0)) {
    throw new Error('「' + SH.CONFIG + '」シートの「枠間の休憩(分)」に0以上の数を入れてください。');
  }
  if (!(cfg.slotsPerDay >= 1)) {
    throw new Error('「' + SH.CONFIG + '」シートの「1日の枠数」に1以上の数を入れてください。');
  }

  // 読み取りから書き込みまでの間に予約が入ると、その予約ごと消してしまう。
  // 保護者の予約処理と同じロックで囲む。
  var result = withLock_(function () {
    var existing = readSlots_();
    var existingMap = {};
    var lost = [];
    for (var i = 0; i < existing.length; i++) {
      var v = existing[i].v;
      var id = String(v[COL.SLOT_ID - 1]);
      existingMap[id] = v;
    }

    var times = dayAllTimes_(cfg);
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
            if (isTakenSlot_(old)) kept++;
          } else {
            rows.push([
              slotId,
              days[d],
              times[t].start,
              times[t].end,
              classes[c].name,
              classes[c].teacher || '',
              times[t].reserve ? STATUS.RESERVE : STATUS.OPEN,
              '', '', '', '', '', ''
            ]);
          }
        }
      }
    }

    for (var k = 0; k < existing.length; k++) {
      var ev = existing[k].v;
      var eid = String(ev[COL.SLOT_ID - 1]);
      if (!newSlotIds[eid] && isTakenSlot_(ev)) {
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
    return { count: rows.length, kept: kept };
  });

  // 日程に合わせて「だめなコマ」シートを作り直し、指定済みの枠を再びブロックにする
  var ng = { blocked: 0, unblocked: 0, conflicts: [] };
  try {
    rebuildNgSheet();
    ng = applyNgSlots();
  } catch (e) {
    console.warn('だめなコマの反映をスキップ:', e);
  }

  logAction_('枠再生成', '', '', '', '', result.count + '枠 / 引継ぎ ' + result.kept + '件');
  return { written: result.count, kept: result.kept, ngBlocked: ng.blocked, ngConflicts: ng.conflicts };
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
 * 学年の見出しに使う色。クラスが多いと列の切れ目が分からなくなるため、
 * 学年ごとに背景を変える。空き(#e6f4ea)・予約済(#ffffff)・面談なし(#f1f3f4)・
 * 警告(#fce8e6) と紛れない色だけを選んである。
 */
var GRADE_HEADER_COLORS = ['#d2e3fc', '#e8d9f7', '#ffe0b2', '#b3e5e0', '#f8d7e3', '#e0d5c1'];
var SPECIAL_HEADER_COLOR = '#dfe1e5';

/**
 * 本文に重ねる学年色の濃さ。0で従来どおり、1で学年色そのもの。
 * 淡い色を「空き」の緑に重ねても隣の学年と見分けが付かないため、
 * 色を重ねるだけでなく、学年の並びが1つおきに少し暗くなるようにしている。
 * 明るさの差のほうが、色の差より確実に伝わる。
 */
var GRADE_BODY_RATIO = 0.16;
var GRADE_BODY_SHADE = 0.955;

/** "#rrggbb" を [r,g,b] に */
function hexToRgb_(hex) {
  var h = String(hex).replace('#', '');
  return [parseInt(h.substring(0, 2), 16),
          parseInt(h.substring(2, 4), 16),
          parseInt(h.substring(4, 6), 16)];
}

function toHex2_(n) {
  var v = Math.max(0, Math.min(255, Math.round(n))).toString(16);
  return v.length === 1 ? '0' + v : v;
}

/**
 * base に tint を ratio の割合で重ねる。
 * 空き・面談なし・⚠ の意味は base の側に残したいので、ratio は小さくしてある。
 */
function mixHex_(base, tint, ratio) {
  if (!tint) return base;
  var a = hexToRgb_(base), b = hexToRgb_(tint);
  return '#' +
    toHex2_(a[0] * (1 - ratio) + b[0] * ratio) +
    toHex2_(a[1] * (1 - ratio) + b[1] * ratio) +
    toHex2_(a[2] * (1 - ratio) + b[2] * ratio);
}

/**
 * 本文セルの色。status の色に学年色を薄く重ね、1つおきの学年はさらに少し暗くする。
 * @param {string} base 状態を表す色
 * @param {?{hex:string, shade:boolean}} band 学年の帯。null なら何もしない
 */
function bodyColor_(base, band) {
  if (!band) return base;
  var c = mixHex_(base, band.hex, GRADE_BODY_RATIO);
  if (!band.shade) return c;
  var rgb = hexToRgb_(c);
  return '#' + toHex2_(rgb[0] * GRADE_BODY_SHADE) +
               toHex2_(rgb[1] * GRADE_BODY_SHADE) +
               toHex2_(rgb[2] * GRADE_BODY_SHADE);
}

/**
 * 学年名 → 見出しの背景色。gradeList の並び順で色を割り当てる。
 * 特別支援は順番に関係なく灰色で固定する。
 */
function gradeHeaderColor_(grade, gradeList) {
  if (grade === '特別支援') return SPECIAL_HEADER_COLOR;
  var i = gradeList.indexOf(grade);
  if (i < 0) return SPECIAL_HEADER_COLOR;
  return GRADE_HEADER_COLORS[i % GRADE_HEADER_COLORS.length];
}

/**
 * 全体ビューを作り直す。
 * 1行目: 学年・各クラス予約進捗サマリーバナー
 * 2行目: 見出しヘッダー
 * 3行目〜: 行 = 日付×時間、列 = クラス。空きが一目で分かる。
 */
function rebuildOverview() {
  var cfg = getConfig();
  var classes = getClasses().slice().sort(function (a, b) {
    var d = gradeOrder_(a.grade) - gradeOrder_(b.grade);
    return d !== 0 ? d : (a.name < b.name ? -1 : (a.name > b.name ? 1 : 0));
  });
  var days = getDays();
  var times = dayAllTimes_(cfg);   // 予備コマも含める
  var slots = readSlots_();
  var roster = getRoster();

  var ngSet = {};
  try { ngSet = readNgSet_(); } catch (e) { ngSet = {}; }
  var ngConflicts = findNgConflicts_(ngSet, slots);

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
    if (isTakenSlot_(sv)) {
      bookedCount++;
      var cName = String(sv[COL.CLASS - 1]);
      if (classStats[cName]) classStats[cName].booked++;
    }
  }

  var totalRate = totalStudents ? Math.round((bookedCount / totalStudents) * 100) : 0;
  // A1:B1 は固定列に入るので、右へスクロールしても消えない全体の進捗をここに置く
  var leftSummary = '📊 ' + bookedCount + '/' + totalStudents + '名 (' + totalRate + '%)';
  var summaryText = '';
  var classSummaries = [];
  for (var cc = 0; cc < classes.length; cc++) {
    var cN = classes[cc].name;
    var st = classStats[cN] || { total: 0, booked: 0 };
    var cRate = st.total ? Math.round((st.booked / st.total) * 100) : 0;
    classSummaries.push(cN + ': ' + st.booked + '/' + st.total + '名 (' + cRate + '%)');
  }
  // クラス数が多い学校では1行に収まらないので、学年ごとにまとめる
  var byGrade = {}, gradeList = [];
  for (var gi = 0; gi < classes.length; gi++) {
    var g = classes[gi].grade || 'その他';
    if (!byGrade[g]) { byGrade[g] = { total: 0, booked: 0 }; gradeList.push(g); }
    var gs = classStats[classes[gi].name] || { total: 0, booked: 0 };
    byGrade[g].total += gs.total;
    byGrade[g].booked += gs.booked;
  }
  gradeList.sort(function (x, y) { return gradeOrder_(x) - gradeOrder_(y); });

  var gradeSummaries = [];
  for (var gk = 0; gk < gradeList.length; gk++) {
    var gname = gradeList[gk];
    var gst = byGrade[gname];
    var grate = gst.total ? Math.round((gst.booked / gst.total) * 100) : 0;
    gradeSummaries.push(gname + ' ' + gst.booked + '/' + gst.total + '名(' + grate + '%)');
  }

  // 学年が1つしかない（従来どおりの使い方）ならクラス別のまま
  summaryText += (gradeList.length > 1 ? gradeSummaries : classSummaries).join(' ｜ ');
  summaryText += '　🕒 ' + Utilities.formatDate(new Date(), TZ, 'M/d HH:mm') + ' 現在';
  if (ngConflicts.length) {
    summaryText += '　⚠ だめなコマに指定した枠に予約が ' + ngConflicts.length + ' 件残っています（「' +
      SH.NG + '」シートの赤いセルを確認してください）';
  }

  var byId = {};
  for (var i = 0; i < slots.length; i++) byId[slots[i].v[COL.SLOT_ID - 1]] = slots[i].v;

  var header = ['日付', '時間'].concat(classes.map(function (c) {
    return (c.grade && gradeList.length > 1 ? c.grade + '\n' : '') +
      c.name + (c.teacher ? '\n' + c.teacher : '');
  })).concat(['空き数']);

  // 各クラスの本文に敷く学年の帯（学年が1つだけなら敷かない）
  var bodyTint = classes.map(function (c) {
    if (gradeList.length <= 1) return null;
    var g = c.grade || 'その他';
    var gi = gradeList.indexOf(g);
    return { hex: gradeHeaderColor_(g, gradeList), shade: gi % 2 === 1 };
  });

  var body = [];
  var colors = [];
  for (var d = 0; d < days.length; d++) {
    for (var t = 0; t < times.length; t++) {
      var isReserve = !!times[t].reserve;
      var line = [dateLabel_(days[d]),
        times[t].start + '–' + times[t].end + (isReserve ? '（予備）' : '')];
      var lineColor = ['#ffffff', isReserve ? '#fef7e0' : '#ffffff'];
      var free = 0;
      for (var c = 0; c < classes.length; c++) {
        var v = byId[makeSlotId_(days[d], classes[c].name, times[t].index)];
        var tint = bodyTint[c];
        if (!v) {
          line.push('—'); lineColor.push(bodyColor_('#f8f9fa', tint)); continue;
        }
        var statusStr = String(v[COL.STATUS - 1]);
        if (statusStr === STATUS.BOOKED) {
          var warn = !!ngSet[String(v[COL.SLOT_ID - 1])];
          line.push((warn ? '⚠ ' : '') + v[COL.NUMBER - 1] + '. ' + v[COL.STUDENT - 1]);
          // ⚠ は見落とすと困るので、赤には学年色を重ねない
          lineColor.push(warn ? '#fce8e6' : bodyColor_('#ffffff', tint));
        } else if (statusStr === STATUS.BLOCKED) {
          line.push('×');
          lineColor.push(bodyColor_('#f1f3f4', tint));
        } else if (statusStr === STATUS.RESERVE) {
          // 予備は保護者に見えない枠。担任が氏名を書き込むと、そこに面談が入る
          var rname = String(v[COL.STUDENT - 1] || '');
          line.push(rname
            ? (v[COL.NUMBER - 1] ? v[COL.NUMBER - 1] + '. ' : '') + rname
            : '予備');
          lineColor.push(rname ? bodyColor_('#ffffff', tint) : '#fef7e0');
        } else {
          line.push('空き');
          lineColor.push(bodyColor_('#e6f4ea', tint));
          free++;
        }
      }
      // 予備の行は「空き数」に数えない。埋まっていないのが普通だから
      line.push(isReserve ? '—' : free);
      lineColor.push(isReserve ? '#fef7e0' : (free === 0 ? '#fce8e6' : '#ffffff'));
      body.push(line);
      colors.push(lineColor);
    }
  }

  var sh = ss_().getSheetByName(SH.OVERVIEW) || ss_().insertSheet(SH.OVERVIEW);
  ensureSheetSize_(sh, body.length + 3, header.length);
  
  // ★ 固定列・固定行を一旦リセットしてからクリア（結合エラー防止）
  try {
    sh.setFrozenRows(0);
    sh.setFrozenColumns(0);
  } catch (e) { /* 無視 */ }

  sh.clear();

  // 1行目: 進捗サマリーバナー
  //
  // A1:B1 と C1:最終列 の2つに分けて結合する。
  // 1行目を全列まとめて結合すると、結合範囲が固定列の境目をまたぐことになり、
  // setFrozenColumns(2) が拒否される。背景色を揃えてあるので、見た目は1本の帯のまま。
  var bannerBg = ngConflicts.length ? '#fce8e6' : '#e8f0fe';
  var bannerFg = ngConflicts.length ? '#c5221f' : '#1a73e8';

  // 前回の結合（旧版では全列1つ）が残っていると結合し直せないので、先にほどく
  sh.getRange(1, 1, 1, header.length).breakApart();

  sh.getRange(1, 1, 1, 2).merge()
    .setValue(leftSummary)
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground(bannerBg)
    .setFontColor(bannerFg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  var rightWidth = header.length - 2;
  var rightBanner = sh.getRange(1, 3, 1, rightWidth);
  if (rightWidth > 1) rightBanner.merge();
  rightBanner
    .setValue(summaryText)
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground(bannerBg)
    .setFontColor(bannerFg)
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

  // クラスの見出しだけ、学年ごとに色を変える（学年が1つなら従来どおり灰色）
  if (classes.length && gradeList.length > 1) {
    var headColors = classes.map(function (c) {
      return gradeHeaderColor_(c.grade || 'その他', gradeList);
    });
    sh.getRange(2, 3, 1, classes.length).setBackgrounds([headColors]);
  }

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
  // 学年が変わる列の左に区切り線を入れる（17クラス並ぶと切れ目が見えないため）
  if (gradeList.length > 1 && body.length) {
    for (var bi = 1; bi < classes.length; bi++) {
      if ((classes[bi].grade || '') === (classes[bi - 1].grade || '')) continue;
      sh.getRange(2, 3 + bi, body.length + 1, 1)
        .setBorder(null, true, null, null, null, null,
          '#5f6368', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }

  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 120);
  for (var colIdx = 0; colIdx < classes.length; colIdx++) sh.setColumnWidth(3 + colIdx, 150);

  // 上部2行（サマリーバナー＋見出し）と、左2列（日付・時間）を固定表示
  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);

  // だめなコマシート側の警告表示も最新にする（予約が取り消されれば自動で消える）
  try {
    markNgConflicts_(ngConflicts);
  } catch (e) {
    console.warn('だめなコマシートの警告表示をスキップ:', e);
  }

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

  // 予約表は毎回作り直すので、担任が予備の行に書いた内容は先に拾って枠マスタへ移す。
  // これをしないと、書いた内容が次の更新で消える。
  try {
    captureReserveEntries_();
  } catch (e) {
    console.warn('予備コマの取り込みをスキップ:', e);
  }

  var slots = readSlots_();
  var ngSet = {};
  try { ngSet = readNgSet_(); } catch (e) { ngSet = {}; }
  var links = {};
  try {
    getLinks_().forEach(function (L) {
      links[L.a.cls + '_' + L.a.no] = L.b;
      links[L.b.cls + '_' + L.b.no] = L.a;
    });
  } catch (e) { links = {}; }

  var headerLeft = CLASS_HEADER_LEFT;
  var headerRight = CLASS_HEADER_RIGHT;
  var rightWidth = headerRight.length + 1;   // 末尾に枠IDの列を足す

  for (var c = 0; c < classes.length; c++) {
    var clsName = classes[c].name;
    var name = CLASS_SHEET_PREFIX + clsName;
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    ensureSheetSize_(sh, 200, 8 + rightWidth);

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
    var reserveByNo = {};
    for (var i = 0; i < slots.length; i++) {
      var v = slots[i].v;
      if (String(v[COL.CLASS - 1]) !== clsName) continue;
      var stt = String(v[COL.STATUS - 1]);
      if (stt === STATUS.BOOKED) {
        bookedByNo[Number(v[COL.NUMBER - 1])] = v;
      } else if (stt === STATUS.RESERVE && v[COL.NUMBER - 1]) {
        // 予備の枠に担任が入れたぶん
        reserveByNo[Number(v[COL.NUMBER - 1])] = v;
      }
    }

    // 左側 (A〜F): 生徒別 予約状況
    var bodyLeft = [];
    var colorsLeft = [];
    for (var s = 0; s < existingStudents.length; s++) {
      var st = existingStudents[s];
      var bk = bookedByNo[st.no];

      // 交流学級のほうで予約している場合は、そちらを表示する
      if (!bk) {
        var partner = links[clsName + '_' + st.no];
        if (partner) {
          for (var pi = 0; pi < slots.length; pi++) {
            var pv = slots[pi].v;
            if (String(pv[COL.STATUS - 1]) !== STATUS.BOOKED) continue;
            if (String(pv[COL.CLASS - 1]) !== partner.cls) continue;
            if (Number(pv[COL.NUMBER - 1]) !== partner.no) continue;
            bodyLeft.push([
              st.no, st.name, '交流学級で予約済',
              dateLabel_(pv[COL.DATE - 1]) + ' ' + pv[COL.START - 1] + '–' + pv[COL.END - 1],
              pv[COL.GUARDIAN - 1] || '',
              partner.cls + ' の担任と面談します'
            ]);
            colorsLeft.push(['#e8f0fe', '#e8f0fe', '#e8f0fe', '#e8f0fe', '#e8f0fe', '#e8f0fe']);
            break;
          }
          if (bodyLeft.length === s + 1) continue;
        }
      }

      if (bk) {
        var bkWarn = !!ngSet[String(bk[COL.SLOT_ID - 1])];
        bodyLeft.push([
          st.no,
          st.name,
          bkWarn ? '⚠ 要調整' : '予約済',
          dateLabel_(bk[COL.DATE - 1]) + ' ' + bk[COL.START - 1] + '–' + bk[COL.END - 1],
          bk[COL.GUARDIAN - 1] || '',
          bkWarn
            ? ('⚠ 「だめなコマ」に指定された時間です。時間の変更をご相談ください。' +
               (bk[COL.NOTE - 1] ? ' ／ ' + bk[COL.NOTE - 1] : ''))
            : (bk[COL.NOTE - 1] || '')
        ]);
        colorsLeft.push(bkWarn
          ? ['#fce8e6', '#fce8e6', '#fce8e6', '#fce8e6', '#fce8e6', '#fce8e6']
          : ['#ffffff', '#ffffff', '#e6f4ea', '#ffffff', '#ffffff', '#ffffff']);
      } else if (reserveByNo[st.no]) {
        var rv = reserveByNo[st.no];
        bodyLeft.push([
          st.no,
          st.name,
          '予約済（予備）',
          dateLabel_(rv[COL.DATE - 1]) + ' ' + rv[COL.START - 1] + '–' + rv[COL.END - 1],
          rv[COL.GUARDIAN - 1] || '',
          rv[COL.NOTE - 1] || '担任が予備の枠に入れました'
        ]);
        colorsLeft.push(['#ffffff', '#ffffff', '#fef7e0', '#ffffff', '#ffffff', '#ffffff']);
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
      var rowWarn = statusStr === STATUS.BOOKED && !!ngSet[String(sv[COL.SLOT_ID - 1])];
      var isRes = statusStr === STATUS.RESERVE;
      bodyRight.push([
        curDate,
        sv[COL.START - 1] + '–' + sv[COL.END - 1] + (isRes ? '（予備）' : ''),
        rowWarn ? '⚠ 予約済（だめなコマ指定）' : statusStr,
        sv[COL.NUMBER - 1] || '',
        sv[COL.STUDENT - 1] || '',
        sv[COL.GUARDIAN - 1] || '',
        sv[COL.CODE - 1] ? "'" + String(sv[COL.CODE - 1]) : '',
        // 取り込みのときにどの枠かを見分けるための列。担任には見せない
        String(sv[COL.SLOT_ID - 1])
      ]);
      var bg = rowWarn ? '#fce8e6'
        : (statusStr === STATUS.OPEN ? '#e6f4ea'
          : (statusStr === STATUS.BLOCKED ? '#f1f3f4'
            : (isRes && !sv[COL.STUDENT - 1] ? '#fef7e0' : '#ffffff')));
      colorsRight.push([bg, bg, bg, bg, bg, bg, bg, bg]);
    }

    try {
      sh.setFrozenRows(0);
      sh.setFrozenColumns(0);
    } catch (e) { /* 無視 */ }

    // 名簿（A・B列）は、このシートにしか無い唯一の控えである。
    // sh.clear() で全部消してから書き戻す作りだと、消したあと書き戻す前に
    // 実行時間の上限や一時的なエラーで止まった瞬間、そのクラスの名簿が失われる。
    // 表示は5分おきのトリガーからも走るため、消さずに上書きし、
    // はみ出した古い行だけをあとから掃除する。
    var oldLast = Math.max(lastRow, 1);

    // 前回の日付区切りの罫線が残らないように、先に消しておく
    if (oldLast >= 2) {
      sh.getRange(2, 9, oldLast - 1, rightWidth)
        .setBorder(false, false, false, false, false, false);
    }

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
    clearSheetTail_(sh, 1, headerLeft.length, 2 + bodyLeft.length, oldLast);

    // 右側ヘッダー (I1:P1)。P列は取り込み用の枠IDで、あとで隠す
    sh.getRange(1, 9, 1, rightWidth).setValues([headerRight.concat(['枠ID'])]);
    sh.getRange(1, 9, 1, rightWidth)
      .setFontWeight('bold')
      .setBackground('#e8eaed')
      .setVerticalAlignment('middle');

    if (bodyRight.length) {
      sh.getRange(2, 9, bodyRight.length, rightWidth).setValues(bodyRight);
      sh.getRange(2, 9, bodyRight.length, rightWidth).setBackgrounds(colorsRight);

      for (var bIdx = 0; bIdx < dateBoundaries.length; bIdx++) {
        var rNum = dateBoundaries[bIdx] + 2;
        sh.getRange(rNum, 9, 1, rightWidth)
          .setBorder(true, null, null, null, null, null, '#5f6368', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    }
    clearSheetTail_(sh, 9, rightWidth, 2 + bodyRight.length, oldLast);

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
    try { sh.hideColumns(9 + headerRight.length); } catch (e) { /* 既に非表示 */ }
  }

  // A・B列を書き直したので、名簿の読み取り結果を作り直させる
  clearRosterCache_();
}

/** キャッシュされた空き枠を読む */
/**
 * 「予約表_〇組」の予備の行に担任が書き込んだ内容を、枠マスタへ取り込む。
 *
 * 予約表は毎回作り直される表示用のシートなので、ここで拾わないと
 * 書いた内容は次の更新で消えてしまう。名簿（A・B列）と同じ扱いにしている。
 * 拾うのは予備の行だけ。通常の枠に手で書いても取り込まない
 * （予約の重複確認を通さずに面談が入ってしまうため）。
 *
 * @return {number} 取り込んだ行数
 */
function captureReserveEntries_() {
  var ss = ss_();
  var slotSh = ss.getSheetByName(SH.SLOTS);
  if (!slotSh) return 0;

  var slots = readSlots_();
  var byId = {};
  var hasReserve = false;
  for (var i = 0; i < slots.length; i++) {
    var sv = slots[i].v;
    byId[String(sv[COL.SLOT_ID - 1])] = slots[i];
    if (String(sv[COL.STATUS - 1]) === STATUS.RESERVE) hasReserve = true;
  }
  if (!hasReserve) return 0;   // 予備を使っていない学校では何もしない

  var roster = getRoster();
  var byName = {}, byNo = {};
  for (var r = 0; r < roster.length; r++) {
    byName[roster[r].cls + '|' + norm_(roster[r].name)] = roster[r];
    byNo[roster[r].cls + '|' + roster[r].no] = roster[r];
  }

  var classes = getClasses();
  var width = CLASS_HEADER_RIGHT.length + 1;   // 末尾が枠IDの列
  var changed = 0;

  for (var c = 0; c < classes.length; c++) {
    var sh = ss.getSheetByName(CLASS_SHEET_PREFIX + classes[c].name);
    if (!sh) continue;
    var last = sh.getLastRow();
    if (last < 2) continue;

    var vals = sh.getRange(2, 9, last - 1, width).getValues();
    for (var v = 0; v < vals.length; v++) {
      var id = String(vals[v][width - 1] || '').trim();
      if (!id) continue;

      var target = byId[id];
      if (!target) continue;
      if (String(target.v[COL.STATUS - 1]) !== STATUS.RESERVE) continue;

      var typedNo = Number(vals[v][3]) || 0;             // 出席番号
      var typedName = String(vals[v][4] || '').trim();   // 生徒氏名
      var typedGuardian = String(vals[v][5] || '').trim();

      // 番号だけ、氏名だけでも通るよう、名簿と突き合わせて足りないほうを補う
      var hit = null;
      if (typedName) hit = byName[classes[c].name + '|' + norm_(typedName)];
      if (!hit && typedNo) hit = byNo[classes[c].name + '|' + typedNo];
      if (hit) { typedNo = hit.no; typedName = hit.name; }

      var curNo = Number(target.v[COL.NUMBER - 1]) || 0;
      var curName = String(target.v[COL.STUDENT - 1] || '');
      var curGuardian = String(target.v[COL.GUARDIAN - 1] || '');
      if (typedNo === curNo && typedName === curName && typedGuardian === curGuardian) continue;

      slotSh.getRange(target.row, COL.NUMBER, 1, 3)
        .setValues([[typedNo || '', typedName, typedGuardian]]);
      changed++;

      logAction_(typedName ? '予備コマに記入' : '予備コマを空に',
        id, classes[c].name, typedNo || '', typedName,
        typedName ? '担任が予約表に記入' : '担任が予約表から消去');
    }
  }

  if (changed) clearSlotCache_();
  return changed;
}

/**
 * 書き込んだ行より下に残っている古い内容を消す。
 * シート全体を clear() せずに済ませるための後始末。
 */
function clearSheetTail_(sh, col, width, fromRow, oldLast) {
  if (!oldLast || oldLast < fromRow) return;
  var rng = sh.getRange(fromRow, col, oldLast - fromRow + 1, width);
  rng.clearContent();
  rng.setBackground(null);
  rng.setBorder(false, false, false, false, false, false);
}

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

/**
 * 予備コマのコマ番号は、通常のコマと衝突しないよう 900 番台にする。
 * 枠IDの見た目は 20261029_1年1組_901 のようになる。
 */
var RESERVE_INDEX_BASE = 900;

/**
 * 予備コマの時刻。最終コマの後ろに、同じ間隔で続けて作る。
 * 保護者には見せないが、時間が決まっていたほうが当日の受付一覧に並べられる。
 */
function dayReserveTimes_(cfg) {
  var out = [];
  var n = Number(cfg.reservePerDay) || 0;
  if (n <= 0) return out;

  var cycle = cfg.slotMin + cfg.breakMin;
  var cur = cfg.startMin + cycle * cfg.slotsPerDay;
  for (var i = 1; i <= n; i++) {
    var end = cur + cfg.slotMin;
    out.push({ index: RESERVE_INDEX_BASE + i, start: fromMinutes_(cur), end: fromMinutes_(end), reserve: true });
    cur = end + cfg.breakMin;
  }
  return out;
}

/** 通常＋予備。表を作るときは、この順で1日分が並ぶ */
function dayAllTimes_(cfg) {
  return daySlotTimes_(cfg).concat(dayReserveTimes_(cfg));
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
