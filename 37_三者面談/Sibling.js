/**
 * きょうだいまとめての予約。
 *
 * 全クラスが同じ時間割で動くので、「同じ日の連続したコマ」を
 * 別々のクラスから1つずつ確保する、という形で扱える。
 * コマとコマの間の休憩が、そのまま教室移動の時間になる。
 *
 * システムは組み合わせを提示するだけで、どれを選ぶかは保護者が決める。
 * 間があいたほうが都合のよい家庭もあるため、順位づけはしない。
 */

var SIBLING_MAX = 6;

/* ================================================================
   受付期間（きょうだい世帯は先に開けられる）
   ================================================================ */

/**
 * きょうだい予約の受付期間。
 * 「きょうだい予約の受付開始」が設定されていれば、通常より早く開く。
 */
function siblingBookingWindow_(cfg) {
  var now = new Date();
  if (!cfg.published) {
    return { ok: false, message: '現在、予約の受付を停止しています。' };
  }

  var openAt = cfg.siblingOpenAt || cfg.openAt;
  if (openAt && now < openAt) {
    return {
      ok: false,
      message: 'きょうだいまとめての予約は ' +
        Utilities.formatDate(openAt, TZ, 'M月d日 HH:mm') + ' から開始します。'
    };
  }
  if (cfg.closeAt && now > cfg.closeAt) {
    return {
      ok: false,
      message: '予約受付は ' + Utilities.formatDate(cfg.closeAt, TZ, 'M月d日 HH:mm') +
        ' で締め切りました。変更は担任までご連絡ください。'
    };
  }
  return { ok: true, message: '' };
}

/* ================================================================
   保護者向け API
   ================================================================ */

/** 最初の画面。学年ごとのクラス一覧と、受付状況を返す。 */
function apiSiblingInit() {
  return safe_(function () {
    var cfg = getConfig();
    var win = siblingBookingWindow_(cfg);
    var classes = getClasses();

    var byGrade = {}, order = [];
    for (var i = 0; i < classes.length; i++) {
      var g = classes[i].grade || 'その他';
      if (!byGrade[g]) { byGrade[g] = []; order.push(g); }
      byGrade[g].push(classes[i].name);
    }
    order.sort(function (a, b) { return gradeOrder_(a) - gradeOrder_(b); });

    return {
      title: cfg.title,
      open: win.ok,
      message: win.message,
      checkName: cfg.checkName,
      maxChildren: SIBLING_MAX,
      grades: order.map(function (g) { return { grade: g, classes: byGrade[g] }; })
    };
  });
}

/**
 * きょうだい全員を照合し、同じ日に取れる組み合わせを返す。
 */
function apiSiblingSlots(p) {
  return safe_(function () {
    var cfg = getConfig();
    var win = siblingBookingWindow_(cfg);
    if (!win.ok) throw new Error(win.message);

    var children = verifySiblings_(p, cfg);
    var slots = readSlots_();

    var already = [];
    for (var i = 0; i < children.length; i++) {
      var linked = linkedIdentity_(children[i].cls, children[i].no);
      var hit = findExistingBookingFor_(slots, children[i].cls, children[i].no, linked);
      if (hit) {
        already.push({
          name: children[i].name,
          no: children[i].no,
          ownCls: children[i].cls,
          cls: String(hit.v[COL.CLASS - 1]),
          dateLabel: dateLabel_(hit.v[COL.DATE - 1]),
          start: String(hit.v[COL.START - 1])
        });
      }
    }

    return {
      children: children.map(function (c) {
        return { cls: c.cls, no: c.no, name: c.name, grade: c.grade, teacher: teacherOf_(c.cls) };
      }),
      already: already,
      days: already.length ? [] : siblingOptions_(children, slots, cfg)
    };
  });
}

/**
 * 選ばれた組み合わせで、全員分を一度に予約する。
 * 1人でも取れなければ、誰も予約しない。
 */
function apiSiblingBook(p) {
  return safe_(function () {
    var cfg = getConfig();
    var win = siblingBookingWindow_(cfg);
    if (!win.ok) throw new Error(win.message);

    var children = verifySiblings_(p, cfg);
    var slotIds = (p && p.slotIds) || [];
    if (slotIds.length !== children.length) {
      throw new Error('選ばれた時間の数が、お子さんの人数と合いません。画面を更新してやり直してください。');
    }

    var seen = {};
    for (var d = 0; d < slotIds.length; d++) {
      var id = String(slotIds[d]);
      if (seen[id]) throw new Error('同じ時間が2回選ばれています。選び直してください。');
      seen[id] = true;
    }

    var guardian = String((p && p.guardian) || '').trim();
    var note = String((p && p.note) || '').trim();

    var pending = [];

    var result = withLock_(function () {
      var sh = sheet_(SH.SLOTS);
      var last = sh.getLastRow();
      var slots = readSlots_();

      var byId = {};
      for (var i = 0; i < slots.length; i++) byId[String(slots[i].v[COL.SLOT_ID - 1])] = slots[i];

      var targets = [];
      for (var c = 0; c < children.length; c++) {
        var t = byId[String(slotIds[c])];
        if (!t) {
          throw new Error(children[c].name + ' さんの時間が見つかりませんでした。画面を更新してやり直してください。');
        }
        if (String(t.v[COL.CLASS - 1]) !== children[c].cls) {
          throw new Error(children[c].name + ' さんの時間が、在籍クラスと合いません。選び直してください。');
        }
        if (String(t.v[COL.STATUS - 1]) !== STATUS.OPEN) {
          throw new Error('申し訳ありません、' + children[c].name +
            ' さんの時間がちょうど埋まりました。もう一度お選びください。');
        }
        var linked = linkedIdentity_(children[c].cls, children[c].no);
        if (findExistingBookingFor_(slots, children[c].cls, children[c].no, linked)) {
          throw new Error(children[c].name + ' さんはすでに予約が入っています。' +
            '「確認・変更・取消」からお手続きください。');
        }
        targets.push(t);
      }

      // 全員分をまとめて書く（1人でも欠けないよう、確認が全部通ってから）
      var code = makeCode_();
      var now = new Date();
      var width = SLOT_LAST_COL - COL.STATUS + 1;
      var block = sh.getRange(2, COL.STATUS, last - 1, width).getValues();

      for (var w = 0; w < targets.length; w++) {
        block[targets[w].row - 2] = [
          STATUS.BOOKED, children[w].no, children[w].name, guardian, note, code, now
        ];
      }
      sh.getRange(2, COL.STATUS, last - 1, width).setValues(block);

      clearSlotCache_();
      markViewsStale_();
      logAction_('きょうだい予約', '', '', '', '',
        children.length + '名（' + children.map(function (x) { return x.cls + x.no + '番'; }).join('・') +
        '） / コード ' + code);

      var bookings = [];
      for (var b = 0; b < targets.length; b++) {
        var tv = targets[b].v;
        pending.push({ child: children[b], slot: targets[b] });
        bookings.push({
          name: children[b].name,
          no: children[b].no,
          cls: children[b].cls,
          teacher: teacherOf_(children[b].cls),
          ymd: ymd_(tv[COL.DATE - 1]),
          dateLabel: dateLabel_(tv[COL.DATE - 1]),
          start: String(tv[COL.START - 1]),
          end: String(tv[COL.END - 1])
        });
      }

      return { code: code, bookings: bookings };
    });

    // 最大6通のメールをロックの中で送ると、その間ほかの保護者が待たされる
    for (var n = 0; n < pending.length; n++) {
      notifyTeacherAfterLock_('予約', pending[n].child,
        { slot: pending[n].slot, guardian: guardian, note: note, code: result.code });
    }
    return result;
  });
}

/* ================================================================
   照合
   ================================================================ */

/**
 * きょうだい全員を名簿と照合する。
 * 誰の入力が合わなかったのかを必ず示す（4人分まとめて弾かれると原因が分からないため）。
 */
function verifySiblings_(p, cfg) {
  var list = (p && p.children) || [];
  if (list.length < 2) throw new Error('きょうだいまとめての予約は、2人以上でご利用ください。');
  if (list.length > SIBLING_MAX) {
    throw new Error('一度にお申し込みいただけるのは ' + SIBLING_MAX + ' 人までです。');
  }

  var guardKey = 'sib_' + String(list[0].cls || '') + '_' + Number(list[0].no || 0);
  guardBruteForce_(guardKey);

  var roster = [];
  try {
    roster = getRoster();
  } catch (err) {
    console.error('getRoster failed:', err);
    throw new Error('名簿データが確認できませんでした。担任までご連絡ください。');
  }

  var index = {};
  for (var r = 0; r < roster.length; r++) index[roster[r].cls + '_' + roster[r].no] = roster[r];

  var gradeOf = {};
  var classes = getClasses();
  for (var g = 0; g < classes.length; g++) gradeOf[classes[g].name] = classes[g].grade;

  var out = [], bad = [], seen = {};

  for (var i = 0; i < list.length; i++) {
    var label = (i + 1) + '人目';
    var cls = String(list[i].cls || '').trim();
    var no = Number(list[i].no || 0);
    var name = String(list[i].name || '').trim();

    if (!cls || !no) { bad.push(label + '：クラスと出席番号を入力してください'); continue; }

    var key = cls + '_' + no;
    if (seen[key]) { bad.push(label + '：同じお子さんが2回入力されています'); continue; }
    seen[key] = true;

    var found = index[key];
    if (!found) { bad.push(label + '（' + cls + ' ' + no + '番）：名簿と一致しません'); continue; }
    if (cfg.checkName && !name) { bad.push(label + '：お子さまの氏名を入力してください'); continue; }
    if (cfg.checkName && norm_(found.name) !== norm_(name)) {
      bad.push(label + '（' + cls + ' ' + no + '番）：氏名が名簿と一致しません'); continue;
    }

    out.push({ cls: found.cls, no: found.no, name: found.name, grade: gradeOf[found.cls] || '' });
  }

  if (bad.length) {
    // まとめて1回の失敗として数える（4人分入力すると一度に4回失敗してしまうため）
    countFailure_(guardKey, list[0].cls, list[0].no, 'きょうだい予約');
    throw new Error('次の入力が確認できませんでした。' + LF + LF + bad.join(LF) + LF + LF +
      'クラス・出席番号・氏名が、学校からのお知らせと合っているかご確認ください。' +
      '（姓と名の間のスペースは、有っても無くても構いません）');
  }

  // 学年の低い順に並べる（面談の順番の既定値）
  out.sort(function (a, b) { return gradeOrder_(a.grade) - gradeOrder_(b.grade); });
  return out;
}

/* ================================================================
   組み合わせの探索
   ================================================================ */

/**
 * 日ごとに、全員を同じ日の連続したコマに入れられる組み合わせを列挙する。
 * @return {Array<{key, dateLabel, options:Array, maxFit:number}>}
 */
function siblingOptions_(children, slots, cfg) {
  var times = daySlotTimes_(cfg);
  var n = children.length;

  // クラス×日×コマ → 空いている枠ID
  var free = {};
  var dayLabel = {}, dayKeys = [], seenDay = {};

  for (var i = 0; i < slots.length; i++) {
    var v = slots[i].v;
    var ymd = ymd_(v[COL.DATE - 1]);
    if (!seenDay[ymd]) { seenDay[ymd] = true; dayKeys.push(ymd); dayLabel[ymd] = dateLabel_(v[COL.DATE - 1]); }
    if (String(v[COL.STATUS - 1]) !== STATUS.OPEN) continue;

    var idx = slotIndexOf_(String(v[COL.SLOT_ID - 1]));
    if (!idx) continue;
    free[String(v[COL.CLASS - 1]) + '|' + ymd + '|' + idx] = {
      slotId: String(v[COL.SLOT_ID - 1]),
      start: String(v[COL.START - 1]),
      end: String(v[COL.END - 1])
    };
  }
  dayKeys.sort();

  var out = [];
  for (var d = 0; d < dayKeys.length; d++) {
    var ymdKey = dayKeys[d];
    var options = [];

    for (var start = 0; start + n <= times.length; start++) {
      var assign = matchConsecutive_(children, ymdKey, times, start, free);
      if (!assign) continue;
      options.push({
        startTime: assign[0].start,
        endTime: assign[assign.length - 1].end,
        picks: assign
      });
    }

    out.push({
      key: ymdKey,
      dateLabel: dayLabel[ymdKey],
      options: options,
      maxFit: options.length ? n : maxFitOnDay_(children, ymdKey, times, free)
    });
  }
  return out;
}

/** 枠IDの末尾のコマ番号 */
function slotIndexOf_(slotId) {
  var q = parseSlotId_(slotId);
  return q ? q.idx : 0;
}

/**
 * start から n コマ連続に、きょうだいを1人ずつ割り当てる。
 * まず学年順（＝children の並び順）で試し、だめなら入れ替えて探す。
 * @return {Array|null} [{name, cls, slotId, start, end}] を時間順で
 */
function matchConsecutive_(children, ymdKey, times, start, free) {
  var n = children.length;
  var wanted = [];
  for (var i = 0; i < n; i++) wanted.push(times[start + i]);

  var used = [];
  for (var u = 0; u < n; u++) used.push(false);
  var picked = [];

  function rec(ci) {
    if (ci === children.length) return true;
    for (var j = 0; j < n; j++) {
      if (used[j]) continue;
      var slot = free[children[ci].cls + '|' + ymdKey + '|' + wanted[j].index];
      if (!slot) continue;
      used[j] = true;
      picked.push({
        name: children[ci].name, cls: children[ci].cls, no: children[ci].no,
        teacher: teacherOf_(children[ci].cls),
        slotId: slot.slotId, start: slot.start, end: slot.end, order: j
      });
      if (rec(ci + 1)) return true;
      picked.pop();
      used[j] = false;
    }
    return false;
  }

  if (!rec(0)) return null;

  var result = picked.slice();
  result.sort(function (a, b) { return a.order - b.order; });
  return result;
}

/** その日に何人まで置けるか（連続でなくてよい） */
function maxFitOnDay_(children, ymdKey, times, free) {
  var usedIdx = {};
  var fit = 0;
  for (var c = 0; c < children.length; c++) {
    for (var t = 0; t < times.length; t++) {
      if (usedIdx[times[t].index]) continue;
      if (!free[children[c].cls + '|' + ymdKey + '|' + times[t].index]) continue;
      usedIdx[times[t].index] = true;
      fit++;
      break;
    }
  }
  return fit;
}
