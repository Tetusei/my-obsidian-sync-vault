/**
 * Web アプリ本体と、保護者向けの予約 API。
 *
 * 匿名アクセスのため、生徒名簿は一切クライアントに送らない。
 * 保護者は「クラス＋出席番号＋生徒氏名」を入力し、サーバ側で名簿と照合する。
 * 空き枠一覧にも予約者名は含めない（空いているかどうかだけ）。
 */

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || '';
  var file = page === 'admin' ? 'admin' : 'index';
  var t = HtmlService.createTemplateFromFile(file);
  var out = t.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  try {
    out.setTitle(getConfig().title);
  } catch (err) {
    out.setTitle('三者面談 予約');
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* 保護者向け API                                                       */
/* ------------------------------------------------------------------ */

function apiInit() {
  return safe_(function () {
    var cfg = getConfig();
    var win = bookingWindow_(cfg);
    return {
      title: cfg.title,
      notice: cfg.notice,
      classes: getClasses().map(function (c) { return c.name; }),
      open: win.ok,
      message: win.message,
      checkName: cfg.checkName
    };
  });
}

/** クラスの空き枠一覧＋この生徒の既存予約 */
function apiSlots(p) {
  return safe_(function () {
    var cfg = getConfig();
    // 受付期間外は名簿の照合自体を行わない（名簿を探る手がかりを残さない）。
    // 予約済みの内容は「確認・変更・取消」から予約コードで確認できる。
    var win = bookingWindow_(cfg);
    if (!win.ok) throw new Error(win.message);
    var student = verifyStudent_(p, cfg);
    var slots = readSlotsCached_();
    var days = {};
    var order = [];
    var existing = null;

    for (var i = 0; i < slots.length; i++) {
      var v = slots[i].v;
      if (String(v[COL.CLASS - 1]) !== student.cls) continue;
      var st = String(v[COL.STATUS - 1]);
      var mine = st === STATUS.BOOKED && Number(v[COL.NUMBER - 1]) === student.no;
      if (mine) {
        existing = {
          slotId: String(v[COL.SLOT_ID - 1]),
          dateLabel: dateLabel_(v[COL.DATE - 1]),
          start: String(v[COL.START - 1]),
          end: String(v[COL.END - 1])
        };
      }
      var key = ymd_(v[COL.DATE - 1]);
      if (!days[key]) {
        days[key] = { dateLabel: dateLabel_(v[COL.DATE - 1]), slots: [] };
        order.push(key);
      }
      days[key].slots.push({
        slotId: String(v[COL.SLOT_ID - 1]),
        start: String(v[COL.START - 1]),
        end: String(v[COL.END - 1]),
        available: st === STATUS.OPEN,
        mine: mine
      });
    }

    return {
      student: { cls: student.cls, no: student.no, name: student.name },
      teacher: teacherOf_(student.cls),
      existing: existing,
      days: order.map(function (k) { return days[k]; })
    };
  });
}

function apiBook(p) {
  return safe_(function () {
    var cfg = getConfig();
    var win = bookingWindow_(cfg);
    if (!win.ok) throw new Error(win.message);
    var student = verifyStudent_(p, cfg);
    var slotId = String(p.slotId || '').trim();
    if (!slotId) throw new Error('時間が選ばれていません。');

    return withLock_(function () {
      var sh = sheet_(SH.SLOTS);
      var slots = readSlots_();
      var target = null;
      var mineCount = 0;

      for (var i = 0; i < slots.length; i++) {
        var v = slots[i].v;
        if (String(v[COL.SLOT_ID - 1]) === slotId) target = slots[i];
        if (String(v[COL.CLASS - 1]) === student.cls &&
          String(v[COL.STATUS - 1]) === STATUS.BOOKED &&
          Number(v[COL.NUMBER - 1]) === student.no) mineCount++;
      }

      if (!target) throw new Error('選ばれた時間が見つかりませんでした。画面を更新してやり直してください。');
      if (String(target.v[COL.CLASS - 1]) !== student.cls) throw new Error('他のクラスの時間は予約できません。');
      if (mineCount >= cfg.maxPerStudent) {
        throw new Error('すでに予約が入っています。時間を変えるときは「確認・変更・取消」から手続きしてください。');
      }
      if (String(target.v[COL.STATUS - 1]) !== STATUS.OPEN) {
        throw new Error('申し訳ありません、その時間はちょうど埋まりました。別の時間を選んでください。');
      }

      var code = makeCode_();
      sh.getRange(target.row, COL.STATUS, 1, SLOT_LAST_COL - COL.STATUS + 1).setValues([[
        STATUS.BOOKED, student.no, student.name,
        String(p.guardian || '').trim(), String(p.note || '').trim(),
        code, new Date()
      ]]);
      clearSlotCache_();
      rebuildOverview();
      rebuildClassSheets();
      logAction_('予約', slotId, student.cls, student.no, student.name, '');

      return {
        code: code,
        booking: {
          dateLabel: dateLabel_(target.v[COL.DATE - 1]),
          start: String(target.v[COL.START - 1]),
          end: String(target.v[COL.END - 1]),
          cls: student.cls,
          teacher: teacherOf_(student.cls),
          name: student.name
        }
      };
    });
  });
}

function apiLookup(p) {
  return safe_(function () {
    var cfg = getConfig();
    var student = verifyStudent_(p, cfg);
    var found = findBookingByCode_(student, p.code);
    return { booking: bookingView_(found.v, student) };
  });
}

function apiCancel(p) {
  return safe_(function () {
    var cfg = getConfig();
    var win = bookingWindow_(cfg);
    if (!win.ok) throw new Error(win.message);
    var student = verifyStudent_(p, cfg);

    return withLock_(function () {
      var found = findBookingByCode_(student, p.code);
      clearSlotRow_(found.row);
      clearSlotCache_();
      rebuildOverview();
      rebuildClassSheets();
      logAction_('取消', String(found.v[COL.SLOT_ID - 1]), student.cls, student.no, student.name, '保護者による取消');
      return { cancelled: true };
    });
  });
}

function apiChange(p) {
  return safe_(function () {
    var cfg = getConfig();
    var win = bookingWindow_(cfg);
    if (!win.ok) throw new Error(win.message);
    var student = verifyStudent_(p, cfg);
    var newId = String(p.slotId || '').trim();
    if (!newId) throw new Error('変更先の時間が選ばれていません。');

    return withLock_(function () {
      var sh = sheet_(SH.SLOTS);
      var slots = readSlots_();
      var old = null, next = null;
      for (var i = 0; i < slots.length; i++) {
        var v = slots[i].v;
        if (String(v[COL.SLOT_ID - 1]) === newId) next = slots[i];
        if (String(v[COL.STATUS - 1]) === STATUS.BOOKED &&
          String(v[COL.CLASS - 1]) === student.cls &&
          Number(v[COL.NUMBER - 1]) === student.no &&
          String(v[COL.CODE - 1]) === String(p.code || '').trim()) old = slots[i];
      }
      if (!old) throw new Error('予約が見つかりませんでした。予約コードをご確認ください。');
      if (!next) throw new Error('変更先の時間が見つかりませんでした。画面を更新してやり直してください。');
      if (String(next.v[COL.CLASS - 1]) !== student.cls) throw new Error('他のクラスの時間には変更できません。');
      if (String(next.v[COL.SLOT_ID - 1]) === String(old.v[COL.SLOT_ID - 1])) {
        return { booking: bookingView_(old.v, student) };
      }
      if (String(next.v[COL.STATUS - 1]) !== STATUS.OPEN) {
        throw new Error('申し訳ありません、その時間はちょうど埋まりました。別の時間を選んでください。');
      }

      var payload = [[
        STATUS.BOOKED, student.no, student.name,
        String(old.v[COL.GUARDIAN - 1] || ''), String(old.v[COL.NOTE - 1] || ''),
        String(old.v[COL.CODE - 1]), new Date()
      ]];
      sh.getRange(next.row, COL.STATUS, 1, SLOT_LAST_COL - COL.STATUS + 1).setValues(payload);
      clearSlotRow_(old.row);
      clearSlotCache_();
      rebuildOverview();
      rebuildClassSheets();
      logAction_('変更', newId, student.cls, student.no, student.name,
        String(old.v[COL.SLOT_ID - 1]) + ' → ' + newId);

      return { booking: bookingView_(next.v, student) };
    });
  });
}

/* ------------------------------------------------------------------ */
/* 内部処理                                                            */
/* ------------------------------------------------------------------ */

/** 例外を {ok:false, error:...} に包む。google.script.run の失敗ハンドラを使わずに済ませる。 */
function safe_(fn) {
  try {
    var data = fn();
    data = data || {};
    data.ok = true;
    return data;
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('ただいま混み合っています。少し待ってからもう一度お試しください。');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/** 名簿と照合する。合わなければ例外。保護者へは親切な案内文を出す。 */
function verifyStudent_(p, cfg) {
  var cls = String((p && p.cls) || '').trim();
  var no = Number((p && p.no) || 0);
  var name = String((p && p.name) || '').trim();
  if (!cls) throw new Error('クラスを選んでください。');
  if (!no) throw new Error('出席番号を入力してください。');
  if (cfg.checkName && !name) throw new Error('生徒の氏名を入力してください。');

  guardBruteForce_(cls + '_' + no);

  var roster = [];
  try {
    roster = getRoster();
  } catch (err) {
    console.error('getRoster failed:', err);
    throw new Error('名簿データが確認できませんでした。入力内容をご確認いただくか、担任までご連絡ください。');
  }

  var foundNo = false;
  for (var i = 0; i < roster.length; i++) {
    if (roster[i].cls !== cls || roster[i].no !== no) continue;
    foundNo = true;
    if (cfg.checkName && norm_(roster[i].name) !== norm_(name)) {
      countFailure_(cls + '_' + no);
      throw new Error('出席番号と氏名が名簿と一致しません。姓と名の間のスペースは無くても構いません。お子様に再度、出席番号をご確認ください。そのあと、ご不明な場合は担任までご連絡ください。');
    }
    return roster[i];
  }
  countFailure_(cls + '_' + no);
  if (foundNo) {
    throw new Error('出席番号と氏名が名簿と一致しません。姓と名の間のスペースは無くても構いません。お子様に再度、出席番号をご確認ください。そのあと、ご不明な場合は担任までご連絡ください。');
  }
  throw new Error('名簿に見つかりませんでした。クラスと出席番号をご確認ください。お子様に再度、出席番号をご確認いただくか、担任までご連絡ください。');
}

function findBookingByCode_(student, code) {
  var c = String(code || '').trim();
  if (!c) throw new Error('予約コードを入力してください。');
  guardBruteForce_('code_' + student.cls + '_' + student.no);

  var slots = readSlots_();
  for (var i = 0; i < slots.length; i++) {
    var v = slots[i].v;
    if (String(v[COL.STATUS - 1]) !== STATUS.BOOKED) continue;
    if (String(v[COL.CLASS - 1]) !== student.cls) continue;
    if (Number(v[COL.NUMBER - 1]) !== student.no) continue;
    if (String(v[COL.CODE - 1]) !== c) continue;
    return slots[i];
  }
  countFailure_('code_' + student.cls + '_' + student.no);
  throw new Error('予約が見つかりませんでした。予約コードをご確認ください。分からない場合は担任までご連絡ください。');
}

function bookingView_(v, student) {
  return {
    slotId: String(v[COL.SLOT_ID - 1]),
    dateLabel: dateLabel_(v[COL.DATE - 1]),
    start: String(v[COL.START - 1]),
    end: String(v[COL.END - 1]),
    cls: student.cls,
    teacher: teacherOf_(student.cls),
    name: student.name,
    guardian: String(v[COL.GUARDIAN - 1] || ''),
    note: String(v[COL.NOTE - 1] || '')
  };
}

function clearSlotRow_(row) {
  sheet_(SH.SLOTS).getRange(row, COL.STATUS, 1, SLOT_LAST_COL - COL.STATUS + 1)
    .setValues([[STATUS.OPEN, '', '', '', '', '', '']]);
}

function teacherOf_(cls) {
  var list = getClasses();
  for (var i = 0; i < list.length; i++) if (list[i].name === cls) return list[i].teacher;
  return '';
}

function makeCode_() {
  var n = Math.floor(Math.random() * 10000);
  return ('000' + n).slice(-4);
}

var MAX_FAILURES = 8;
var FAILURE_WINDOW_SEC = 900;

function guardBruteForce_(key) {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get('fail_' + key) || 0);
  if (n >= MAX_FAILURES) {
    throw new Error('入力の誤りが続いたため、しばらく受け付けを停止しています。15分ほど時間をおくか、担任までご連絡ください。');
  }
}

function countFailure_(key) {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get('fail_' + key) || 0) + 1;
  cache.put('fail_' + key, String(n), FAILURE_WINDOW_SEC);
}
