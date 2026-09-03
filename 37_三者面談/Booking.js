/**
 * Web アプリ本体と、保護者向けの予約 API。
 *
 * 匿名アクセスのため、生徒名簿は一切クライアントに送らない。
 * 保護者は「クラス＋出席番号＋生徒氏名」を入力し、サーバ側で名簿と照合する。
 * 空き枠一覧にも予約者名は含めない（空いているかどうかだけ）。
 */

function doGet(e) {
  // 案内プリントを作る前の疎通確認用。
  // HTMLが表示できるだけでなく、Webアプリから設定シートまで読めることを確かめる。
  if (e && e.parameter && e.parameter.health === '1') {
    var health;
    try {
      var cfg = getConfig();
      health = { ok: true, version: VERSION, title: cfg.title };
    } catch (err) {
      health = { ok: false, version: VERSION, error: String(err && err.message || err) };
    }
    return ContentService.createTextOutput(JSON.stringify(health))
      .setMimeType(ContentService.MimeType.JSON);
  }

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
    var win = bookingWindow_(cfg);
    if (!win.ok) throw new Error(win.message);
    var student = verifyStudent_(p, cfg);
    var slots = readSlotsCached_();
    var linked = linkedIdentity_(student.cls, student.no);
    var days = {};
    var order = [];
    var existing = null;
    var linkedBooking = null;

    for (var i = 0; i < slots.length; i++) {
      var v = slots[i].v;
      if (String(v[COL.CLASS - 1]) !== student.cls) continue;
      var st = String(v[COL.STATUS - 1]);

      // 交流学級のほうで予約している場合も拾う（別クラスなので上の絞り込みの外）
      var mine = (st === STATUS.BOOKED || st === STATUS.RESERVE) &&
        Number(v[COL.NUMBER - 1]) === student.no;
      if (mine) {
        existing = {
          slotId: String(v[COL.SLOT_ID - 1]),
          dateLabel: dateLabel_(v[COL.DATE - 1]),
          start: String(v[COL.START - 1]),
          end: String(v[COL.END - 1])
        };
      }

      // 予備は担任が手で使う枠。保護者の一覧には出さない（存在に気づかせない）
      if (st === STATUS.RESERVE) continue;

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
        blocked: st === STATUS.BLOCKED,
        mine: mine
      });
    }

    if (linked) {
      for (var j = 0; j < slots.length; j++) {
        var lv = slots[j].v;
        if (!isTakenSlot_(lv)) continue;
        if (String(lv[COL.CLASS - 1]) !== linked.cls) continue;
        if (Number(lv[COL.NUMBER - 1]) !== linked.no) continue;
        linkedBooking = {
          cls: linked.cls,
          teacher: teacherOf_(linked.cls),
          dateLabel: dateLabel_(lv[COL.DATE - 1]),
          start: String(lv[COL.START - 1]),
          end: String(lv[COL.END - 1])
        };
        break;
      }
    }

    return {
      student: { cls: student.cls, no: student.no, name: student.name },
      teacher: teacherOf_(student.cls),
      existing: existing,
      linked: linked ? { cls: linked.cls, teacher: teacherOf_(linked.cls) } : null,
      linkedBooking: linkedBooking,
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

    // メール送信は数秒かかることがある。ロックの中で送ると、その間ほかの保護者が待たされる。
    // ロックを離してから送るため、必要な材料だけ持ち出す。
    // 交流学級の紐づけは参照データなので、ロックを取る前に読んでおく
    var linked = linkedIdentity_(student.cls, student.no);
    var pending = null;

    var result = withLock_(function () {
      var sh = sheet_(SH.SLOTS);
      var slots = readSlots_();
      var target = null;
      var mineCount = 0;

      for (var i = 0; i < slots.length; i++) {
        var v = slots[i].v;
        if (String(v[COL.SLOT_ID - 1]) === slotId) target = slots[i];
        var vst = String(v[COL.STATUS - 1]);
        if (String(v[COL.CLASS - 1]) === student.cls &&
          (vst === STATUS.BOOKED || vst === STATUS.RESERVE) &&
          Number(v[COL.NUMBER - 1]) === student.no) mineCount++;
      }

      if (!target) throw new Error('選ばれた時間が見つかりませんでした。画面を更新してやり直してください。');
      if (String(target.v[COL.CLASS - 1]) !== student.cls) throw new Error('他のクラスの時間は予約できません。');
      if (mineCount >= cfg.maxPerStudent) {
        throw new Error('すでに予約が入っています。時間を変えるときは「確認・変更・取消」から手続きしてください。');
      }

      // 交流学級のほうで予約済みなら、二重にはできない
      if (linked) {
        var already = findExistingBookingFor_(slots, linked.cls, linked.no, null);
        if (already) {
          throw new Error('すでに ' + linked.cls + ' の担任と面談を予約されています（' +
            dateLabel_(already.v[COL.DATE - 1]) + ' ' + already.v[COL.START - 1] + '）。' +
            '変更するときは「確認・変更・取消」から、' + linked.cls + ' を選んで手続きしてください。');
        }
      }
      if (String(target.v[COL.STATUS - 1]) !== STATUS.OPEN) {
        throw new Error('申し訳ありません、その時間はちょうど埋まりました。別の時間を選んでください。');
      }

      var code = makeCode_();
      var guardian = String(p.guardian || '').trim();
      var note = String(p.note || '').trim();

      sh.getRange(target.row, COL.STATUS, 1, SLOT_LAST_COL - COL.STATUS + 1).setValues([[
        STATUS.BOOKED, student.no, student.name,
        guardian, note, code, new Date()
      ]]);
      clearSlotCache_();
      markViewsStale_();
      logAction_('予約', slotId, student.cls, student.no, student.name, 'コード ' + code);
      pending = { slot: target, guardian: guardian, note: note, code: code };

      return {
        code: code,
        booking: {
          ymd: ymd_(target.v[COL.DATE - 1]),
          dateLabel: dateLabel_(target.v[COL.DATE - 1]),
          start: String(target.v[COL.START - 1]),
          end: String(target.v[COL.END - 1]),
          cls: student.cls,
          teacher: teacherOf_(student.cls),
          name: student.name
        }
      };
    });

    notifyTeacherAfterLock_('予約', student, pending);
    return result;
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

    var pending = null;

    var result = withLock_(function () {
      var found = findBookingByCode_(student, p.code);
      clearSlotRow_(found.row, String(found.v[COL.SLOT_ID - 1]));
      clearSlotCache_();
      markViewsStale_();
      // 取り消すと枠マスタの行は空になるので、あとから追えるようコードをログに残す
      logAction_('取消', String(found.v[COL.SLOT_ID - 1]), student.cls, student.no, student.name,
        '保護者による取消 / コード ' + String(found.v[COL.CODE - 1] || ''));
      pending = { slot: found, guardian: '', note: '', code: String(p.code || '') };

      return { cancelled: true };
    });

    notifyTeacherAfterLock_('取消', student, pending);
    return result;
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

    var pending = null;

    var result = withLock_(function () {
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

      var guardian = String(old.v[COL.GUARDIAN - 1] || '');
      var note = String(old.v[COL.NOTE - 1] || '');
      var code = String(old.v[COL.CODE - 1]);

      var payload = [[
        STATUS.BOOKED, student.no, student.name,
        guardian, note, code, new Date()
      ]];
      sh.getRange(next.row, COL.STATUS, 1, SLOT_LAST_COL - COL.STATUS + 1).setValues(payload);
      clearSlotRow_(old.row, String(old.v[COL.SLOT_ID - 1]));
      clearSlotCache_();
      markViewsStale_();
      logAction_('変更', newId, student.cls, student.no, student.name,
        String(old.v[COL.SLOT_ID - 1]) + ' → ' + newId + ' / コード ' + code);
      pending = { slot: next, guardian: guardian, note: note, code: code };

      // next.v は書き込む前に読んだ内容なので、予約欄はまだ空のまま。
      // 引き継いだ保護者氏名と連絡事項が確認画面から消えないよう、書いた内容を反映させる
      var updated = next.v.slice();
      updated[COL.STATUS - 1] = STATUS.BOOKED;
      updated[COL.NUMBER - 1] = student.no;
      updated[COL.STUDENT - 1] = student.name;
      updated[COL.GUARDIAN - 1] = guardian;
      updated[COL.NOTE - 1] = note;
      updated[COL.CODE - 1] = code;

      return { booking: bookingView_(updated, student) };
    });

    notifyTeacherAfterLock_('変更', student, pending);
    return result;
  });
}

/* ---------------- 担任メール通知 ---------------- */

/**
 * ロックを離してから担任へ通知する。
 * 通知に失敗しても予約そのものは成立しているので、保護者にはエラーを見せない。
 */
function notifyTeacherAfterLock_(action, student, pending) {
  if (!pending) return;
  try {
    sendTeacherNotification_(action, student, pending.slot, pending.guardian, pending.note, pending.code);
  } catch (e) {
    console.warn('担任通知をスキップ:', e);
  }
}

function sendTeacherNotification_(action, student, targetSlot, guardian, note, code) {
  var cfg = getConfig();
  if (!cfg.notifyTeacher) return;
  var classes = getClasses();
  var teacherEmail = '';
  var teacherName = '';
  for (var i = 0; i < classes.length; i++) {
    if (classes[i].name === student.cls) {
      teacherEmail = classes[i].email;
      teacherName = classes[i].teacher;
    }
  }
  if (!teacherEmail) return;

  var dateStr = targetSlot.v ? dateLabel_(targetSlot.v[COL.DATE - 1]) + ' ' + targetSlot.v[COL.START - 1] + '〜' + targetSlot.v[COL.END - 1] : '';
  var subject = '【三者面談通知】' + student.cls + ' ' + student.no + '番 ' + student.name + ' 様の' + action;
  var body = (teacherName ? teacherName + ' 先生\n\n' : '') +
    '三者面談の' + action + 'がありましたのでお知らせいたします。\n\n' +
    '----------------------------------------\n' +
    '■ 対象生徒: ' + student.cls + ' ' + student.no + '番 ' + student.name + '\n' +
    (dateStr ? '■ 面談日時: ' + dateStr + '\n' : '') +
    (guardian ? '■ 保護者名: ' + guardian + '\n' : '') +
    (note ? '■ 連絡事項: ' + note + '\n' : '') +
    (code ? '■ 予約コード: ' + code + '\n' : '') +
    '----------------------------------------\n\n' +
    '※スプレッドシートの「予約表_' + student.cls + '」シートにも自動反映されています。';

  try {
    MailApp.sendEmail(teacherEmail, subject, body);
  } catch (err) {
    console.warn('担任メール通知スキップ:', err);
  }
}

/* ---------------- 内部処理 ---------------- */

/** 例外を {ok:false, error:...} に包む */
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
  // 受付開始の直後は予約が集中する。順番待ちの余裕を少し長めに取る
  if (!lock.tryLock(30000)) {
    throw new Error('ただいま混み合っています。少し待ってからもう一度お試しください。');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/** 名簿と照合する */
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
      countFailure_(cls + '_' + no, cls, no, '生徒氏名');
      throw new Error('入力情報が名簿と一致しませんでした。お子様に出席番号をご確認の上、もう一度お試しください。解消しない場合は担任までご連絡ください。（※姓と名の間のスペースは無くても構いません）');
    }
    return roster[i];
  }
  countFailure_(cls + '_' + no, cls, no, foundNo ? '生徒氏名' : '出席番号');
  if (foundNo) {
    throw new Error('入力情報が名簿と一致しませんでした。お子様に出席番号をご確認の上、もう一度お試しください。解消しない場合は担任までご連絡ください。（※姓と名の間のスペースは無くても構いません）');
  }
  throw new Error('入力情報が名簿と一致しませんでした。お子様に出席番号をご確認の上、もう一度お試しください。解消しない場合は担任までご連絡ください。（※姓と名の間のスペースは無くても構いません）');
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
  countFailure_('code_' + student.cls + '_' + student.no, student.cls, student.no, '予約コード');
  throw new Error('予約が見つかりませんでした。予約コードをご確認ください。分からない場合は担任までご連絡ください。');
}

function bookingView_(v, student) {
  return {
    slotId: String(v[COL.SLOT_ID - 1]),
    ymd: ymd_(v[COL.DATE - 1]),
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

/**
 * 予約内容を消して枠を空ける。
 * その枠が「だめなコマ」に指定されている場合は、空きではなくブロックに戻す
 * （担任が面談を入れないと決めた枠が、取消をきっかけに再び予約されるのを防ぐ）。
 */
function clearSlotRow_(row, slotId) {
  var status = STATUS.OPEN;
  if (slotId) {
    try {
      if (readNgSet_()[String(slotId)]) status = STATUS.BLOCKED;
    } catch (e) {
      console.warn('だめなコマの確認をスキップ:', e);
    }
  }
  sheet_(SH.SLOTS).getRange(row, COL.STATUS, 1, SLOT_LAST_COL - COL.STATUS + 1)
    .setValues([[status, '', '', '', '', '', '']]);
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
var FAILURE_NOTICE_AT = 3;   // 何回続いたら学校側に記録するか

function guardBruteForce_(key) {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get('fail_' + key) || 0);
  if (n >= MAX_FAILURES) {
    throw new Error('入力の誤りが続いたため、しばらく受け付けを停止しています。15分ほど時間をおくか、担任までご連絡ください。');
  }
}

/**
 * 入力の誤りを数える。
 *
 * 保護者が予約できずにいても、これまで学校側には何も伝わらなかった。
 * 誤りが続いた時点で予約ログに残し、担任から声をかけられるようにする。
 */
function countFailure_(key, cls, no, kind) {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get('fail_' + key) || 0) + 1;
  cache.put('fail_' + key, String(n), FAILURE_WINDOW_SEC);

  if (n === FAILURE_NOTICE_AT || n === MAX_FAILURES) {
    try {
      logAction_('入力エラー', '', cls || '', no || '', '',
        n + '回続けて名簿と一致しませんでした（' + (kind || '') + '）' +
        (n >= MAX_FAILURES ? ' ／ 15分間の受付停止中' : ''));
    } catch (e) {
      console.warn('入力エラーの記録をスキップ:', e);
    }
  }
}
