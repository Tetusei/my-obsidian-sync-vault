/**
 * 担任・管理者向け — スプレッドシートのメニューと、Web管理画面の API。
 */

function onOpen() {
  buildMenu_();
}

function buildMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('三者面談')
    .addItem('① 初期セットアップ', 'menuSetup')
    .addItem('② 枠を再生成', 'menuGenerateSlots')
    .addItem('③ ダミー生徒10名を作成', 'menuGenerateDummyRoster')
    .addSeparator()
    .addItem('全体ビュー・クラス別予約表を更新', 'menuRefreshViews')
    .addItem('未予約の生徒を表示', 'menuUnbooked')
    .addSeparator()
    .addItem('選択した枠をブロック（面談を入れない）', 'menuBlock')
    .addItem('選択した枠のブロックを解除', 'menuUnblock')
    .addItem('選択した枠の予約を取り消す', 'menuCancel')
    .addSeparator()
    .addItem('予約受付を開始する', 'menuPublish')
    .addItem('予約受付を停止する', 'menuUnpublish')
    .addItem('保護者用URLを表示', 'menuShowUrl')
    .addToUi();
}

/**
 * 単体（非バインド）スクリプトの場合はこれを1度だけ手動実行して、
 * スプレッドシートを開いたときにメニューが出るようにする。
 */
function createOnOpenTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onOpen') return 'すでに設定済みです。';
  }
  ScriptApp.newTrigger('onOpen').forSpreadsheet(SPREADSHEET_ID).onOpen().create();
  return 'メニューを表示するトリガーを設定しました。スプレッドシートを開き直してください。';
}

/* ---------------- メニュー処理 ---------------- */

function menuSetup() {
  var ui = SpreadsheetApp.getUi();
  var created = setupSystem();
  ui.alert('初期セットアップ',
    (created.length ? '作成したシート: ' + created.join(', ') : '必要なシートはすべて揃っています。') +
    '\n\n次の順で入力してください。\n' +
    '1. 「クラス」シートに 4クラス分の担任名とメールを入力\n' +
    '2. 「生徒名簿」シートにクラス・出席番号・生徒氏名を貼り付け\n' +
    '3. 「面談日」「設定」を確認\n' +
    '4. メニューの「② 枠を再生成」を実行\n' +
    '5. Webアプリをデプロイし、「予約受付を開始する」',
    ui.ButtonSet.OK);
}

function menuGenerateDummyRoster() {
  var ui = SpreadsheetApp.getUi();
  try {
    var res = generateDummyRoster();
    var msg = res.count + ' 名のダミー生徒を「生徒名簿」シートに追加しました。';
    if (res.slotsCreated) {
      msg += '\n\n※「枠マスタ」シートが空だったため、面談枠も自動的に再生成しました！';
    } else {
      msg += '\n\n※面談枠がまだない場合は、メニューの「② 枠を再生成」を実行してください。';
    }
    ui.alert('ダミー生徒の作成', msg, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('エラー', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuGenerateSlots() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = generateSlots();
    rebuildOverview();
    rebuildClassSheets();
    ui.alert('枠を再生成しました',
      r.written + ' 枠を作成しました（既存の予約 ' + r.kept + ' 件を引き継ぎ）。', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('再生成できません', String(err.message || err), ui.ButtonSet.OK);
  }
}

function menuRefreshViews() {
  rebuildOverview();
  rebuildClassSheets();
  ss_().toast('全体ビューとクラス別予約表を更新しました。', '三者面談', 5);
}

function menuUnbooked() {
  var ui = SpreadsheetApp.getUi();
  var list = unbookedStudents(null);
  if (!list.length) {
    ui.alert('未予約の生徒', '全員の予約が入っています。', ui.ButtonSet.OK);
    return;
  }
  var byClass = {};
  list.forEach(function (s) {
    if (!byClass[s.cls]) byClass[s.cls] = [];
    byClass[s.cls].push(s.no + '. ' + s.name);
  });
  var lines = Object.keys(byClass).sort().map(function (c) {
    return '【' + c + '】(' + byClass[c].length + '名)\n' + byClass[c].join('、');
  });
  ui.alert('未予約の生徒（計 ' + list.length + '名）', lines.join('\n\n'), ui.ButtonSet.OK);
}

function menuBlock() { setStatusForSelection_(STATUS.BLOCKED); }
function menuUnblock() { setStatusForSelection_(STATUS.OPEN); }

function setStatusForSelection_(newStatus) {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SH.SLOTS) {
    ui.alert('「' + SH.SLOTS + '」シートで、対象の行を選んでから実行してください。');
    return;
  }
  var rng = SpreadsheetApp.getActiveRange();
  var start = Math.max(rng.getRow(), 2);
  var end = rng.getRow() + rng.getNumRows() - 1;
  if (end < 2) { ui.alert('見出し行以外の行を選んでください。'); return; }

  var changed = 0, skipped = 0;
  for (var r = start; r <= end; r++) {
    var status = String(sh.getRange(r, COL.STATUS).getValue());
    if (status === STATUS.BOOKED) { skipped++; continue; }
    if (status === newStatus) continue;
    sh.getRange(r, COL.STATUS).setValue(newStatus);
    changed++;
  }
  clearSlotCache_();
  rebuildOverview();
  rebuildClassSheets();
  ui.alert(changed + ' 件を「' + newStatus + '」にしました。' +
    (skipped ? '\n予約が入っている ' + skipped + ' 件は変更していません（先に取り消してください）。' : ''));
}

function menuCancel() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SH.SLOTS) {
    ui.alert('「' + SH.SLOTS + '」シートで、取り消したい予約の行を選んでから実行してください。');
    return;
  }
  var rng = SpreadsheetApp.getActiveRange();
  var start = Math.max(rng.getRow(), 2);
  var end = rng.getRow() + rng.getNumRows() - 1;

  var targets = [];
  for (var r = start; r <= end; r++) {
    var v = sh.getRange(r, 1, 1, SLOT_LAST_COL).getValues()[0];
    if (String(v[COL.STATUS - 1]) === STATUS.BOOKED) targets.push({ row: r, v: v });
  }
  if (!targets.length) { ui.alert('選んだ範囲に予約はありません。'); return; }

  var names = targets.map(function (t) {
    return dateLabel_(t.v[COL.DATE - 1]) + ' ' + t.v[COL.START - 1] + '　' +
      t.v[COL.CLASS - 1] + ' ' + t.v[COL.NUMBER - 1] + '. ' + t.v[COL.STUDENT - 1];
  }).join('\n');
  var res = ui.alert('次の予約を取り消します。よろしいですか？', names, ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;

  targets.forEach(function (t) {
    clearSlotRow_(t.row);
    logAction_('取消', String(t.v[COL.SLOT_ID - 1]), String(t.v[COL.CLASS - 1]),
      t.v[COL.NUMBER - 1], String(t.v[COL.STUDENT - 1]), '担任による取消');
  });
  clearSlotCache_();
  rebuildOverview();
  rebuildClassSheets();
  ui.alert(targets.length + ' 件の予約を取り消しました。');
}

function menuPublish() { setPublished_(true); }
function menuUnpublish() { setPublished_(false); }

function setPublished_(flag) {
  var sh = sheet_(SH.CONFIG);
  var vals = sh.getDataRange().getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === '公開') {
      sh.getRange(i + 1, 2).setValue(flag);
      SpreadsheetApp.getUi().alert('予約受付を' + (flag ? '開始' : '停止') + 'しました。');
      return;
    }
  }
  SpreadsheetApp.getUi().alert('「設定」シートに「公開」の行がありません。初期セットアップを実行してください。');
}

function menuShowUrl() {
  var ui = SpreadsheetApp.getUi();
  var url;
  try {
    url = ScriptApp.getService().getUrl();
  } catch (err) {
    url = '';
  }
  if (!url) {
    ui.alert('URLがまだありません', 'Apps Script エディタの「デプロイ ▸ 新しいデプロイ ▸ ウェブアプリ」で\n' +
      '　次のユーザーとして実行: 自分\n　アクセスできるユーザー: 全員\n' +
      'を選んでデプロイしてください。', ui.ButtonSet.OK);
    return;
  }
  ui.alert('保護者用URL', url + '\n\n担任用の管理画面:\n' + url + '?page=admin', ui.ButtonSet.OK);
}

/* ---------------- Web管理画面 API ---------------- */

function requireAdmin_(pass) {
  var cfg = getConfig();
  if (!cfg.adminPasscode) throw new Error('管理パスコードが未設定です。「設定」シートの「管理パスコード」を入力してください。');
  if (String(pass || '') !== cfg.adminPasscode) {
    Utilities.sleep(800);
    throw new Error('パスコードが違います。');
  }
  return cfg;
}

function apiAdminInit(pass) {
  return safe_(function () {
    var cfg = requireAdmin_(pass);
    return {
      title: cfg.title,
      published: cfg.published,
      classes: getClasses(),
      days: getDays().map(function (d) { return { key: ymd_(d), label: dateLabel_(d) }; })
    };
  });
}

/** 指定日の全クラス分の枠（担任向けなので氏名込み） */
function apiAdminBoard(pass, dayKey) {
  return safe_(function () {
    requireAdmin_(pass);
    var classes = getClasses().map(function (c) { return c.name; });
    var slots = readSlots_();
    var rows = {};
    var order = [];

    for (var i = 0; i < slots.length; i++) {
      var v = slots[i].v;
      if (dayKey && ymd_(v[COL.DATE - 1]) !== dayKey) continue;
      var time = String(v[COL.START - 1]) + '–' + String(v[COL.END - 1]);
      var key = ymd_(v[COL.DATE - 1]) + ' ' + time;
      if (!rows[key]) {
        rows[key] = { dateLabel: dateLabel_(v[COL.DATE - 1]), time: time, cells: {} };
        order.push(key);
      }
      rows[key].cells[String(v[COL.CLASS - 1])] = {
        slotId: String(v[COL.SLOT_ID - 1]),
        status: String(v[COL.STATUS - 1]),
        no: v[COL.NUMBER - 1],
        name: String(v[COL.STUDENT - 1] || ''),
        guardian: String(v[COL.GUARDIAN - 1] || ''),
        note: String(v[COL.NOTE - 1] || '')
      };
    }
    order.sort();
    return { classes: classes, rows: order.map(function (k) { return rows[k]; }) };
  });
}

function apiAdminSetStatus(pass, slotId, status) {
  return safe_(function () {
    requireAdmin_(pass);
    if (status !== STATUS.OPEN && status !== STATUS.BLOCKED) throw new Error('指定できない状態です。');
    return withLock_(function () {
      var found = findSlotRow_(slotId);
      if (String(found.v[COL.STATUS - 1]) === STATUS.BOOKED) {
        throw new Error('予約が入っています。先に取り消してください。');
      }
      sheet_(SH.SLOTS).getRange(found.row, COL.STATUS).setValue(status);
      clearSlotCache_();
      logAction_(status === STATUS.BLOCKED ? 'ブロック' : 'ブロック解除', slotId, String(found.v[COL.CLASS - 1]), '', '', '担任Web');
      return {};
    });
  });
}

function apiAdminCancel(pass, slotId) {
  return safe_(function () {
    requireAdmin_(pass);
    return withLock_(function () {
      var found = findSlotRow_(slotId);
      if (String(found.v[COL.STATUS - 1]) !== STATUS.BOOKED) throw new Error('その枠に予約はありません。');
      clearSlotRow_(found.row);
      clearSlotCache_();
      logAction_('取消', slotId, String(found.v[COL.CLASS - 1]), found.v[COL.NUMBER - 1],
        String(found.v[COL.STUDENT - 1]), '担任Webによる取消');
      return {};
    });
  });
}

function apiAdminUnbooked(pass, cls) {
  return safe_(function () {
    requireAdmin_(pass);
    return { students: unbookedStudents(cls || null) };
  });
}

function apiAdminRefreshViews(pass) {
  return safe_(function () {
    requireAdmin_(pass);
    rebuildOverview();
    rebuildClassSheets();
    return {};
  });
}

function findSlotRow_(slotId) {
  var slots = readSlots_();
  for (var i = 0; i < slots.length; i++) {
    if (String(slots[i].v[COL.SLOT_ID - 1]) === String(slotId)) return slots[i];
  }
  throw new Error('枠が見つかりません: ' + slotId);
}
