/**
 * データの点検（健康診断）。
 *
 * 予約が始まってから発覚すると困る入力ミスを、事前にまとめて洗い出す。
 * 特に「保護者だけが困っていて学校が気づけない」種類の問題を重視している。
 * 例）出席番号が重複していると、片方の家庭はいつまでも予約できないが、
 *     学校側には何も表示されない。
 */

/** 深刻度 */
var CHECK_LEVEL = {
  ERROR: 'error',   // このままでは予約が正しく動かない
  WARN: 'warn'      // 動くが、確認したほうがよい
};

/**
 * すべての点検を実行する。
 * @return {{errors:Array<Object>, warns:Array<Object>}}
 */
function checkData() {
  var found = [];

  checkClasses_(found);
  checkDays_(found);
  checkConfigValues_(found);
  checkRoster_(found);
  checkSlotState_(found);
  checkOrphanSheets_(found);
  checkInputErrors_(found);
  checkAutoRefresh_(found);
  checkLinkSheet_(found);

  var errors = [], warns = [];
  for (var i = 0; i < found.length; i++) {
    (found[i].level === CHECK_LEVEL.ERROR ? errors : warns).push(found[i]);
  }
  return { errors: errors, warns: warns };
}

function addFinding_(list, level, title, detail, fix) {
  list.push({ level: level, title: title, detail: detail || '', fix: fix || '' });
}

/* ---------------- 個別の点検 ---------------- */

function checkClasses_(found) {
  var classes;
  try {
    classes = getClasses();
  } catch (e) {
    addFinding_(found, CHECK_LEVEL.ERROR, '「' + SH.CLASSES + '」シートがありません',
      String(e.message || e), '「' + MENU.SETUP + '」を実行してください。');
    return;
  }

  if (!classes.length) {
    addFinding_(found, CHECK_LEVEL.ERROR, 'クラスが1つも登録されていません', '',
      '「' + SH.CLASSES + '」シートにクラス名を入力してください。');
    return;
  }

  var names = classes.map(function (c) { return c.name; });
  var dup = findDuplicates_(names);
  if (dup.length) {
    addFinding_(found, CHECK_LEVEL.ERROR, 'クラス名が重複しています', dup.join('、'),
      '重複した行を消してください。このままでは面談枠が二重に作られます。');
  }

  // シート名に使えない文字
  for (var i = 0; i < names.length; i++) {
    if (/[\[\]\*\?\/\\:]/.test(names[i])) {
      addFinding_(found, CHECK_LEVEL.ERROR, 'クラス名に使えない文字が入っています', names[i],
        '[ ] * ? / \\ : は名簿シートの名前に使えません。別の表記にしてください。');
    }
  }

  var cfg = null;
  try { cfg = getConfig(); } catch (e) { cfg = null; }
  if (cfg && cfg.notifyTeacher) {
    var noMail = [];
    for (var m = 0; m < classes.length; m++) {
      if (!classes[m].email) noMail.push(classes[m].name);
    }
    if (noMail.length) {
      addFinding_(found, CHECK_LEVEL.WARN, '担任メールが未入力のクラスがあります', noMail.join('、'),
        '「担任メール通知」がTRUEですが、このクラスには通知が届きません。');
    }
  }
}

function checkDays_(found) {
  var days;
  try {
    days = getDays();
  } catch (e) {
    addFinding_(found, CHECK_LEVEL.ERROR, '「' + SH.DAYS + '」シートがありません', String(e.message || e), '');
    return;
  }

  if (!days.length) {
    addFinding_(found, CHECK_LEVEL.ERROR, '面談日が1日も設定されていません', '',
      '「' + SH.DAYS + '」シートで、実施する日の「実施する」にチェックを入れてください。');
    return;
  }

  var dup = findDuplicates_(days.map(function (d) { return dateLabel_(d); }));
  if (dup.length) {
    addFinding_(found, CHECK_LEVEL.ERROR, '同じ日付が複数あります', dup.join('、'),
      '重複した行を消すか、片方の「実施する」のチェックを外してください。');
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var past = [];
  for (var i = 0; i < days.length; i++) {
    if (days[i] < today) past.push(dateLabel_(days[i]));
  }
  if (past.length) {
    addFinding_(found, CHECK_LEVEL.WARN, '過ぎた日付が実施日になっています', past.join('、'),
      '前回の面談日が残っている可能性があります。');
  }
}

function checkConfigValues_(found) {
  var cfg;
  try {
    cfg = getConfig();
  } catch (e) {
    addFinding_(found, CHECK_LEVEL.ERROR, '「' + SH.CONFIG + '」シートが読めません', String(e.message || e), '');
    return;
  }

  if (!(cfg.slotMin >= 1)) {
    addFinding_(found, CHECK_LEVEL.ERROR, '「面談枠の長さ(分)」が正しくありません', String(cfg.slotMin),
      '1以上の数を入れてください。');
  }
  if (!(cfg.breakMin >= 0)) {
    addFinding_(found, CHECK_LEVEL.ERROR, '「枠間の休憩(分)」が正しくありません', String(cfg.breakMin),
      '0以上の数を入れてください（0なら休憩なし）。');
  }
  if (!(cfg.slotsPerDay >= 1)) {
    addFinding_(found, CHECK_LEVEL.ERROR, '「1日の枠数」が正しくありません', String(cfg.slotsPerDay),
      '1以上の数を入れてください。');
  }
  if (cfg.siblingOpenAt) {
    addFinding_(found, CHECK_LEVEL.WARN, '「きょうだい予約の受付開始」に日時が入っています',
      Utilities.formatDate(cfg.siblingOpenAt, TZ, 'M/d HH:mm'),
      'きょうだい世帯だけ先に受け付けると、他の家庭から不公平に見えます。' +
      '全家庭を同時に受け付ける方針であれば、空欄に戻してください。');
  }
  if (cfg.parentUrl) {
    if (!/^https:\/\/script\.google\.com\/.+\/exec(\?.*)?$/.test(cfg.parentUrl)) {
      addFinding_(found, CHECK_LEVEL.ERROR, '「保護者用URL」の書き方が正しくありません', cfg.parentUrl,
        'https://script.google.com/macros/s/……/exec の形で貼り付けてください（末尾は /exec）。');
    }
  } else {
    addFinding_(found, CHECK_LEVEL.WARN, '「保護者用URL」が未入力です', '',
      'いまは自動取得しています。デプロイが複数あると、案内プリントのQRコードが' +
      '実際に配るURLと違うものを指すことがあります。「' + SH.CONFIG + '」シートに貼り付けておいてください。');
  }
  if (!cfg.adminPasscode) {
    addFinding_(found, CHECK_LEVEL.WARN, '管理パスコードが未設定です', '',
      '担任用のWeb管理画面が使えません。使う場合は「' + SH.CONFIG + '」シートに入力してください。');
  }
  if (cfg.openAt && cfg.closeAt && cfg.openAt > cfg.closeAt) {
    addFinding_(found, CHECK_LEVEL.ERROR, '予約受付の開始が締切より後になっています',
      Utilities.formatDate(cfg.openAt, TZ, 'M/d HH:mm') + ' → ' +
      Utilities.formatDate(cfg.closeAt, TZ, 'M/d HH:mm'),
      'このままでは誰も予約できません。');
  }
}

function checkRoster_(found) {
  var classes;
  try { classes = getClasses(); } catch (e) { return; }

  var ss = ss_();
  var empty = [];

  for (var c = 0; c < classes.length; c++) {
    var clsName = classes[c].name;
    var sh = ss.getSheetByName(CLASS_SHEET_PREFIX + clsName);
    if (!sh) {
      addFinding_(found, CHECK_LEVEL.ERROR, '名簿シートがありません', CLASS_SHEET_PREFIX + clsName,
        '「' + MENU.SYNC_CLASSES + '」を実行すると作られます。');
      continue;
    }

    var last = sh.getLastRow();
    var rows = last >= 2 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
    var nos = [], blank = 0;

    for (var r = 0; r < rows.length; r++) {
      var no = Number(rows[r][0]);
      var nm = String(rows[r][1] || '').trim();
      if (!no && !nm) continue;
      if (!no || !nm) { blank++; continue; }
      nos.push(no);
    }

    if (!nos.length) { empty.push(clsName); continue; }

    var dupNo = findDuplicates_(nos);
    if (dupNo.length) {
      addFinding_(found, CHECK_LEVEL.ERROR, clsName + ' の出席番号が重複しています',
        dupNo.join('、') + ' 番',
        '同じ番号の生徒がいると、あとの生徒は予約できません（名簿と一致しない、と表示されます）。番号を直してください。');
    }
    if (blank) {
      addFinding_(found, CHECK_LEVEL.WARN, clsName + ' に番号か氏名が欠けた行があります', blank + ' 行',
        '出席番号と氏名の両方がそろっていない行は、名簿として読み込まれません。');
    }
  }

  if (empty.length) {
    addFinding_(found, CHECK_LEVEL.ERROR, '名簿が空のクラスがあります', empty.join('、'),
      '該当の「' + CLASS_SHEET_PREFIX + '〇組」シートのA・B列に、出席番号と生徒氏名を貼り付けてください。' +
      'このままでは、そのクラスの保護者は予約できません。');
  }
}

function checkSlotState_(found) {
  var slots;
  try { slots = readSlots_(); } catch (e) { slots = []; }

  if (!slots.length) {
    addFinding_(found, CHECK_LEVEL.ERROR, '面談枠がまだ作られていません', '',
      '「' + MENU.GENERATE + '」を実行してください。');
    return;
  }

  // 枠IDの重複（過去の重複入力の名残など）
  var dupId = findDuplicates_(slots.map(function (s) { return String(s.v[COL.SLOT_ID - 1]); }));
  if (dupId.length) {
    addFinding_(found, CHECK_LEVEL.ERROR, '面談枠が二重に登録されています',
      dupId.slice(0, 5).join('、') + (dupId.length > 5 ? ' ほか' : ''),
      'クラス名か日付が重複した状態で枠が作られたようです。重複を直してから「' + MENU.GENERATE + '」を実行してください。');
  }

  // だめなコマと予約の衝突
  var ngSet = {};
  try { ngSet = readNgSet_(); } catch (e) { ngSet = {}; }
  var conflicts = findNgConflicts_(ngSet, slots);
  if (conflicts.length) {
    var lines = [];
    for (var i = 0; i < conflicts.length && i < 5; i++) lines.push(ngConflictLabel_(conflicts[i]));
    addFinding_(found, CHECK_LEVEL.WARN, 'だめなコマに予約が残っています',
      lines.join(' / ') + (conflicts.length > 5 ? ' ほか' : ''),
      '保護者に連絡して時間を移すか取り消し、「' + MENU.NG_APPLY + '」を実行してください。');
  }

  // 予約はあるのに、全体が非公開のまま
  var cfg = null;
  try { cfg = getConfig(); } catch (e) { cfg = null; }
  var booked = 0, open = 0;
  for (var s = 0; s < slots.length; s++) {
    var st = String(slots[s].v[COL.STATUS - 1]);
    if (isTakenSlot_(slots[s].v)) booked++;   // 予備に入れたぶんも数える
    else if (st === STATUS.OPEN) open++;
  }
  if (cfg && cfg.published && open === 0 && booked < slots.length) {
    addFinding_(found, CHECK_LEVEL.WARN, '空き枠が1つもありません', '',
      '受付中ですが、選べる時間がありません。枠数か面談日を増やすことを検討してください。');
  }

  // 名簿の人数に対して枠が足りない
  var roster = 0;
  try { roster = getRoster().length; } catch (e) { roster = 0; }
  if (roster && (booked + open) < roster) {
    addFinding_(found, CHECK_LEVEL.WARN, '枠の数が生徒の人数より少なくなっています',
      '予約できる枠 ' + (booked + open) + ' に対して 生徒 ' + roster + ' 名',
      '全員分の枠がありません。面談日か1日の枠数を増やしてください。');
  }
}

function checkOrphanSheets_(found) {
  var classes;
  try { classes = getClasses(); } catch (e) { return; }

  var known = {};
  for (var c = 0; c < classes.length; c++) known[classes[c].name] = true;

  var ss = ss_();
  var sheets = ss.getSheets();
  var orphans = [];

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf(CLASS_SHEET_PREFIX) !== 0) continue;
    var cls = name.slice(CLASS_SHEET_PREFIX.length);
    if (known[cls]) continue;
    orphans.push(name + '（' + countStudentsInSheet_(sheets[i]) + '名）');
  }

  if (orphans.length) {
    addFinding_(found, CHECK_LEVEL.WARN, '使われていない名簿シートがあります', orphans.join('、'),
      'クラスを減らした・名前を変えた名残です。個人情報が残っているため、確認のうえ手動で削除してください。');
  }
}

/**
 * 予約ログから「入力が続けて一致しなかった家庭」を拾う。
 * 予約できずに困っている家庭を、学校側から見つけるための手がかり。
 * @return {Array<{cls:string, no:number, count:number, last:Date, detail:string}>}
 */
function recentInputErrors_(days) {
  var sh = ss_().getSheetByName(SH.LOG);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];

  var since = new Date();
  since.setDate(since.getDate() - (days || 14));

  var vals = sh.getRange(2, 1, last - 1, 7).getValues();
  var byKey = {}, order = [];

  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][1]) !== '入力エラー') continue;
    var at = vals[i][0];
    if (!(at instanceof Date) || at < since) continue;

    var cls = String(vals[i][3] || '');
    var no = Number(vals[i][4]) || 0;
    var key = cls + '_' + no;
    if (!byKey[key]) {
      byKey[key] = { cls: cls, no: no, count: 0, last: at, detail: String(vals[i][6] || '') };
      order.push(key);
    }
    byKey[key].count++;
    // ログは古い順に並ぶ。同じ時刻なら後の行を採用する
    if (at >= byKey[key].last) {
      byKey[key].last = at;
      byKey[key].detail = String(vals[i][6] || '');
    }
  }

  var out = [];
  for (var k = 0; k < order.length; k++) out.push(byKey[order[k]]);
  out.sort(function (a, b) { return b.last - a.last; });
  return out;
}

/** 表示用の1行 */
function inputErrorLabel_(e) {
  return e.cls + ' ' + e.no + '番　' +
    Utilities.formatDate(e.last, TZ, 'M/d HH:mm') + '　' + e.detail;
}

function checkLinkSheet_(found) {
  var problems = [];
  try { problems = checkLinks_(); } catch (e) { return; }
  if (!problems.length) return;

  addFinding_(found, CHECK_LEVEL.ERROR, '「' + SH.LINK + '」シートの内容に誤りがあります',
    problems.slice(0, 5).join(' ／ ') + (problems.length > 5 ? ' ほか' : ''),
    '紐づけが正しくないと、同じお子さんが両方のクラスで予約できてしまいます。' +
    'クラス名と出席番号が名簿と一致しているか確認してください。');
}

function checkAutoRefresh_(found) {
  var cfg = null;
  try { cfg = getConfig(); } catch (e) { return; }
  if (!cfg || !cfg.published) return;

  var on = false;
  try { on = autoRefreshEnabled_(); } catch (e) { return; }
  if (on) return;

  var pending = 0;
  try { pending = pendingViewUpdates_(); } catch (e) { pending = 0; }

  addFinding_(found, CHECK_LEVEL.WARN, '表示の自動更新がオフです',
    pending ? '未反映の予約が ' + pending + ' 件たまっています' : '',
    '受付中は「' + MENU.AUTO_REFRESH + '」でオンにしておくと、' +
    '全体ビューとクラス別予約表が自動で最新になります。' +
    '（オフのままでも予約は正しく処理されます）');
}

function checkInputErrors_(found) {
  var errs = [];
  try { errs = recentInputErrors_(14); } catch (e) { return; }
  if (!errs.length) return;

  var lines = [];
  for (var i = 0; i < errs.length && i < 6; i++) lines.push(inputErrorLabel_(errs[i]));

  addFinding_(found, CHECK_LEVEL.WARN, '予約できずに困っている家庭がありそうです',
    lines.join(' ／ ') + (errs.length > 6 ? ' ほか' : ''),
    '入力が続けて名簿と一致しなかった記録です（過去14日）。出席番号や氏名の表記を確認し、担任から連絡してください。');
}

/* ---------------- 表示用 ---------------- */

/** 点検結果を、ダイアログに出せる文章にする */
function formatCheckResult_(res) {
  var nl = String.fromCharCode(10);

  if (!res.errors.length && !res.warns.length) {
    return '問題は見つかりませんでした。' + nl + nl +
      'クラス・面談日・設定・名簿・面談枠のいずれにも、予約の妨げになる入力ミスはありません。';
  }

  var out = [];

  if (res.errors.length) {
    out.push('■ 直したほうがよいもの（' + res.errors.length + '件）');
    out.push('このままでは、予約が正しく動かない可能性があります。');
    out.push('');
    for (var i = 0; i < res.errors.length; i++) out.push(findingText_(res.errors[i], nl));
  }

  if (res.warns.length) {
    if (out.length) out.push('');
    out.push('■ 確認しておきたいもの（' + res.warns.length + '件）');
    out.push('');
    for (var w = 0; w < res.warns.length; w++) out.push(findingText_(res.warns[w], nl));
  }

  return out.join(nl);
}

function findingText_(f, nl) {
  var t = '・' + f.title;
  if (f.detail) t += nl + '　　' + f.detail;
  if (f.fix) t += nl + '　　→ ' + f.fix;
  return t + nl;
}
