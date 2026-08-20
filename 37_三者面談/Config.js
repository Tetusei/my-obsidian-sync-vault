/**
 * 三者面談 予約システム — 共通設定・ユーティリティ
 *
 * データはすべてスプレッドシート「三者面談」に置く。
 * 保護者用 Web アプリは「実行:自分 / アクセス:全員(匿名)」でデプロイするため、
 * メールアドレスは一切収集しない。本人確認はクラス＋出席番号＋氏名＋予約コードで行う。
 */

/** Antigravity 管理バージョン */
var VERSION = '3.3.0';

/** 対象スプレッドシート。バインドでも単体スクリプトでも動くよう ID を明示する。 */
var SPREADSHEET_ID = '1nvbdoNcZvwCrPi48GdxG_Q6eveCBDvk9s7HF4V10BM0';

/** シート名 */
var SH = {
  CONFIG: '設定',
  DAYS: '面談日',
  CLASSES: 'クラス',
  SLOTS: '枠マスタ',
  OVERVIEW: '全体ビュー',
  LOG: '予約ログ'
};

/** フォルダ名定数 */
var BACKUP_FOLDER_NAME = '📦_バックアップ保存箱';
var PDF_FOLDER_NAME = '📄_三者面談PDF';

/** 枠マスタの列番号(1始まり) */
var COL = {
  SLOT_ID: 1,
  DATE: 2,
  START: 3,
  END: 4,
  CLASS: 5,
  TEACHER: 6,
  STATUS: 7,
  NUMBER: 8,
  STUDENT: 9,
  GUARDIAN: 10,
  NOTE: 11,
  CODE: 12,
  BOOKED_AT: 13
};
var SLOT_LAST_COL = 13;

/** 枠の状態 */
var STATUS = {
  OPEN: '空き',
  BOOKED: '予約済',
  BLOCKED: 'ブロック'
};

var TZ = 'Asia/Tokyo';

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。メニューの「初期セットアップ」を実行してください。');
  return sh;
}

/** 空白・全角半角・大文字小文字を無視して比較するための正規化 */
function norm_(v) {
  return String(v == null ? '' : v)
    .normalize('NFKC')
    .replace(/[\s　]/g, '')
    .toLowerCase();
}

function ymd_(d) {
  return Utilities.formatDate(toDate_(d), TZ, 'yyyy-MM-dd');
}

function ymdCompact_(d) {
  return Utilities.formatDate(toDate_(d), TZ, 'yyyyMMdd');
}

var WEEKDAY_JA = ['日', '月', '火', '水', '木', '金', '土'];

function dateLabel_(d) {
  var dt = toDate_(d);
  return Utilities.formatDate(dt, TZ, 'M月d日') + '(' + WEEKDAY_JA[dt.getDay()] + ')';
}

function toDate_(d) {
  if (d instanceof Date) return d;
  var s = String(d).trim().replace(/\//g, '-');
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) throw new Error('日付として読めません: ' + d);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/** "13:40" / Date / シリアル値 を分(0時からの通算)に変換 */
function toMinutes_(v) {
  if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})\s*[:時]\s*(\d{1,2})/);
  if (m) return Number(m[1]) * 60 + Number(m[2]);
  m = s.match(/^(\d{3,4})$/); // 1340 のような書き方も許す
  if (m) {
    var n = Number(m[1]);
    return Math.floor(n / 100) * 60 + (n % 100);
  }
  throw new Error('時刻として読めません: ' + v);
}

function fromMinutes_(min) {
  var h = Math.floor(min / 60), m = min % 60;
  return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
}

function truthy_(v) {
  var s = norm_(v);
  return s === 'true' || s === 'はい' || s === '○' || s === 'yes' || s === '1' || s === 'on';
}

/**
 * 設定シートを key/value で読む。
 * @return {Object} 正規化済みの設定
 */
function getConfig() {
  var sh = sheet_(SH.CONFIG);
  var values = sh.getDataRange().getValues();
  var raw = {};
  for (var i = 0; i < values.length; i++) {
    var k = String(values[i][0] || '').trim();
    if (k) raw[k] = values[i][1];
  }
  var slotMin = Number(raw['面談枠の長さ(分)'] || 15);
  var breakMin = Number(raw['枠間の休憩(分)'] || 10);
  return {
    title: String(raw['見出し'] || '三者面談 予約'),
    notice: String(raw['案内文'] || ''),
    startMin: toMinutes_(raw['面談開始時刻'] || '13:40'),
    slotMin: slotMin,
    breakMin: breakMin,
    slotsPerDay: Number(raw['1日の枠数'] || 6),
    openAt: raw['予約受付開始'] instanceof Date ? raw['予約受付開始'] : null,
    closeAt: raw['予約受付締切'] instanceof Date ? raw['予約受付締切'] : null,
    maxPerStudent: Number(raw['1家庭あたり予約可能数'] || 1),
    checkName: raw['氏名照合'] === '' || raw['氏名照合'] == null ? true : truthy_(raw['氏名照合']),
    published: truthy_(raw['公開']),
    notifyTeacher: truthy_(raw['担任メール通知']),
    adminPasscode: String(raw['管理パスコード'] || '')
  };
}

/** 受付期間内かどうか。{ok:boolean, message:string} */
function bookingWindow_(cfg) {
  var now = new Date();
  if (!cfg.published) {
    return { ok: false, message: '現在、予約の受付を停止しています。' };
  }
  if (cfg.openAt && now < cfg.openAt) {
    return { ok: false, message: '予約受付は ' + Utilities.formatDate(cfg.openAt, TZ, 'M月d日 HH:mm') + ' から開始します。' };
  }
  if (cfg.closeAt && now > cfg.closeAt) {
    return { ok: false, message: '予約受付は ' + Utilities.formatDate(cfg.closeAt, TZ, 'M月d日 HH:mm') + ' で締め切りました。変更は担任までご連絡ください。' };
  }
  return { ok: true, message: '' };
}

/** クラス一覧 [{name, teacher, email}] */
function getClasses() {
  var values = sheet_(SH.CLASSES).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var name = String(values[i][0] || '').trim();
    if (!name) continue;
    out.push({
      name: name,
      teacher: String(values[i][1] || '').trim(),
      email: String(values[i][2] || '').trim()
    });
  }
  return out;
}

/** 実施日一覧 [Date]（「実施する」がTRUEの行のみ、昇順） */
function getDays() {
  var values = sheet_(SH.DAYS).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    if (!truthy_(values[i][2])) continue;
    out.push(toDate_(values[i][0]));
  }
  out.sort(function (a, b) { return a - b; });
  return out;
}

/** 
 * 各クラスの「予約表_〇組」シートのA・B列から生徒名簿を取得 [{cls, no, name}]
 */
function getRoster() {
  var classes = getClasses();
  var ss = ss_();
  var out = [];

  for (var c = 0; c < classes.length; c++) {
    var clsName = classes[c].name;
    var sheetName = '予約表_' + clsName;
    var sh = ss.getSheetByName(sheetName);
    if (!sh) continue;

    var lastRow = sh.getLastRow();
    if (lastRow < 2) continue;

    var vals = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      var no = Number(vals[i][0]);
      var name = String(vals[i][1] || '').trim();
      if (no && name) {
        out.push({ cls: clsName, no: no, name: name });
      }
    }
  }

  // 旧「生徒名簿」シートがまだ存在する場合のバックアップフォールバック
  if (!out.length) {
    try {
      var oldSh = ss.getSheetByName('生徒名簿');
      if (oldSh && oldSh.getLastRow() >= 2) {
        var oldVals = oldSh.getRange(2, 1, oldSh.getLastRow() - 1, 3).getValues();
        for (var k = 0; k < oldVals.length; k++) {
          var oCls = String(oldVals[k][0] || '').trim();
          var oNo = Number(oldVals[k][1]) || 0;
          var oName = String(oldVals[k][2] || '').trim();
          if (oCls && oName) out.push({ cls: oCls, no: oNo, name: oName });
        }
      }
    } catch (e) {
      console.warn('旧生徒名簿読み込みスキップ:', e);
    }
  }

  return out;
}

/** 予約ログに1行追記 */
function logAction_(action, slotId, cls, no, name, detail) {
  try {
    var sh = ss_().getSheetByName(SH.LOG);
    if (!sh) return;
    sh.appendRow([new Date(), action, slotId || '', cls || '', no || '', name || '', detail || '']);
  } catch (err) {
    console.error('logAction_ failed: ' + err);
  }
}

/** システム用内部シート（枠マスタのみ）を非表示化し、予約ログは表示状態を保つ */
function hideInternalSheets() {
  var ss = ss_();
  // 枠マスタのみ非表示化
  var slotSh = ss.getSheetByName(SH.SLOTS);
  if (slotSh && !slotSh.isSheetHidden()) {
    try { slotSh.hideSheet(); } catch (e) { /* 無視 */ }
  }
  // 予約ログシートは確認用として表示状態（Un-hide）を維持
  var logSh = ss.getSheetByName(SH.LOG);
  if (logSh && logSh.isSheetHidden()) {
    try { logSh.showSheet(); } catch (e) { /* 無視 */ }
  }
}
