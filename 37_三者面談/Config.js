/**
 * 三者面談 予約システム — 共通設定・ユーティリティ
 *
 * データはすべてスプレッドシート「三者面談」に置く。
 * 保護者用 Web アプリは「実行:自分 / アクセス:全員(匿名)」でデプロイするため、
 * メールアドレスは一切収集しない。本人確認はクラス＋出席番号＋氏名＋予約コードで行う。
 */

/**
 * Antigravity 管理バージョン。
 * 「予約の受付・URL ▸ いまの受付状態を確認する」に表示されるので、
 * 手元のファイルがどの版かを担任が確かめられる。
 *
 * 4.5.0 の変更
 *   ・メニューに「いま何がオンか」を表示（受付／だめなコマの自動反映／自動更新／
 *     自動リマインド／自動バックアップ）。設定シートからも切り替え可能に
 *   ・設定シートの行を意味ごとに並べ替え。管理画面URLを自動で書き出し、
 *     管理パスコードと並べて末尾に。1家庭あたり→1生徒あたりに改名
 *   ・自動バックアップを専用フォルダに分離。30世代でゴミ箱へ
 *   ・予約ログに予約コードを記録
 *   ・ウェブ管理画面：枠の切替を即時反映、まとめて指定の説明、処理中の表示
 *   ・修正 予約表の作り直しで名簿を失う恐れがあったのを解消
 *   ・修正 枠の再生成を排他制御下に（再生成中の予約が消える恐れ）
 *   ・修正 ファイルを複製したとき、複製側の操作が元のファイルへ向かう問題
 *   ・修正 setConfigValue_ の二重定義
 *   ・一斉アクセスに備え、参照データを短時間キャッシュ
 *
 * 4.6.0 の変更
 *   ・予備のコマ。保護者には見せず、担任が「予約表_〇組」または管理画面から入れる
 *   ・設定シートの並びを見直し、管理画面URLと管理パスコードを末尾へ
 *   ・リセットに「配布用にまっさらにする」を追加（他校へ渡すときの初期化）
 *   ・保護者用URLを書き替えると、管理画面URLもその場で作り直す
 *   ・Web管理画面で日付の区切りを見やすく
 *   ・受付状態の確認にバージョンを表示
 *   ・修正 ファイルを複製したとき、複製側の操作が元のファイルへ向かう問題
 *
 * 4.6.1 の変更
 *   ・保護者用URLが未設定・不正・到達不能なら、誤ったQR付きプリントを作らず停止する
 *
 * 4.6.2 の変更
 *   ・Webアプリ実行時にも、紐づくスプレッドシートをScript Propertiesから確実に開く
 *   ・案内プリント作成前に、予約画面から設定シートまで読めることを確認する
 *
 * 4.6.3 の変更
 *   ・案内プリントURLへ版番号を付け、端末に残った旧エラーページを回避する
 *
 * 4.6.4 の変更
 *   ・一部のAndroid端末が長いURLを誤復号するQRマスクを選ばないようにする
 *
 * 4.6.5 の変更
 *   ・印刷用QRを誤り訂正Lにし、長いURLでもマスを大きくしてスマホで読みやすくする
 *
 * 4.6.6 の変更
 *   ・AndroidがURLへ無効な /u/1/ を挿入しないよう、QRでGoogleアカウント0を明示する
 *
 * 4.6.7 の変更
 *   ・設定の参照キャッシュが実際には保存されない不具合を修正
 *   ・クラス変更の反映時に古いキャッシュを読む不具合を修正
 *
 * 4.6.8 の変更
 *   ・QRの形式情報のビット順を修正し、読み取れないQRが作られる不具合を修正
 *
 * 4.6.9 の変更
 *   ・Androidで予約画面が開かない場合の、シークレットタブを使う具体的な操作を案内プリントへ追加
 *
 * 4.6.10 の変更
 *   ・案内プリント全体の文字を拡大し、下部の説明を1文ずつ分けて行間を広げた
 *
 * 4.6.11 の変更
 *   ・案内プリントの曖昧な注意書きを、予約画面の表示内容が明確に伝わる文へ変更
 *
 * 4.6.12 の変更
 *   ・案内プリントの利用手順へGoogleカレンダー登録を追加し、重要語句を部分的に太字化
 *
 * 4.6.13 の変更
 *   ・シートのメニュー「この枠に生徒を入れる」で、予備の枠にも生徒を入れられるようにした
 *   ・修正 予備の枠に入れた生徒が、次の表示更新で消えることがあったのを解消
 *
 * 4.6.14 の変更
 *   ・自動で作り直される場所に、書き込む前の警告（保護：警告のみ）を掛けた
 *   ・全体ビューへ手で書いたときも、その場で知らせるようにした
 *   ・登録されないまま消える手入力を予約ログに残し、データ点検で拾えるようにした
 *   ・予備の枠への手入力にも、名簿の照合と二重予約の確認を通すようにした
 *
 * 4.6.15 の変更
 *   ・予約表への手入力を、予備の行も含めてすべて「書いても残らない」に統一
 *   ・予備の枠の記入は、担任用の管理画面か「この枠に生徒を入れる」だけに一本化
 *   ・「この枠に生徒を入れる」に、予備の枠を空にするボタンを足した
 *
 * 4.6.16 の変更
 *   ・修正 予約表の「予約済（予備）」が未予約と同じ色で、埋まったのが分からなかった
 *
 * 4.6.17 の変更
 *   ・全体ビューのクラス名から、そのクラスの予約表へ飛べるようにした
 *   ・各予約表のG1から、全体ビューへ戻れるようにした
 *
 * 4.6.18 の変更
 *   ・黄色い予備行のL〜N列へ、担任が直接入力できるように戻した
 *   ・直接入力にも名簿照合と二重予約確認を通し、誤入力は理由を表示して元へ戻す
 *   ・通常枠の誤編集防止と、管理画面・小窓からの入力方法はそのまま維持
 */
var VERSION = '4.6.18';

/**
 * ふだんは空にしておく。操作先は必ず ss_() / ssId_() から取る。
 *
 * ここにIDを書くと、ファイルをコピーして使い始めたときに、
 * コピー側での操作が元のファイルへ向かってしまう。
 * 実際にそれで名簿を消す事故が起きたので、空のままにしてある。
 *
 * スプレッドシートに紐づかない単体のスクリプトとして動かす場合にだけ、
 * 対象のファイルIDを入れる。
 */
var SPREADSHEET_ID = '';

/**
 * Webアプリでは SpreadsheetApp.getActiveSpreadsheet() が null になるため、
 * コンテナバインドだけに頼ると、画面は出ても予約データを開けない。
 * スプレッドシート上のメニューから動いたときに、そのファイルIDをプロジェクト固有の
 * Script Propertiesへ記憶し、Webアプリ側はそこから開く。
 */
var BOUND_SPREADSHEET_ID_KEY = 'BOUND_SPREADSHEET_ID';

/**
 * 対象IDを今回の実行でもう記憶したか。
 * setProperty は毎回リモートへ書きに行くため、sheet_() 経由で何十回も呼ばれる
 * ss_() の中で毎度実行すると、表の作り直しが目に見えて遅くなる。
 * グローバルは1回の実行が終われば消えるので、次の実行では必ず書き直される。
 */
var SS_ID_REMEMBERED_ = false;

/**
 * Webアプリの対象を初期化する保守用関数。
 * 通常はスプレッドシートを開けば ss_() が自動記憶するため、手動実行は不要。
 * @param {string} spreadsheetId このスクリプトに紐づくスプレッドシートID
 * @return {{ok:boolean, spreadsheetId:string, title:string}}
 */
function initializeWebAppTarget(spreadsheetId) {
  var id = String(spreadsheetId || '').trim();
  if (!/^[A-Za-z0-9_-]{20,}$/.test(id)) {
    throw new Error('スプレッドシートIDの書き方が正しくありません。');
  }
  var target = SpreadsheetApp.openById(id); // 権限と実在を確認してから保存する
  PropertiesService.getScriptProperties().setProperty(BOUND_SPREADSHEET_ID_KEY, id);
  return { ok: true, spreadsheetId: id, title: target.getName() };
}

/** シート名 */
var SH = {
  CONFIG: '設定',
  DAYS: '面談日',
  CLASSES: 'クラス',
  SLOTS: '枠マスタ',
  OVERVIEW: '全体ビュー',
  NG: 'だめなコマ',
  LINK: '交流学級',
  LOG: '予約ログ'
};

/**
 * メニューの表示名。案内メッセージでも同じ文字列を使うため、ここにまとめておく。
 * メニュー構成を変えたら、この定数だけ直せば全部の案内文が追従する。
 */
var MENU = {
  ROOT: '三者面談',
  NG: '🚫 だめなコマ',
  NG_SHEET: '🚫 だめなコマ ▸ 一覧シートを作る・更新する',
  NG_APPLY: '🚫 だめなコマ ▸ 指定を枠に反映する',
  RECEPTION: '予約の受付・URL',
  PUBLISH: '予約の受付・URL ▸ 受付を開始する',
  URL: '予約の受付・URL ▸ 保護者用URLを表示',
  REFRESH: '全体ビュー・予約表を更新',
  PDF: '📄 印刷用PDFを作る',
  ROWOPS: '選んだ行を操作',
  MAINT: '準備・メンテナンス',
  SETUP: '準備・メンテナンス ▸ ① 初期セットアップ',
  GENERATE: '準備・メンテナンス ▸ ② 枠を再生成',
  SYNC_CLASSES: '準備・メンテナンス ▸ クラスの増減を反映する',
  ORDER_SHEETS: '準備・メンテナンス ▸ シートの並び順を整える',
  CHECK: '準備・メンテナンス ▸ データを点検する',
  NG_AUTO: '🚫 だめなコマ ▸ チェックを自動で反映する',
  REMIND: '予約の受付・URL ▸ 未予約の生徒を担任にメールする',
  AUTO_REFRESH: '準備・メンテナンス ▸ 表示の自動更新を設定する',
  BACKUP: '準備・メンテナンス ▸ 📦 バックアップを作成',
  RESET: '🧹 リセット',
  RESET_NEXT: '🧹 リセット ▸ 次の面談にそなえて初期化する',
  RESET_ALL: '🧹 リセット ▸ 年度末：すべて初期化する',
  AUTO_BACKUP: '準備・メンテナンス ▸ 毎日の自動バックアップを設定する',
  DAY_LIST: '📄 印刷用PDFを作る ▸ 当日の受付一覧（全校・時間順）'
};

/**
 * 自動で作り直される場所に掛ける「保護（警告のみ）」の目印。
 * この文字列で始まる保護だけを、システムが付け外しする。
 * 担任や管理職が自分で付けた保護には触れない。
 */
var EDIT_GUARD_TAG = '⚠三者面談：自動で作り直されます';

/**
 * 登録されないまま消える手入力を、予約ログへ残すときの「操作」名。
 * データ点検（Check.js）からも同じ名前で拾う。
 */
var STRAY_EDIT_ACTION = '未登録の手入力';

/** フォルダ名定数（PdfExport.js からもここを参照する） */
var BACKUP_FOLDER_NAME = '📦_バックアップ保存箱';
/**
 * 毎日の自動バックアップは、保存箱の中のサブフォルダに分ける。
 * 手動バックアップは「枠を再生成する直前」「リセットの直前」など、
 * 戻したい一点を狙って取るもの。日々の複製に埋もれると探せなくなるうえ、
 * 増えすぎた自動バックアップを整理するときに、消してはいけないものと混ざる。
 */
var AUTO_BACKUP_FOLDER_NAME = '🕒_自動（毎日）';
var AUTO_BACKUP_KEEP_KEY = '自動バックアップの保存世代数';
/**
 * だめなコマの自動反映がオンかどうかは、トリガーの有無で決まる。
 * ところが ScriptApp.getProjectTriggers() は認可が必要で、
 * メニューを組み立てる単純トリガーの onOpen からは呼べない。
 * そこで実態を「設定」シートに書き写しておき、メニューのラベルはそちらを読む。
 */
var NG_AUTO_KEY = 'だめなコマの自動反映';
var VIEW_AUTO_KEY = '表示の自動更新';
var REMINDER_AUTO_KEY = '締切前の自動リマインド';
var BACKUP_AUTO_KEY = '毎日の自動バックアップ';

/** 「設定」シートに書き写すときの説明文 */
var AUTO_FLAG_DESC = {};
AUTO_FLAG_DESC[NG_AUTO_KEY] = 'だめなコマのチェックを、付け外しした瞬間に枠へ反映するか。ここで切り替えてもメニューから切り替えてもよい';
AUTO_FLAG_DESC[VIEW_AUTO_KEY] = 'たまった予約を5分おきに全体ビュー・クラス別予約表へ反映するか。ここで切り替えてもメニューから切り替えてもよい';
AUTO_FLAG_DESC[REMINDER_AUTO_KEY] = '締切の指定日数前に、未予約者の一覧を担任へ自動送信するか。ここで切り替えてもメニューから切り替えてもよい';
AUTO_FLAG_DESC[BACKUP_AUTO_KEY] = '受付中に毎日バックアップを取るか。ここで切り替えてもメニューから切り替えてもよい';

/**
 * 数えている単位は家庭ではなく生徒1人（クラス＋出席番号）なので、
 * 「1家庭あたり」から名前を直した。古い名前のシートも読めるようにしてある。
 */
var MAX_PER_STUDENT_KEY = '1生徒あたり予約可能数';

/**
 * 担任用の管理画面は、保護者用URLに ?page=admin を付けただけの入口。
 * 知らないと開きようがないので、設定シートに書き出しておく。
 */
var ADMIN_URL_KEY = '管理画面URL';
var MAX_PER_STUDENT_KEY_OLD = '1家庭あたり予約可能数';

/** 設定シートのキー名の付け替え。値はそのまま残す */
var CONFIG_KEY_RENAMES = [
  {
    from: MAX_PER_STUDENT_KEY_OLD,
    to: MAX_PER_STUDENT_KEY,
    description: '同じ生徒が同時に持てる予約の数。通常は1（きょうだいは別々に数えるので1のままでよい）'
  }
];

/**
 * 「設定」シートの行を並べる順。
 *
 * 行は必要になった時点で末尾へ足していくため、放っておくと
 * 「リマインド日数」と「締切前の自動リマインド」のように、
 * 対で使う項目が離れてしまう。意味のまとまりごとに並べ直す。
 * ここに無いキーは、いまの順のまま後ろに残す。
 */
var CONFIG_ORDER = [
  // 保護者に見えるもの
  '見出し',
  '案内文',
  '保護者用URL',
  // 面談の時間割
  '面談開始時刻',
  '面談枠の長さ(分)',
  '枠間の休憩(分)',
  '1日の枠数',
  RESERVE_COUNT_KEY,
  // 受付
  '公開',
  '予約受付開始',
  '予約受付締切',
  'きょうだい予約の受付開始',
  MAX_PER_STUDENT_KEY,
  '氏名照合',
  // 担任への連絡
  '担任メール通知',
  REMINDER_AUTO_KEY,
  'リマインド日数(締切の何日前)',
  // 自動処理
  NG_AUTO_KEY,
  VIEW_AUTO_KEY,
  BACKUP_AUTO_KEY,
  AUTO_BACKUP_KEEP_KEY,
  // 担任用の管理画面。URLと合言葉は対で使うので、いちばん下にまとめる
  ADMIN_URL_KEY,
  '管理パスコード'
];
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

/** 「予約表_〇組」シートの見出し（左＝生徒別、右＝時間枠別） */
var CLASS_HEADER_LEFT = ['出席番号', '生徒氏名', '予約状況', '予約日時', '保護者氏名', '連絡事項'];
var CLASS_HEADER_RIGHT = ['日付', '時間', '状態', '出席番号', '生徒氏名', '保護者氏名', '予約コード'];

/** 枠の状態 */
/**
 * 予備コマの枠数を決める設定キー。
 * 保護者には見せず、担任が「もしものとき」に手で入れるための枠。
 */
var RESERVE_COUNT_KEY = '予備の枠数(1日あたり)';

var STATUS = {
  OPEN: '空き',
  BOOKED: '予約済',
  BLOCKED: 'ブロック',
  /**
   * 予備。保護者の画面には一覧にも出さない。
   * 担任用の管理画面、行メニュー、または「予約表_〇組」の黄色い行にある
   * L〜N列（出席番号・生徒氏名・保護者氏名）から入れる。
   */
  RESERVE: '予備'
};

/** 改行。メッセージの組み立てで使う */
var LF = String.fromCharCode(10);

var TZ = 'Asia/Tokyo';

/**
 * 対象のスプレッドシート。
 *
 * openById() は認可を必要とするので、単純トリガーの onOpen からは呼べない。
 * そのためメニューを組み立てる時点では設定を読めず、受付状態をラベルに出せなかった。
 * バインドされているときは getActive() を使う（単純トリガーでも読める・こちらのほうが速い）。
 */
/* ---------------- 参照データの短時間キャッシュ ----------------
 *
 * 予約受付が開く瞬間、保護者は一斉に集まる。
 * 予約1件ごとに「設定」「クラス」「交流学級」と13クラス分の名簿シートを読み直していると、
 * 1リクエストで20回近くシートを読むことになり、開始直後に詰まる。
 *
 * どれも面談期間中はほとんど変わらない参照データなので、
 * スクリプト全体で共有する短時間キャッシュに載せる。
 * 書き換える側では必ずキャッシュを捨てるので、担任の操作が反映されないことはない。
 */
var CACHE_KEY = {
  CONFIG: 'ref_config',
  CLASSES: 'ref_classes',
  ROSTER: 'ref_roster',
  LINKS: 'ref_links'
};

var CACHE_SEC = {
  CONFIG: 20,    // 受付の開始・停止をすぐ反映したいので短くする
  CLASSES: 120,
  ROSTER: 120,
  LINKS: 120
};

/** CacheService の1キーあたりの上限は100KB。余裕を見て手前で諦める */
var CACHE_MAX_CHARS = 95000;

function cacheGet_(key) {
  try {
    var hit = CacheService.getScriptCache().get(key);
    return hit ? JSON.parse(hit) : null;
  } catch (e) {
    return null;   // 読めなければシートから読み直すだけ
  }
}

function cachePut_(key, value, sec) {
  try {
    var text = JSON.stringify(value);
    if (text.length > CACHE_MAX_CHARS) return;
    CacheService.getScriptCache().put(key, text, sec);
  } catch (e) { /* 載せられなくても動作に影響はない */ }
}

/** 参照データのキャッシュを捨てる。もとのシートを書き換えたら必ず呼ぶ。 */
function dropRefCaches_() {
  try {
    CacheService.getScriptCache().removeAll([
      CACHE_KEY.CONFIG, CACHE_KEY.CLASSES, CACHE_KEY.ROSTER, CACHE_KEY.LINKS
    ]);
  } catch (e) { /* 無視 */ }
}

function ss_() {
  var active = null;
  try {
    active = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    active = null; // スプレッドシート以外から実行された経路
  }
  // シート上のメニュー・トリガーから実行されたときは、いま開いているファイルが正解。
  // コピー後に古いIDが残っていても、ここで必ず新しいIDへ上書きされる。
  if (active) {
    // 単純トリガー（onOpen / onEdit）では PropertiesService を呼べず必ず失敗する。
    // 先に印を立てておき、同じ実行の中で何度も失敗しにいかないようにする。
    if (!SS_ID_REMEMBERED_) {
      SS_ID_REMEMBERED_ = true;
      try {
        PropertiesService.getScriptProperties()
          .setProperty(BOUND_SPREADSHEET_ID_KEY, active.getId());
      } catch (e) { /* 記憶に失敗しても、今回のシート操作は続けられる */ }
    }
    return active;
  }

  // Webアプリ・時間主導トリガー・Execution APIでは active が無い。
  // 直前にシート側で記憶した、このスクリプト専用のIDを使う。
  var remembered = '';
  try {
    remembered = PropertiesService.getScriptProperties()
      .getProperty(BOUND_SPREADSHEET_ID_KEY) || '';
  } catch (e2) {
    remembered = '';
  }
  if (remembered) return SpreadsheetApp.openById(remembered);

  // 単体スクリプトとして明示IDを設定している場合だけ、そのIDを使う。
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);

  // まだ一度もスプレッドシート側から実行されていない。取り違えを避けて止める。
  if (!remembered) {
    throw new Error(
      '対象のスプレッドシートをまだ確認できません。' +
      'スプレッドシートを再読み込みし、「三者面談」メニューを一度開いてから、もう一度お試しください。');
  }
}

/**
 * いま操作している対象のファイルID。
 * トリガーの取り付け先、バックアップの複製元、PDFの出力元にも、必ずこちらを使う。
 */
function ssId_() {
  return ss_().getId();
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。メニューの「' + MENU.SETUP + '」を実行してください。');
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
 * 設定シートの数値を読む。
 * 空欄なら既定値。0 は「入力された0」として尊重する
 * （|| を使うと 0 が既定値に化けるため、この関数を通す）。
 */
function numConfig_(raw, key, def) {
  var v = raw[key];
  if (v === '' || v == null) return def;
  var n = Number(v);
  return isNaN(n) ? def : n;
}

/** 同じ値が2回以上出てくるものを返す */
function findDuplicates_(list) {
  var seen = {}, dup = {}, out = [];
  for (var i = 0; i < list.length; i++) {
    var k = String(list[i]);
    if (seen[k] && !dup[k]) { dup[k] = true; out.push(k); }
    seen[k] = true;
  }
  return out;
}

/**
 * 設定シートを key/value で読む。
 * @return {Object} 正規化済みの設定
 */
function getConfig() {
  var hit = cacheGet_(CACHE_KEY.CONFIG);
  if (hit) return reviveConfigDates_(hit);

  var sh = sheet_(SH.CONFIG);
  var values = sh.getDataRange().getValues();
  var raw = {};
  for (var i = 0; i < values.length; i++) {
    var k = String(values[i][0] || '').trim();
    if (k) raw[k] = values[i][1];
  }
  var slotMin = numConfig_(raw, '面談枠の長さ(分)', 15);
  var breakMin = numConfig_(raw, '枠間の休憩(分)', 10);   // 0（休憩なし）も有効
  var cfg = {
    title: String(raw['見出し'] || '三者面談 予約'),
    notice: String(raw['案内文'] || ''),
    startMin: toMinutes_(raw['面談開始時刻'] || '13:40'),
    slotMin: slotMin,
    breakMin: breakMin,
    slotsPerDay: numConfig_(raw, '1日の枠数', 6),
    reservePerDay: numConfig_(raw, RESERVE_COUNT_KEY, 0),
    openAt: raw['予約受付開始'] instanceof Date ? raw['予約受付開始'] : null,
    siblingOpenAt: raw['きょうだい予約の受付開始'] instanceof Date ? raw['きょうだい予約の受付開始'] : null,
    closeAt: raw['予約受付締切'] instanceof Date ? raw['予約受付締切'] : null,
    maxPerStudent: (raw[MAX_PER_STUDENT_KEY] === '' || raw[MAX_PER_STUDENT_KEY] == null)
      ? numConfig_(raw, MAX_PER_STUDENT_KEY_OLD, 1)
      : numConfig_(raw, MAX_PER_STUDENT_KEY, 1),
    checkName: raw['氏名照合'] === '' || raw['氏名照合'] == null ? true : truthy_(raw['氏名照合']),
    published: truthy_(raw['公開']),
    notifyTeacher: truthy_(raw['担任メール通知']),
    reminderDays: numConfig_(raw, 'リマインド日数(締切の何日前)', 3),
    ngAutoApply: truthy_(raw[NG_AUTO_KEY]),
    viewAutoRefresh: truthy_(raw[VIEW_AUTO_KEY]),
    reminderAuto: truthy_(raw[REMINDER_AUTO_KEY]),
    backupAuto: truthy_(raw[BACKUP_AUTO_KEY]),
    autoBackupKeep: numConfig_(raw, AUTO_BACKUP_KEEP_KEY, 30),
    adminPasscode: String(raw['管理パスコード'] || ''),
    parentUrl: String(raw['保護者用URL'] || '').trim()
  };

  cachePut_(CACHE_KEY.CONFIG, cfg, CACHE_SEC.CONFIG);
  return cfg;
}

/** JSON を通すと Date が文字列になるので、日時の項目だけ戻す */
function reviveConfigDates_(cfg) {
  cfg.openAt = cfg.openAt ? new Date(cfg.openAt) : null;
  cfg.siblingOpenAt = cfg.siblingOpenAt ? new Date(cfg.siblingOpenAt) : null;
  cfg.closeAt = cfg.closeAt ? new Date(cfg.closeAt) : null;
  return cfg;
}

/**
 * 保護者用ページのURL。
 *
 * ScriptApp.getService().getUrl() は「いちばん新しいデプロイ」を返すため、
 * デプロイをいくつも作っていると、実際に配っているURLと食い違う。
 * そこで「設定」シートに書いてあればそちらを優先する。
 *
 * @return {{url:string, fromConfig:boolean}}
 */
function webAppUrlInfo_() {
  var fromConfig = '';
  try {
    fromConfig = getConfig().parentUrl;
  } catch (e) { /* 設定が読めなくても、自動取得は試す */ }

  if (fromConfig) return { url: fromConfig, fromConfig: true };

  var auto = '';
  try {
    auto = ScriptApp.getService().getUrl() || '';
  } catch (e) {
    auto = '';
  }
  return { url: auto, fromConfig: false };
}

/** 保護者用ページのURL（文字列だけ必要なとき） */
function webAppUrl_() {
  return webAppUrlInfo_().url;
}

/** HTMLに埋め込む文字列の記号を打ち消す */
function escHtml_(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * 「設定」シートに指定のキーが無ければ、末尾に空欄で足す。
 * 既にある値には触らない。
 * 初期セットアップを実行済みのシートに、あとから増えたキーを届けるために使う。
 * @return {boolean} 足したら true
 */
function ensureConfigKey_(key, description, defaultValue) {
  var sh = ss_().getSheetByName(SH.CONFIG);
  if (!sh) return false;

  var last = sh.getLastRow();
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0] || '').trim() === key) return false;
    }
  }

  ensureSheetSize_(sh, last + 1, 3);
  sh.getRange(last + 1, 1, 1, 3)
    .setValues([[key, defaultValue == null ? '' : defaultValue, description || '']]);
  dropRefCaches_();
  return true;
}

/**
 * 設定シートのキー名を古い名前から新しい名前へ付け替える。
 * 値はそのまま残し、説明だけ新しい文面にする。
 * 新しい名前の行がすでにあるときは、古いほうには触らない。
 *
 * @return {Array<string>} 付け替えた内容
 */
function migrateConfigKeys_() {
  var sh = ss_().getSheetByName(SH.CONFIG);
  if (!sh) return [];

  var last = sh.getLastRow();
  if (last < 2) return [];

  var keys = sh.getRange(2, 1, last - 1, 1).getValues();
  var rowOf = {};
  for (var i = 0; i < keys.length; i++) {
    var k = String(keys[i][0] || '').trim();
    if (k) rowOf[k] = i + 2;
  }

  var done = [];
  for (var r = 0; r < CONFIG_KEY_RENAMES.length; r++) {
    var rule = CONFIG_KEY_RENAMES[r];
    if (rowOf[rule.to] || !rowOf[rule.from]) continue;

    sh.getRange(rowOf[rule.from], 1).setValue(rule.to);
    if (rule.description) sh.getRange(rowOf[rule.from], 3).setValue(rule.description);
    done.push(rule.from + ' → ' + rule.to);
  }
  return done;
}

/**
 * メニューの表示用に、自動処理の状態をまとめて1回で読む。
 *
 * getConfig() は時刻の解析などを行うため、設定が未完成だと例外になる。
 * メニューは設定が揃っていなくても開けなければならないので、
 * ここでは必要な行だけを素直に読む。
 *
 * @return {Object} キー → boolean
 */
function readAutoFlags_() {
  var flags = {};
  try {
    var sh = ss_().getSheetByName(SH.CONFIG);
    if (!sh) return flags;
    var last = sh.getLastRow();
    if (last < 2) return flags;
    var vals = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      var k = String(vals[i][0] || '').trim();
      if (k) flags[k] = truthy_(vals[i][1]);
    }
  } catch (e) {
    // メニュー生成中に読めないことがある。既定の表示のままにする
  }
  return flags;
}

/**
 * 「設定」シートの値を1つ書き換える。キーが無ければ何もしない。
 * @return {boolean} 実際に書き換えたら true
 */
function setConfigValue_(key, value) {
  var sh = ss_().getSheetByName(SH.CONFIG);
  if (!sh) return false;

  var last = sh.getLastRow();
  if (last < 2) return false;

  var vals = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() !== key) continue;
    if (vals[i][1] === value) return false;
    sh.getRange(i + 2, 2).setValue(value);
    dropRefCaches_();
    return true;
  }
  return false;
}

/**
 * 「設定」シートの説明（C列）を書き換える。文面を直したときに、
 * すでにある行へも反映させるために使う。異なるときだけ書き込む。
 */
function setConfigDescription_(key, description) {
  if (!description) return false;
  var sh = ss_().getSheetByName(SH.CONFIG);
  if (!sh) return false;

  var last = sh.getLastRow();
  if (last < 2) return false;

  var vals = sh.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() !== key) continue;
    if (String(vals[i][2] || '') === description) return false;
    sh.getRange(i + 2, 3).setValue(description);
    return true;
  }
  return false;
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

/** 見出し行から列の位置を探す。見つからなければ -1。 */
function headerIndex_(header, names) {
  for (var i = 0; i < header.length; i++) {
    var h = norm_(header[i]);
    for (var j = 0; j < names.length; j++) {
      if (h === norm_(names[j])) return i;
    }
  }
  return -1;
}

/** 「3年1組」から「3年」を読み取る。読めなければ空文字。 */
function inferGrade_(className) {
  var s = String(className || '');
  var m = s.match(/^\s*(\d+)\s*年/);
  if (m) return m[1] + '年';
  if (/特別支援|特支|支援学級/.test(s)) return '特別支援';
  return '';
}

/**
 * クラス一覧 [{name, grade, teacher, email, special}]
 *
 * 列の位置ではなく見出しの名前で読む。
 * 「学年」列を後から足しても、順番が違っても動くようにするため。
 */
function getClasses() {
  var hit = cacheGet_(CACHE_KEY.CLASSES);
  if (hit) return hit;

  var values = sheet_(SH.CLASSES).getDataRange().getValues();
  if (!values.length) return [];

  var header = values[0];
  var iName = headerIndex_(header, ['クラス', 'クラス名']);
  var iGrade = headerIndex_(header, ['学年']);
  var iTeacher = headerIndex_(header, ['担任名', '担任']);
  var iMail = headerIndex_(header, ['担任メール', 'メール', 'メールアドレス']);
  if (iName < 0) iName = 0;
  if (iTeacher < 0) iTeacher = 1;
  if (iMail < 0) iMail = 2;

  var out = [];
  for (var i = 1; i < values.length; i++) {
    var name = String(values[i][iName] || '').trim();
    if (!name) continue;

    var grade = iGrade >= 0 ? String(values[i][iGrade] || '').trim() : '';
    if (!grade) grade = inferGrade_(name);

    out.push({
      name: name,
      grade: grade,
      teacher: String(values[i][iTeacher] || '').trim(),
      email: String(values[i][iMail] || '').trim(),
      special: grade === '特別支援'
    });
  }

  cachePut_(CACHE_KEY.CLASSES, out, CACHE_SEC.CLASSES);
  return out;
}

/** 学年の並び順。数字の学年を先に、特別支援を最後に。 */
function gradeOrder_(grade) {
  var m = String(grade || '').match(/^(\d+)/);
  if (m) return Number(m[1]);
  if (grade === '特別支援') return 900;
  return 800;
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
 * 名簿のキャッシュ用に残している名前。
 * 一度キャッシュを入れたが、名簿を貼り付けた直後に予約できない時間が生まれ、
 * 読み取り回数も全体の5%程度だったため、素直に毎回読む形に戻した。
 */
function clearRosterCache_() {
  try {
    CacheService.getScriptCache().remove(CACHE_KEY.ROSTER);
  } catch (e) { /* 無視 */ }
}

/**
 * 各クラスの「予約表_〇組」シートのA・B列から生徒名簿を取得 [{cls, no, name}]
 */
function getRoster() {
  var hit = cacheGet_(CACHE_KEY.ROSTER);
  if (hit) return hit;

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

  cachePut_(CACHE_KEY.ROSTER, out, CACHE_SEC.ROSTER);
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

/**
 * シートの行数・列数が足りなければ広げる。
 * 新しいシートは既定 1000行×26列しかないため、クラス数が増えると
 * 「範囲が無効です」で落ちる。書き込む前に必ず通す。
 */
function ensureSheetSize_(sh, rows, cols) {
  var maxRows = sh.getMaxRows();
  if (rows > maxRows) sh.insertRowsAfter(maxRows, rows - maxRows);
  var maxCols = sh.getMaxColumns();
  if (cols > maxCols) sh.insertColumnsAfter(maxCols, cols - maxCols);
  return sh;
}

/**
 * 同じファイルの中の、別のシートへ飛ぶリンクの数式。
 *
 * クラスが17もあると、下のタブから目的の予約表を探すだけで手間がかかる。
 * 全体ビューの見出しと、各予約表の先頭に置いて行き来できるようにする。
 *
 * 数式の文字列リテラルには改行を書けないので、改行は CHAR(10) でつなぐ。
 *
 * @param {number} gid 飛び先のシートID（Sheet#getSheetId）
 * @param {string} label セルに出す文字。改行を含んでよい
 * @return {string} setValues / setFormula に渡す数式
 */
function sheetLinkFormula_(gid, label) {
  var parts = String(label == null ? '' : label).split('\n').map(function (t) {
    return '"' + t.replace(/"/g, '""') + '"';
  });
  return '=HYPERLINK("#gid=' + gid + '",' + parts.join('&CHAR(10)&') + ')';
}

/**
 * 自動で作り直される場所に、書き込む前の警告を付ける。
 *
 * 「予約表_〇組」の予約状況や「全体ビュー」は枠マスタから毎回作り直す表示用の場所で、
 * 手で書いても次の更新で消える。onEdit のトーストは15秒で流れてしまい、
 * 手引きを読まずに感覚で操作する担任には届かない。
 * Googleスプレッドシートの「保護」を**警告のみ**で掛け、書き込む瞬間に確認を出す。
 *
 * 警告のみなので、担任は［OK］を押せばそのまま書ける。作業が止まることはない。
 * 黄色い予備行のL〜N列は担任が直接入力する正規の欄なので、警告を掛けない。
 * それ以外の作り直される場所を守り、正しい操作に警告を出さないようにする。
 *
 * すでに同じ目印の保護があれば作り直さない（5分おきの更新から呼ばれるので、
 * 毎回付け替えると無駄な書き込みになる）。
 * 守る場所を変えたときは、いらなくなった目印付きの保護をここで片付ける。
 *
 * `range` を省いたときはシート全体を守る。全体ビューのように `clear()` で
 * 丸ごと作り直すシートは、範囲で持つと作り直しのたびに付け直しになりかねない。
 *
 * @param {Sheet} sh 対象シート
 * @param {Array<{key:string, range:(Range|undefined), note:string}>} specs 守る範囲
 */
function ensureEditGuards_(sh, specs) {
  try {
    var wholeSheet = false;
    for (var w = 0; w < specs.length; w++) if (!specs[w].range) wholeSheet = true;

    var list = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    if (wholeSheet) {
      list = list.concat(sh.getProtections(SpreadsheetApp.ProtectionType.SHEET));
    }

    var keep = {};

    for (var s = 0; s < specs.length; s++) {
      var prefix = EDIT_GUARD_TAG + '／' + specs[s].key;
      var found = null;
      for (var i = 0; i < list.length; i++) {
        if (String(list[i].getDescription() || '').indexOf(prefix) !== 0) continue;
        // 更新が同時に2つ走ると、同じ保護が二重に作られることがある。見つけたら片付ける
        if (!found) { found = list[i]; keep[i] = true; continue; }
        try { list[i].remove(); } catch (e2) { /* 消せなくても実害はない */ }
      }
      if (found) continue;

      (specs[s].range || sh).protect()
        .setDescription(prefix + '：' + specs[s].note)
        .setWarningOnly(true);
    }

    // 守る場所を変えた前の版が残っていると、いらない所に警告が出続ける
    for (var k = 0; k < list.length; k++) {
      if (keep[k]) continue;
      if (String(list[k].getDescription() || '').indexOf(EDIT_GUARD_TAG) !== 0) continue;
      try { list[k].remove(); } catch (e3) { /* 消せなくても実害はない */ }
    }
  } catch (e) {
    // 保護を付けられなくても、表そのものは正しく作れている。
    // 権限の都合で付けられない環境でも、更新が止まらないようにする
    console.warn('編集の警告を付けられませんでした: ' + sh.getName(), e);
  }
}

/**
 * シートまわりの整理。
 * 枠マスタだけを非表示にし、予約ログは確認用に表示のまま残したうえで、
 * タブの並び順を決められた順に戻す。
 */
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

  try {
    orderSheets();
  } catch (e) {
    console.warn('シートの並び替えをスキップ:', e);
  }

}
