/**
 * 初期セットアップ — 必要なシートを作成し、既定値を書き込む。
 * 既存シートは壊さない（無い列・無いキーだけ追記する）。
 * 「基礎データ」など元からあるシートには一切触れない。
 */

var DEFAULT_CONFIG_ROWS = [
  ['見出し', '三者面談 予約', '保護者の画面に出るタイトル'],
  ['案内文', '希望する日時を1つ選んでご予約ください。予約後に表示される4桁の予約コードは、変更・取消に必要です。必ず控えてください。', '保護者の画面に出る案内。改行可'],
  ['面談開始時刻', '13:40', '1日の最初の枠の開始時刻'],
  ['面談枠の長さ(分)', 15, '1件あたりの面談時間'],
  ['枠間の休憩(分)', 10, '面談と面談の間の休憩'],
  ['1日の枠数', 6, '13:40開始・15分面談・10分休憩なら6枠(最終 15:45-16:00)'],
  ['予約受付開始', '', '空欄なら即時開始。日時で入力(例 2026/10/01 9:00)'],
  ['予約受付締切', '', '空欄なら締切なし。日時で入力'],
  ['1家庭あたり予約可能数', 1, '通常は1'],
  ['氏名照合', true, 'TRUEなら出席番号に加えて生徒氏名の一致も確認する'],
  ['公開', false, 'TRUEにすると保護者の予約受付が始まる'],
  ['管理パスコード', '', '担任用Web管理画面のパスコード。空欄だとWeb管理画面は使えない']
];

var DEFAULT_DAYS = [
  ['2026-10-29', '', true],
  ['2026-10-30', '', true],
  ['2026-11-02', '', true],
  ['2026-11-03', '文化の日のため実施しない', false],
  ['2026-11-04', '', true],
  ['2026-11-05', '', true],
  ['2026-11-06', '', true]
];

var DEFAULT_CLASSES = [
  ['1組', '', ''],
  ['2組', '', ''],
  ['3組', '', ''],
  ['4組', '', '']
];

function setupSystem() {
  var ss = ss_();
  var created = [];

  created = created.concat(ensureConfigSheet_(ss));
  created = created.concat(ensureTableSheet_(ss, SH.DAYS, ['日付', '備考', '実施する'], DEFAULT_DAYS));
  created = created.concat(ensureTableSheet_(ss, SH.CLASSES, ['クラス', '担任名', '担任メール'], DEFAULT_CLASSES));
  created = created.concat(ensureTableSheet_(ss, SH.ROSTER, ['クラス', '出席番号', '生徒氏名', '備考'], []));
  created = created.concat(ensureTableSheet_(ss, SH.LOG, ['日時', '操作', '枠ID', 'クラス', '出席番号', '生徒氏名', '詳細'], []));
  created = created.concat(ensureSlotSheet_(ss));

  if (!ss.getSheetByName(SH.OVERVIEW)) {
    ss.insertSheet(SH.OVERVIEW);
    created.push(SH.OVERVIEW);
  }

  // 日付列を日付書式に
  var days = ss.getSheetByName(SH.DAYS);
  days.getRange(2, 1, Math.max(days.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy/mm/dd');

  return created;
}

function ensureConfigSheet_(ss) {
  var sh = ss.getSheetByName(SH.CONFIG);
  var created = [];
  if (!sh) {
    sh = ss.insertSheet(SH.CONFIG);
    created.push(SH.CONFIG);
    sh.getRange(1, 1, 1, 3).setValues([['キー', '値', '説明']]);
    styleHeader_(sh, 3);
  }
  // 既存の値は保持し、足りないキーだけ追記する
  var existing = {};
  var last = sh.getLastRow();
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) existing[String(vals[i][0] || '').trim()] = true;
  }
  var toAdd = DEFAULT_CONFIG_ROWS.filter(function (r) { return !existing[r[0]]; });
  if (toAdd.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 3).setValues(toAdd);
  }
  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 260);
  sh.setColumnWidth(3, 420);
  sh.setFrozenRows(1);
  return created;
}

function ensureTableSheet_(ss, name, header, defaults) {
  if (ss.getSheetByName(name)) return [];
  var sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  styleHeader_(sh, header.length);
  if (defaults && defaults.length) {
    sh.getRange(2, 1, defaults.length, defaults[0].length).setValues(defaults);
  }
  sh.setFrozenRows(1);
  return [name];
}

function ensureSlotSheet_(ss) {
  if (ss.getSheetByName(SH.SLOTS)) return [];
  var header = ['枠ID', '日付', '開始', '終了', 'クラス', '担任', '状態',
    '出席番号', '生徒氏名', '保護者氏名', '連絡事項', '予約コード', '予約日時'];
  var sh = ss.insertSheet(SH.SLOTS);
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  styleHeader_(sh, header.length);
  sh.setFrozenRows(1);
  sh.setColumnWidth(COL.SLOT_ID, 170);
  return [SH.SLOTS];
}

function styleHeader_(sh, cols) {
  sh.getRange(1, 1, 1, cols)
    .setFontWeight('bold')
    .setBackground('#e8eaed')
    .setVerticalAlignment('middle');
}

/**
 * テスト用ダミー生徒データ10名を「生徒名簿」シートに追加生成し、
 * 枠が未生成の場合は自動で枠も生成する。
 * @return {{count:number, slotsCreated:boolean}}
 */
function generateDummyRoster() {
  setupSystem();
  var sh = sheet_(SH.ROSTER);
  var dummyData = [
    ['1組', 1, '佐藤 勝利', 'ダミー生徒'],
    ['1組', 2, '鈴木 一朗', 'ダミー生徒'],
    ['1組', 3, '高橋 咲', 'ダミー生徒'],
    ['1組', 4, '田中 蓮', 'ダミー生徒'],
    ['1組', 5, '伊藤 結衣', 'ダミー生徒'],
    ['1組', 6, '渡辺 翔太', 'ダミー生徒'],
    ['1組', 7, '山本 凛', 'ダミー生徒'],
    ['1組', 8, '中村 陽翔', 'ダミー生徒'],
    ['1組', 9, '小林 葵', 'ダミー生徒'],
    ['1組', 10, '加藤 陸', 'ダミー生徒']
  ];
  
  var lastRow = sh.getLastRow();
  sh.getRange(lastRow + 1, 1, dummyData.length, 4).setValues(dummyData);

  var slotsCreated = false;
  var slotSh = sheet_(SH.SLOTS);
  if (slotSh.getLastRow() < 2) {
    try {
      generateSlots();
      rebuildOverview();
      rebuildClassSheets();
      slotsCreated = true;
    } catch (e) {
      console.warn('自動枠生成スキップ:', e);
    }
  }

  return { count: dummyData.length, slotsCreated: slotsCreated };
}
