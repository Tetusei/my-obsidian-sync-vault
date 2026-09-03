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
  [RESERVE_COUNT_KEY, 0, '保護者には見せない予備の枠。最終コマの後ろに作られ、担任が管理画面か「この枠に生徒を入れる」で埋める。0で無し'],
  ['予約受付開始', '', '空欄なら即時開始。日時で入力(例 2026/10/01 9:00)'],
  ['きょうだい予約の受付開始', '', '【空欄のままにすること】きょうだい世帯だけ先に受け付ける機能。先に受け付けると他の家庭から不公平に見えるため使わない'],
  ['予約受付締切', '', '空欄なら締切なし。日時で入力'],
  [MAX_PER_STUDENT_KEY, 1, '同じ生徒が同時に持てる予約の数。通常は1（きょうだいは別々に数えるので1のままでよい）'],
  ['氏名照合', true, 'TRUEなら出席番号に加えて生徒氏名の一致も確認する'],
  ['担任メール通知', false, 'TRUEにすると、予約・変更・取消のたびに「クラス」シートの担任メールへ通知が届く'],
  ['リマインド日数(締切の何日前)', 3, '自動リマインドを有効にしたとき、締切の何日前に未予約者を担任へ知らせるか'],
  ['公開', false, 'TRUEにすると保護者の予約受付が始まる'],
  ['管理パスコード', '', '担任用Web管理画面のパスコード。空欄だとWeb管理画面は使えない'],
  ['保護者用URL', '', '実際に保護者へ配るURL(末尾 /exec)。空欄なら自動取得するが、デプロイが複数あると別のURLになることがある']
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
  ['1組', '', '', ''],
  ['2組', '', '', ''],
  ['3組', '', '', ''],
  ['4組', '', '', '']
];

function setupSystem() {
  dropRefCaches_();

  // 「設定」シートの自動処理の行を、そのまま入力として扱えるようにする
  try {
    ensureConfigWatcher();
  } catch (e) {
    console.warn('設定シートの見張り役を取り付けられませんでした:', e);
  }

  // キー名の付け替えと、行の並べ直し
  try {
    orderConfigRows();
  } catch (e) {
    console.warn('設定シートの行の並び替えをスキップ:', e);
  }

  var ss = ss_();
  var created = [];

  created = created.concat(ensureConfigSheet_(ss));
  created = created.concat(ensureTableSheet_(ss, SH.DAYS, ['日付', '備考', '実施する'], DEFAULT_DAYS));
  created = created.concat(ensureTableSheet_(ss, SH.CLASSES, ['クラス', '担任名', '担任メール', '学年'], DEFAULT_CLASSES));
  ensureGradeColumn_(ss);
  created = created.concat(ensureTableSheet_(ss, SH.LOG, ['日時', '操作', '枠ID', 'クラス', '出席番号', '生徒氏名', '詳細'], []));
  created = created.concat(ensureSlotSheet_(ss));

  if (!ss.getSheetByName(SH.OVERVIEW)) {
    ss.insertSheet(SH.OVERVIEW);
    created.push(SH.OVERVIEW);
  }

  // 旧「生徒名簿」シートが残っている場合はクリーンアップ削除
  var oldRoster = ss.getSheetByName('生徒名簿');
  if (oldRoster) {
    try { ss.deleteSheet(oldRoster); } catch (e) { /* 無視 */ }
  }

  // 各クラスの「予約表_〇組」シートを準備・初期作成
  var classes = getClasses();
  for (var c = 0; c < classes.length; c++) {
    var name = CLASS_SHEET_PREFIX + classes[c].name;
    if (!ss.getSheetByName(name)) {
      createClassSheet_(ss, name);
      created.push(name);
    }
  }

  // 「交流学級」シート（特別支援学級と通常学級の在籍を結びつける）
  if (ensureLinkSheet_()) created.push(SH.LINK);

  // 「だめなコマ」シート（担任が面談を入れられない枠を指定する）
  var hadNg = !!ss.getSheetByName(SH.NG);
  try {
    rebuildNgSheet();
    if (!hadNg && ss.getSheetByName(SH.NG)) created.push(SH.NG);
  } catch (e) {
    console.warn('だめなコマシートの作成をスキップ:', e);
  }

  // 日付列を日付書式に
  var days = ss.getSheetByName(SH.DAYS);
  days.getRange(2, 1, Math.max(days.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy/mm/dd');

  // システム用内部シート（枠マスタ・予約ログ）を非表示化
  hideInternalSheets();

  return created;
}

/**
 * タブを左から並べる順番。
 * ここに無いシート（「基礎データ」など元からあるもの）は動かさず、後ろに残す。
 */
function sheetOrder_() {
  var order = [SH.CONFIG, SH.DAYS, SH.CLASSES, SH.LINK, SH.NG, SH.LOG, SH.OVERVIEW];

  var classes = [];
  try { classes = getClasses(); } catch (e) { classes = []; }
  for (var i = 0; i < classes.length; i++) {
    order.push(CLASS_SHEET_PREFIX + classes[i].name);
  }

  order.push(SH.SLOTS);   // 内部用。非表示なのでいちばん後ろ
  return order;
}

/**
 * シートを決められた順に並べ替える。
 * @return {number} 並べ替えの対象になったシート数
 */
/**
 * 「設定」シートの行を CONFIG_ORDER の順に並べ直す。
 *
 * 値だけを書き換えると TRUE/FALSE のプルダウンや色が行に残ったままずれるので、
 * moveRows で行ごと動かす。必要な行を上から順に引き上げるだけなので、
 * 移動は常に上向きになる。
 *
 * @return {number} 動かした行数
 */
function orderConfigRows() {
  var sh = ss_().getSheetByName(SH.CONFIG);
  if (!sh) return 0;

  // 並び順は新しいキー名で書いてあるので、先に名前を直しておく
  try {
    migrateConfigKeys_();
  } catch (e) {
    console.warn('設定シートのキー名の付け替えをスキップ:', e);
  }

  var moved = 0;
  var target = 2; // 1行目は見出し

  for (var i = 0; i < CONFIG_ORDER.length; i++) {
    var last = sh.getLastRow();
    if (target >= last) break;

    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    var at = -1;
    for (var r = 0; r < keys.length; r++) {
      if (String(keys[r][0] || '').trim() === CONFIG_ORDER[i]) { at = r + 2; break; }
    }
    if (at < 0) continue; // その行がまだ無い

    if (at !== target) {
      sh.moveRows(sh.getRange(at, 1), target);
      moved++;
    }
    target++;
  }
  return moved;
}

function orderSheets() {
  var ss = ss_();
  var order = sheetOrder_();

  var active = null;
  try { active = ss.getActiveSheet(); } catch (e) { active = null; }

  var pos = 1;
  for (var i = 0; i < order.length; i++) {
    var sh = ss.getSheetByName(order[i]);
    if (!sh) continue;
    moveSheetTo_(ss, sh, pos);
    pos++;
  }

  // 並べ替えでアクティブシートが変わるので、元のシートに戻す
  if (active) {
    try { ss.setActiveSheet(active); } catch (e) { /* 非表示になった場合は戻せない */ }
  }
  return pos - 1;
}

/** シート1枚を指定位置へ動かす。非表示のものは一時的に表示してから戻す。 */
function moveSheetTo_(ss, sh, pos) {
  if (sh.getIndex() === pos) return;

  var wasHidden = sh.isSheetHidden();
  try {
    if (wasHidden) sh.showSheet();
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(pos);
  } catch (e) {
    console.warn('シートの移動をスキップ: ' + sh.getName(), e);
  }
  if (wasHidden) {
    try { sh.hideSheet(); } catch (e2) { /* 無視 */ }
  }
}

/** 「予約表_〇組」シートを見出しつきで作る */
function createClassSheet_(ss, name) {
  var sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, CLASS_HEADER_LEFT.length).setValues([CLASS_HEADER_LEFT])
    .setFontWeight('bold').setBackground('#d9ead3');
  sh.getRange(1, 9, 1, CLASS_HEADER_RIGHT.length).setValues([CLASS_HEADER_RIGHT])
    .setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);
  return sh;
}

/** 名簿シートに入っている生徒数を数える */
function countStudentsInSheet_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, 1, last - 1, 2).getValues();
  var n = 0;
  for (var i = 0; i < vals.length; i++) {
    if (Number(vals[i][0]) && String(vals[i][1] || '').trim()) n++;
  }
  return n;
}

/**
 * 「クラス」シートの内容に合わせて、名簿シートと面談枠をそろえる。
 *
 * クラス名は枠IDにも名簿シート名にも使われているため、増減や改名は
 * 「クラス」シートを直すだけでは反映されない。この関数がその橋渡しをする。
 * 使われなくなった名簿シートは、中の個人情報を勝手に消さないよう報告だけする。
 *
 * @return {{created:Array<string>, orphans:Array<Object>, slots:number, kept:number, ngConflicts:Array<Object>}}
 */
function syncClasses() {
  // 「クラス」シートを編集した直後に実行される操作なので、
  // 直前の一覧をキャッシュから読まないよう先に捨てる。
  dropRefCaches_();
  var ss = ss_();
  var classes = getClasses();
  if (!classes.length) throw new Error('「' + SH.CLASSES + '」シートにクラス名を入力してください。');

  var created = [];
  var known = {};
  for (var c = 0; c < classes.length; c++) {
    known[classes[c].name] = true;
    var name = CLASS_SHEET_PREFIX + classes[c].name;
    if (!ss.getSheetByName(name)) {
      createClassSheet_(ss, name);
      created.push(name);
    }
  }

  // クラス一覧に無い名簿シート（減らした・名前を変えた場合に残る）
  var orphans = [];
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sname = sheets[i].getName();
    if (sname.indexOf(CLASS_SHEET_PREFIX) !== 0) continue;
    var cls = sname.slice(CLASS_SHEET_PREFIX.length);
    if (known[cls]) continue;
    orphans.push({ name: sname, cls: cls, students: countStudentsInSheet_(sheets[i]) });
  }

  clearRosterCache_();
  var gen = generateSlots();
  rebuildOverview();
  rebuildClassSheets();
  hideInternalSheets();

  logAction_('クラス増減の反映', '', '', '', '',
    'クラス ' + classes.length + '／シート作成 ' + created.length +
    '／未使用シート ' + orphans.length);

  return {
    created: created,
    orphans: orphans,
    slots: gen.written,
    kept: gen.kept,
    ngConflicts: gen.ngConflicts || []
  };
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

/**
 * 「クラス」シートに「学年」列が無ければ、末尾に足してクラス名から推定して埋める。
 * 既存の列は動かさない（順番を変えると入力済みの内容がずれるため）。
 */
function ensureGradeColumn_(ss) {
  var sh = ss.getSheetByName(SH.CLASSES);
  if (!sh) return false;

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headerIndex_(header, ['学年']) >= 0) return false;

  var col = lastCol + 1;
  ensureSheetSize_(sh, Math.max(sh.getLastRow(), 2), col);
  sh.getRange(1, col).setValue('学年').setFontWeight('bold').setBackground('#e8eaed');

  var last = sh.getLastRow();
  if (last >= 2) {
    var names = sh.getRange(2, 1, last - 1, 1).getValues();
    var grades = [];
    for (var i = 0; i < names.length; i++) {
      grades.push([inferGrade_(names[i][0])]);
    }
    sh.getRange(2, col, grades.length, 1).setValues(grades);
  }
  return true;
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
 * 各クラス（1組:35名、2組:35名、3組:35名、4組:35名）のダミー生徒データを
 * 各クラスの「予約表_〇組」シートのA・B列に自動生成・書き込みする。
 * @return {{count:number, slotsCreated:boolean}}
 */
function generateDummyRoster() {
  setupSystem();
  var ss = ss_();
  var classes = getClasses();

  var familyNames = [
    '佐藤', '鈴木', '高橋', '田中', '伊藤', '渡辺', '山本', '中村', '小林', '加藤',
    '吉田', '山田', '佐々木', '山口', '松本', '井上', '木村', '林', '斎藤', '清水',
    '山崎', '森', '池田', '橋本', '阿部', '石川', '山下', '中川', '中島', '前田',
    '藤田', '小川', '岡田', '後藤', '長谷川'
  ];

  var firstNames = [
    '勝利', '一朗', '咲', '蓮', '結衣', '翔太', '凛', '陽翔', '葵', '陸',
    '悠斗', '陽菜', '奏太', '莉子', '大翔', '結菜', '颯太', '芽依', '樹', '咲良',
    '蒼', '杏', '湊', '心春', '大和', '楓', '新', '紬', '暖', '澪',
    '瑛太', '詩', '律', '花', '朝陽'
  ];

  var totalAdded = 0;

  for (var c = 0; c < classes.length; c++) {
    var clsName = classes[c].name;
    var sheetName = '予約表_' + clsName;
    var sh = ss.getSheetByName(sheetName);
    if (!sh) sh = createClassSheet_(ss, sheetName);

    var dummyList = [];
    for (var i = 1; i <= 35; i++) {
      var fn = familyNames[(i - 1 + c * 3) % familyNames.length];
      var gn = firstNames[(i - 1 + c * 7) % firstNames.length];
      dummyList.push([i, fn + ' ' + gn]);
    }

    sh.getRange(2, 1, dummyList.length, 2).setValues(dummyList);
    totalAdded += dummyList.length;
    clearRosterCache_();
  }

  var slotsCreated = false;
  var slotSh = sheet_(SH.SLOTS);
  if (slotSh.getLastRow() < 2) {
    try {
      generateSlots();
      slotsCreated = true;
    } catch (e) {
      console.warn('自動枠生成スキップ:', e);
    }
  }

  rebuildOverview();
  rebuildClassSheets();

  // 内部シートを非表示化
  hideInternalSheets();

  return { count: totalAdded, slotsCreated: slotsCreated };
}
