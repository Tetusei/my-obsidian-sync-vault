/**
 * 保護者へ配る案内プリント（A4・1枚）をPDFで作る。
 *
 * QRコードはスクリプト内で生成したPNG画像をシートへ貼る。
 * 外部サービスにURLを送らずに済み、印刷でもかすれない。
 */

/**
 * QRのねらいの大きさ（px）。1マスの大きさは、この値をマス数で割って決める。
 * 大きすぎると紙面を圧迫し、小さすぎると読み取れない。
 * スマホでの読み取りは 250px 前後（印刷で約6.5cm）あれば十分に通る。
 */
var HANDOUT_QR_TARGET_PX = 340;

/**
 * 1マスの下限は8px。
 * スプレッドシートの列幅には下限があり、6pxを指定しても7pxになる。
 * 行の高さは指定どおりになるため、小さくするとマスが長方形になり、
 * 読み取り機が位置検出パターンの比率を測れず復号できなくなる。
 */
/** PDF出力時に列へ加算される幅（実測値） */
var HANDOUT_COL_RENDER_PAD = 1;

var HANDOUT_QR_MODULE_MIN = 8;
var HANDOUT_QR_MODULE_MAX = 10;

/** A4の印刷幅のめやす。列の数はここから決める（PDFは幅に合わせて等倍で縮む） */
var HANDOUT_PAGE_PX = 900;
var HANDOUT_QR_QUIET = 4;        // QRの周囲に必要な余白（規格で4マス以上）
var HANDOUT_QR_COL = 5;          // QRを描き始める列（余白の分だけ左に空きが要る）
var HANDOUT_QR_ROW = 18;         // QRを描き始める行
var HANDOUT_SPACER_PX = 10;      // 何も置かない「すき間」の行の高さ

/**
 * 案内プリントのPDFを作る。
 * @return {{fileName:string, fileUrl:string, folderName:string, folderUrl:string, url:string}}
 */
function exportHandoutPdf() {
  var cfg = getConfig();
  var url = verifiedHandoutUrl_(cfg);

  var ss = ss_();
  var temp = ss.insertSheet('__配布プリント__' + new Date().getTime());

  try {
    buildHandoutSheet_(temp, cfg, url);

    var folder = getOrCreatePdfFolder_();
    var fileName = '三者面談_保護者用案内_' +
      Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm') + '.pdf';

    var blob = handoutPdfBlob_(ss, temp, fileName);
    var file = folder.createFile(blob);

    return {
      fileName: file.getName(),
      fileUrl: file.getUrl(),
      folderName: PDF_FOLDER_NAME,
      folderUrl: folder.getUrl(),
      url: url
    };
  } finally {
    try { ss.deleteSheet(temp); } catch (e) { /* 無視 */ }
  }
}

/**
 * 案内プリントへ入れてよい保護者用URLを返す。
 *
 * ScriptApp.getService().getUrl() は、デプロイが複数あると実際に配るURLとは
 * 別のものを返すことがある。削除済みデプロイのURLを紙へ印刷すると回収できないため、
 * プリントだけは自動取得へフォールバックせず、設定シートの明示URLを必須にする。
 */
function verifiedHandoutUrl_(cfg) {
  var url = String(cfg && cfg.parentUrl || '').trim();
  if (!url) {
    throw new Error(
      '案内プリントは作成しませんでした。' + LF +
      '「' + SH.CONFIG + '」シートの「保護者用URL」が空欄です。' + LF +
      'Apps Script エディタの「デプロイ ▸ デプロイを管理」で、保護者へ配るURLを確認し、' + LF +
      '末尾が /exec のURLを設定してから、もう一度作成してください。');
  }

  if (!/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(url)) {
    throw new Error(
      '案内プリントは作成しませんでした。' + LF +
      '「' + SH.CONFIG + '」シートの「保護者用URL」の書き方が正しくありません。' + LF +
      'https://script.google.com/macros/s/……/exec の形で、末尾が /exec のURLを設定してください。' + LF +
      '現在の値: ' + url);
  }

  // URLをQRへ焼き込む前に、ログインしていない保護者と同じ条件で到達確認する。
  // 既存の判定を共用し、削除済みデプロイ・権限違い・ログイン要求を紙に残さない。
  if (typeof parentUrlReachability_ === 'function') {
    var reach = parentUrlReachability_();
    if (!reach || reach.level !== 'ok') {
      throw new Error(
        '案内プリントは作成しませんでした。' + LF +
        (reach && reach.message ? reach.message : '保護者用URLへ到達できることを確認できませんでした。') + LF +
        'URLを直した後、ブラウザのシークレットモードで開けることを確かめてから、もう一度作成してください。');
    }
  }

  return url;
}

/**
 * QRへ焼き込むURL。紙に印字する文字列とは分けてある。
 *
 * AndroidのGoogleアプリは、QRから開くときに複数アカウント用の /u/1/ を
 * パスへ挿入することがあり、Apps Scriptの公開URLではそれが404になる。
 * authuser=0 を明示するのは、挿入そのものを起こさせないための予防。
 * すでに挿入された後のURLはクエリでは救えない（/u/1/…?authuser=0 も404）ので、
 * これで必ず直るわけではない。
 * 版番号は、設定変更前のGoogleのエラーページが端末に残るのを防ぐために付ける。
 *
 * 印字するほうへは付けない。手で入力する経路では /u/1/ の挿入は起きないし、
 * 「読み取れない場合はこちらを入力してください」の宛先に20文字も足すのは本末転倒になる。
 */
function handoutQrUrl_(url) {
  return url + '?authuser=0&v=' +
    encodeURIComponent(String(VERSION || '').replace(/\s+/g, ''));
}

/** デプロイ済みのウェブアプリURL。未デプロイなら空文字。 */
/** プリントの中身を組み立てる */
function buildHandoutSheet_(sh, cfg, url) {
  var qr = qrMatrix_(handoutQrUrl_(url));
  var qrSize = qr.length;

  // 1マスの大きさ。QRのマス数が増えても、印刷される大きさが変わらないようにする
  var mod = Math.max(HANDOUT_QR_MODULE_MIN,
    Math.min(HANDOUT_QR_MODULE_MAX, Math.round(HANDOUT_QR_TARGET_PX / qrSize)));


  // QRのマス目を正方形にするため、列幅をそろえて細かく刻む。
  // 右側の説明文のぶんも含めて、A4の幅におさまる列数を用意する。
  var textCol = HANDOUT_QR_COL + qrSize + HANDOUT_QR_QUIET + 2;
  var width = Math.max(Math.ceil(HANDOUT_PAGE_PX / mod), textCol + 50);
  var height = HANDOUT_QR_ROW + qrSize + HANDOUT_QR_QUIET + 34;

  // 新しいシートは既定で26列しかないので、足りない分を足す
  if (sh.getMaxColumns() < width) {
    sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  }
  if (sh.getMaxRows() < height) {
    sh.insertRowsAfter(sh.getMaxRows(), height - sh.getMaxRows());
  }

  sh.getRange(1, 1, sh.getMaxRows(), width)
    .setBackground('#ffffff')
    .setFontFamily('Meiryo');

  for (var c = 1; c <= width; c++) sh.setColumnWidth(c, mod);

  // PDFに出力すると、列だけが指定より1px広く描かれる（行は指定どおり）。
  //   指定6px → 列7px・行6px ／ 指定8px → 列9px・行8px
  // getColumnWidth() は指定値をそのまま返すので、読み戻しでは気づけない。
  // マスが正方形でないと読み取り機が位置検出パターンの比率を測れず、
  // データが正しくても復号できないため、行の高さを1px足して合わせる。
  var cell = mod + HANDOUT_COL_RENDER_PAD;

  // QRの脇に置く文章は、行の高さが1マス分に固定される。
  // 文字を大きくするぶんは、またぐ行数で高さを確保する
  var line = Math.max(2, Math.ceil(26 / cell));

  // ── 見出し
  mergeText_(sh, 2, 2, 1, width - 2, cfg.title, {
    size: 28, bold: true, align: 'center', height: 54
  });
  mergeText_(sh, 3, 2, 1, width - 2, '保護者の皆さまへ', {
    size: 17, align: 'center', color: '#5f6368', height: 32
  });

  sh.getRange(4, 2, 1, width - 2).setBorder(null, null, true, null, null, null,
    '#1a73e8', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // ── 案内文
  var notice = cfg.notice ||
    '希望する日時を1つ選んでご予約ください。予約後に表示される4桁の予約コードは、変更・取消に必要です。必ず控えてください。';
  mergeText_(sh, 6, 2, 3, width - 2, notice, { size: 16, wrap: true, height: 32 });

  // ── 受付期間
  var period = [];
  if (cfg.openAt) period.push(Utilities.formatDate(cfg.openAt, TZ, 'M月d日(') +
    WEEKDAY_JA[cfg.openAt.getDay()] + ') ' + Utilities.formatDate(cfg.openAt, TZ, 'HH:mm') + ' から');
  if (cfg.closeAt) period.push(Utilities.formatDate(cfg.closeAt, TZ, 'M月d日(') +
    WEEKDAY_JA[cfg.closeAt.getDay()] + ') ' + Utilities.formatDate(cfg.closeAt, TZ, 'HH:mm') + ' まで');
  mergeText_(sh, 10, 2, 1, width - 2,
    period.length ? '受付期間　' + period.join('　〜　') : '受付期間　担任からのお知らせをご確認ください',
    { size: 17, bold: true, height: 38, bg: '#e8f0fe' });

  mergeText_(sh, 12, 2, 1, width - 2, '▼ スマートフォンで読み取ってください', {
    size: 17, bold: true, height: 34
  });

  // ── QRコード（セルの背景色で描く）
  drawQr_(sh, qr, HANDOUT_QR_ROW, HANDOUT_QR_COL, cell);
  placeQrImage_(sh, qr, HANDOUT_QR_ROW, HANDOUT_QR_COL, cell);

  var afterQr = HANDOUT_QR_ROW + qrSize + 1;

  // QRと同じ行を使うため、行の高さは1マス分(7px)に固定されている。
  // 文字を大きくするぶんは、またぐ行数を増やして確保する。
  var side = HANDOUT_QR_ROW + 1;
  mergeText_(sh, side, textCol, line, width - textCol,
    '読み取れない場合は、こちらのアドレスを入力してください',
    { size: 15, wrap: true, color: '#5f6368' });

  side += line + 1;
  mergeText_(sh, side, textCol, line * 5, width - textCol, url,
    { size: 14, wrap: true, bg: '#f1f3f4' });

  // ── 手順
  var steps = [
    '① クラス・出席番号・お子さまの氏名を入力',
    '② 空いている時間から希望の時間を選ぶ',
    '③ 保護者氏名を入力して予約を確定',
    '④ 表示された4桁の予約コードを控える'
  ];
  side += line * 5 + 1;
  mergeText_(sh, side, textCol, line, width - textCol, 'ご利用の手順',
    { size: 17, bold: true });

  side += line + 1;
  for (var s = 0; s < steps.length; s++) {
    mergeText_(sh, side + s * (line + 1), textCol, line, width - textCol, steps[s],
      { size: 16, wrap: true });
  }

  // ── ごきょうだいがいる場合
  //    QRの右側は行の高さを変えられない（変えるとQRが歪む）ので、
  //    まとまった文章はQRの下の全幅に置く。ここなら行の高さを自由にできる。
  // 脇の文章がQRの帯より下に伸びることがある。
  // その場合は、続きの見出しをさらに下へずらして重ならないようにする
  var sideEnd = side + (steps.length - 1) * (line + 1) + line;
  var sibRow = Math.max(afterQr + HANDOUT_QR_QUIET + 2, sideEnd + 2);
  mergeText_(sh, sibRow, 2, 1, width - 2,
    'ごきょうだいが本校にいる場合', { size: 17, bold: true, height: 36 });

  var sibLines = [
    '「きょうだいまとめて」を選ぶと、同じ日の続いた時間で全員分をまとめてご予約いただけます。',
    '学年 → クラスの順にお選びください。'
  ];
  // 特別支援学級のある学校だけ、学年の欄の選び方を添える
  if (hasSpecialClass_()) {
    sibLines.push('（特別支援学級のお子さまは、学年の欄で「特別支援」をお選びください）');
  }
  sibLines.push('予約コードは、ごきょうだい全員で同じものが1つ出ます。');

  // 行を1つの結合セルにまとめず、1文ずつ別のセルにする。
  // Googleスプレッドシートには行間の設定がないため、これで文と文の間を確保する。
  var sibBlockRows = 2;
  for (var sl = 0; sl < sibLines.length; sl++) {
    mergeText_(sh, sibRow + 2 + sl * sibBlockRows, 2, sibBlockRows, width - 2,
      sibLines[sl], { size: 15, wrap: true, height: 21 });
  }

  // ── AndroidでGoogleアカウントの切替に引っかかった場合
  // 「シークレットモード」とだけ書いても操作が伝わらないため、
  // 保護者が紙を見ながら進められる順番で案内する。
  var sibBodyEnd = sibRow + 2 + sibLines.length * sibBlockRows;
  var helpRow = sibBodyEnd + 2;
  mergeText_(sh, helpRow, 2, 1, width - 2,
    '予約画面が開かない場合', {
      size: 17, bold: true, height: 36, bg: '#fff4ce'
    });

  var androidHelpLines = [
    'Androidスマートフォンで「ファイルを開くことができません」と表示された場合は、次の操作をお試しください。',
    '1．画面上部のアドレス欄を押し、「コピー」のマークを押します。',
    '2．Chromeの右上にある「︙」を押し、「新しいシークレット タブ」を選びます。',
    '3．新しく開いた画面上部の入力欄を長押しして「貼り付け」を選び、画面を開きます。',
    '操作が難しい場合は、別のスマートフォンやタブレットでお試しください。',
    'それでも予約画面が開かない場合は、学校までご連絡ください。'
  ];
  // 最初の説明は長いため3行分、それ以外は2行分の高さを確保する。
  var androidHelpRows = [3, 2, 2, 2, 2, 2];
  var helpCursor = helpRow + 2;
  for (var ah = 0; ah < androidHelpLines.length; ah++) {
    var helpRows = androidHelpRows[ah] || 2;
    mergeText_(sh, helpCursor, 2, helpRows, width - 2,
      androidHelpLines[ah],
      { size: 15, wrap: true, height: 18, color: '#3c4043' });
    helpCursor += helpRows;
  }

  // ── 注意書き
  var notes = [
    '※ 予約画面には、予約できる時間だけが表示されます。',
    '※ 予約の変更・取り消しには、4桁の予約コードが必要です。控えをなくされた場合は担任までご連絡ください。'
  ];
  notes.push('※ ごきょうだいが同じクラスの場合や、続いた時間が取れない場合は、お一人ずつご予約ください。');

  var helpBodyEnd = helpCursor;
  var notesRow = helpBodyEnd + 2;
  var noteBlockRows = 3;
  for (var nt = 0; nt < notes.length; nt++) {
    mergeText_(sh, notesRow + nt * noteBlockRows, 2, noteBlockRows, width - 2,
      notes[nt],
      { size: 15, wrap: true, height: 16, color: '#3c4043' });
  }

  // ── すき間の行を詰める
  //    何も置いていない行は既定の21pxのままで、合計するとA4を超えてしまう。
  //    QRの帯（行の高さ＝1マス分）には触らないこと。触るとQRが歪んで読めなくなる。
  var spacers = [1, 4, 5, 9, 11, 13];
  for (var sp = Math.max(afterQr + HANDOUT_QR_QUIET - 1, sideEnd + 1); sp < sibRow; sp++) spacers.push(sp);
  spacers.push(sibRow + 1);
  for (var sp2 = sibBodyEnd; sp2 < helpRow; sp2++) spacers.push(sp2);
  spacers.push(helpRow + 1);
  for (var sp3 = helpBodyEnd; sp3 < notesRow; sp3++) spacers.push(sp3);
  for (var sp4 = 0; sp4 < spacers.length; sp4++) {
    try { sh.setRowHeight(spacers[sp4], HANDOUT_SPACER_PX); } catch (e) { /* 無視 */ }
  }

  sh.setHiddenGridlines(true);
  try { sh.setFrozenRows(0); } catch (e) { /* 無視 */ }
}

/** 特別支援学級が登録されているか */
function hasSpecialClass_() {
  try {
    var classes = getClasses();
    for (var i = 0; i < classes.length; i++) {
      if (classes[i].special) return true;
    }
  } catch (e) { /* 読めないときは触れない */ }
  return false;
}

/** セルを結合して文字を置く */
function mergeText_(sh, row, col, rows, cols, text, opt) {
  opt = opt || {};
  var range = sh.getRange(row, col, rows, cols);
  range.merge().setValue(text);
  range.setFontSize(opt.size || 11);
  range.setFontWeight(opt.bold ? 'bold' : 'normal');
  range.setHorizontalAlignment(opt.align || 'left');
  range.setVerticalAlignment('middle');
  range.setWrap(!!opt.wrap);
  if (opt.color) range.setFontColor(opt.color);
  if (opt.bg) range.setBackground(opt.bg);
  if (opt.height) {
    for (var r = 0; r < rows; r++) sh.setRowHeight(row + r, opt.height);
  }
  return range;
}

/** QRのマトリクスを、セルの背景色として描く */
function drawQr_(sh, qr, top, left, mod) {
  var size = qr.length;
  var quiet = HANDOUT_QR_QUIET;
  var px = mod || HANDOUT_QR_MODULE_MAX;

  // 余白がシートの外に出ないことを確かめる（外に出ると実行時エラーになる）
  if (top - quiet < 1 || left - quiet < 1) {
    throw new Error('QRコードの描画位置が上端・左端に近すぎます。');
  }

  for (var r = 0; r < size + quiet * 2; r++) {
    sh.setRowHeight(top - quiet + r, px);
  }

  // 背景は白のまま。QRそのものは画像として上に貼る（placeQrImage_）。
  // セルの背景色で描くと、PDF出力で黒マスが1pxはみ出し、
  // 読み取り機が倍率によって復号できなくなる。
  sh.getRange(top - quiet, left - quiet, size + quiet * 2, size + quiet * 2)
    .setBackground('#ffffff');
}

/**
 * QRを画像としてシートに貼る。
 * セルの塗り分けと違い、輪郭が正確に出るので確実に読み取れる。
 */
function placeQrImage_(sh, qr, top, left, cell) {
  var quiet = HANDOUT_QR_QUIET;
  var size = qr.length;
  var side = (size + quiet * 2) * cell;

  // 貼る大きさより十分に細かく作る。粗いまま引き伸ばすと輪郭がぼけ、
  // 印刷や画面の倍率によっては読み取れなくなる。
  // ただし insertImage には 100万画素・2MB の上限がある。
  // 一辺950画素までに収まる範囲で、いちばん細かい倍率を選ぶ。
  var span = size + quiet * 2;
  var px = Math.max(4, Math.floor(950 / span));
  var blob = qrPngBlob_(qr, px, quiet, 'qr.png');
  var img = sh.insertImage(blob, left - quiet, top - quiet, 0, 0);
  img.setWidth(side).setHeight(side);
  return img;
}

/** 一時シートだけをPDFにする */
function handoutPdfBlob_(ss, sheet, fileName) {
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' + [
    'format=pdf',
    'gid=' + sheet.getSheetId(),
    'portrait=true',
    'size=A4',
    'fitw=true',
    'gridlines=false',
    'printtitle=false',
    'sheetnames=false',
    'pagenum=UNDEFINED',
    'attachment=true',
    'top_margin=0.4',
    'bottom_margin=0.4',
    'left_margin=0.4',
    'right_margin=0.4'
  ].join('&');

  var token = ScriptApp.getOAuthToken();
  var res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  return res.getBlob().setName(fileName);
}
