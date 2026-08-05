/**
 * 備品台帳写真撮影システム - バックエンドスクリプト (Code.js)
 *
 * GoogleスプレッドシートおよびGoogleドライブと連携し、
 * ウェブフォームからのデータ取得および写真アップロード要求を処理します。
 *
 * 列レイアウト（個別シート 18列）:
 *   A(0):検査済  B(1):No  C(2):登載番号  D(3):台帳番号  E(4):台帳分類
 *   F(5):取得年月日  G(6):品名番号  H(7):番号  I(8):品　　名  J(9):規　　格
 *   K(10):数量  L(11):単価  M(12):納品業者  N(13):保管場所等  O(14):理振・算振
 *   P(15):特認事項欄  Q(16):QRコード  R(17):写真のリンク
 */

// ============================================================
// 個別シートの統一ヘッダー定義 (18列)
// ============================================================
const TARGET_HEADERS = [
  "検査済", "No", "登載番号", "台帳番号", "台帳分類", "取得年月日",
  "品名番号", "番号", "品\u3000\u3000名", "規\u3000\u3000格",
  "数量", "単価", "納品業者", "保管場所等", "理振・算振",
  "特認事項欄", "QRコード", "写真のリンク"
];
const TARGET_COL_COUNT = TARGET_HEADERS.length; // 18

// ============================================================
// doGet: Webアプリ起動
// ============================================================
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.queryParamId  = (e && e.parameter && e.parameter.id)  || '';
  template.queryParamRow = (e && e.parameter && e.parameter.row) || '';
  return template.evaluate()
    .setTitle('備品台帳 点検・撮影')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// onOpen: カスタムメニュー追加
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 備品台帳 点検・撮影')
    .addItem('📱 点検・撮影アプリを開く', 'showWebAppLink')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 データ管理')
      .addItem('🔄 各シートを一括同期（データ更新）',   'syncAllIndividualSheets')
      .addItem('📐 列幅を１算数からコピー',             'applyColumnWidthsFromSansuu')
      .addItem('🔳 QRコードを一括作成（マスタ用）',     'generateQRCodesForMaster')
    )
    .addSubMenu(ui.createMenu('☑️ 検査チェック')
      .addItem('☑️ チェックボックスを全シートに追加',   'addCheckboxesToAllSheets')
      .addItem('🔲 チェックを全てリセット（一括オフ）', 'resetAllCheckboxes')
    )
    .addSeparator()
    .addItem('⚙️ アプリ設定',      'showSheetSettings')
    .addItem('💾 バックアップ作成', 'createSpreadsheetBackup')
    .addToUi();
}

// ============================================================
// onEdit: チェックボックス変更をリアルタイムで対向シートへ同期
// ============================================================
function onEdit(e) {
  if (!e) return;
  const col   = e.range.getColumn();
  const row   = e.range.getRow();
  const sheet = e.range.getSheet();
  const ss    = e.source; // simple trigger では e.source を使用

  // A列（検査済チェック列）の変更のみ処理
  if (col !== 1) return;

  const { headerRow, startRow } = findHeaderAndStartRow(sheet);
  if (row < startRow) return; // ヘッダー行以上はスキップ

  // A列ヘッダーが「検査済」でない場合はスキップ
  const firstHeader = String(sheet.getRange(headerRow, 1).getValue()).trim().replace(/\s+/g, '');
  if (firstHeader !== '検査済') return;

  const isChecked = (e.value === 'TRUE' || e.value === true || e.value === 'true');
  const masterSheetName = getTargetSheetName();

  // 登載番号列インデックスを動的に取得
  const lastCol  = Math.min(sheet.getLastColumn(), 20);
  const headers  = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const idColIdx = getMasterIdColumnIdx(headers);
  if (idColIdx === -1) return;

  const masterId = String(sheet.getRange(row, idColIdx + 1).getValue()).trim();
  if (!masterId) return;

  if (sheet.getName() === masterSheetName) {
    // マスタ → 個別シート へ伝播
    _syncCheckboxToIndividual(ss, masterId, isChecked, masterSheetName);
  } else {
    // 個別シート → マスタ へ伝播
    _syncCheckboxToMasterSheet(ss, masterId, isChecked, masterSheetName);
  }
}

// ============================================================
// ヘルパー: ヘッダー配列から「登載番号」列の 0-indexed インデックスを返す
// ============================================================
function getMasterIdColumnIdx(headers) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i]).trim().replace(/\s+/g, '');
    if (h === '登載番号') return i;
  }
  return -1;
}

// ============================================================
// ヘルパー: 個別シート → マスタ へ 1 件分のチェック状態を書き込む
//           (onEdit 用: e.source を受け取る)
// ============================================================
function _syncCheckboxToMasterSheet(ss, masterId, isChecked, masterSheetName) {
  const masterSheet = ss.getSheetByName(masterSheetName);
  if (!masterSheet) return;
  _writeCheckboxByMasterId(masterSheet, masterId, isChecked);
}

// ============================================================
// ヘルパー: マスタ → 全個別シート へ 1 件分のチェック状態を書き込む
//           (onEdit 用: e.source を受け取る)
// ============================================================
function _syncCheckboxToIndividual(ss, masterId, isChecked, masterSheetName) {
  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    if (sheet.getName() === masterSheetName) continue;
    const { headerRow } = findHeaderAndStartRow(sheet);
    const fh = String(sheet.getRange(headerRow, 1).getValue()).trim().replace(/\s+/g, '');
    if (fh !== '検査済') continue;
    _writeCheckboxByMasterId(sheet, masterId, isChecked);
  }
}

// ============================================================
// ヘルパー: 指定シートの「登載番号」が masterId と一致する行のA列に isChecked を書く
// ============================================================
function _writeCheckboxByMasterId(sheet, masterId, isChecked) {
  const { startRow } = findHeaderAndStartRow(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const { headerRow } = findHeaderAndStartRow(sheet);
  const lastCol  = Math.min(sheet.getLastColumn(), 20);
  const headers  = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const idColIdx = getMasterIdColumnIdx(headers);
  if (idColIdx === -1) return;

  const numRows  = lastRow - startRow + 1;
  const ids      = sheet.getRange(startRow, idColIdx + 1, numRows, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === masterId) {
      sheet.getRange(startRow + i, 1).setValue(isChecked);
      break;
    }
  }
}

// ============================================================
// Webアプリ用: 特定備品の検査済チェック状態を取得
// ============================================================
function getCheckboxState(masterId) {
  try {
    const ss          = getSpreadsheet();
    const masterSheet = ss.getSheetByName(getTargetSheetName());
    if (!masterSheet) return false;

    const { headerRow, startRow } = findHeaderAndStartRow(masterSheet);
    const lastCol = Math.min(masterSheet.getLastColumn(), 20);
    const headers = masterSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

    const fh = String(headers[0]).trim().replace(/\s+/g, '');
    if (fh !== '検査済') return false;

    const idColIdx = getMasterIdColumnIdx(headers);
    if (idColIdx === -1) return false;

    const lastRow = masterSheet.getLastRow();
    if (lastRow < startRow) return false;

    const numRows = lastRow - startRow + 1;
    const data    = masterSheet.getRange(startRow, 1, numRows, Math.max(idColIdx + 1, 1)).getValues();
    for (const row of data) {
      if (String(row[idColIdx]).trim() === String(masterId).trim()) {
        return row[0] === true;
      }
    }
    return false;
  } catch(e) {
    Logger.log('getCheckboxState エラー: ' + e.toString());
    return false;
  }
}

// ============================================================
// Webアプリ用: 特定備品の検査済チェック状態をマスタ＋全個別シートへ書き込む
// ============================================================
function setCheckboxState(masterId, isChecked) {
  try {
    const ss              = getSpreadsheet();
    const masterSheetName = getTargetSheetName();
    const masterSheet     = ss.getSheetByName(masterSheetName);

    if (masterSheet) {
      _writeCheckboxByMasterId(masterSheet, String(masterId).trim(), isChecked);
    }

    const sheets = ss.getSheets();
    for (const sheet of sheets) {
      if (sheet.getName() === masterSheetName) continue;
      const { headerRow } = findHeaderAndStartRow(sheet);
      const fh = String(sheet.getRange(headerRow, 1).getValue()).trim().replace(/\s+/g, '');
      if (fh !== '検査済') continue;
      _writeCheckboxByMasterId(sheet, String(masterId).trim(), isChecked);
    }

    return { success: true, isChecked: isChecked };
  } catch(e) {
    Logger.log('setCheckboxState エラー: ' + e.toString());
    throw new Error('チェック状態の保存に失敗しました: ' + e.message);
  }
}

// ============================================================
// 全シートにチェックボックス列を追加する
// ============================================================
function addCheckboxesToAllSheets() {
  try {
    const ss              = getSpreadsheet();
    const masterSheetName = getTargetSheetName();
    const sheets          = ss.getSheets();
    let addedCount = 0, skippedCount = 0;

    for (const sheet of sheets) {
      const { headerRow, startRow } = findHeaderAndStartRow(sheet);
      if (sheet.getLastColumn() < 1) continue;

      const firstHeader = String(sheet.getRange(headerRow, 1).getValue()).trim().replace(/\s+/g, '');

      if (firstHeader === '検査済') {
        // 既存チェックボックスのデータ検証を確実に設定
        const lastRow = sheet.getLastRow();
        if (lastRow >= startRow) {
          const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
          sheet.getRange(startRow, 1, lastRow - startRow + 1, 1).setDataValidation(rule);
        }
        skippedCount++;
        continue;
      }

      // チェックボックスを追加してよい既知シートかどうかを判定
      const isMaster = (sheet.getName() === masterSheetName);
      const knownHeaders = ['No','一連番号','登載番号','台帳番号','品名番号','教科番号','マスタ'];
      const isKnown = knownHeaders.some(k => k === firstHeader);
      if (!isMaster && !isKnown) { skippedCount++; continue; }

      // 列Aの前に新規列を挿入して「検査済」ヘッダーをセット
      sheet.insertColumnBefore(1);
      sheet.getRange(headerRow, 1).setValue('検査済');

      // データ行にチェックボックスを挿入
      const lastRow = sheet.getLastRow();
      if (lastRow >= startRow) {
        const cbRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, 1);
        cbRange.setValue(false);
        cbRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
      }
      addedCount++;
    }

    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert(
      `チェックボックスの追加が完了しました。\n\n` +
      `・追加したシート数: ${addedCount} シート\n` +
      `・スキップ（設定済み等）: ${skippedCount} シート\n\n` +
      `続けて「🔄 各台帳シートを一括同期」を実行すると全シートのレイアウトが最新化されます。`
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('チェックボックスの追加中にエラー: ' + e.message);
  }
}

// ============================================================
// 全シートのチェックボックスを一括でオフにする
// ============================================================
function resetAllCheckboxes() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '確認',
    '全シートのチェックボックスを一括でオフにしますか？\nこの操作は元に戻せません。',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    const ss     = getSpreadsheet();
    const sheets = ss.getSheets();
    let resetCount = 0;

    for (const sheet of sheets) {
      const { headerRow, startRow } = findHeaderAndStartRow(sheet);
      const fh = String(sheet.getRange(headerRow, 1).getValue()).trim().replace(/\s+/g, '');
      if (fh !== '検査済') continue;

      const lastRow = sheet.getLastRow();
      if (lastRow < startRow) continue;

      const numRows    = lastRow - startRow + 1;
      const falseVals  = Array.from({ length: numRows }, () => [false]);
      sheet.getRange(startRow, 1, numRows, 1).setValues(falseVals);
      resetCount++;
    }

    SpreadsheetApp.flush();
    ui.alert(`完了しました。${resetCount} シートのチェックをオフにしました。`);
  } catch(e) {
    ui.alert('一括オフ処理中にエラー: ' + e.message);
  }
}

// ============================================================
// showWebAppLink
// ============================================================
function showWebAppLink() {
  const props    = PropertiesService.getScriptProperties();
  const savedUrl = props.getProperty('CUSTOM_WEBAPP_URL');
  let autoUrl = '';
  try { autoUrl = ScriptApp.getService().getUrl(); } catch(e) {}

  let url = savedUrl || autoUrl;
  try {
    const email = Session.getActiveUser().getEmail();
    if (email && url) {
      const sep = url.indexOf('?') === -1 ? '?' : '&';
      url = url + sep + 'authuser=' + encodeURIComponent(email);
    }
  } catch(e) {}

  if (!url) {
    SpreadsheetApp.getUi().alert(
      'ウェブアプリのURLを取得できませんでした。\n\n' +
      '「⚙️ アプリの設定」からデプロイURLを手動で登録してください。'
    );
    return;
  }

  const html = `
    <div style="font-family:sans-serif;padding:16px;color:#1e293b;text-align:center;">
      <p style="font-weight:bold;margin-bottom:8px;font-size:1.05rem;">備品撮影用ウェブアプリを開く</p>
      <a href="${url}" target="_blank" rel="noopener noreferrer"
         style="display:inline-block;background:#6366f1;color:white;padding:12px 24px;text-decoration:none;border-radius:8px;font-weight:bold;margin-bottom:16px;box-shadow:0 4px 12px rgba(99,102,241,.3);font-size:.95rem;">
        アプリ画面を開く ↗
      </a>
      <div style="margin-top:10px;text-align:left;">
        <p style="font-size:.8rem;color:#64748b;margin-bottom:4px;">アプリのURL（スマートフォン等への共有用）:</p>
        <input type="text" value="${url}" readonly
               style="width:100%;padding:8px;border-radius:6px;border:1px solid #cbd5e1;font-size:.8rem;text-align:center;background:#f8fafc;"
               onclick="this.select()">
      </div>
    </div>`;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(450).setHeight(240),
    '備品撮影ウェブアプリ'
  );
}

// ============================================================
// showSheetSettings
// ============================================================
function showSheetSettings() {
  const ss              = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet    = getTargetSheetName();
  const savedUrl        = PropertiesService.getScriptProperties().getProperty('CUSTOM_WEBAPP_URL') || '';
  let autoUrl = '';
  try { autoUrl = ScriptApp.getService().getUrl() || ''; } catch(e) {}
  const displayUrl = savedUrl || autoUrl;

  const optionsHtml = ss.getSheets().map(s => {
    const n = s.getName();
    const sel = n === currentSheet ? 'selected' : '';
    return `<option value="${escapeHtml(n)}" ${sel}>${escapeHtml(n)}</option>`;
  }).join('');

  const html = `
    <div style="font-family:sans-serif;padding:12px;color:#1e293b;font-size:.9rem;">
      <div style="margin-bottom:12px;">
        <label style="font-weight:bold;display:block;margin-bottom:6px;">1. 対象スプレッドシート名</label>
        <select id="sheetSelect" style="width:100%;padding:8px;border-radius:6px;border:1px solid #cbd5e1;font-size:.9rem;">
          ${optionsHtml}
        </select>
      </div>
      <div style="margin-bottom:16px;">
        <label style="font-weight:bold;display:block;margin-bottom:6px;">2. アプリのURL（自動取得できない場合に入力）</label>
        <input type="text" id="urlInput" value="${escapeHtml(displayUrl)}"
               placeholder="https://script.google.com/macros/s/.../exec"
               style="width:100%;padding:8px;border-radius:6px;border:1px solid #cbd5e1;font-size:.85rem;">
        <small style="color:#64748b;font-size:.75rem;display:block;margin-top:4px;">※ 自動取得URLで問題ない場合は不要です。</small>
      </div>
      <div style="text-align:right;display:flex;gap:8px;justify-content:flex-end;">
        <button onclick="google.script.host.close()"
                style="padding:8px 16px;border-radius:6px;border:1px solid #cbd5e1;background:white;cursor:pointer;">キャンセル</button>
        <button onclick="saveSettings()"
                style="padding:8px 16px;border-radius:6px;border:none;background:#6366f1;color:white;font-weight:bold;cursor:pointer;">保存して反映</button>
      </div>
    </div>
    <script>
      function saveSettings() {
        const sheetName  = document.getElementById('sheetSelect').value;
        const webappUrl  = document.getElementById('urlInput').value.trim();
        google.script.run
          .withSuccessHandler(() => { alert('設定を保存しました。'); google.script.host.close(); })
          .withFailureHandler(err => alert('保存失敗: ' + err.message))
          .saveAppSettings(sheetName, webappUrl);
      }
    </script>`;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(450).setHeight(260),
    '⚙️ アプリの設定'
  );
}

function getTargetSheetName() {
  return PropertiesService.getScriptProperties().getProperty('TARGET_SHEET_NAME')
    || 'R8全備品台帳【3万以上】  ＋理振＋算振【1万円以上】';
}

function saveAppSettings(sheetName, webappUrl) {
  if (!sheetName) throw new Error('シート名が空です。');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('TARGET_SHEET_NAME', sheetName);
  if (webappUrl) props.setProperty('CUSTOM_WEBAPP_URL', webappUrl);
  else            props.deleteProperty('CUSTOM_WEBAPP_URL');
  return true;
}

// ============================================================
// getItems: Webアプリ向け備品データ取得（検査済状態を含む）
// ============================================================
function getItems() {
  try {
    const ss        = getSpreadsheet();
    const sheetName = getTargetSheetName();
    const sheet     = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('シート「' + sheetName + '」が見つかりませんでした。');

    const { headerRow, startRow } = findHeaderAndStartRow(sheet);
    const lastRow     = sheet.getLastRow();
    if (lastRow < startRow) return { headers: [], items: [] };

    const masterLastCol = sheet.getLastColumn();
    const headersVal    = sheet.getRange(headerRow, 1, 1, masterLastCol).getValues()[0];
    const headers       = headersVal.map((h, idx) => {
      const letter = String.fromCharCode(65 + idx);
      return h ? `${letter}列: ${String(h).trim()}` : `${letter}列 (無題)`;
    });

    const cleanMaster = headersVal.map(h => String(h).trim().replace(/\s+/g, ''));
    const numRows     = lastRow - startRow + 1;
    const values      = sheet.getRange(startRow, 1, numRows, masterLastCol).getValues();

    // 各列インデックスを動的に検出
    let colCheck = -1, colMasterId = -1, colB = -1, colC = -1;
    let colD = -1, colE = -1, colF = -1, colG = -1, colH = -1, colPhoto = -1, colLocation = -1;
    for (let col = 0; col < cleanMaster.length; col++) {
      const m = cleanMaster[col];
      if (m === '検査済')   colCheck    = col;
      else if (m === '登載番号') colMasterId = col;
      else if (m === '台帳番号') colB        = col;
      else if (m === '台帳分類') colC        = col;
      else if (m === '取得年月日') colD      = col;
      else if (m === '品名番号') colE        = col;
      else if (m === '番号')     colF        = col;
      else if (m === '規格')     colG        = col;
      else if (m === '品名')     colH        = col;
      else if (m === '保管場所等' || m.includes('場所') || m.includes('所在')) colLocation = col;
      else if (col > 7 && colPhoto === -1 &&
               (m.includes('写真') || m.includes('画像') || m.includes('リンク'))) colPhoto = col;
    }

    const get     = (row, idx) => idx !== -1 && idx < row.length ? String(row[idx]).trim() : '';
    const getDate = (row, idx) => {
      if (idx === -1 || idx >= row.length) return '';
      const v = row[idx];
      if (!v) return '';
      // GAS が Date オブジェクトとして返した場合
      if (v instanceof Date) {
        const y  = v.getFullYear();
        const mo = String(v.getMonth() + 1).padStart(2, '0');
        const d  = String(v.getDate()).padStart(2, '0');
        return `${y}/${mo}/${d}`;
      }
      // 既に文字列の場合（例: "2002/07/30" や "Tue Jul 30..."）
      const s = String(v).trim();
      if (!s) return '';
      // 「yyyy/MM/dd」形式ならそのまま返す
      if (/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(s)) return s;
      // その他の日付文字列をパースして整形
      const parsed = new Date(s);
      if (!isNaN(parsed.getTime())) {
        const y  = parsed.getFullYear();
        const mo = String(parsed.getMonth() + 1).padStart(2, '0');
        const d  = String(parsed.getDate()).padStart(2, '0');
        return `${y}/${mo}/${d}`;
      }
      return s;
    };

    const items = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const b   = get(row, colB);
      const c   = get(row, colC);
      const g   = get(row, colG);
      if (!b && !c && !g) continue;

      items.push({
        row:      i + startRow,
        a:        get(row, colMasterId),
        b:        b,
        c:        c,
        d:        getDate(row, colD),
        e:        get(row, colE),
        f:        get(row, colF),
        g:        g,
        h:        get(row, colH),
        location: get(row, colLocation),
        photoUrl: get(row, colPhoto),
        checked:  colCheck !== -1 && colCheck < row.length ? row[colCheck] === true : false
      });
    }

    return { headers, items };
  } catch(e) {
    Logger.log('getItems エラー: ' + e.toString());
    throw new Error('備品データの読み込みに失敗しました: ' + e.message);
  }
}

// ============================================================
// uploadPhoto: 写真をドライブに保存してシートに書き込む
// ============================================================
function uploadPhoto(rowNumber, base64Data, fileName) {
  try {
    if (!rowNumber)   throw new Error('行番号が指定されていません。');
    if (!base64Data)  throw new Error('画像データがありません。');

    const matches = base64Data.match(/^data:(.*);base64,(.*)$/);
    if (!matches || matches.length !== 3) throw new Error('無効な画像データ形式です。');

    const blob   = Utilities.newBlob(Utilities.base64Decode(matches[2]), matches[1], fileName);
    const folder = (() => {
      const it = DriveApp.getFoldersByName('備品台帳写真');
      return it.hasNext() ? it.next() : DriveApp.createFolder('備品台帳写真');
    })();
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = file.getUrl();

    const ss        = getSpreadsheet();
    const sheetName = getTargetSheetName();
    const sheet     = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('シート「' + sheetName + '」が見つかりませんでした。');

    const { headerRow } = findHeaderAndStartRow(sheet);
    const masterLastCol = sheet.getLastColumn();
    const masterHdrs    = sheet.getRange(headerRow, 1, 1, masterLastCol).getValues()[0];

    let photoColIdx = -1;
    for (let col = 0; col < masterHdrs.length; col++) {
      const h = String(masterHdrs[col]).trim().replace(/\s+/g, '');
      if (col > 7 && (h.includes('写真') || h.includes('画像') || h.includes('リンク'))) {
        photoColIdx = col; break;
      }
    }
    if (photoColIdx === -1) {
      photoColIdx = masterLastCol;
      sheet.getRange(headerRow, photoColIdx + 1).setValue('写真のリンク');
      SpreadsheetApp.flush();
    }

    sheet.getRange(rowNumber, photoColIdx + 1).setValue(fileUrl);

    const itemVals = sheet.getRange(rowNumber, 1, 1, Math.min(masterLastCol, 10)).getValues()[0];
    const cleanHdr = masterHdrs.map(h => String(h).trim().replace(/\s+/g, ''));
    const idIdx    = cleanHdr.indexOf('登載番号');
    const bIdx     = cleanHdr.indexOf('台帳番号');
    const cIdx     = cleanHdr.indexOf('台帳分類');
    updateIndividualSheetPhoto(ss, {
      a: idIdx !== -1 ? String(itemVals[idIdx]).trim() : '',
      b: bIdx  !== -1 ? String(itemVals[bIdx]).trim()  : '',
      c: cIdx  !== -1 ? String(itemVals[cIdx]).trim()  : ''
    }, fileUrl);

    return { success: true, url: fileUrl, fileName };
  } catch(e) {
    Logger.log('uploadPhoto エラー: ' + e.toString());
    throw new Error('画像のアップロードに失敗しました: ' + e.message);
  }
}

// ============================================================
// syncAllIndividualSheets: 同期ダイアログを表示
// ============================================================
function syncAllIndividualSheets() {
  const html = `
    <div style="font-family:sans-serif;padding:16px;color:#1e293b;font-size:.9rem;">
      <p style="font-weight:bold;margin-bottom:12px;font-size:1rem;">台帳シートの一括同期を実行中...</p>
      <div style="background:#e2e8f0;border-radius:9999px;height:14px;overflow:hidden;margin-bottom:12px;border:1px solid #cbd5e1;">
        <div id="progressBar" style="width:0%;height:100%;background:linear-gradient(to right,#6366f1,#10b981);transition:width .4s ease;border-radius:9999px;"></div>
      </div>
      <p id="progressStatus" style="font-size:.85rem;color:#64748b;margin-bottom:20px;line-height:1.4;max-height:120px;overflow-y:auto;">
        接続中... マスタデータを読み込んでいます。
      </p>
      <div style="text-align:right;">
        <button id="closeBtn" disabled onclick="google.script.host.close()"
                style="padding:8px 18px;border-radius:6px;border:1px solid #cbd5e1;background:#f1f5f9;color:#94a3b8;font-weight:bold;cursor:not-allowed;font-size:.85rem;">
          処理中...
        </button>
      </div>
    </div>
    <script>
      window.onload = function() {
        const pBar = document.getElementById('progressBar');
        const pStatus = document.getElementById('progressStatus');
        const btn = document.getElementById('closeBtn');
        pBar.style.width = '15%';
        google.script.run
          .withSuccessHandler(function(r) {
            pBar.style.width = '100%';
            pStatus.innerHTML =
              '<span style="color:#10b981;font-weight:bold;font-size:.95rem;">✓ 同期が完了しました！</span><br>' +
              '<span style="display:inline-block;margin-top:6px;color:#334155;">' +
              '・新規追加: <b>' + r.newAddedCount + '</b> 件<br>' +
              '・同期・更新: <b>' + r.successCount + '</b> 件</span>';
            btn.disabled = false;
            btn.innerText = '閉じる';
            btn.style.cssText = 'padding:8px 18px;border-radius:6px;border:none;background:#6366f1;color:white;font-weight:bold;cursor:pointer;font-size:.85rem;';
          })
          .withFailureHandler(function(err) {
            pBar.style.width = '100%';
            pBar.style.background = '#ef4444';
            pStatus.innerHTML =
              '<span style="color:#ef4444;font-weight:bold;">✕ エラー</span><br>' +
              '<pre style="color:#ef4444;font-size:.8rem;background:#fef2f2;padding:8px;border-radius:4px;white-space:pre-wrap;margin-top:8px;">' + err.message + '</pre>';
            btn.disabled = false;
            btn.innerText = '閉じる';
            btn.style.cssText = 'padding:8px 18px;border-radius:6px;background:#e2e8f0;color:#1e293b;cursor:pointer;font-size:.85rem;';
          })
          .executeSyncProcess();
      };
    </script>`;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(450).setHeight(250),
    '一括同期の進捗'
  );
}

// ============================================================
// findHeaderAndStartRow: ヘッダー行とデータ開始行を動的に検出
// ============================================================
function findHeaderAndStartRow(sheet) {
  const lastRow  = Math.min(sheet.getLastRow(), 15);
  if (lastRow < 1) return { headerRow: 7, startRow: 8 };

  const values = sheet.getRange(1, 1, lastRow, Math.min(sheet.getLastColumn(), 10)).getValues();
  const keywords = ['一連番号','品名番号','規格・型式','規格','品名','教科番号',
                    '個別連番','No','登載番号','台帳番号','検査済'];
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    for (let c = 0; c < row.length; c++) {
      const val = String(row[c]).trim().replace(/\s+/g, '');
      if (keywords.includes(val)) {
        return { headerRow: r + 1, startRow: r + 2 };
      }
    }
  }
  return { headerRow: 7, startRow: 8 };
}

// ============================================================
// adjustSheetLayoutIfNeeded (旧レイアウト対応)
// ============================================================
function adjustSheetLayoutIfNeeded(sheet) {
  const { headerRow } = findHeaderAndStartRow(sheet);
  if (sheet.getLastColumn() < 1) return;

  const firstHeader = String(sheet.getRange(headerRow, 1).getValue()).trim().replace(/\s+/g, '');
  if (/一連番号|マスタ|ID|台帳番号/i.test(firstHeader)) {
    sheet.insertColumnsBefore(1, 1);
    sheet.getRange(headerRow, 1).setValue('No');
    SpreadsheetApp.flush();
  }
}

// ============================================================
// buildColumnMapping: マスタ↔個別シートのヘッダー対応マップを構築
// ============================================================
function buildColumnMapping(masterHeaders, indHeaders) {
  const mapping     = {};
  const cleanMaster = masterHeaders.map(h => String(h).trim().replace(/\s+/g, ''));

  for (let i = 0; i < indHeaders.length; i++) {
    const h = String(indHeaders[i]).trim().replace(/\s+/g, '');
    // No列・検査済列はマッピングから除外（別途管理）
    if (!h || h === 'No' || h === '検査済') continue;

    let matched = -1;

    if (h === '登載番号') {
      matched = cleanMaster.indexOf('登載番号');
    } else if (h === '台帳番号' || h === '一連番号' || h === 'ID') {
      matched = cleanMaster.indexOf('台帳番号');
      if (matched === -1) matched = cleanMaster.indexOf('一連番号');
    } else if (h === '台帳分類' || h === '教科名') {
      matched = cleanMaster.indexOf('台帳分類');
      if (matched === -1) matched = cleanMaster.indexOf('教科名');
    } else if (h === '取得年月日' || h === '購入年月日') {
      matched = cleanMaster.indexOf('取得年月日');
      if (matched === -1) matched = cleanMaster.indexOf('購入年月日');
    } else if (h === '品名' || h === '物品名' || h === '名称') {
      matched = cleanMaster.indexOf('品名');
      if (matched === -1) matched = cleanMaster.indexOf('物品名');
    } else if (h === '規格' || h === '規格・型式') {
      matched = cleanMaster.indexOf('規格');
      if (matched === -1) matched = cleanMaster.indexOf('規格・型式');
    } else if (h === 'QRコード' || h === 'QR') {
      matched = cleanMaster.indexOf('QRコード');
    } else if (h === '写真のリンク' || h === '写真' || h === '画像' || h === 'リンク') {
      matched = cleanMaster.indexOf('写真のリンク');
      if (matched === -1) matched = cleanMaster.indexOf('写真');
    } else {
      matched = cleanMaster.findIndex(m => m !== '' && (m.includes(h) || h.includes(m)));
    }

    if (matched !== -1) mapping[i] = matched;
  }
  return mapping;
}

// ============================================================
// executeSyncProcess: 一括同期のメイン処理
// ============================================================
function executeSyncProcess() {
  try {
    const ss              = getSpreadsheet();
    const masterSheetName = getTargetSheetName();
    const masterSheet     = ss.getSheetByName(masterSheetName);
    if (!masterSheet) throw new Error(`マスタシート「${masterSheetName}」が見つかりません。`);

    const { headerRow: masterHeaderRow, startRow: masterStartRow } = findHeaderAndStartRow(masterSheet);
    const lastRow = masterSheet.getLastRow();
    if (lastRow < masterStartRow) return { newAddedCount: 0, successCount: 0 };

    const masterLastCol  = masterSheet.getLastColumn();
    const masterHeaders  = masterSheet.getRange(masterHeaderRow, 1, 1, masterLastCol).getValues()[0];
    const numRows        = lastRow - masterStartRow + 1;
    const masterValues   = masterSheet.getRange(masterStartRow, 1, numRows, masterLastCol).getValues();

    // マスタの登載番号・台帳番号・台帳分類のインデックスを動的に検出
    const cleanMasterHdr = masterHeaders.map(h => String(h).trim().replace(/\s+/g, ''));
    const mIdIdx  = cleanMasterHdr.indexOf('登載番号');  // 登載番号列
    const mBIdx   = cleanMasterHdr.indexOf('台帳番号');  // 台帳番号列
    const mCIdx   = cleanMasterHdr.indexOf('台帳分類');  // 台帳分類列
    if (mIdIdx === -1) throw new Error('マスタシートに「登載番号」列が見つかりません。');

    const sheets       = ss.getSheets();
    const sheetNames   = sheets.map(s => s.getName());
    const sheetCache   = {};
    const validIdsPerSheet = {};

    let successCount = 0, newAddedCount = 0;

    // WebアプリURL取得
    const props    = PropertiesService.getScriptProperties();
    const savedUrl = props.getProperty('CUSTOM_WEBAPP_URL');
    let autoUrl = '';
    try { autoUrl = ScriptApp.getService().getUrl() || ''; } catch(e) {}
    const webappUrl      = savedUrl || autoUrl;
    const encodedBaseUrl = webappUrl ? encodeURIComponent(webappUrl + '?id=') : '';

    // ----------------------------------------------------------------
    // マスタ各行を走査して個別シートへ反映
    // ----------------------------------------------------------------
    for (let i = 0; i < masterValues.length; i++) {
      const row   = masterValues[i];
      const colA  = mIdIdx < row.length ? String(row[mIdIdx]).trim() : '';   // 登載番号
      const colBv = mBIdx  < row.length ? String(row[mBIdx]).trim()  : '';   // 台帳番号
      const colCv = mCIdx  < row.length ? String(row[mCIdx]).trim()  : '';   // 台帳分類
      if (!colCv) continue;

      const targetSheetName = findIndividualSheetName(sheetNames, colBv, colCv);
      if (!targetSheetName) continue;

      if (!validIdsPerSheet[targetSheetName]) validIdsPerSheet[targetSheetName] = new Set();
      if (colA) validIdsPerSheet[targetSheetName].add(colA);

      // ----------------------------------------------------------------
      // 個別シートキャッシュ構築（初回のみ）
      // ----------------------------------------------------------------
      if (!sheetCache[targetSheetName]) {
        const sh = ss.getSheetByName(targetSheetName);

        // チェックボックス列（検査済）が無ければ挿入
        const { headerRow: hRow } = findHeaderAndStartRow(sh);
        const firstHdr = String(sh.getRange(hRow, 1).getValue()).trim().replace(/\s+/g, '');
        if (firstHdr !== '検査済') {
          sh.insertColumnBefore(1);
          sh.getRange(hRow, 1).setValue('検査済');
        }

        // 18列に拡張
        const curCols = sh.getMaxColumns();
        if (curCols < TARGET_COL_COUNT) {
          sh.insertColumnsAfter(curCols, TARGET_COL_COUNT - curCols);
        }

        const { headerRow, startRow } = findHeaderAndStartRow(sh);
        sh.getRange(headerRow, 1, 1, TARGET_COL_COUNT).setValues([TARGET_HEADERS]);

        const shLastRow = sh.getLastRow();
        const colMax    = Math.max(sh.getLastColumn(), TARGET_COL_COUNT);
        let vals        = [];

        if (shLastRow >= startRow) {
          vals = sh.getRange(startRow, 1, shLastRow - startRow + 1, colMax).getValues();
        }

        // 既存チェック状態を「登載番号」でキャッシュ
        // 個別シートの登載番号は index 2 (C列)
        const checkboxStates = {};
        for (const v of vals) {
          const mid = v.length > 2 ? String(v[2]).trim() : '';
          if (mid && v[0] === true) checkboxStates[mid] = true;
        }

        const mapping = buildColumnMapping(masterHeaders, TARGET_HEADERS);

        // 個別シートの「登載番号」列インデックスを取得（index=2 のはず）
        let colIndexMasterId = 2;
        let colIndexC        = 4; // 台帳分類
        for (const [indIdxStr, masterIdx] of Object.entries(mapping)) {
          const indIdx = parseInt(indIdxStr, 10);
          if (masterIdx === mIdIdx) colIndexMasterId = indIdx;
          if (masterIdx === mCIdx)  colIndexC        = indIdx;
        }

        let defaultClassification = '';
        if (colIndexC !== -1 && vals.length > 0) {
          for (const v of vals) {
            const val = String(v[colIndexC] || '').trim();
            if (val && !/^[0-9]+$/.test(val)) { defaultClassification = val; break; }
          }
        }
        if (!defaultClassification) {
          defaultClassification = targetSheetName.replace(/^[0-9１２３４５６７８９０\s]+/g, '');
        }

        sheetCache[targetSheetName] = {
          sheet: sh, values: vals, checkboxStates, mapping,
          colIndexMasterId, colIndexC, defaultClassification,
          headerRow, startRow, numCols: colMax
        };
      }

      const sc = sheetCache[targetSheetName];

      // 既存行の検索（登載番号で照合）
      let matchedIndex = -1;
      for (let k = 0; k < sc.values.length; k++) {
        const idVal = sc.colIndexMasterId < sc.values[k].length
          ? String(sc.values[k][sc.colIndexMasterId]).trim() : '';
        if (colA && idVal && colA === idVal) { matchedIndex = k; break; }
      }

      if (matchedIndex !== -1) {
        const exRow     = sc.values[matchedIndex];
        let rowChanged  = false;
        const upd = (idx, newVal) => {
          if (idx < exRow.length && String(exRow[idx]).trim() !== String(newVal).trim()) {
            exRow[idx] = newVal; rowChanged = true;
          }
        };
        upd(1, matchedIndex + 1); // B列 No 更新（A列チェックは触らない）
        for (const [indIdxStr, masterIdx] of Object.entries(sc.mapping)) {
          const ii = parseInt(indIdxStr, 10);
          if (ii < exRow.length && masterIdx < row.length) {
            const v = ii === sc.colIndexC ? sc.defaultClassification : row[masterIdx];
            upd(ii, v);
          }
        }
        if (rowChanged) successCount++;
      } else {
        const newRow = new Array(sc.numCols).fill('');
        newRow[0] = false; // 検査済 = 未チェック
        newRow[1] = sc.values.length + 1; // No
        for (const [indIdxStr, masterIdx] of Object.entries(sc.mapping)) {
          const ii = parseInt(indIdxStr, 10);
          if (ii < newRow.length && masterIdx < row.length) {
            newRow[ii] = ii === sc.colIndexC ? sc.defaultClassification : row[masterIdx];
          }
        }
        sc.values.push(newRow);
        newAddedCount++;
      }
    }

    // ----------------------------------------------------------------
    // 各個別シートへ書き込み
    // ----------------------------------------------------------------
    for (const [targetSheetName, sc] of Object.entries(sheetCache)) {
      const validIds = validIdsPerSheet[targetSheetName] || new Set();

      // マスタから消えた行を除去
      sc.values = sc.values.filter(v => {
        const mid = sc.colIndexMasterId < v.length ? String(v[sc.colIndexMasterId]).trim() : '';
        return mid === '' || validIds.has(mid);
      });

      // チェック状態を復元
      for (let k = 0; k < sc.values.length; k++) {
        const mid = sc.colIndexMasterId < sc.values[k].length
          ? String(sc.values[k][sc.colIndexMasterId]).trim() : '';
        sc.values[k][0] = sc.checkboxStates[mid] || false;
        sc.values[k][1] = k + 1; // No 再割り当て
      }

      // データエリアをクリア
      const shLastRow = sc.sheet.getLastRow();
      if (shLastRow >= sc.startRow) {
        sc.sheet.getRange(sc.startRow, 1, shLastRow - sc.startRow + 1, sc.numCols).clearContent();
      }

      // QRコード数式を設定（C列 = 登載番号 を参照）
      if (encodedBaseUrl) {
        for (let k = 0; k < sc.values.length; k++) {
          const targetRowNum = sc.startRow + k;
          sc.values[k][16] = `=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodedBaseUrl}" & ENCODEURL(C${targetRowNum}))`;
        }
      }

      if (sc.values.length > 0) {
        const dataRange = sc.sheet.getRange(sc.startRow, 1, sc.values.length, sc.numCols);
        dataRange.setValues(sc.values);
        dataRange.setWrap(true);

        // A列（検査済）にチェックボックス検証を設定（値は保持）
        sc.sheet.getRange(sc.startRow, 1, sc.values.length, 1)
          .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

        // R列（写真のリンク, 18列目）をはみ出しに設定
        sc.sheet.getRange(sc.startRow, 18, sc.values.length, 1)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
      }

      // ソートと連番再設定
      const newLastRow = sc.sheet.getLastRow();
      sortSheetNumericallyByMasterId(sc.sheet, newLastRow, sc.colIndexMasterId);

      // 行の高さを自動調整
      if (newLastRow >= sc.startRow) {
        sc.sheet.autoResizeRows(sc.startRow, newLastRow - sc.startRow + 1);
      }
    }

    return { newAddedCount, successCount };
  } catch(e) {
    Logger.log('executeSyncProcess エラー: ' + e.toString());
    throw new Error(e.message);
  }
}

// ============================================================
// updateIndividualSheetPhoto: 写真URLを個別シートに書き込む
// ============================================================
function updateIndividualSheetPhoto(ss, item, photoUrl) {
  try {
    const sheetNames      = ss.getSheets().map(s => s.getName());
    const b = item.b ? String(item.b).trim() : '';
    const c = item.c ? String(item.c).trim() : '';
    if (!c) return;

    const targetSheetName = findIndividualSheetName(sheetNames, b, c);
    if (!targetSheetName) return;

    const sheet = ss.getSheetByName(targetSheetName);
    adjustSheetLayoutIfNeeded(sheet);

    const { headerRow, startRow } = findHeaderAndStartRow(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return;

    const masterSheet = ss.getSheetByName(getTargetSheetName());
    const { headerRow: mHR } = findHeaderAndStartRow(masterSheet);
    const masterHeaders = masterSheet.getRange(mHR, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    let masterPhotoIdx  = -1;
    for (let col = 0; col < masterHeaders.length; col++) {
      const h = String(masterHeaders[col]).trim().replace(/\s+/g, '');
      if (col > 7 && (h.includes('写真') || h.includes('画像') || h.includes('リンク'))) {
        masterPhotoIdx = col; break;
      }
    }

    const indHeaders    = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const mapping       = buildColumnMapping(masterHeaders, indHeaders);
    let colIndexMasterId = 2, colIndexPhoto = 17;
    const mIdIdx = masterHeaders.map(h => String(h).trim().replace(/\s+/g, '')).indexOf('登載番号');
    for (const [indIdxStr, masterIdx] of Object.entries(mapping)) {
      const ii = parseInt(indIdxStr, 10);
      if (masterIdx === mIdIdx)        colIndexMasterId = ii;
      if (masterIdx === masterPhotoIdx && masterPhotoIdx !== -1) colIndexPhoto = ii;
    }

    if (lastRow < startRow) return;
    const numRows  = lastRow - startRow + 1;
    const numCols  = sheet.getLastColumn();
    const values   = sheet.getRange(startRow, 1, numRows, numCols).getValues();
    const targetA  = item.a ? String(item.a).trim() : '';
    let matchedRow = -1;
    for (let i = 0; i < values.length; i++) {
      const id = colIndexMasterId < values[i].length ? String(values[i][colIndexMasterId]).trim() : '';
      if (targetA && id && targetA === id) { matchedRow = startRow + i; break; }
    }

    if (matchedRow !== -1) {
      sheet.getRange(matchedRow, colIndexPhoto + 1).setValue(photoUrl);
      // R列（写真リンク）をはみ出しに設定
      sheet.getRange(startRow, 18, numRows, 1)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
      sortSheetNumericallyByMasterId(sheet, lastRow, colIndexMasterId);
    }
  } catch(e) {
    Logger.log('updateIndividualSheetPhoto エラー: ' + e.toString());
  }
}

// ============================================================
// sortSheetNumericallyByMasterId: 登載番号で昇順ソート、Noに連番を設定
// ============================================================
function sortSheetNumericallyByMasterId(sheet, lastRow, colIndexMasterId) {
  const { headerRow, startRow } = findHeaderAndStartRow(sheet);
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const numCols = sheet.getLastColumn();
  if (numCols < 1) return;

  try {
    // 登載番号を数値に変換してソート
    const idRange  = sheet.getRange(startRow, colIndexMasterId + 1, numRows, 1);
    const idValues = idRange.getValues();
    idRange.setValues(idValues.map(([v]) => {
      const s = String(v).trim();
      if (s === '') return [''];
      const n = parseFloat(s);
      return [isNaN(n) ? v : n];
    }));

    sheet.getRange(startRow, 1, numRows, numCols)
         .sort({ column: colIndexMasterId + 1, ascending: true });

    // B列 (No) に連番を設定 — A列 (検査済) は触らない
    const headers = sheet.getRange(headerRow, 1, 1, Math.min(numCols, 3)).getValues()[0];
    const noColIdx = headers.findIndex(h => String(h).trim().replace(/\s+/g, '') === 'No');
    if (noColIdx !== -1) {
      sheet.getRange(startRow, noColIdx + 1, numRows, 1)
           .setValues(Array.from({ length: numRows }, (_, i) => [i + 1]));
    }
  } catch(err) {
    throw new Error(
      `ソート処理中にエラー: ${err.message}\n` +
      `対象列: ${colIndexMasterId + 1}列目, 範囲: ${startRow}〜${lastRow}行目`
    );
  }
}

// ============================================================
// generateQRCodesForMaster: マスタQ列にQRコードを一括生成
// ============================================================
function generateQRCodesForMaster() {
  const ss          = getSpreadsheet();
  const masterSheet = ss.getSheetByName(getTargetSheetName());
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('マスタシートが見つかりません。');
    return;
  }

  const { headerRow, startRow } = findHeaderAndStartRow(masterSheet);
  const lastRow = masterSheet.getLastRow();
  if (lastRow < startRow) { SpreadsheetApp.getUi().alert('データ行がありません。'); return; }

  const masterLastCol = masterSheet.getLastColumn();
  const headers       = masterSheet.getRange(headerRow, 1, 1, masterLastCol).getValues()[0];
  const cleanHdr      = headers.map(h => String(h).trim().replace(/\s+/g, ''));

  // QRコード列・写真リンク列・登載番号列を動的に検出
  let qrColIdx    = cleanHdr.indexOf('QRコード');
  let photoColIdx = cleanHdr.indexOf('写真のリンク');
  const idColIdx  = cleanHdr.indexOf('登載番号');
  if (idColIdx === -1) { SpreadsheetApp.getUi().alert('マスタシートに「登載番号」列が見つかりません。'); return; }

  // 列が存在しなければ末尾に追加
  if (qrColIdx === -1) {
    qrColIdx = masterLastCol;
    masterSheet.getRange(headerRow, qrColIdx + 1).setValue('QRコード');
  }
  if (photoColIdx === -1) {
    photoColIdx = Math.max(qrColIdx + 1, masterLastCol);
    masterSheet.getRange(headerRow, photoColIdx + 1).setValue('写真のリンク');
  }

  // WebアプリURL取得
  const props    = PropertiesService.getScriptProperties();
  const savedUrl = props.getProperty('CUSTOM_WEBAPP_URL');
  let autoUrl = '';
  try { autoUrl = ScriptApp.getService().getUrl() || ''; } catch(e) {}
  let webappUrl = savedUrl || autoUrl;
  if (!webappUrl) {
    SpreadsheetApp.getUi().alert('WebアプリのURLが設定されていません。「⚙️ アプリの設定」から登録してください。');
    return;
  }

  // authuser 付与
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) {
      const sep = webappUrl.indexOf('?') === -1 ? '?' : '&';
      webappUrl = webappUrl + sep + 'authuser=' + encodeURIComponent(email);
    }
  } catch(e) {}

  const baseWithId     = webappUrl + (webappUrl.indexOf('?') === -1 ? '?' : '&') + 'id=';
  const encodedBaseUrl = encodeURIComponent(baseWithId);

  // 登載番号列の列文字（A=1, B=2, ... → 文字に変換）
  const idColLetter = colIndexToLetter(idColIdx);

  const numRows   = lastRow - startRow + 1;
  const masterIds = masterSheet.getRange(startRow, idColIdx + 1, numRows, 1).getValues();
  const qrValues  = masterSheet.getRange(startRow, qrColIdx + 1, numRows, 1).getValues();

  let count = 0;
  for (let i = 0; i < numRows; i++) {
    const rowNum   = startRow + i;
    const masterId = String(masterIds[i][0]).trim();
    const curQr    = String(qrValues[i][0]).trim();
    if (masterId && !curQr) {
      masterSheet.getRange(rowNum, qrColIdx + 1).setFormula(
        `=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodedBaseUrl}" & ENCODEURL(${idColLetter}${rowNum}))`
      );
      count++;
    }
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(`QRコードの作成が完了しました。\n\n新規作成: ${count} 件`);
}

// ============================================================
// ヘルパー: 列インデックス(0始まり)を列文字(A,B,...,Z,AA...)に変換
// ============================================================
function colIndexToLetter(idx) {
  let letter = '';
  let n = idx + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

// ============================================================
// createSpreadsheetBackup: バックアップ作成
// ============================================================
function createSpreadsheetBackup() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const file     = DriveApp.getFileById(ss.getId());
    const parents  = file.getParents();
    const parent   = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    const folders  = parent.getFoldersByName('バックアップ');
    const folder   = folders.hasNext() ? folders.next() : parent.createFolder('バックアップ');

    const timestamp  = Utilities.formatDate(new Date(), 'GMT+9', 'yyyyMMdd_HHmmss');
    const backupName = `${ss.getName()}_バックアップ_${timestamp}`;
    file.makeCopy(backupName, folder);

    ui.alert('バックアップ作成完了',
      `保存先: ${parent.getName()} / バックアップ\nファイル名: ${backupName}`,
      ui.ButtonSet.OK
    );
  } catch(e) {
    ui.alert('バックアップ作成中にエラー: ' + e.message);
  }
}

// ============================================================
// findIndividualSheetName: 教科番号・教科名で個別シートを検索
// ============================================================
function findIndividualSheetName(sheetNames, b, c) {
  if (!c) return null;
  const bInt     = parseInt(b, 10);
  const bHalf    = isNaN(bInt) ? '' : String(bInt);
  const bFull    = isNaN(bInt) ? '' : String(bInt).replace(/[0-9]/g, s => String.fromCharCode(s.charCodeAt(0) + 65248));
  const bZero    = isNaN(bInt) ? '' : (bInt < 10 ? '0' + bInt : String(bInt));
  const patterns = [];
  if (!isNaN(bInt)) {
    [bHalf, bFull, bZero].forEach(pfx => {
      patterns.push(pfx + c, pfx + ' ' + c);
    });
  }
  patterns.push(c);

  for (const pat of patterns) {
    const m = sheetNames.find(n => n.trim() === pat);
    if (m) return m;
  }
  return sheetNames.find(n => n.includes(c)) || null;
}

// ============================================================
// escapeHtml
// ============================================================
function escapeHtml(str) {
  if (typeof str !== 'string') return str;
  return str.replace(/[&<>"']/g, m => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  })[m]);
}

// ============================================================
// getSpreadsheet
// ============================================================
function getSpreadsheet() {
  const SPREADSHEET_ID = '1k-06SrGaT0HZGnmdTZKNg2tZNFl9W47Y73ARD3amaW0';
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch(e) {
    Logger.log('IDでのオープンに失敗: ' + e.toString());
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ============================================================
// applyColumnWidthsFromSansuu: 「１算数」の列幅を全個別シートへコピー
// ============================================================
function applyColumnWidthsFromSansuu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss              = getSpreadsheet();
    const masterSheetName = getTargetSheetName();
    const sheets          = ss.getSheets();

    // ソースシートを探す（「算数」を含む最初のシート）
    const srcSheet = sheets.find(s => s.getName().includes('算数'));
    if (!srcSheet) {
      ui.alert('「算数」という名前を含むシートが見つかりませんでした。');
      return;
    }

    // ソースシートの列幅を取得（最大 20 列分）
    const maxCols = Math.min(srcSheet.getMaxColumns(), 20);
    const widths  = [];
    for (let col = 1; col <= maxCols; col++) {
      widths.push(srcSheet.getColumnWidth(col));
    }

    // 個別シート（マスタ・ソース以外）へ順番に適用
    let appliedCount = 0;
    for (const sheet of sheets) {
      const name = sheet.getName();
      if (name === masterSheetName || name === srcSheet.getName()) continue;

      // 個別台帳シートかどうかをヘッダーで判定
      const { headerRow } = findHeaderAndStartRow(sheet);
      const fh = String(sheet.getRange(headerRow, 1).getValue()).trim().replace(/\s+/g, '');
      const knownHeaders = ['検査済', 'No', '一連番号', '登載番号', '台帳番号', '品名番号', '教科番号'];
      if (!knownHeaders.some(k => k === fh)) continue;

      // 列幅を適用
      const sheetCols = Math.min(sheet.getMaxColumns(), maxCols);
      for (let col = 1; col <= sheetCols; col++) {
        sheet.setColumnWidth(col, widths[col - 1]);
      }
      appliedCount++;
    }

    SpreadsheetApp.flush();
    ui.alert(
      `完了しました。\n\n` +
      `ソース: 「${srcSheet.getName()}」\n` +
      `適用したシート数: ${appliedCount} シート`
    );
  } catch(e) {
    ui.alert('列幅コピー中にエラー: ' + e.message);
  }
}

// ============================================================
// deletePhoto: 写真リンクをマスタと個別シートから消去する
// ============================================================
function deletePhoto(masterId, rowNumber) {
  try {
    const ss = getSpreadsheet();
    const masterSheetName = getTargetSheetName();
    const masterSheet = ss.getSheetByName(masterSheetName);
    
    // 1. マスタシートの写真リンクを消去
    if (masterSheet && rowNumber) {
      const { headerRow } = findHeaderAndStartRow(masterSheet);
      const masterLastCol = masterSheet.getLastColumn();
      const masterHdrs = masterSheet.getRange(headerRow, 1, 1, masterLastCol).getValues()[0];
      
      let photoColIdx = -1;
      for (let col = 0; col < masterHdrs.length; col++) {
        const h = String(masterHdrs[col]).trim().replace(/\s+/g, '');
        if (col > 7 && (h.includes('写真') || h.includes('画像') || h.includes('リンク'))) {
          photoColIdx = col; break;
        }
      }
      
      if (photoColIdx !== -1) {
        masterSheet.getRange(rowNumber, photoColIdx + 1).setValue('');
      }
    }
    
    // 2. 個別シートの写真リンクを消去
    if (masterId) {
      const sheets = ss.getSheets();
      for (const sheet of sheets) {
        if (sheet.getName() === masterSheetName) continue;
        const { headerRow, startRow } = findHeaderAndStartRow(sheet);
        const lastRow = sheet.getLastRow();
        if (lastRow < startRow) continue;
        
        const lastCol = sheet.getLastColumn();
        const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
        const idColIdx = getMasterIdColumnIdx(headers);
        if (idColIdx === -1) continue;
        
        let photoColIdx = -1;
        for (let col = 0; col < headers.length; col++) {
          const h = String(headers[col]).trim().replace(/\s+/g, '');
          if (col > 7 && (h.includes('写真') || h.includes('画像') || h.includes('リンク'))) {
            photoColIdx = col; break;
          }
        }
        if (photoColIdx === -1) continue;
        
        const numRows = lastRow - startRow + 1;
        const ids = sheet.getRange(startRow, idColIdx + 1, numRows, 1).getValues();
        for (let i = 0; i < ids.length; i++) {
          if (String(ids[i][0]).trim() === String(masterId).trim()) {
            sheet.getRange(startRow + i, photoColIdx + 1).setValue('');
            break;
          }
        }
      }
    }
    
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    Logger.log('deletePhoto エラー: ' + e.toString());
    throw new Error('写真リンクの削除に失敗しました: ' + e.message);
  }
}

