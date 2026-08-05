// Version: 2.0.0 (APIバッチ処理対応版 - 2026-06-13)
// キャッシュ用オブジェクト（スプレッドシートやフォルダなどの情報を保持）
const Cache = {
  ss: null,
  ssId: null,
  shId: null,
  folderId: null
};

/**
 * スプレッドシートや親フォルダの情報を遅延評価・取得します。
 * グローバルスコープでの重いAPI実行によるエラーや遅延を防ぎます。
 */
function getSpreadsheetInfo() {
  if (!Cache.ss) {
    Cache.ss = SpreadsheetApp.getActiveSpreadsheet();
    Cache.ssId = Cache.ss.getId();
    Cache.shId = Cache.ss.getActiveSheet().getSheetId();
    
    try {
      // スプレッドシートの親フォルダを取得
      const parentFolder = DriveApp.getFileById(Cache.ssId).getParents();
      if (parentFolder.hasNext()) {
        Cache.folderId = parentFolder.next().getId();
      }
    } catch (e) {
      console.warn("親フォルダの取得に失敗しました: " + e.toString());
    }
  }
  return Cache;
}

/**
 * 印刷の確認メッセージを表示し、実行が許可されればPDF作成を開始します。
 */
function jikkou() {
  const result = Browser.msgBox("本当に印刷（ＰＤＦ作成）しますか？（１人分作成に10秒かかります。）", Browser.Buttons.OK_CANCEL);
  if (result === "ok") {
    savePdf();
  } 
}

/**
 * 出席番号を1つ戻します（前へ）。
 */
function mae() {
  const info = getSpreadsheetInfo();
  const range = info.ss.getRange("an2");
  const currentValue = Number(range.getValue()) || 1;
  if (currentValue > 1) {
    range.setValue(currentValue - 1);
  }
}

/**
 * 出席番号を1つ進めます（次へ）。
 */
function tugi() {
  const info = getSpreadsheetInfo();
  const range = info.ss.getRange("an2");
  const currentValue = Number(range.getValue()) || 1;
  range.setValue(currentValue + 1);
}

/**
 * 指定された出席番号の範囲でPDFを作成し、最後にそれらを結合します。
 */
function savePdf() {
  const info = getSpreadsheetInfo();
  
  // 印刷開始位置と終わりの位置を取得
  const syoki = info.ss.getRange("v5"); // 開始番号
  const owari = info.ss.getRange('aa5'); // 終了番号
  const startNum = syoki.getValue();
  const endNum = owari.getValue();

  const mergearr = [];

  // PDFを開始番号から終了番号まで出力する
  for (let i = startNum; i <= endNum; i++) {
    info.ss.getRange('an2').setValue(i);

    // スプレッドシートに即時反映させる
    SpreadsheetApp.flush();

    // PDFの一時ファイル名
    const fileName = info.ss.getName() + " _" + String(i);

    // 関数createPdfを実行し、PDFを作成してリストに追加
    const file = createPdf(info.folderId, info.ssId, info.shId, fileName);
    if (file) {
      mergearr.push(file);
    }
    Utilities.sleep(8000);
  }

  // PDFの結合を実行
  pdf_merge(info.folderId, info.ss, mergearr);
}

/**
 * PDFを作成し指定したフォルダーに保存する関数
 * @param {string} folderId 保存先フォルダーのID
 * @param {string} ssId スプレッドシートのID
 * @param {string} shId シートのID
 * @param {string} fileName ファイル名
 * @return {File} 作成されたPDFファイルオブジェクト
 */
function createPdf(folderId, ssId, shId, fileName) {
  // PDFを作成するためのベースとなるURL
  const baseUrl = "https://docs.google.com/spreadsheets/d/"
          + ssId
          + "/export?gid="
          + shId;
 
  // PDFのオプションを指定
  const pdfOptions = "&exportFormat=pdf&format=pdf"
              + "&size=A4" // 用紙サイズ (A4)
              + "&portrait=true" // 用紙の向き true: 縦向き / false: 横向き
              + "&range=A2:ao55" // セル範囲を指定
              + "&fitw=true" // ページ幅を用紙にフィットさせるか true: フィットさせる / false: 原寸大
              + "&top_margin=0.50" // 上の余白
              + "&right_margin=0.50" // 右の余白
              + "&bottom_margin=0.50" // 下の余白
              + "&left_margin=1.00" // 左の余白
              + "&horizontal_alignment=CENTER" // 水平方向の位置
              + "&vertical_alignment=TOP" // 垂直方向の位置
              + "&printtitle=false" // スプレッドシート名の表示有無
              + "&sheetnames=false" // シート名の表示有無
              + "&gridlines=false" // グリッドラインの表示有無
              + "&fzr=false" // 固定行の表示有無
              + "&fzc=false"; // 固定列の表示有無

  // PDFを作成するためのURL
  const url = baseUrl + pdfOptions;

  // アクセストークンを取得する
  const token = ScriptApp.getOAuthToken();

  // headersにアクセストークンを格納する
  const options = {
    headers: {
        'Authorization': 'Bearer ' + token
    }
  };
 
  // PDFを作成する
  const blob = UrlFetchApp.fetch(url, options).getBlob().setName(fileName + '.pdf');

  // PDFの保存先フォルダー
  const folder = DriveApp.getFolderById(folderId);

  // PDFを指定したフォルダに保存する
  const file = folder.createFile(blob);
  return file;
}

/**
 * 分割して作成された複数のPDFを1つに結合します。
 * @param {string} folderId 保存先フォルダーのID
 * @param {Spreadsheet} ss スプレッドシートオブジェクト
 * @param {Array<File>} mergearr 結合対象のPDFファイルオブジェクトの配列
 */
function pdf_merge(folderId, ss, mergearr) {
  const ui = SpreadsheetApp.getUi();
  const folder = DriveApp.getFolderById(folderId);
  
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + 1;
  const date = today.getDate();
 
  // 結合後ファイル名の指定
  const filename = String(year) + String(month) + String(date) + "_" + ss.getName(); 
  
  // PDF結合実行
  const mergedFile = mergePdfs(folder, filename, mergearr);
  
  // 結合前PDFをゴミ箱へ移動
  for (let i = 0; i < mergearr.length; i++) {
    mergearr[i].setTrashed(true);
  }

  // 終了メッセージの表示と確認ダイアログ
  const res = ui.alert("同じフォルダ内に開始番号から終了番号までのPDFを作成しました。", "フォルダーを開きますか？", ui.ButtonSet.OK_CANCEL);

  if (res == ui.Button.OK) {
    const url = "https://docs.google.com/drive/folders/" + folderId;
    const script = "<a href='" + url + "' target='_blank'>" + url + "</a>";
    const html = HtmlService.createHtmlOutput(script);
    SpreadsheetApp.getUi().showModalDialog(html, 'URLをクリックしてください。');
    Logger.log("OKボタンが押されました");
  } else if (res == ui.Button.CANCEL) {
    Logger.log("キャンセルボタンが押されました");
  } else if (res == ui.Button.CLOSE) {
    Logger.log("閉じるボタンが押されました");
  }
}

/**
 * Merges all given PDF files into one.
 *
 * @param {Folder} directory the folder to store the output file
 * @param {string} name the desired name of the output file
 * @param {File} pdf1 the first PDF file
 * @param {File} pdf2 the second PDF file
 * @param {File} opt_pdf3 [optional] the third PDF file; add as many more as you like
 *
 * @return {File} the merged file
 */
function mergePdfs(directory, name, pdflist) {
  
  if (name.slice(-4) != '.pdf') {
    
    name = name + '.pdf';
    
  }
  var newObjects = ['1 0 obj\r\n<</Type/Catalog/Pages 2 0 R >>\r\nendobj'];
  var pageAddresses = [];
  for (var i = 0;i<pdflist.length;i++) {
    
    var bytes = pdflist[i].getBlob().getBytes();
    var xrefByteOffset = '';
    var byteIndex = bytes.length - 1;
    while (!/\sstartxref\s/.test(xrefByteOffset)) {
      
      xrefByteOffset = String.fromCharCode(bytes[byteIndex]) + xrefByteOffset;
      byteIndex--;
      
    }
    xrefByteOffset = +(/\s\d+\s/.exec(xrefByteOffset)[0]);
    var objectByteOffsets = [];
    var trailerDictionary = '';
    var rootAddress = '';
    do {
      
      var xrefTable = '';
      var trailerEndByteOffset = byteIndex;
      byteIndex = xrefByteOffset;
      for (byteIndex; byteIndex <= trailerEndByteOffset; byteIndex++) {
        
        xrefTable = xrefTable + String.fromCharCode(bytes[byteIndex]);
        
      }
      xrefTable = xrefTable.split(/\s*trailer\s*/);
      trailerDictionary = xrefTable[1];
      if (objectByteOffsets.length < 1) {
        
        rootAddress = /\d+\s+\d+\s+R/.exec(/\/Root\s*\d+\s+\d+\s+R/.exec(trailerDictionary)[0])[0].replace('R', 'obj');
        
      }
      xrefTable = xrefTable[0].split('\n');
      xrefTable.shift();
      while (xrefTable.length > 0) {
        
        var xrefSectionHeader = xrefTable.shift().split(/\s+/);
        var objectNumber = +xrefSectionHeader[0];
        var numberObjects = +xrefSectionHeader[1];
        for (var entryIndex = 0; entryIndex < numberObjects; entryIndex++) {
          
          var entry = xrefTable.shift().split(/\s+/);
          objectByteOffsets.push([[objectNumber, +entry[1], 'obj'], +entry[0]]);
          objectNumber++;
          
        }
        
      }
      if (/\s*\/Prev/.test(trailerDictionary)) {
        
        xrefByteOffset = +(/\s*\d+\s/.exec(/\s*\/Prev\s*\d+\s/.exec(trailerDictionary)[0])[0]);
        
      }
      
    } while (/\s*\/Prev/.test(trailerDictionary));
    var rootObject = getObject(rootAddress, objectByteOffsets, bytes);
    var pagesAddress = /\d+\s+\d+\s+R/.exec(/\/Pages\s*\d+\s+\d+\s+R/.exec(rootObject)[0])[0].replace('R', 'obj');
    var pagesObject = getObject(pagesAddress, objectByteOffsets, bytes);
    var objects = getDependencies(pagesObject, objectByteOffsets, bytes);
    var newObjectsInsertionIndex = newObjects.length;
    for (var objectIndex = 0; objectIndex < objects.length; objectIndex++) {
      
      var newObjectAddress = [(newObjects.length + 3) + '', 0 + '', 'obj'];
      if (!Array.isArray(objects[objectIndex])) {
        
        objects[objectIndex] = [objects[objectIndex]];
          
      }
      objects[objectIndex].unshift(newObjectAddress);
      var objectAddress = objects[objectIndex][1].match(/\d+\s+\d+\s+obj/)[0].split(/\s+/);
      objects[objectIndex].splice(1, 0, objectAddress);
      if (/\/Type\s*\/Page[^s]/.test(objects[objectIndex][2])) {
        
        objects[objectIndex][2] = objects[objectIndex][2].replace(/\/Parent\s*\d+\s+\d+\s+R/.exec(objects[objectIndex][2])[0], '/Parent 2 0 R');
        pageAddresses.push(newObjectAddress.join(' ').replace('obj', 'R'));
        
      }
      var addressRegExp = new RegExp(objectAddress[0] + '\\s+' + objectAddress[1] + '\\s+' + 'obj');
      objects[objectIndex][2] = objects[objectIndex][2].replace(addressRegExp.exec(objects[objectIndex][2])[0], newObjectAddress.join(' '));
      newObjects.push(objects[objectIndex]);
      
    }
    for (var referencingObjectIndex = newObjectsInsertionIndex; referencingObjectIndex < newObjects.length; referencingObjectIndex++) {
      
      var references = newObjects[referencingObjectIndex][2].match(/\d+\s+\d+\s+R/g);
      if (references != null) {
        
        var string = newObjects[referencingObjectIndex][2];
        var referenceIndices = [];
        var currentIndex = 0;
        for (var referenceIndex = 0; referenceIndex < references.length; referenceIndex++) {
          
          referenceIndices.push([]);
          referenceIndices[referenceIndex].push(string.slice(currentIndex).indexOf(references[referenceIndex]) + currentIndex);
          referenceIndices[referenceIndex].push(references[referenceIndex].length);
          currentIndex += string.slice(currentIndex).indexOf(references[referenceIndex]);
          
        }
        for (var referenceIndex = 0; referenceIndex < references.length; referenceIndex++) {
          
          var objectAddress = references[referenceIndex].replace('R', 'obj').split(/\s+/);
          for (var objectIndex = newObjectsInsertionIndex; objectIndex < newObjects.length; objectIndex++) {
            
            if (arrayEquals(objectAddress, newObjects[objectIndex][1])) {
              
              var length = string.length;
              newObjects[referencingObjectIndex][2] = string.slice(0, referenceIndices[referenceIndex][0]) + newObjects[objectIndex][0].join(' ').replace('obj', 'R') +
                string.slice(referenceIndices[referenceIndex][0] + referenceIndices[referenceIndex][1]);
              string = newObjects[referencingObjectIndex][2];
              var newLength = string.length;
              if (!(length == newLength)) {
                
                for (var subsequentReferenceIndex = referenceIndex + 1; subsequentReferenceIndex < references.length; subsequentReferenceIndex++) {
                  
                  referenceIndices[subsequentReferenceIndex][0] += (newLength - length);
                  
                }
                
              }
              break;
              
            }
            
          }
          
        }
        
      }
      
    }
    for (var objectIndex = newObjectsInsertionIndex; objectIndex < newObjects.length; objectIndex++) {
      
      if (Array.isArray(newObjects[objectIndex])) {
        
        if (newObjects[objectIndex][3] != undefined) {
          
          newObjects[objectIndex] = newObjects[objectIndex].slice(2);
          
        } else {
          
          newObjects[objectIndex] = newObjects[objectIndex][2];
          
        }
        
      }
      
    }
    
  }
  newObjects.splice(1, 0, '2 0 obj\r\n<</Type/Pages/Count ' + pageAddresses.length + ' /Kids [' + pageAddresses.join(' ') + ' ]>>\r\nendobj');
  newObjects.splice(2, 0, '3 0 obj\r\n<</Title (' + name + ') /Producer (PdfManipulation.mergePdfs\\(\\), a Google Apps Script project by Jarom Luker \\(pricebook@hbboys.com\\)) /CreationDate (D' +
       Utilities.formatDate(new Date(), CalendarApp.getDefaultCalendar().getTimeZone(), 'yyyyMMddHHmmssZ').slice(0, -2) + "'00) /ModDate (D" + Utilities.formatDate(new Date(),
       CalendarApp.getDefaultCalendar().getTimeZone(), 'yyyyMMddHHmmssZ').slice(0, -2) + "'00)>>\r\nendobj");
  var byteOffsets = [0];
  var bytes = [];
  var header = '%PDF-1.3\r\n';
  for (var headerIndex = 0; headerIndex < header.length; headerIndex++) {
    
    bytes.push(header.charCodeAt(headerIndex));
    
  }
  bytes.push('%'.charCodeAt(0));
  for (var characterCode = -127; characterCode < -123; characterCode++) {
    
    bytes.push(characterCode);
    
  }
  bytes.push('\r'.charCodeAt(0));
  bytes.push('\n'.charCodeAt(0));
  while (newObjects.length > 0) {
    
    byteOffsets.push(bytes.length);
    var object = newObjects.shift();
    if (Array.isArray(object)) {
      
      var streamKeyword = /stream\s*\n/.exec(object[0])[0];
      if (streamKeyword.indexOf('\n\n') > streamKeyword.length - 3) {
        
        streamKeyword = streamKeyword.slice(0, -1);
        
      } else if (streamKeyword.indexOf('\r\n\r\n') > streamKeyword.length - 5) {
        
        streamKeyword = streamKeyword.slice(0, -2);
        
      }
      var streamIndex = object[0].indexOf(streamKeyword) + streamKeyword.length;
      for (var objectIndex = 0; objectIndex < streamIndex; objectIndex++) {
        
        bytes.push(object[0].charCodeAt(objectIndex))
        
      }
      bytes = bytes.concat(object[1]);
      for (var objectIndex = streamIndex; objectIndex < object[0].length; objectIndex++) {
        
        bytes.push(object[0].charCodeAt(objectIndex));
        
      }
      
    } else {
      
      for (var objectIndex = 0; objectIndex < object.length; objectIndex++) {
        
        bytes.push(object.charCodeAt(objectIndex));
        
      }
      
    }
    bytes.push('\r'.charCodeAt(0));
    bytes.push('\n'.charCodeAt(0));
    
  }
  var xrefByteOffset = bytes.length;
  var xrefHeader = 'xref\r\n';
  for (var xrefHeaderIndex = 0; xrefHeaderIndex < xrefHeader.length; xrefHeaderIndex++) {
    
    bytes.push(xrefHeader.charCodeAt(xrefHeaderIndex));
    
  }
  var xrefSectionHeader = '0 ' + byteOffsets.length + '\r\n';
  for (var xrefSectionHeaderIndex = 0; xrefSectionHeaderIndex < xrefSectionHeader.length; xrefSectionHeaderIndex++) {
    
    bytes.push(xrefSectionHeader.charCodeAt(xrefSectionHeaderIndex));
    
  }
  for (var byteOffsetIndex = 0; byteOffsetIndex < byteOffsets.length; byteOffsetIndex++) {
    
    for (var byteOffsetStringIndex = 0; byteOffsetStringIndex < 10; byteOffsetStringIndex++) {
      
      bytes.push(Utilities.formatString('%010d', byteOffsets[byteOffsetIndex]).charCodeAt(byteOffsetStringIndex));
      
    }
    bytes.push(' '.charCodeAt(0));
    if (byteOffsetIndex == 0) {
      
      for (var generationStringIndex = 0; generationStringIndex < 5; generationStringIndex++) {
        
        bytes.push('65535'.charCodeAt(generationStringIndex));
        
      }
      for (var keywordIndex = 0; keywordIndex < 2; keywordIndex++) {
        
        bytes.push(' f'.charCodeAt(keywordIndex));
        
      }
      
    } else {
      
      for (var generationStringIndex = 0; generationStringIndex < 5; generationStringIndex++) {
        
        bytes.push('0'.charCodeAt(0));
        
      }
      for (var keywordIndex = 0; keywordIndex < 2; keywordIndex++) {
        
        bytes.push(' n'.charCodeAt(keywordIndex));
        
      }
    
    }
    bytes.push('\r'.charCodeAt(0));
    bytes.push('\n'.charCodeAt(0));
    
  }
  for (var trailerHeaderIndex = 0; trailerHeaderIndex < 9; trailerHeaderIndex++) {
    
    bytes.push('trailer\r\n'.charCodeAt(trailerHeaderIndex));
    
  }
  var idBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, (new Date).toString());
  var id = '';
  for (var idByteIndex = 0; idByteIndex < idBytes.length; idByteIndex++) {
    
    id = id + ('0' + (idBytes[idByteIndex] & 0xFF).toString(16)).slice(-2);
    
  }
  var trailer = '<</Size ' + (byteOffsets.length) + ' /Root 1 0 R /Info 2 0 R /ID [<' + id + '> <' + id + '>]>>\r\nstartxref\r\n' + xrefByteOffset + '\r\n%%EOF';
  for (var trailerIndex = 0; trailerIndex < trailer.length; trailerIndex++) {
    
    bytes.push(trailer.charCodeAt(trailerIndex));
    
  }
  return directory.createFile(Utilities.newBlob(bytes, 'application/pdf', name));
  function getObject(objectAddress, objectByteOffsets, bytes) {
    
    objectAddress = objectAddress.split(/\s+/);
    for (var addressIndex = 0; addressIndex < 2; addressIndex++) {
      
      objectAddress[addressIndex] = +objectAddress[addressIndex];
      
    }
    var object = [];
    var byteIndex = 0;
    for (var index in objectByteOffsets){
      
      var offset = objectByteOffsets[index];
      
      
      if (arrayEquals(objectAddress, offset[0])) {
        
        byteIndex = offset[1];
        break;
        
      }
      
    }
    object.push('');
    while (object[0].indexOf('endobj') <= -1) {
      
      if (/stream\s*\n/.test(object[0])) {
        
        var streamLength;
        var lengthFinder = object[0].slice(object[0].indexOf(/\/Length/.exec(object[0])[0]));
        if (/\/Length\s*\d+\s+\d+\s+R/.test(lengthFinder)) {
          
          var lengthObjectAddress = /\d+\s+\d+\s+R/.exec(/\/Length\s*\d+\s+\d+\s+R/.exec(lengthFinder)[0])[0].split(/\s+/);
          lengthObjectAddress[2] = 'obj';
          for (var addressIndex = 0; addressIndex < 2; addressIndex++) {
            
            lengthObjectAddress[addressIndex] = +lengthObjectAddress[addressIndex];
            
          }
          var lengthObject = ''
          var lengthByteIndex = 0;
          for (var index in objectByteOffsets) {
            
            var offset = objectByteOffsets[index];
            if (arrayEquals(lengthObjectAddress, offset[0])) {
              
              lengthByteIndex = offset[1];
              break;
              
            }
            
          }
          while (lengthObject.indexOf('endobj') <= -1) {
            
            lengthObject = lengthObject + String.fromCharCode(bytes[lengthByteIndex]);
            lengthByteIndex++;
            
          }
          streamLength = +(lengthObject.match(/obj\s*\n\s*\d+\s*\n\s*endobj/)[0].match(/\d+/)[0]);
          
        } else {
          
          streamLength = +(/\d+/.exec(lengthFinder)[0]);
          
        }
        var streamBytes = bytes.slice(byteIndex, byteIndex + streamLength);
        object.push(streamBytes);
        byteIndex += streamLength;
        while (object[0].indexOf('endobj') <= -1) {
          
          object[0] = object[0] + String.fromCharCode(bytes[byteIndex]);
          byteIndex++;
          
        }
        return object;
        
      }
      object[0] = object[0] + String.fromCharCode(bytes[byteIndex]);
      byteIndex++;
      
    }
    return object[0];
    
  }
  function arrayEquals(array1, array2) {
    
    if (array1 == array2) {
      
      return true;
      
    }
    if (array1 == null && array2 == null) {
      
      return true;
      
    } else if (array1 == null || array2 == null) {
      
      return false;
      
    }
    if (array1.length != array2.length) {
      
      return false;
      
    }
    for (var index = 0; index < array1.length; index++) {
      
      if (Array.isArray(array1[index])) {
        
        if (!arrayEquals(array1[index], array2[index])) {
          
          return false;
          
        }
        continue;
        
      }
      if (array1[index] != array2[index]) {
        
        return false;
        
      }
      
    }
    return true;
    
  }
  function getDependencies(objectString, objectByteOffsets, bytes) {
    
    var dependencies = [];
    var references = objectString.match(/\d+\s+\d+\s+R/g);
    if (references != null) {
      
      while (references.length > 0) {
        
        if (/\/Parent/.test(objectString.slice(objectString.indexOf(references[0]) - 8, objectString.indexOf(references[0])))) {
          
          references.shift();
          continue;
          
        }
        var dependency = getObject(references.shift().replace('R', 'obj'), objectByteOffsets, bytes);
        var dependencyExists = false;
        for (var index in dependencies) {
          
          var entry = dependencies[index];
          dependencyExists = (arrayEquals(dependency, entry)) ? true : dependencyExists;
          
        }
        if (!dependencyExists) {
          
          dependencies.push(dependency);
          
        }
        if (Array.isArray(dependency)) {
          
          dependencies = dependencies.concat(getDependencies(dependency[0], objectByteOffsets, bytes));
          
        } else {
          
          dependencies = dependencies.concat(getDependencies(dependency, objectByteOffsets, bytes));
          
        }
        
      }
      
    }
    return dependencies;
    
  }
  
}

/**
 * セル編集時の自動トリガー（シンプルトリガー）
 * 「基礎データ」シート of J1セルにAPIキーが入力された場合、プロパティサービスに保存して表示をマスクします。
 */
function onEdit(e) {
  if (!e) return;
  try {
    const range = e.range;
    const sheet = range.getSheet();
    
    // 「基礎データ」シート of J1セル（1行目、10列目） of 編集を判定
    if (sheet.getName() === "基礎データ" && range.getRow() === 1 && range.getColumn() === 10) {
      const val = range.getValue().toString().trim();
      
      // すでに「格納済み」と表示されている、または空の場合は処理しない
      if (val && !val.includes('格納済み')) {
        // 改行、スペース、カンマ、セミコロン、引用符、スラッシュなどの区切り文字で分割
        // ピリオド (.) は新規格のキー（AQ.〜）に含まれるため、区切り文字から除外します。
        const keys = val.split(/[\s,，、;；:："'"`\/\n\r\t]+/)
                        .map(k => k.trim())
                        .filter(k => k.length >= 35 && k.length <= 70 && /^[a-zA-Z0-9_.-]+$/.test(k));
        
        if (keys.length > 0) {
          // 秘密 of 内部（ScriptPropertiesおよびUserProperties）に格納
          PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEYS', JSON.stringify(keys));
          PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', keys[0]);
          PropertiesService.getUserProperties().setProperty('GEMINI_API_KEY', keys[0]);
          
          // 全角数字に変換して「〇本格納済み」と表示
          const numStr = toFullWidth(keys.length);
          range.setValue(`${numStr}本格納済み`);
          
          e.source.toast(`APIキーを${keys.length}本、秘密 of 内部に格納しました。`, '🔑 格納完了');
        }
      }
    }
  } catch (err) {
    console.error("onEdit error:", err);
  }
}

/**
 * 半角数字を全角数字に変換するヘルパー関数
 */
function toFullWidth(num) {
  const halfWidth = String(num);
  return halfWidth.replace(/[0-9]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) + 0xFEE0);
  });
}


/**
 * スプレッドシートが開かれたときに自動で実行される関数（シンプルトリガー）
 * カスタムメニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 PDF読込ツール')
    .addItem('📝 児童のPDF振り返りを取り込む', 'importPdfReflections')
    .addItem('🔍 データの再確認（異なるセルを着色）', 'recheckPdfReflections')
    .addItem('🎨 セルの着色をクリアする', 'clearAllHighlights')
    .addSeparator()
    .addItem('🔑 APIキーを登録・更新する', 'registerApiKeyFromSheet')
    .addItem('👁️ APIキーの登録状況を確認する', 'showCurrentApiKeyStatus')
    .addSeparator()
    .addItem('📁 バックアップ作成', 'createBackup')
    .addToUi();
}

/**
 * 「基礎データ」シートのJ1セルに入力されたAPIキーを手動で登録・更新します。
 * メニューから手動実行するため、権限エラーを回避し、確実に入力状況を確認できます。
 */
function registerApiKeyFromSheet() {
  const ui = SpreadsheetApp.getUi();
  const info = getSpreadsheetInfo();
  
  const keySheet = info.ss.getSheetByName("基礎データ");
  if (!keySheet) {
    ui.alert("エラー", "「基礎データ」シートが見つかりません。", ui.ButtonSet.OK);
    return;
  }
  
  const range = keySheet.getRange("J1");
  const val = range.getValue().toString().trim();
  
  if (!val) {
    ui.alert("エラー", "J1セルが空です。APIキーを入力してから実行してください。", ui.ButtonSet.OK);
    return;
  }
  
  if (val.includes("格納済み")) {
    ui.alert("情報", "すでにAPIキーは格納されています。キーを変更したい場合は、J1セルに新しいAPIキーを入力してから、再度このメニューを実行してください。", ui.ButtonSet.OK);
    return;
  }
  
  // 改行、スペース、カンマなどの区切り文字で分割してキー配列を抽出
  const keys = val.split(/[\s,，、;；:："'"`\/\n\r\t]+/)
                  .map(k => k.trim())
                  .filter(k => k.length >= 35 && k.length <= 70 && /^[a-zA-Z0-9_.-]+$/.test(k));
  
  if (keys.length === 0) {
    ui.alert("エラー", "有効なAPIキーが見つかりませんでした。\nキーは35文字〜70文字の英数字・記号である必要があります。入力内容を確認してください。\n（現在の入力長: " + val.length + "文字）", ui.ButtonSet.OK);
    return;
  }
  
  try {
    // 秘密の内部ストレージ（ScriptPropertiesおよびUserProperties）に格納
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEYS', JSON.stringify(keys));
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', keys[0]);
    PropertiesService.getUserProperties().setProperty('GEMINI_API_KEY', keys[0]);
    
    // 全角数字に変換して「〇本格納済み」と表示を更新
    const numStr = toFullWidth(keys.length);
    range.setValue(`${numStr}本格納済み`);
    
    ui.alert("登録完了", `APIキーを ${keys.length} 件、正常に登録・更新しました。\n以前の古いキーは上書きクリアされ、新しいキーが有効になりました。`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("エラー", "APIキーの保存中にエラーが発生しました:\n" + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 現在登録されているAPIキーの数と、マスクされたキーの内容を表示します。
 */
function showCurrentApiKeyStatus() {
  const ui = SpreadsheetApp.getUi();
  let keys = [];
  
  try {
    const keysStr = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEYS');
    if (keysStr) {
      keys = JSON.parse(keysStr);
    }
  } catch (e) {
    ui.alert("エラー", "登録情報の読み込みに失敗しました: " + e.toString(), ui.ButtonSet.OK);
    return;
  }
  
  // 互換性フォールバック
  if (keys.length === 0) {
    const singleKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') ||
                      PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
    if (singleKey) {
      keys = [singleKey];
    }
  }
  
  if (keys.length === 0) {
    ui.alert("APIキー状況", "現在、登録されているAPIキーはありません。\n「基礎データ」シートのJ1セルにキーを入力し、メニューから登録を行ってください。", ui.ButtonSet.OK);
    return;
  }
  
  let msg = `現在、${keys.length}本のAPIキーが登録されています。\n\n`;
  keys.forEach((key, index) => {
    const masked = key.substring(0, 6) + "..." + key.substring(key.length - 4);
    msg += `${index + 1}本目: ${masked} (長さ: ${key.length}文字)\n`;
  });
  
  msg += "\n※キーを変更または追加したい場合は、「基礎データ」シートのJ1セルに新しいキーを貼り付け、メニューから再度「登録・更新」を実行してください。";
  ui.alert("🔑 APIキー登録状況", msg, ui.ButtonSet.OK);
}


/**
 * スプレッドシートと同じフォルダに「バックアップ」フォルダを作成し、コピーを作成する関数
 */
function createBackup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    const fileId = spreadsheet.getId();
    const file = DriveApp.getFileById(fileId);
    const parents = file.getParents();
    
    if (!parents.hasNext()) {
      ui.alert('エラー: スプレッドシートの親フォルダが見つかりません。');
      return;
    }
    
    const parentFolder = parents.next();
    const folderName = "バックアップ";
    
    // 「バックアップ」フォルダが存在するか確認、なければ作成
    const folders = parentFolder.getFoldersByName(folderName);
    let backupFolder;
    if (folders.hasNext()) {
      backupFolder = folders.next();
    } else {
      backupFolder = parentFolder.createFolder(folderName);
    }
    
    // 現在日時を取得してファイル名を作成
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const backupName = `${spreadsheet.getName()}_バックアップ_${timestamp}`;
    
    // バックアップを作成
    spreadsheet.toast('バックアップを作成しています...', '📁 処理中');
    file.makeCopy(backupName, backupFolder);
    spreadsheet.toast('バックアップを作成しました！', '✅ 完了', 5);
    ui.alert('完了', 'バックアップの作成が完了しました。', ui.ButtonSet.OK);
    
  } catch (e) {
    console.error(e);
    ui.alert('バックアップ作成エラー: ' + e.message);
  }
}

/**
 * シート上の判定用背景着色をすべてクリア（初期化）します。
 */
function clearAllHighlights() {
  const ui = SpreadsheetApp.getUi();
  const info = getSpreadsheetInfo();
  const sheet = info.ss.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (!sheetName.includes("回目")) {
    ui.alert("エラー", "現在アクティブなシート（" + sheetName + "）はクリア対象外です。「１回目」「２回目」「３回目」などのシートを開いた状態で実行してください。", ui.ButtonSet.OK);
    return;
  }
  
  const startRow = 7;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < startRow || lastCol < 4) {
    info.ss.toast("クリア対象のデータセルがありません。", "⚠️ 情報", 3);
    return;
  }
  
  // 選択肢付きの確認メッセージを表示
  const response = ui.alert(
    "データのクリア確認",
    "どのようなクリア処理を行いますか？\n\n" +
    "【はい (YES)】: 「データ値（回答・振り返り等）とセルの着色」の両方を完全に消去する（枠線やプルダウンは残ります）\n" +
    "【いいえ (NO)】: 「セルの着色（オレンジ色等）」のみをクリアし、データ値は残す\n" +
    "【キャンセル】: 処理を中止する",
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  const targetRange = sheet.getRange(startRow, 4, lastRow - startRow + 1, lastCol - 3);
  
  if (response === ui.Button.YES) {
    // データ値と背景色の両方をクリア
    targetRange.clearContent();
    targetRange.setBackground(null);
    info.ss.toast("データと着色をすべてクリアしました。", "🎨 完全クリア完了", 3);
  } else if (response === ui.Button.NO) {
    // 背景色のみクリア
    targetRange.setBackground(null);
    info.ss.toast("セルの着色をすべてクリアしました（データ値は残っています）。", "🎨 クリア完了", 3);
  } else {
    info.ss.toast("処理をキャンセルしました。", "⚠️ 中止");
  }
}

/**
 * 登録されたGemini APIキーを交互に切り替えるローテータークラス
 */
class ApiKeyRotator {
  constructor() {
    this.keys = [];
    try {
      const keysStr = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEYS');
      if (keysStr) {
        this.keys = JSON.parse(keysStr);
      }
    } catch (e) {
      console.warn("Failed to parse GEMINI_API_KEYS:", e);
    }
    
    // 互換性のための単一キーフォールバック
    if (this.keys.length === 0) {
      const singleKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') ||
                        PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
      if (singleKey) {
        this.keys = [singleKey];
      }
    }
    
    if (this.keys.length === 0) {
      throw new Error("APIキーが設定されていません。基礎データシートのJ1セルに入力して保存してください。");
    }
    
    this.currentIndex = 0;
  }
  
  getNextKey() {
    const key = this.keys[this.currentIndex];
    this.currentIndex = (this.currentIndex + 1) % this.keys.length;
    return key;
  }
  
  getKeyCount() {
    return this.keys.length;
  }
}

/**
 * J2セルで指定されたPDFから手書きデータを読み込み、「１回目」シートへ取り込むメイン関数
 */
/**
 * J2セルで指定されたPDFから手書きデータを読み込み、「１回目」シートへ取り込むメイン関数
 */
function importPdfReflections() {
  try {
    processPdfReflections(false).catch(err => {
      showFatalError(err);
    });
  } catch (err) {
    showFatalError(err);
  }
}

/**
 * 既存のシートデータとPDFの手書きデータを再確認（照合）し、異なる箇所を着色・上書きする関数
 */
function recheckPdfReflections() {
  try {
    processPdfReflections(true).catch(err => {
      showFatalError(err);
    });
  } catch (err) {
    showFatalError(err);
  }
}

/**
 * 処理中に発生した致命的なエラーをUIでユーザーに通知するヘルパー関数
 */
function showFatalError(err) {
  const ui = SpreadsheetApp.getUi();
  const errMsg = err.stack ? err.stack : err.toString();
  console.error("Fatal Error:", errMsg);
  ui.alert("⚠️ 致命的なエラーが発生しました", "処理中にエラーが発生しました。詳細なエラー内容は以下の通りです:\n\n" + errMsg, ui.ButtonSet.OK);
}

/**
 * PDFの読み込みと取り込み処理を行うコアロジック
 * @param {boolean} isRecheckMode 再確認モード（異なる箇所を着色）の場合はtrue
 */
async function processPdfReflections(isRecheckMode = false) {
  const ui = SpreadsheetApp.getUi();
  const info = getSpreadsheetInfo();
  
  // 1. 基礎データシートからPDFのURLを取得
  const keySheet = info.ss.getSheetByName("基礎データ");
  if (!keySheet) {
    ui.alert("エラー", "「基礎データ」シートが見つかりません。J2セルにPDFのURLを入力してください。", ui.ButtonSet.OK);
    return;
  }
  
  const urlRange = keySheet.getRange("J2");
  const pdfUrl = urlRange.getValue().toString().trim();
  if (!pdfUrl) {
    ui.alert("エラー", "基礎データシートのJ2セルにPDFのURL（Googleドライブのリンク）を入力してください。", ui.ButtonSet.OK);
    return;
  }
  
  // 2. アクティブなシートを開いて列インデックスを取得
  const sheet = info.ss.getActiveSheet();
  const sheetName = sheet.getName();
  if (!sheetName.includes("回目")) {
    ui.alert("エラー", "現在アクティブなシート（" + sheetName + "）は取り込み対象外です。「１回目」「２回目」「３回目」などの振り返り用シートを開いた状態で実行してください。", ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  const startRow = 7; // データ行は7行目から開始
  if (lastRow < startRow) {
    ui.alert("確認", "「" + sheetName + "」シートに生徒データ（7行目以降の名簿）が入力されていません。", ui.ButtonSet.OK);
    return;
  }
  
  const cols = getColumnIndices(sheet);
  if (cols.numberCol === -1) {
    ui.alert("エラー", "「" + sheetName + "」シートに出席番号列が見つかりませんでした。ヘッダー行を確認してください。", ui.ButtonSet.OK);
    return;
  }
  
  // D列（4列目）から最大28問（AE列/31列目）の質問列をスキャンして入力規則（選択肢）と質問テキストを取得
  const startCol = 4;
  const numQuestions = 28;
  const lastCol = sheet.getLastColumn();
  const questionsConfig = [];
  
  for (let q = 1; q <= numQuestions; q++) {
    const c = startCol + q - 1;
    if (c > lastCol) {
      break;
    }
    
    const rule = sheet.getRange(startRow, c).getDataValidation();
    let allowedValues = [];
    if (rule) {
      const criteria = rule.getCriteriaType();
      if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        allowedValues = rule.getCriteriaValues()[0];
      }
    }
    
    // ヘッダー行から質問テキストを取得
    const headerText = sheet.getRange(cols.headerRow, c).getValue().toString().trim();
    
    questionsConfig.push({
      colIndex: c,
      qNum: q,
      headerText: headerText,
      allowedValues: allowedValues
    });
  }
  
  // 自己評価の入力規則（選択肢）を取得
  let allowedEvaluationValues = [];
  if (cols.evalCol !== -1) {
    const rule = sheet.getRange(startRow, cols.evalCol).getDataValidation();
    if (rule) {
      const criteria = rule.getCriteriaType();
      if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        allowedEvaluationValues = rule.getCriteriaValues()[0];
      }
    }
  }

  // 判定結果をトースト表示してユーザーに通知
  
  
  // 3. 名簿マッピング（出席番号 ⇄ 行番号、名簿上の名前）を構築
  const studentRowMap = {};
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowNum = i + startRow;
    const sNum = parseStudentNumber(values[i][cols.numberCol - 1]);
    if (!isNaN(sNum)) {
      studentRowMap[sNum] = {
        rowNum: rowNum,
        nameInSheet: cols.nameCol !== -1 ? values[i][cols.nameCol - 1].toString().trim() : ""
      };
    }
  }

  // 4. APIキーローテーターの初期化
  let rotator;
  try {
    rotator = new ApiKeyRotator();
  } catch (e) {
    ui.alert("エラー", e.message, ui.ButtonSet.OK);
    return;
  }
  
  // 判定結果をトースト表示してユーザーに通知
  info.ss.toast(`列判定: 番号=${cols.numberCol}列, 質問数=${questionsConfig.length}問, 登録キー数=${rotator.getKeyCount()}本`, "📊 準備完了", 5);
  
  // 5. PDFのダウンロード
  info.ss.toast("PDFファイルを読み込んでいます...", "⏳ 読み込み中");
  
  let pdfBlob;
  try {
    pdfBlob = getPdfBlobFromUrl(pdfUrl);
  } catch (e) {
    ui.alert("エラー", "PDFファイルの読み込みに失敗しました:\n" + e.message, ui.ButtonSet.OK);
    return;
  }
  
  const pdfBytes = new Uint8Array(pdfBlob.getBytes());
  
  // PDFの総ページ数を確認
  let totalPages = 0;
  try {
    const lib = loadPdfLib();
    const pdfDoc = await lib.PDFDocument.load(pdfBytes);
    totalPages = pdfDoc.getPageCount();
    info.ss.toast(`PDFの読み込み完了: 総ページ数=${totalPages}ページ`, "📄 PDF解析結果", 5);
  } catch (e) {
    ui.alert("エラー", "PDFの解析に失敗しました:\n" + e.toString(), ui.ButtonSet.OK);
    return;
  }
  
  const numPagePairs = Math.floor(totalPages / 2);
  const BATCH_SIZE = 5; // 1回につき5人分（10ページ）をまとめて処理
  let processedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;
  const skipDetails = [];
  
  // 6. PDFのページペアをバッチごとにスキャン実行
  for (let pairIndex = 0; pairIndex < numPagePairs; pairIndex += BATCH_SIZE) {
    const currentBatchSize = Math.min(BATCH_SIZE, numPagePairs - pairIndex);
    const pageIndices = [];
    for (let b = 0; b < currentBatchSize; b++) {
      const idx = pairIndex + b;
      pageIndices.push(idx * 2, idx * 2 + 1);
    }
    
    const pageRangeStr = `${pageIndices[0] + 1}〜${pageIndices[pageIndices.length - 1] + 1}`;
    info.ss.toast(`PDFの ${pageRangeStr} ページ目をスキャン中（${currentBatchSize}人分）...`, "⏳ 処理中");
    
    try {
      // PDFから複数ページをスライス抽出
      const extractedBytes = await extractPdfPages(pdfBytes, pageIndices);
      const base64Pdf = Utilities.base64Encode(extractedBytes);
      
      // 質問ごとのプロンプト文を構築
      let questionsPromptList = "";
      questionsConfig.forEach(cfg => {
        questionsPromptList += `- 質問番号 ${cfg.qNum}: 「${cfg.headerText}」\n`;
      });

      // 振り返り・自己評価の指示書を動的に作成
      let optionalInstructions = "";
      if (cols.contentCol !== -1) {
        optionalInstructions += `4. 振り返り記述（感想や記述内容）があれば抽出してください。\n`;
      }
      if (cols.evalCol !== -1) {
        optionalInstructions += `5. 自己評価欄（◎・◯・△や数値など）の記入があれば抽出してください。\n`;
      }

      // 動的JSONスキーマの構築
      const answersProperties = {};
      const answersRequired = [];
      questionsConfig.forEach(cfg => {
        const propKey = cfg.qNum.toString();
        answersRequired.push(propKey);
        if (cfg.allowedValues && cfg.allowedValues.length > 0) {
          answersProperties[propKey] = {
            type: "STRING",
            enum: cfg.allowedValues,
            description: `Answer for Question ${cfg.qNum}`
          };
        } else {
          answersProperties[propKey] = {
            type: "STRING",
            description: `Answer for Question ${cfg.qNum}`
          };
        }
      });

      const studentSchema = {
        type: "OBJECT",
        properties: {
          detectedNumber: {
            type: "INTEGER",
            description: "The handwritten student attendance number (出席番号) read from the worksheet header (e.g., 1, 2, 15, 24). Do not guess or copy placeholders."
          },
          detectedName: {
            type: "STRING",
            description: "The handwritten student name read from the worksheet header."
          },
          answers: {
            type: "OBJECT",
            properties: answersProperties,
            required: answersRequired
          }
        },
        required: ["detectedNumber", "detectedName", "answers"]
      };

      if (cols.contentCol !== -1) {
        studentSchema.properties.reflection = {
          type: "STRING",
          description: "Student's handwritten reflection text (振り返り内容). If empty, output empty string."
        };
      }
      if (cols.evalCol !== -1) {
        if (allowedEvaluationValues && allowedEvaluationValues.length > 0) {
          studentSchema.properties.evaluation = {
            type: "STRING",
            enum: allowedEvaluationValues
          };
        } else {
          studentSchema.properties.evaluation = {
            type: "STRING"
          };
        }
      }

      const responseSchema = {
        type: "OBJECT",
        properties: {
          students: {
            type: "ARRAY",
            description: `An array containing exactly ${currentBatchSize} items, corresponding to the ${currentBatchSize} students in the PDF.`,
            items: studentSchema
          }
        },
        required: ["students"]
      };

      // Geminiプロンプトの構成
      const prompt = `This PDF represents ${currentBatchSize * 2} pages of handwritten worksheets by ${currentBatchSize} elementary school students.
Each student has 2 consecutive pages (e.g. Pages 1-2 is Student A, Pages 3-4 is Student B, and so on).

Please perform the following tasks for each of the ${currentBatchSize} students in order:
1. Read the handwritten student number (出席番号) written in the header of the worksheet. (Usually located in a field like "( )年 ( )組 ( N )番" - look for the number N inside the parentheses or near "番").
2. Read the handwritten student name written in the header of the worksheet. (Usually next to "名前" or "なまえ").
3. Extract the checked option for each of the 28 questions (numbered 1 to 28).
${optionalInstructions}

=== IMPORTANT INSTRUCTIONS ===
- The student is a 1st to 3rd-grade elementary school child. Their circles/marks inside the square boxes "[ ]" might be messy, faint, or slightly offset.
- Please carefully examine each square box "[ ]" and its immediate surroundings for any handwritten circles '○', check marks '✓', dots, or scribbles that indicate a selection.
- Only one box is checked per question. Choose the option that has the clearest handwritten mark.
- If a box is checked in the PDF, map it to the corresponding allowed option defined in the answers properties enum.
- The worksheet has an example question labeled "(れい)" or "（れい）" (e.g. "さんすうの　じゅぎょうは　たのしい　ですか。") at the top. Do NOT extract the answer for this example. The example question is NOT Question 1.
- The actual 28 questions start below the example section and are numbered 1 to 28.
- Below is the list of questions:
${questionsPromptList}

=== OUTPUT FORMAT ===
Please output your findings in a single, strictly valid JSON object conforming exactly to the responseSchema definition.
Do not include any other text, markdown blocks, or explanation. The response must be valid JSON.`;

      // API呼び出し（ローテーション＆リトライ付）
      const result = callGeminiWithRetry(rotator, prompt, base64Pdf, responseSchema);
      const studentsData = result.students || [];

      if (!Array.isArray(studentsData)) {
        throw new Error("APIから返されたデータ構造が不正です（studentsが配列ではありません）。");
      }

      // 各生徒のデータを順に処理
      for (let s = 0; s < studentsData.length; s++) {
        const studentResult = studentsData[s];
        const detectedNum = parseStudentNumber(studentResult.detectedNumber);
        const detectedName = studentResult.detectedName ? studentResult.detectedName.toString().trim() : "";
        
        if (isNaN(detectedNum)) {
          skippedCount++;
          skipDetails.push(`バッチ内順序 ${s + 1}人目: 出席番号の認識失敗（AI判定値: "${studentResult.detectedNumber}"）`);
          continue;
        }
        
        // 出席番号が有効な数値かチェック
        if (detectedNum < 1) {
          skippedCount++;
          skipDetails.push(`出席番号 ${detectedNum}番: 無効な番号のためスキップ`);
          continue;
        }
        
        // 名簿マップから該当行情報を取得
        const matchInfo = studentRowMap[detectedNum];
        if (!matchInfo) {
          skippedCount++;
          skipDetails.push(`${detectedNum}番 (${detectedName || "名前不明"}): 名簿に見つからない`);
          continue;
        }
        
        const rowNum = matchInfo.rowNum;
        const rowIdx = rowNum - startRow;
        const rowData = values[rowIdx];
        
        // 1. 在籍確認（B列にメールアドレス「@」が入っているか）
        let studentExists = false;
        if (rowData) {
          const bCellVal = rowData[1]; // B列（インデックス1）の値
          if (bCellVal !== undefined && bCellVal !== null && bCellVal.toString().includes("@")) {
            studentExists = true;
          }
        }
        
        if (!studentExists) {
          skippedCount++;
          skipDetails.push(`${detectedNum}番: B列にメールアドレスがないため無視（児童が存在しません）`);
          continue;
        }
        
        // 2. すでにデータが書き込まれている場合はスキップ（通常取り込みモードのみ）
        if (!isRecheckMode) {
          let hasExistingData = false;
          
          if (rowData) {
            if (questionsConfig.length > 0) {
              const firstQCol = questionsConfig[0].colIndex;
              const cellVal = rowData[firstQCol - 1];
              if (cellVal !== undefined && cellVal !== null && cellVal.toString().trim() !== "") {
                hasExistingData = true;
              }
            }
            if (!hasExistingData && cols.contentCol >= 4) {
              const contentVal = rowData[cols.contentCol - 1];
              if (contentVal !== undefined && contentVal !== null && contentVal.toString().trim() !== "") {
                hasExistingData = true;
              }
            }
          }
          
          if (hasExistingData) {
            skippedCount++;
            skipDetails.push(`${detectedNum}番 (${matchInfo.nameInSheet || detectedName}): すでにデータ値が書き込み済み`);
            continue;
          }
        }
        
        info.ss.toast(`出席番号 ${detectedNum} 番 (${matchInfo.nameInSheet || detectedName}) のデータを書き込み中...`, "⏳ 処理中");
        
        // セル書込＆比較関数
        const writeValue = (cellCol, val, allowedList) => {
          let finalVal = val !== undefined && val !== null ? val.toString().trim() : "";
          if (finalVal && allowedList && allowedList.length > 0) {
            finalVal = findClosestMatch(finalVal, allowedList);
          }
          
          const cellRange = sheet.getRange(rowNum, cellCol);
          const currentVal = cellRange.getValue().toString().trim();
          
          if (isRecheckMode) {
            if (finalVal !== currentVal) {
              if (currentVal === "") {
                cellRange.setValue(finalVal);
                cellRange.setBackground(null);
              } else {
                cellRange.setBackground("#ffe599");
              }
            } else {
              cellRange.setBackground(null);
            }
          } else {
            if (finalVal) {
              cellRange.setValue(finalVal);
            }
          }
        };
        
        // 各質問の書き込み
        questionsConfig.forEach(cfg => {
          const answerVal = studentResult.answers && studentResult.answers[cfg.qNum.toString()];
          writeValue(cfg.colIndex, answerVal, cfg.allowedValues);
        });
        
        // 振り返りの書き込み
        if (cols.contentCol !== -1) {
          writeValue(cols.contentCol, studentResult.reflection, []);
        }
        
        // 自己評価の書き込み
        if (cols.evalCol !== -1) {
          writeValue(cols.evalCol, studentResult.evaluation, allowedEvaluationValues);
        }
        
        processedCount++;
      }
      
      // 即時反映
      SpreadsheetApp.flush();
      
      // API制限回避のウェイト（無料枠の15RPM制限を回避するため6秒に設定）
      Utilities.sleep(6000);
      
    } catch (err) {
      console.error(`Error processing batch starting at page pair ${pairIndex}:`, err);
      errorCount++;
      info.ss.toast(`ページ ${pageIndices[0] + 1}〜${pageIndices[pageIndices.length - 1] + 1} の読み取りバッチ処理に失敗しました: ${err.message}`, "❌ エラー", 5);
      SpreadsheetApp.flush();
    }
  }
  
  // 終了通知
  const modeStr = isRecheckMode ? "データの再確認" : "PDFデータの取り込み";
  let detailMsg = "";
  if (skipDetails.length > 0) {
    detailMsg = "\n\n【スキップの詳細（先頭20件まで）】\n" + skipDetails.slice(0, 20).join("\n");
    if (skipDetails.length > 20) {
      detailMsg += `\n...他 ${skipDetails.length - 20} 件がスキップされました。`;
    }
  }
  ui.alert("処理完了", `${modeStr}が完了しました。\n\n・処理人数: ${processedCount}人分\n・スキップ: ${skippedCount}件\n・エラー件数: ${errorCount}件${detailMsg}`, ui.ButtonSet.OK);
}

/**
 * シート内の列インデックスをヘッダー名から自動特定するヘルパー
 */
function getColumnIndices(sheet) {
  let cols = { numberCol: -1, nameCol: -1, contentCol: -1, evalCol: -1, headerRow: 6 };
  
  // 1行目から6行目まで順次チェックし、最もマッチ度の高い行をヘッダー行とする
  let bestHeaderRow = 6;
  let maxMatchedCount = -1;
  const numCols = sheet.getLastColumn();
  
  for (let r = 1; r <= 6; r++) {
    const headers = sheet.getRange(r, 1, 1, numCols).getValues()[0];
    const currentCols = detectCols(headers);
    
    // 番号列が見つかることが最優先
    let score = 0;
    if (currentCols.numberCol !== -1) score += 10;
    if (currentCols.nameCol !== -1) score += 5;
    if (currentCols.contentCol !== -1) score += 3;
    if (currentCols.evalCol !== -1) score += 2;
    
    if (score > maxMatchedCount) {
      maxMatchedCount = score;
      bestHeaderRow = r;
      cols = currentCols;
      cols.headerRow = r;
    }
  }
  
  // もしキーワードから番号列が見つからなかった場合のデータスキャンによるフォールバック
  if (cols.numberCol === -1) {
    const scanStartRow = 7;
    const scanEndRow = Math.min(sheet.getLastRow(), 15);
    if (scanEndRow >= scanStartRow) {
      const scanValues = sheet.getRange(scanStartRow, 1, scanEndRow - scanStartRow + 1, numCols).getValues();
      
      let bestCol = -1;
      let maxNumericCount = 0;
      
      for (let c = 0; c < numCols; c++) {
        let numericCount = 0;
        for (let r = 0; r < scanValues.length; r++) {
          const val = scanValues[r][c];
          const num = parseStudentNumber(val);
          if (!isNaN(num) && num > 0 && num < 100) {
            numericCount++;
          }
        }
        if (numericCount > maxNumericCount) {
          maxNumericCount = numericCount;
          bestCol = c + 1;
        }
      }
      
      if (maxNumericCount >= (scanValues.length / 2)) {
        cols.numberCol = bestCol;
        // 名前列が未検出の場合は番号列の隣を仮指定
        if (cols.nameCol === -1 && bestCol < numCols) {
          cols.nameCol = bestCol + 1;
        }
      }
    }
  }
  
  return cols;
}

/**
 * ヘッダー配列から各列の位置を特定する内部ヘルパー
 */
function detectCols(headers) {
  let numberCol = -1;
  let nameCol = -1;
  let contentCol = -1;
  let evalCol = -1;
  
  for (let i = 0; i < headers.length; i++) {
    const colIndex = i + 1;
    // 質問列（4〜31列目）は、振り返り列や自己評価列、番号列として誤判定されないように除外
    if (colIndex >= 4 && colIndex <= 31) {
      continue;
    }
    
    const header = headers[i].toString().trim().toLowerCase();
    
    // 豊富なマッチングパターンを適用
    if (header.includes("番号") || header.includes("出席") || header === "no" || header.startsWith("no.") || header === "id" || header === "番" || header === "№" || header === "code" || header.includes("コード")) {
      numberCol = colIndex;
    } else if (header.includes("名前") || header.includes("氏名") || header.includes("なまえ") || header.includes("氏") || header === "name" || header.includes("児童")) {
      nameCol = colIndex;
    } else if (header.includes("振り返り") || header.includes("内容") || header.includes("感想") || header.includes("記述") || header.includes("所見") || header.includes("記入") || header.includes("ふりかえり")) {
      contentCol = colIndex;
    } else if (header.includes("評価") || header.includes("じこ") || header.includes("自己") || header === "score" || header.includes("ひょうか")) {
      evalCol = colIndex;
    }
  }
  return { numberCol, nameCol, contentCol, evalCol };
}

/**
 * URLからPDFのBlobを取得する。フォルダURLの場合は最初のファイルを対象とする。
 */
function getPdfBlobFromUrl(url) {
  // file IDのパターンマッチ
  const fileIdMatch = url.match(/\/file\/d\/([-\w]{25,})/);
  if (fileIdMatch) {
    return DriveApp.getFileById(fileIdMatch[1]).getBlob();
  }
  
  // foldersのパターンマッチ
  const folderIdMatch = url.match(/\/folders\/([-\w]{25,})/);
  if (folderIdMatch) {
    const folder = DriveApp.getFolderById(folderIdMatch[1]);
    const files = folder.getFilesByType(MimeType.PDF);
    if (files.hasNext()) {
      return files.next().getBlob();
    }
    throw new Error("指定されたフォルダ内にPDFファイルが見つかりませんでした。");
  }
  
  // 汎用IDの抽出と試行
  const idMatch = url.match(/[-\w]{25,}/);
  if (idMatch) {
    try {
      return DriveApp.getFileById(idMatch[0]).getBlob();
    } catch (e) {
      try {
        const folder = DriveApp.getFolderById(idMatch[0]);
        const files = folder.getFilesByType(MimeType.PDF);
        if (files.hasNext()) {
          return files.next().getBlob();
        }
      } catch (err) {
        throw new Error("URLからファイルまたはフォルダを特定できませんでした: " + e.toString());
      }
    }
  }
  
  throw new Error("有効なGoogleドライブのURL（ファイルまたはフォルダ）が見つかりませんでした。");
}

// Global cache for PDFLib to avoid re-evaluating
let PDFLibInstance = null;

/**
 * pdf-lib ライブラリをCDN（またはGoogleドライブキャッシュ）からロードして初期化するヘルパー
 */
function loadPdfLib() {
  if (PDFLibInstance) return PDFLibInstance;
  
  const info = getSpreadsheetInfo();
  let content = "";
  
  // Googleドライブからキャッシュの読み込みを試行
  if (info.folderId) {
    try {
      const folder = DriveApp.getFolderById(info.folderId);
      const files = folder.getFilesByName("pdf-lib-cache.txt");
      if (files.hasNext()) {
        content = files.next().getBlob().getDataAsString();
      }
    } catch (e) {
      console.warn("キャッシュの読み込みに失敗しました: " + e.toString());
    }
  }
  
  if (!content) {
    // キャッシュがない場合はCDNから取得し、裏でキャッシュファイルを作成
    const pdfLibUrl = "https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js";
    content = UrlFetchApp.fetch(pdfLibUrl).getContentText();
    
    if (info.folderId) {
      try {
        const folder = DriveApp.getFolderById(info.folderId);
        const oldFiles = folder.getFilesByName("pdf-lib-cache.txt");
        while (oldFiles.hasNext()) {
          oldFiles.next().setTrashed(true);
        }
        folder.createFile("pdf-lib-cache.txt", content);
      } catch (e) {
        console.warn("キャッシュの保存に失敗しました: " + e.toString());
      }
    }
  }
  
  // setTimeout shim for GAS environment
  const shimmedContent = "var setTimeout = function(f, t) { Utilities.sleep(t); f(); };\n" + content;
  eval(shimmedContent);
  PDFLibInstance = PDFLib;
  return PDFLibInstance;
}

/**
 * 指定されたPDFバイト配列から指定されたページ（0開始）を抽出したPDFバイト配列を返すヘルパー
 */
async function extractPdfPages(pdfBytes, pageIndices) {
  const lib = loadPdfLib();
  const pdfDoc = await lib.PDFDocument.load(pdfBytes);
  const newPdfDoc = await lib.PDFDocument.create();
  
  const copiedPages = await newPdfDoc.copyPages(pdfDoc, pageIndices);
  copiedPages.forEach(page => newPdfDoc.addPage(page));
  
  const newPdfBytes = await newPdfDoc.save();
  return newPdfBytes;
}

/**
 * APIキーを切り替えながらGemini APIをリトライ実行するヘルパー
 */
function callGeminiWithRetry(rotator, prompt, base64Pdf, schema = null) {
  let attempts = 0;
  const maxAttempts = Math.max(rotator.getKeyCount() * 2, 4);
  
  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        {
          inlineData: {
            mimeType: "application/pdf",
            data: base64Pdf
          }
        }
      ]
    }],
    generationConfig: {
      responseMimeType: "application/json",
      temperature: 0.0
    }
  };
  
  if (schema) {
    payload.generationConfig.responseSchema = schema;
  }
  
  let consecutive429s = 0;
  let lastErrorDetail = "";
  
  while (attempts < maxAttempts) {
    const apiKey = rotator.getNextKey();
    const model = 'gemini-2.5-flash';
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    
    try {
      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const text = response.getContentText();
      
      if (code === 200) {
        const json = JSON.parse(text);
        if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts) {
          const contentText = json.candidates[0].content.parts[0].text;
          return JSON.parse(contentText);
        }
      } else {
        lastErrorDetail = `HTTP ${code}: ${text}`;
        console.warn(`API call failed (Attempt ${attempts + 1}/${maxAttempts}) with status ${code}: ${text}`);
        if (code === 429) {
          consecutive429s++;
          const info = getSpreadsheetInfo();
          // 注意：同じGoogle AI Studioプロジェクトで複数キーを作成した場合、制限は共有されます。
          if (consecutive429s >= rotator.getKeyCount()) {
            info.ss.toast("すべてのAPIキーが制限に達しました。制限解除まで60秒間待機します...", "⏳ 一時停止", 60);
            console.warn("All keys rate-limited. Waiting 60 seconds...");
            Utilities.sleep(60000);
            consecutive429s = 0;
          } else {
            // 同一プロジェクト内のキーで制限が共有されている可能性があるため、10秒待機して切り替えます
            info.ss.toast("API制限(429)を検出しました。別キー切り替えのため10秒待機します...", "⏳ 待機中", 10);
            Utilities.sleep(10000);
          }
        } else {
          Utilities.sleep(2000);
        }
      }
    } catch (e) {
      lastErrorDetail = e.toString();
      console.warn(`API call exception (Attempt ${attempts + 1}/${maxAttempts}): ${e.toString()}`);
      Utilities.sleep(1000);
    }
    
    attempts++;
  }
  
  throw new Error(`Gemini APIへの問い合わせがすべてエラーまたは上限に達しました。\n最後の詳細: ${lastErrorDetail}`);
}

/**
 * セル値から出席番号（数値）を安全にパースするヘルパー
 * 全角数字、テキスト混じり（「15番」「No.15」など）に対応します。
 */
function parseStudentNumber(val) {
  if (val === null || val === undefined) return NaN;
  const str = val.toString().trim();
  if (!str) return NaN;
  
  // 全角数字を半角数字に変換
  const normalized = str.replace(/[０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  
  // 数字の部分のみを正規表現で抽出
  const match = normalized.match(/\d+/);
  return match ? Number(match[0]) : NaN;
}

/**
 * 抽出された値に最も近い入力規則（選択肢）の値を返すヘルパー
 * 入力規則にない値をセットしてエラーになるのを防ぎます。
 */
function findClosestMatch(value, allowedValues) {
  if (!value) return "";
  const cleaned = value.toString().trim();
  if (allowedValues.includes(cleaned)) return cleaned;
  
  // 記号や空白を取り除く正規化関数
  const normalize = (str) => {
    return str.toString()
              .replace(/[\s\u3000\r\n\t、。.,!?！？]/g, "")
              .toLowerCase();
  };
  
  const normalizedValue = normalize(cleaned);
  if (!normalizedValue) return "";
  
  // 1. 完全一致（正規化後）の検証
  for (const allowed of allowedValues) {
    if (normalize(allowed) === normalizedValue) {
      return allowed;
    }
  }
  
  // 2. 部分一致（正規化後）の検証
  for (const allowed of allowedValues) {
    const normalizedAllowed = normalize(allowed);
    if (normalizedValue.includes(normalizedAllowed) || normalizedAllowed.includes(normalizedValue)) {
      return allowed;
    }
  }
  
  return ""; // 一致しない場合は空文字にして入力規則違反エラーを防ぐ
}
