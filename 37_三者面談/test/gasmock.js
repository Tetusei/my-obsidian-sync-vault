/**
 * Apps Script を実際に動かさずに、Node.js だけでロジックを検証するための土台。
 *
 * 1つ上のフォルダの `.js` をそのまま読み込んで動かす。
 *
 * **実機の制約はできるだけ再現してある。** 甘くすると、テストが通るのに本番で落ちる。
 * 過去に、モックを厳しくしたことで次の不具合が見つかっている。
 *   ・シートは既定 1000行 × 26列しかない → 24クラスで範囲外エラー
 *   ・非表示シートはアクティブにできない
 *   ・clearDataValidations() は Range のメソッドで、Sheet には無い
 * 緩めたくなったら、まず実機がどうなのかを確認すること。
 */

'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const PROJECT_DIR = path.join(__dirname, '..');

/* ================================================================
   Range
   ================================================================ */

/** 書式まわりは動作に影響しないので、鎖のようにつなげられる空実装にする */
const NOOP_RANGE_METHODS = [
  'setFontWeight', 'setFontWeights',
  'setFontSize', 'setFontColor', 'setFontFamily', 'setHorizontalAlignment',
  'setHorizontalAlignments', 'setVerticalAlignment', 'setWrap', 'setBorder',
  'setNumberFormat', 'insertCheckboxes', 'clearDataValidations',
  'setRichTextValue', 'setRichTextValues', 'setShowHyperlink'
];

class MockRange {
  constructor(sheet, row, col, numRows, numCols) {
    if (!(numRows >= 1) || !(numCols >= 1)) {
      throw new Error('行数と列数は 1 以上でなければなりません: ' + numRows + '×' + numCols);
    }
    if (row < 1 || col < 1) {
      throw new Error('範囲が無効です（行・列は1から始まります）: ' + row + ',' + col);
    }
    // 実機は、シートの大きさを超える範囲を取ろうとすると例外になる
    if (row + numRows - 1 > sheet._maxRows || col + numCols - 1 > sheet._maxCols) {
      throw new Error(
        '範囲が無効です。シート「' + sheet.getName() + '」は ' +
        sheet._maxRows + '行×' + sheet._maxCols + '列ですが、' +
        (row + numRows - 1) + '行×' + (col + numCols - 1) + '列を指定しました。');
    }
    this._sheet = sheet;
    this._row = row;
    this._col = col;
    this._rows = numRows;
    this._cols = numCols;

    NOOP_RANGE_METHODS.forEach((name) => { this[name] = () => this; });
  }

  getRow() { return this._row; }
  getColumn() { return this._col; }
  getNumRows() { return this._rows; }
  getNumColumns() { return this._cols; }
  getSheet() { return this._sheet; }

  getCell(r, c) {
    return new MockRange(this._sheet, this._row + r - 1, this._col + c - 1, 1, 1);
  }

  getA1Notation() {
    const colName = (n) => {
      let s = '';
      while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
      return s;
    };
    const a = colName(this._col) + this._row;
    if (this._rows === 1 && this._cols === 1) return a;
    return a + ':' + colName(this._col + this._cols - 1) + (this._row + this._rows - 1);
  }

  getValues() {
    const out = [];
    for (let r = 0; r < this._rows; r++) {
      const row = [];
      for (let c = 0; c < this._cols; c++) {
        row.push(this._sheet._get(this._row + r, this._col + c));
      }
      out.push(row);
    }
    return out;
  }

  getValue() { return this.getValues()[0][0]; }

  setValues(values) {
    if (!Array.isArray(values) || values.length !== this._rows) {
      throw new Error('範囲の行数と、渡された配列の行数が違います: ' +
        this._rows + ' に対して ' + (values ? values.length : 'null'));
    }
    for (let r = 0; r < this._rows; r++) {
      if (!Array.isArray(values[r]) || values[r].length !== this._cols) {
        throw new Error('範囲の列数と、渡された配列の列数が違います: ' +
          this._cols + ' に対して ' + (values[r] ? values[r].length : 'null'));
      }
      for (let c = 0; c < this._cols; c++) {
        this._sheet._set(this._row + r, this._col + c, values[r][c]);
      }
    }
    return this;
  }

  setValue(v) {
    for (let r = 0; r < this._rows; r++) {
      for (let c = 0; c < this._cols; c++) this._sheet._set(this._row + r, this._col + c, v);
    }
    return this;
  }

  setFormula(f) { return this.setValue(f); }
  getFormula() { const v = this.getValue(); return typeof v === 'string' && v[0] === '=' ? v : ''; }

  clearContent() { return this.setValue(''); }

  getNotes() {
    const out = [];
    for (let r = 0; r < this._rows; r++) {
      const row = [];
      for (let c = 0; c < this._cols; c++) {
        row.push(this._sheet._notes.get((this._row + r) + ',' + (this._col + c)) || '');
      }
      out.push(row);
    }
    return out;
  }

  setNotes(notes) {
    for (let r = 0; r < this._rows; r++) {
      for (let c = 0; c < this._cols; c++) {
        this._sheet._notes.set((this._row + r) + ',' + (this._col + c), notes[r][c] || '');
      }
    }
    return this;
  }

  /**
   * 背景色は「状態を色で伝える」ための情報なので、本当に覚えておく。
   * 空実装にすると、予約済と未予約が同じ色でもテストが通ってしまう
   * （v4.6.16 の「予約済（予備）が未予約と同じ色」がまさにそれ）。
   */
  setBackground(color) {
    for (let r = 0; r < this._rows; r++) {
      for (let c = 0; c < this._cols; c++) {
        this._sheet._bg.set((this._row + r) + ',' + (this._col + c), color || '#ffffff');
      }
    }
    return this;
  }

  setBackgrounds(grid) {
    for (let r = 0; r < this._rows; r++) {
      for (let c = 0; c < this._cols; c++) {
        this._sheet._bg.set((this._row + r) + ',' + (this._col + c),
          (grid[r] && grid[r][c]) || '#ffffff');
      }
    }
    return this;
  }

  getBackgrounds() {
    const out = [];
    for (let r = 0; r < this._rows; r++) {
      const row = [];
      for (let c = 0; c < this._cols; c++) {
        row.push(this._sheet._bg.get((this._row + r) + ',' + (this._col + c)) || '#ffffff');
      }
      out.push(row);
    }
    return out;
  }

  getBackground() { return this.getBackgrounds()[0][0]; }

  merge() { this._sheet._merges.push(this.getA1Notation()); return this; }
  breakApart() {
    const a1 = this.getA1Notation();
    this._sheet._merges = this._sheet._merges.filter((m) => m !== a1);
    return this;
  }

  protect() {
    const p = new MockProtection(this._sheet, 'RANGE', this);
    this._sheet._protections.push(p);
    return p;
  }
}

/* ================================================================
   Protection
   ================================================================ */

class MockProtection {
  constructor(sheet, type, range) {
    this._sheet = sheet;
    this._type = type;
    this._range = range || null;
    this._description = '';
    this._warningOnly = false;
    this._removed = false;
  }
  getDescription() { return this._description; }
  setDescription(d) { this._description = String(d == null ? '' : d); return this; }
  setWarningOnly(b) { this._warningOnly = !!b; return this; }
  isWarningOnly() { return this._warningOnly; }
  getRange() { return this._range; }
  setRange(r) { this._range = r; return this; }
  getProtectionType() { return this._type; }
  remove() {
    this._removed = true;
    const i = this._sheet._protections.indexOf(this);
    if (i >= 0) this._sheet._protections.splice(i, 1);
    return this;
  }
}

/* ================================================================
   Sheet
   ================================================================ */

/** 新しいシートは既定 1000行 × 26列しかない。ここを緩めると本番で範囲外エラーになる */
const DEFAULT_ROWS = 1000;
const DEFAULT_COLS = 26;

let SHEET_ID_SEQ = 1;

class MockSheet {
  constructor(spreadsheet, name) {
    this._ss = spreadsheet;
    this._name = name;
    this._id = SHEET_ID_SEQ++;
    this._maxRows = DEFAULT_ROWS;
    this._maxCols = DEFAULT_COLS;
    this._cells = new Map();       // "r,c" → 値
    this._notes = new Map();
    this._bg = new Map();
    this._merges = [];
    this._protections = [];
    this._hidden = false;
    this._frozenRows = 0;
    this._frozenCols = 0;
    this._colWidths = new Map();
    this._rowHeights = new Map();
    this._hiddenCols = new Set();

    ['setColumnWidth', 'setRowHeight', 'setFrozenRows', 'setFrozenColumns',
      'hideColumns', 'setHiddenGridlines', 'setTabColor', 'autoResizeColumn']
      .forEach((n) => {
        this[n] = (...args) => {
          if (n === 'setFrozenRows') this._frozenRows = args[0];
          if (n === 'setFrozenColumns') this._frozenCols = args[0];
          if (n === 'setColumnWidth') this._colWidths.set(args[0], args[1]);
          if (n === 'setRowHeight') this._rowHeights.set(args[0], args[1]);
          if (n === 'hideColumns') this._hiddenCols.add(args[0]);
          return this;
        };
      });
  }

  _key(r, c) { return r + ',' + c; }
  _get(r, c) { const v = this._cells.get(this._key(r, c)); return v === undefined ? '' : v; }
  _set(r, c, v) { this._cells.set(this._key(r, c), v === undefined || v === null ? '' : v); }

  getName() { return this._name; }
  setName(n) { this._name = n; return this; }
  getSheetId() { return this._id; }
  getParent() { return this._ss; }
  getMaxRows() { return this._maxRows; }
  getMaxColumns() { return this._maxCols; }
  getIndex() { return this._ss._sheets.indexOf(this) + 1; }

  isSheetHidden() { return this._hidden; }
  hideSheet() { this._hidden = true; return this; }
  showSheet() { this._hidden = false; return this; }

  getLastRow() {
    let last = 0;
    for (const [k, v] of this._cells) {
      if (v === '' || v === null || v === undefined) continue;
      const r = Number(k.split(',')[0]);
      if (r > last) last = r;
    }
    return last;
  }

  getLastColumn() {
    let last = 0;
    for (const [k, v] of this._cells) {
      if (v === '' || v === null || v === undefined) continue;
      const c = Number(k.split(',')[1]);
      if (c > last) last = c;
    }
    return last;
  }

  getRange(a, b, c, d) {
    if (typeof a === 'string') throw new Error('このモックはA1形式に対応していません: ' + a);
    return new MockRange(this, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d);
  }

  getDataRange() {
    return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1));
  }

  appendRow(values) {
    const row = this.getLastRow() + 1;
    if (row > this._maxRows) this._maxRows = row;
    values.forEach((v, i) => this._set(row, i + 1, v));
    return this;
  }

  insertRowsAfter(after, howMany) { this._maxRows += howMany; return this; }
  insertColumnsAfter(after, howMany) { this._maxCols += howMany; return this; }

  clear() {
    this._cells.clear();
    this._notes.clear();
    this._bg.clear();
    this._merges = [];
    // 保護は clear() では消えない前提でモデル化している。
    // 実機の挙動が違うと分かったら、ここを直すこと
    return this;
  }

  /**
   * 行を動かす。orderConfigRows が使う「上へ引き上げる」動きを再現する。
   * @param {MockRange} rowSpec 動かす行
   * @param {number} destinationIndex 移動先の行番号
   */
  moveRows(rowSpec, destinationIndex) {
    const from = rowSpec.getRow();
    const count = rowSpec.getNumRows();
    if (from === destinationIndex) return this;

    const width = this._maxCols;
    const grab = [];
    for (let i = 0; i < count; i++) {
      const row = [];
      for (let c = 1; c <= width; c++) row.push(this._get(from + i, c));
      grab.push(row);
    }

    if (from > destinationIndex) {
      // 上へ引き上げる。あいだの行を下へずらす
      for (let r = from - 1; r >= destinationIndex; r--) {
        for (let c = 1; c <= width; c++) this._set(r + count, c, this._get(r, c));
      }
      for (let i = 0; i < count; i++) {
        for (let c = 1; c <= width; c++) this._set(destinationIndex + i, c, grab[i][c - 1]);
      }
    } else {
      for (let r = from + count; r < destinationIndex; r++) {
        for (let c = 1; c <= width; c++) this._set(r - count, c, this._get(r, c));
      }
      for (let i = 0; i < count; i++) {
        for (let c = 1; c <= width; c++) this._set(destinationIndex - count + i, c, grab[i][c - 1]);
      }
    }
    return this;
  }

  getProtections(type) {
    const want = type && type.toString ? String(type) : 'RANGE';
    return this._protections.filter((p) => p._type === want && !p._removed);
  }

  protect() {
    const p = new MockProtection(this, 'SHEET', null);
    this._protections.push(p);
    return p;
  }

  insertImage() { return { setWidth: () => ({ setHeight: () => ({}) }) }; }
}

/* ================================================================
   Spreadsheet
   ================================================================ */

class MockSpreadsheet {
  constructor(name) {
    this._name = name || '三者面談';
    this._id = 'ss_' + Math.random().toString(36).slice(2, 12);
    this._sheets = [];
    this._active = null;
    this.toasts = [];
  }

  getId() { return this._id; }
  getName() { return this._name; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this._id + '/edit'; }

  getSheets() { return this._sheets.slice(); }
  getSheetByName(name) { return this._sheets.find((s) => s.getName() === name) || null; }

  insertSheet(name) {
    const sh = new MockSheet(this, name || ('シート' + (this._sheets.length + 1)));
    this._sheets.push(sh);
    this._active = sh;
    return sh;
  }

  deleteSheet(sh) {
    const i = this._sheets.indexOf(sh);
    if (i >= 0) this._sheets.splice(i, 1);
    return this;
  }

  getActiveSheet() { return this._active || this._sheets[0] || null; }

  setActiveSheet(sh) {
    // 実機では、非表示のシートはアクティブにできない
    if (sh.isSheetHidden()) throw new Error('非表示のシートはアクティブにできません: ' + sh.getName());
    this._active = sh;
    return sh;
  }

  moveActiveSheet(pos) {
    const sh = this._active;
    const i = this._sheets.indexOf(sh);
    if (i < 0) return;
    this._sheets.splice(i, 1);
    this._sheets.splice(Math.max(0, Math.min(pos - 1, this._sheets.length)), 0, sh);
  }

  toast(msg, title, sec) { this.toasts.push({ msg: msg, title: title, sec: sec }); }
}

/* ================================================================
   そのほかのサービス
   ================================================================ */

function makeUtilities() {
  const PAD = (n, w) => ('0000' + n).slice(-w);
  return {
    formatDate(date, tz, fmt) {
      const d = new Date(date);
      // Java の SimpleDateFormat 風。MM=月 / mm=分 に注意
      return String(fmt).replace(/yyyy|MM|dd|HH|mm|ss|M|d|H|h/g, (t) => {
        switch (t) {
          case 'yyyy': return String(d.getFullYear());
          case 'MM': return PAD(d.getMonth() + 1, 2);
          case 'M': return String(d.getMonth() + 1);
          case 'dd': return PAD(d.getDate(), 2);
          case 'd': return String(d.getDate());
          case 'HH': return PAD(d.getHours(), 2);
          case 'H': case 'h': return String(d.getHours());
          case 'mm': return PAD(d.getMinutes(), 2);
          case 'ss': return PAD(d.getSeconds(), 2);
          default: return t;
        }
      });
    },
    sleep() { },
    newBlob(data, type, name) { return { getName: () => name, setName: (n) => ({ getName: () => n }) }; }
  };
}

function makeCacheService() {
  const store = new Map();
  const cache = {
    get(k) {
      const hit = store.get(k);
      if (!hit) return null;
      if (hit.until < Date.now()) { store.delete(k); return null; }
      return hit.value;
    },
    put(k, v, sec) { store.set(k, { value: String(v), until: Date.now() + (sec || 600) * 1000 }); },
    remove(k) { store.delete(k); },
    removeAll(keys) { (keys || []).forEach((k) => store.delete(k)); },
    _store: store
  };
  return { getScriptCache: () => cache, getUserCache: () => cache, _cache: cache };
}

function makePropertiesService() {
  const store = new Map();
  const props = {
    getProperty(k) { return store.has(k) ? store.get(k) : null; },
    setProperty(k, v) { store.set(k, String(v)); return props; },
    deleteProperty(k) { store.delete(k); return props; },
    getProperties() { return Object.fromEntries(store); },
    _store: store
  };
  return { getScriptProperties: () => props, getUserProperties: () => props };
}

function makeScriptApp(state) {
  return {
    getProjectTriggers() { return state.triggers.slice(); },
    deleteTrigger(t) {
      const i = state.triggers.indexOf(t);
      if (i >= 0) state.triggers.splice(i, 1);
    },
    newTrigger(fn) {
      const trigger = {
        getHandlerFunction: () => fn,
        getUniqueId: () => 'trg_' + fn + '_' + state.triggers.length
      };
      const builder = {
        forSpreadsheet: () => builder,
        timeBased: () => builder,
        onEdit: () => builder,
        onOpen: () => builder,
        atHour: () => builder,
        everyDays: () => builder,
        everyMinutes: () => builder,
        create: () => { state.triggers.push(trigger); return trigger; }
      };
      return builder;
    },
    getService: () => ({ getUrl: () => state.webAppUrl }),
    getOAuthToken: () => 'test-token'
  };
}

function makeUi(state) {
  const Button = { OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO' };
  const ButtonSet = { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO' };
  const menu = () => {
    const m = {
      addItem: () => m, addSeparator: () => m, addSubMenu: () => m, addToUi: () => m
    };
    return m;
  };
  return {
    Button, ButtonSet,
    alert(...args) { state.alerts.push(args); return state.nextButton; },
    prompt(...args) {
      state.prompts.push(args);
      return {
        getSelectedButton: () => state.nextButton,
        getResponseText: () => state.nextPromptText
      };
    },
    createMenu: menu,
    showModalDialog(...args) { state.dialogs.push(args); },
    showSidebar() { }
  };
}

/* ================================================================
   読み込み
   ================================================================ */

/**
 * プロジェクトの .js を読み込んで、テストから呼べるようにする。
 *
 * 読み込む順はファイル名順。Apps Script のエディタもこの順で評価するため、
 * `Setup.js` の `DEFAULT_CONFIG_ROWS` が `Config.js` の定数を参照できる。
 * 順番を変えると、トップレベルで別ファイルの定数を使っている箇所が undefined になる。
 *
 * @param {Object} opt
 * @param {string} opt.webAppUrl ScriptApp.getService().getUrl() が返す値
 * @param {Function} opt.fetch UrlFetchApp.fetch の中身
 * @return {Object} すべてのグローバル（プロジェクトの関数を直接呼べる）
 */
function load(opt) {
  installCrashReport();
  opt = opt || {};

  const ss = new MockSpreadsheet('三者面談');
  const state = {
    triggers: [],
    alerts: [],
    prompts: [],
    dialogs: [],
    mails: [],
    warnings: [],
    lockAvailable: true,
    nextButton: 'OK',
    nextPromptText: '',
    webAppUrl: opt.webAppUrl || 'https://script.google.com/macros/s/TESTDEPLOY/exec'
  };

  const cacheService = makeCacheService();
  const ui = makeUi(state);

  const SpreadsheetApp = {
    getActiveSpreadsheet: () => ss,
    openById: (id) => {
      if (id !== ss.getId()) throw new Error('そのIDのスプレッドシートは開けません: ' + id);
      return ss;
    },
    create: (name) => new MockSpreadsheet(name),
    getActiveSheet: () => ss.getActiveSheet(),
    getActiveRange: () => state.activeRange,
    getUi: () => ui,
    flush: () => { },
    BorderStyle: { SOLID: 'SOLID', SOLID_MEDIUM: 'SOLID_MEDIUM', SOLID_THICK: 'SOLID_THICK' },
    ProtectionType: { RANGE: 'RANGE', SHEET: 'SHEET' },
    newTextStyle: () => {
      const b = {
        setBold: () => b, setFontSize: () => b, setFontFamily: () => b,
        setForegroundColor: () => b, build: () => ({})
      };
      return b;
    },
    newRichTextValue: () => {
      const b = {
        setText: () => b, setTextStyle: () => b, setLinkUrl: () => b, build: () => ({})
      };
      return b;
    }
  };

  const sandbox = {
    SpreadsheetApp,
    PropertiesService: makePropertiesService(),
    CacheService: cacheService,
    LockService: {
      getScriptLock: () => ({
        tryLock: (waitMs) => {
          state.lastLockWaitMs = waitMs;
          return state.lockAvailable && (opt.tryLock ? opt.tryLock(waitMs) : true);
        },
        releaseLock: () => { },
        waitLock: () => { }
      })
    },
    Utilities: makeUtilities(),
    ScriptApp: makeScriptApp(state),
    MailApp: { sendEmail: (...a) => { state.mails.push(a); } },
    DriveApp: {
      getFileById: () => ({
        getName: () => '三者面談',
        getUrl: () => 'https://drive.google.com/file/test',
        getParents: () => ({ hasNext: () => false, next: () => null }),
        makeCopy: (name) => ({ getName: () => name, getUrl: () => 'https://drive/test' }),
        setTrashed: () => { }
      }),
      getRootFolder: () => ({
        getFoldersByName: () => ({ hasNext: () => false, next: () => null }),
        createFolder: (n) => ({ getName: () => n, getUrl: () => 'https://drive/folder', getFiles: () => ({ hasNext: () => false }) })
      })
    },
    UrlFetchApp: {
      fetch: opt.fetch || (() => ({
        getResponseCode: () => 200,
        getContentText: () => '{"ok":true,"version":"test","title":"三者面談 予約"}',
        getAllHeaders: () => ({}),
        getBlob: () => ({ setName: (n) => ({ getName: () => n }) })
      }))
    },
    HtmlService: {
      createTemplateFromFile: () => ({ evaluate: () => htmlOutput() }),
      createHtmlOutput: () => htmlOutput(),
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' }
    },
    ContentService: {
      createTextOutput: (t) => ({ setMimeType: () => ({ getContent: () => t }) }),
      MimeType: { JSON: 'JSON' }
    },
    console: {
      log: () => { },
      info: () => { },
      warn: (...a) => { state.warnings.push(a.map(String).join(' ')); },
      error: (...a) => { state.warnings.push(a.map(String).join(' ')); }
    },
    Date, Math, JSON, String, Number, Boolean, Array, Object, RegExp, Error, isNaN, parseInt, parseFloat
  };

  function htmlOutput() {
    const o = {
      addMetaTag: () => o, setXFrameOptionsMode: () => o, setTitle: () => o,
      setWidth: () => o, setHeight: () => o, getContent: () => ''
    };
    return o;
  }

  const context = vm.createContext(sandbox);

  fs.readdirSync(PROJECT_DIR)
    .filter((f) => f.endsWith('.js'))
    .sort()
    .forEach((f) => {
      const code = fs.readFileSync(path.join(PROJECT_DIR, f), 'utf8');
      vm.runInContext(code, context, { filename: f });
    });

  sandbox.__ss = ss;
  sandbox.__state = state;
  sandbox.__cache = cacheService._cache;
  return sandbox;
}

/* ================================================================
   テスト用の小さな道具
   ================================================================ */

let passed = 0;
let failed = 0;
const failures = [];

function ok(cond, label) {
  if (cond) { passed++; return true; }
  failed++;
  failures.push(label);
  return false;
}

function eq(actual, expected, label) {
  const a = JSON.stringify(actual);
  const b = JSON.stringify(expected);
  if (a === b) { passed++; return true; }
  failed++;
  failures.push(label + '\n      期待: ' + b + '\n      実際: ' + a);
  return false;
}

/** 例外が投げられ、その文言に needle を含むこと */
function throwsWith(fn, needle, label) {
  try {
    fn();
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    if (msg.indexOf(needle) >= 0) { passed++; return true; }
    failed++;
    failures.push(label + '\n      「' + needle + '」を含む例外を期待、実際: ' + msg);
    return false;
  }
  failed++;
  failures.push(label + '\n      例外が投げられませんでした');
  return false;
}

/** safe_() でくるまれた API の戻り。失敗なら error を投げ直して読みやすくする */
function unwrap(res, label) {
  if (!res || !res.ok) {
    throw new Error((label ? label + ': ' : '') + (res && res.error ? res.error : '結果がありません'));
  }
  return res;
}

/** API が {ok:false} を返し、error に needle を含むこと */
function failsWith(res, needle, label) {
  if (res && res.ok) {
    failed++;
    failures.push(label + '\n      失敗するはずが成功しました');
    return false;
  }
  const msg = String(res && res.error ? res.error : '');
  if (msg.indexOf(needle) >= 0) { passed++; return true; }
  failed++;
  failures.push(label + '\n      「' + needle + '」を含むエラーを期待、実際: ' + msg);
  return false;
}

let reported = false;

function report(title) {
  if (reported) return failed === 0;
  reported = true;

  const name = title || path.basename(process.argv[1] || 'test', '.js');
  const line = failed === 0
    ? '✅ ' + name + '  ' + passed + '件すべて通過'
    : '❌ ' + name + '  ' + passed + '件通過 / ' + failed + '件失敗';
  if (failed) {
    failures.forEach((f, i) => console.log('  ' + (i + 1) + '. ' + f));
  }
  console.log(line);
  if (failed) process.exitCode = 1;
  return failed === 0;
}

/**
 * 途中で例外が出ても、そこまでの結果と止まった場所を必ず出す。
 *
 * 予約が消えるような回帰では、以降の下ごしらえが軒並み失敗して例外になる。
 * スタックだけが流れて何も報告されないと、どこから壊れたのかが読めない。
 * load() から自動で仕掛けるので、スイート側で書くことは無い。
 */
function installCrashReport() {
  process.on('uncaughtException', (e) => {
    failed++;
    failures.push('途中で止まりました: ' + String(e && e.message ? e.message : e));
    report(null);
    process.exit(1);
  });
  process.on('exit', () => { report(null); });
}

/* ================================================================
   よく使う下ごしらえ
   ================================================================ */

/**
 * 学校を1つ組み立てる。プロジェクト自身の setupSystem() を通すので、
 * 本番と同じ経路でシートができる。
 *
 * @param {Object} g load() の戻り
 * @param {Object} spec
 * @param {Array} spec.classes [{name, teacher, email, grade}]
 * @param {Array<string>} spec.days 'YYYY-MM-DD'
 * @param {Object} spec.config 設定シートへ上書きする値
 * @param {number} spec.perClass 1クラスあたりの生徒数
 */
function seedSchool(g, spec) {
  spec = spec || {};
  const classes = spec.classes || [
    { name: '1年1組', teacher: '山田', email: 'y@example.jp', grade: '1年' },
    { name: '1年2組', teacher: '鈴木', email: 's@example.jp', grade: '1年' }
  ];
  const days = spec.days || ['2026-10-29', '2026-10-30'];
  const perClass = spec.perClass === undefined ? 3 : spec.perClass;

  g.setupSystem();
  const ss = g.__ss;

  // 既定の 1組〜4組 の予約表は使わないので、先に片付ける
  ss.getSheets()
    .filter((s) => s.getName().indexOf(g.CLASS_SHEET_PREFIX) === 0)
    .forEach((s) => ss.deleteSheet(s));

  // クラス
  const clsSh = ss.getSheetByName(g.SH.CLASSES);
  const lastCls = clsSh.getLastRow();
  if (lastCls >= 2) clsSh.getRange(2, 1, lastCls - 1, 4).clearContent();
  clsSh.getRange(2, 1, classes.length, 4).setValues(
    classes.map((c) => [c.name, c.teacher || '', c.email || '', c.grade || '']));

  // 面談日
  const daySh = ss.getSheetByName(g.SH.DAYS);
  const lastDay = daySh.getLastRow();
  if (lastDay >= 2) daySh.getRange(2, 1, lastDay - 1, 3).clearContent();
  daySh.getRange(2, 1, days.length, 3).setValues(days.map((d) => [d, '', true]));

  // 設定
  const cfg = Object.assign({
    '公開': true,
    '1日の枠数': 3,
    '面談開始時刻': '13:40',
    '面談枠の長さ(分)': 15,
    '枠間の休憩(分)': 10,
    '予約受付開始': '',
    '予約受付締切': '',
    '担任メール通知': false
  }, spec.config || {});
  Object.keys(cfg).forEach((k) => {
    if (!g.setConfigValue_(k, cfg[k])) {
      g.ensureConfigKey_(k, '', cfg[k]);
    }
  });

  g.dropRefCaches_();

  // 「予約表_〇組」を作り、枠も揃える（本番と同じ経路を通す）
  g.syncClasses();

  // 名簿（予約表_〇組 の A・B列）
  if (perClass > 0) {
    classes.forEach((c, ci) => {
      const sh = ss.getSheetByName(g.CLASS_SHEET_PREFIX + c.name);
      const rows = [];
      for (let i = 1; i <= perClass; i++) rows.push([i, '生徒' + (ci + 1) + '_' + i]);
      sh.getRange(2, 1, rows.length, 2).setValues(rows);
    });
    g.clearRosterCache_();
  }

  g.dropRefCaches_();
  return { classes, days, perClass };
}

/** クラス・コマ番号から枠IDを引く */
function slotIdOf(g, ymd, cls, index) {
  return g.makeSlotId_(new Date(ymd), cls, index);
}

/** 枠マスタの1行を読む */
function slotRow(g, slotId) {
  const slots = g.readSlots_();
  return slots.find((s) => String(s.v[g.COL.SLOT_ID - 1]) === slotId) || null;
}

function statusOf(g, slotId) {
  const s = slotRow(g, slotId);
  return s ? String(s.v[g.COL.STATUS - 1]) : null;
}

/**
 * 枠マスタの1セルを読む。枠そのものが消えていれば '(枠が無い)' を返す。
 *
 * 回帰したときに「枠が消えた」ことがそのまま失敗の文言に出るようにするため。
 * 素直に `slotRow(...).v[...]` と書くと、消えた瞬間に例外で止まり、
 * どのテストが何を期待していたのか分からなくなる。
 */
function slotValue(g, slotId, col) {
  const s = slotRow(g, slotId);
  return s ? String(s.v[col - 1]) : '(枠が無い)';
}

module.exports = {
  load,
  ok, eq, throwsWith, unwrap, failsWith, report,
  seedSchool, slotIdOf, slotRow, statusOf, slotValue,
  MockSpreadsheet, MockSheet, MockRange
};
