/**
 * だめなコマ（担任が面談を入れられない枠）の指定と、枠への反映。
 *
 * 担任は「だめなコマ」シートで自分のクラスの列にチェックを入れるだけでよい。
 * チェックした枠は「ブロック」になり、保護者の画面から選べなくなる。
 *
 * ここが効かないと、出張や会議で不在の時間に保護者の予約が入る。
 * 逆に、外したのに戻らないと、使えるはずの枠が埋まらない。
 */

'use strict';

const m = require('./gasmock');

const g = m.load();
m.seedSchool(g, { perClass: 3 });

const C1 = '1年1組';
const C2 = '1年2組';
const DAY1 = '2026-10-29';
const DAY2 = '2026-10-30';

g.setConfigValue_(g.RESERVE_COUNT_KEY, 1);
g.dropRefCaches_();
g.generateSlots();

const ngSh = () => g.__ss.getSheetByName(g.SH.NG);

/** だめなコマシートで (クラス, 日, コマ) のセル位置を探す */
function ngPos(cls, ymd, idx) {
  const sh = ngSh();
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(g.NG_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  let col = 0;
  for (let c = g.NG_COL.FIRST_CLASS - 1; c < lastCol; c++) {
    if (g.ngClassNameOf_(header[c]) === cls) { col = c + 1; break; }
  }
  const last = sh.getLastRow();
  const rows = sh.getRange(g.NG_FIRST_ROW, 1, last - g.NG_FIRST_ROW + 1, 4).getValues();
  let row = 0;
  for (let r = 0; r < rows.length; r++) {
    if (!rows[r][0]) continue;
    if (g.ymd_(rows[r][0]) === ymd && Number(rows[r][2]) === idx) {
      row = r + g.NG_FIRST_ROW;
      break;
    }
  }
  return { row: row, col: col };
}

function checkNg(cls, ymd, idx, value) {
  const p = ngPos(cls, ymd, idx);
  ngSh().getRange(p.row, p.col).setValue(value);
  return p;
}

function readNg(cls, ymd, idx) {
  const p = ngPos(cls, ymd, idx);
  return ngSh().getRange(p.row, p.col).getValue();
}

/* ---------------- シートの作られ方 ---------------- */

const rowCount = ngSh().getLastRow() - g.NG_FIRST_ROW + 1;
m.eq(rowCount, 6, '行は 2日 × 3コマ = 6（予備のコマは含めない）');

const header = ngSh().getRange(g.NG_HEADER_ROW, 1, 1, ngSh().getLastColumn()).getValues()[0];
m.eq(header.slice(0, 4).map(String), ['日付', '曜日', 'コマ', '時間'], '左4列は日付・曜日・コマ・時間');
m.eq(g.ngClassNameOf_(header[4]), C1, '5列目から先はクラス（見出しの1行目がクラス名）');
m.ok(String(header[4]).indexOf('山田') >= 0, '見出しに担任名も出す（自分の列を探しやすいように）');

m.ok(String(ngSh().getRange(g.NG_BANNER_ROW, 1).getValue()).indexOf('チェックを入れて') >= 0,
  '1行目に使い方の案内を出す');

const firstRow = ngSh().getRange(g.NG_FIRST_ROW, 1, 1, 4).getValues()[0];
m.eq(g.ymd_(firstRow[0]), DAY1, '1行目のデータは1日目');
m.eq(String(firstRow[1]), '木', '曜日を出す（日付だけでは分かりにくい）');
m.eq(Number(firstRow[2]), 1, 'コマ番号');
m.eq(String(firstRow[3]), '13:40–13:55', '時間帯');

// 予備のコマは、担任が指定する対象ではない
m.eq(ngPos(C1, DAY1, g.RESERVE_INDEX_BASE + 1).row, 0,
  '予備のコマは、だめなコマシートに行を作らない');

/* ---------------- チェックの読み取り ---------------- */

const slot2 = m.slotIdOf(g, DAY1, C1, 2);

m.eq(g.readNgSet_()[slot2], undefined, '最初は何も指定されていない');

checkNg(C1, DAY1, 2, true);
m.eq(g.readNgSet_()[slot2], true, 'チェックを入れると読み取れる');
m.eq(g.readNgSet_()[m.slotIdOf(g, DAY1, C2, 2)], undefined,
  '同じ行でも、他のクラスの列は影響しない');

/* ---------------- 手書きの記号も受け付ける ---------------- */
/* 配られたシートに「×」を書き込む担任がいる。チェックボックスに限らない */

const marks = ['×', '✕', '✖', 'x', 'X', 'NG', 'ng', '●', '■', '✓', '☑', '1', 'TRUE'];
marks.forEach((mark) => {
  checkNg(C1, DAY2, 1, mark);
  m.ok(g.readNgSet_()[m.slotIdOf(g, DAY2, C1, 1)] === true,
    '「' + mark + '」も「面談を入れない」として読む');
});

['', false, '○', '〇', 'OK', '0'].forEach((mark) => {
  checkNg(C1, DAY2, 1, mark);
  m.eq(g.readNgSet_()[m.slotIdOf(g, DAY2, C1, 1)], undefined,
    '「' + String(mark) + '」は指定とみなさない');
});

checkNg(C1, DAY2, 1, false);

/* ---------------- 枠への反映 ---------------- */

let res = g.applyNgSlots();
m.eq(res.blocked, 1, 'チェックした1枠をブロックにした、と報告する');
m.eq(res.unblocked, 0, '戻した枠は無い');
m.eq(m.statusOf(g, slot2), 'ブロック', '枠マスタがブロックになる');
m.eq(m.statusOf(g, m.slotIdOf(g, DAY1, C1, 1)), '空き', '指定していない枠は空きのまま');

// 保護者の画面から選べなくなる
const view = m.unwrap(g.apiSlots({ cls: C1, no: 1, name: '生徒1_1' }), '保護者の一覧');
const cell = view.days[0].slots.find((s) => s.slotId === slot2);
m.ok(cell, '一覧には残る（時間割として見えている必要がある）');
m.ok(!cell.available, 'ただし選べない');
m.ok(cell.blocked, '「面談なし」として表示するための印が立つ');

m.failsWith(g.apiBook({ cls: C1, no: 1, name: '生徒1_1', slotId: slot2 }),
  'ちょうど埋まりました', 'ブロックの枠は予約できない');

/* ---------------- もう一度実行しても変わらない ---------------- */

res = g.applyNgSlots();
m.eq(res.blocked, 0, '同じ状態で実行しても、二重にブロックしない');
m.eq(res.unblocked, 0, '余計に戻しもしない');

/* ---------------- チェックを外すと空きに戻る ---------------- */

checkNg(C1, DAY1, 2, false);
res = g.applyNgSlots();
m.eq(res.unblocked, 1, 'チェックを外した1枠を空きに戻した、と報告する');
m.eq(m.statusOf(g, slot2), '空き', '空きに戻る');

/* ---------------- 予約が入っている枠は変えない ---------------- */

const booked = m.unwrap(g.apiBook({ cls: C1, no: 1, name: '生徒1_1', slotId: slot2 }), '予約を入れる');
checkNg(C1, DAY1, 2, true);
res = g.applyNgSlots();

m.eq(res.blocked, 0, '予約済の枠はブロックにしない');
m.eq(m.statusOf(g, slot2), '予約済', '予約はそのまま残る');
m.eq(res.conflicts.length, 1, 'ぶつかっている枠として報告する');
m.eq(res.conflicts[0].name, '生徒1_1', 'どの生徒かを返す');
m.eq(res.conflicts[0].cls, C1, 'どのクラスかを返す');

// 予約を取り消すと、指定どおりブロックになる
m.unwrap(g.apiCancel({ cls: C1, no: 1, name: '生徒1_1', code: booked.code }), '予約を取消');
m.eq(m.statusOf(g, slot2), 'ブロック',
  'だめなコマに指定された枠は、取り消した時点でブロックへ戻る');
res = g.applyNgSlots();
m.eq(res.conflicts.length, 0, '取り消せば、ぶつかりは解消する');

checkNg(C1, DAY1, 2, false);
g.applyNgSlots();

/* ---------------- 予備の枠には触らない ---------------- */

const reserveId = m.slotIdOf(g, DAY1, C1, g.RESERVE_INDEX_BASE + 1);
m.eq(m.statusOf(g, reserveId), '予備', '反映の前は予備');
g.applyNgSlots();
m.eq(m.statusOf(g, reserveId), '予備',
  'だめなコマの反映は、予備の枠を空きにしてしまわない');

/* ---------------- 管理画面・メニューからの指定と同期する ---------------- */

const slot3 = m.slotIdOf(g, DAY2, C2, 3);
g.setNgFlag_(slot3, true);
m.eq(readNg(C2, DAY2, 3), true, '管理画面で×にすると、シートのチェックも入る');
m.eq(g.readNgSet_()[slot3], true, '読み取りでも一致する');

g.setNgFlag_(slot3, false);
m.eq(readNg(C2, DAY2, 3), false, '管理画面で戻すと、シートのチェックも外れる');

// まとめて指定
g.setNgFlags_([
  { slotId: m.slotIdOf(g, DAY2, C1, 1), flag: true },
  { slotId: m.slotIdOf(g, DAY2, C1, 2), flag: true },
  { slotId: m.slotIdOf(g, DAY2, C1, 3), flag: true }
]);
m.eq([1, 2, 3].map((i) => readNg(C1, DAY2, i)), [true, true, true],
  'まとめて指定すると、その列がまとめてチェックされる');

res = g.applyNgSlots();
m.eq(res.blocked, 3, '半日まるごとブロックできる（出張などの想定）');

/* ---------------- 枠を再生成してもチェックは残る ---------------- */

g.generateSlots();
m.eq([1, 2, 3].map((i) => readNg(C1, DAY2, i)), [true, true, true],
  'シートを作り直しても、入っていたチェックは引き継がれる');
m.eq(m.statusOf(g, m.slotIdOf(g, DAY2, C1, 1)), 'ブロック',
  '再生成のあとも、指定どおりブロックのまま');

/* ---------------- 枠数を増やすと、行も増える ---------------- */

g.setConfigValue_('1日の枠数', 4);
g.dropRefCaches_();
g.generateSlots();
m.eq(ngSh().getLastRow() - g.NG_FIRST_ROW + 1, 8, '枠数を増やすと 2日 × 4コマ = 8 行になる');
m.eq([1, 2, 3].map((i) => readNg(C1, DAY2, i)), [true, true, true],
  '増やしても、もとのチェックは残る');
m.eq(readNg(C1, DAY2, 4), false, '増えたコマは未チェック');

g.setConfigValue_('1日の枠数', 3);
g.dropRefCaches_();
g.generateSlots();

/* ---------------- すべて外す（リセット） ---------------- */

const cleared = g.clearAllNg_();
m.eq(cleared, 3, '外した数を返す');
m.eq([1, 2, 3].map((i) => readNg(C1, DAY2, i)), [false, false, false], 'チェックが全部外れる');

res = g.applyNgSlots();
m.eq(res.unblocked, 3, 'ブロックしていた枠が空きに戻る');
m.eq(m.statusOf(g, m.slotIdOf(g, DAY2, C1, 1)), '空き', '空きに戻っている');

/* ---------------- 枠IDの読み解き ---------------- */

m.eq(g.parseSlotId_('20261029_1年1組_2'), { ymd: '20261029', cls: '1年1組', idx: 2 },
  '枠IDを 日付・クラス・コマ番号 に分解できる');
m.eq(g.parseSlotId_('20261029_ひまわり学級_901').idx, 901, '予備の900番台も読める');
m.eq(g.parseSlotId_('20261029_1年_1組_2').cls, '1年_1組',
  'クラス名にアンダースコアが入っていても、コマ番号を取り違えない');
m.eq(g.parseSlotId_('こわれた枠ID'), null, '読めない枠IDは null');

/* ---------------- ログ ---------------- */

const log = g.__ss.getSheetByName(g.SH.LOG);
const rows = log.getRange(2, 1, log.getLastRow() - 1, 7).getValues();
const applied = rows.filter((r) => String(r[1]) === 'だめなコマ反映');
m.ok(applied.length > 0, 'だめなコマの反映がログに残る');
m.ok(String(applied[applied.length - 1][6]).indexOf('ブロック') >= 0,
  'ログに件数が残る');

m.report('test_ng');
