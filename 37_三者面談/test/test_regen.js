/**
 * 枠の再生成 — 予約の引き継ぎと、失う枠があるときの中止。
 *
 * 面談日や枠数を直したあとに走る処理で、ここが壊れると
 * **すでに入っている予約が黙って消える。** しかも消えたことに誰も気づけない。
 *
 * 枠IDは「日付＋クラス名＋コマ番号」なので、そのどれを触っても影響が出る。
 */

'use strict';

const m = require('./gasmock');

const g = m.load();
m.seedSchool(g, { perClass: 3 });

const CLS = '1年1組';
const OTHER = '1年2組';
const DAY1 = '2026-10-29';
const DAY2 = '2026-10-30';

const A = { cls: CLS, no: 1, name: '生徒1_1' };
const B = { cls: CLS, no: 2, name: '生徒1_2' };
const C = { cls: OTHER, no: 1, name: '生徒2_1' };

/** 「面談日」シートの「実施する」を切り替える */
function setDayEnabled(ymd, on) {
  const sh = g.__ss.getSheetByName(g.SH.DAYS);
  const last = sh.getLastRow();
  const vals = sh.getRange(2, 1, last - 1, 3).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (g.ymd_(vals[i][0]) === ymd) { sh.getRange(i + 2, 3).setValue(on); return true; }
  }
  return false;
}

/** 「クラス」シートのクラス名を書き換える */
function renameClass(from, to) {
  const sh = g.__ss.getSheetByName(g.SH.CLASSES);
  const last = sh.getLastRow();
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === from) { sh.getRange(i + 2, 1).setValue(to); break; }
  }
  g.dropRefCaches_();
}

function setConfig(key, value) {
  g.setConfigValue_(key, value);
  g.dropRefCaches_();
}

/* ---------------- コマの時刻 ---------------- */

const t = g.daySlotTimes_(g.getConfig());
m.eq(t.map((x) => x.start + '-' + x.end),
  ['13:40-13:55', '14:05-14:20', '14:30-14:45'],
  '13:40開始・15分・休憩10分で3コマの時刻が並ぶ');
m.eq(t.map((x) => x.index), [1, 2, 3], 'コマ番号は1から');

m.eq(m.slotIdOf(g, DAY1, CLS, 1), '20261029_1年1組_1', '枠IDは 日付_クラス_コマ番号');

/* ---------------- 変更が無ければ、そのまま引き継ぐ ---------------- */

const s1 = m.slotIdOf(g, DAY1, CLS, 1);
const s3 = m.slotIdOf(g, DAY1, CLS, 3);
const day2s1 = m.slotIdOf(g, DAY2, CLS, 1);

const bookedA = m.unwrap(g.apiBook(Object.assign({}, A, {
  slotId: s1, guardian: '保護者A', note: '進路のこと'
})), 'Aを予約');

let res = g.generateSlots();
m.eq(res.written, 12, '再生成しても枠数は変わらない');
m.eq(res.kept, 1, '予約1件を引き継いだと報告する');

let row = m.slotRow(g, s1);
m.eq(String(row.v[g.COL.STATUS - 1]), '予約済', '状態を引き継ぐ');
m.eq(String(row.v[g.COL.STUDENT - 1]), A.name, '生徒氏名を引き継ぐ');
m.eq(String(row.v[g.COL.GUARDIAN - 1]), '保護者A', '保護者氏名を引き継ぐ');
m.eq(String(row.v[g.COL.NOTE - 1]), '進路のこと', '連絡事項を引き継ぐ');
m.eq(String(row.v[g.COL.CODE - 1]), bookedA.code, '予約コードを引き継ぐ');
m.ok(row.v[g.COL.BOOKED_AT - 1] !== '', '予約日時を引き継ぐ');

// 引き継いだあとも、保護者が自分の予約を引けること
m.unwrap(g.apiLookup(Object.assign({}, A, { code: bookedA.code })), '再生成後も予約コードで引ける');

/* ---------------- 枠を増やす ---------------- */

setConfig('1日の枠数', 4);
res = g.generateSlots();
m.eq(res.written, 16, '枠数を増やすと 2クラス×2日×4コマ = 16');
m.eq(res.kept, 1, '増やしても予約は残る');
m.eq(m.statusOf(g, s1), '予約済', '増やしても予約済のまま');
m.eq(m.statusOf(g, m.slotIdOf(g, DAY1, CLS, 4)), '空き', '増えたコマは空き');

/* ---------------- 予約のある枠が消える変更は中止する ---------------- */
/* ここが本丸。中止せずに書き込むと、予約が黙って消える */

const bookedB = m.unwrap(g.apiBook(Object.assign({}, B, { slotId: s3 })), 'Bを3コマ目に予約');
m.eq(m.statusOf(g, s3), '予約済', '3コマ目が埋まった');

const before = g.readSlots_().length;
setConfig('1日の枠数', 2);   // 3コマ目・4コマ目が消える設定
m.throwsWith(() => g.generateSlots(),
  '予約が入っているため再生成を中止しました',
  '予約のあるコマが消える設定では、再生成を中止する');

m.eq(g.readSlots_().length, before, '中止したので枠は1つも書き換わっていない');
m.eq(m.statusOf(g, s3), '予約済', '中止したので予約は残っている');
m.eq(m.slotValue(g, s3, g.COL.STUDENT), B.name, '中止したので生徒氏名も残っている');

// 中止のメッセージに、どの予約が邪魔しているかが出ること（出ないと直しようがない）
try {
  g.generateSlots();
} catch (e) {
  const msg = String(e.message);
  m.ok(msg.indexOf(B.name) >= 0, '中止のメッセージに生徒氏名が出る');
  m.ok(msg.indexOf('10月29日') >= 0, '中止のメッセージに日付が出る');
}

// 予約を取り消せば、同じ設定でも通る
m.unwrap(g.apiCancel(Object.assign({}, B, { code: bookedB.code })), 'Bの予約を取消');
res = g.generateSlots();
m.eq(res.written, 8, '取り消したあとは 2クラス×2日×2コマ = 8 に減らせる');
m.eq(res.kept, 1, 'Aの予約は残っている');

setConfig('1日の枠数', 3);
g.generateSlots();

/* ---------------- 面談日を減らす ---------------- */

const bookedDay2 = m.unwrap(g.apiBook(Object.assign({}, B, { slotId: day2s1 })), 'Bを2日目に予約');

setDayEnabled(DAY2, false);
m.throwsWith(() => g.generateSlots(),
  '予約が入っているため再生成を中止しました',
  '予約のある日を実施しないに変えたら、再生成を中止する');
m.eq(m.statusOf(g, day2s1), '予約済', '中止したので2日目の予約は残っている');

m.unwrap(g.apiCancel(Object.assign({}, B, { code: bookedDay2.code })), '2日目の予約を取消');
res = g.generateSlots();
m.eq(res.written, 6, '日を1つ減らすと 2クラス×1日×3コマ = 6');

setDayEnabled(DAY2, true);
res = g.generateSlots();
m.eq(res.written, 12, '日を戻すと12に戻る');
m.eq(m.statusOf(g, s1), '予約済', '日を出し入れしてもAの予約は残る');

/* ---------------- クラス名を変える ---------------- */
/* 枠IDにクラス名が入っているので、改名は「消えて別のができる」に等しい */

const bookedC = m.unwrap(g.apiBook(Object.assign({}, C, {
  slotId: m.slotIdOf(g, DAY1, OTHER, 1)
})), 'Cを1年2組に予約');

renameClass(OTHER, '1年3組');
m.throwsWith(() => g.generateSlots(),
  '予約が入っているため再生成を中止しました',
  '予約のあるクラスを改名したら、再生成を中止する');
m.eq(m.statusOf(g, m.slotIdOf(g, DAY1, OTHER, 1)), '予約済', '中止したので改名前の予約は残る');

renameClass('1年3組', OTHER);
g.generateSlots();
m.eq(m.statusOf(g, m.slotIdOf(g, DAY1, OTHER, 1)), '予約済', '名前を戻せば予約はそのまま');

/* ---------------- 入力の誤りは、書き込む前に止める ---------------- */

const clsSh = g.__ss.getSheetByName(g.SH.CLASSES);
clsSh.getRange(4, 1).setValue(CLS);      // クラス名を重複させる
g.dropRefCaches_();
m.throwsWith(() => g.generateSlots(), '同じクラス名が複数あります',
  'クラス名が重複していたら、枠を作らずに止める');
clsSh.getRange(4, 1).setValue('');
g.dropRefCaches_();

const daySh = g.__ss.getSheetByName(g.SH.DAYS);
daySh.getRange(4, 1, 1, 3).setValues([[DAY1, '', true]]);   // 同じ日を2回
m.throwsWith(() => g.generateSlots(), '同じ日付が複数あります',
  '面談日が重複していたら、枠を作らずに止める');
daySh.getRange(4, 1, 1, 3).setValues([['', '', '']]);

setConfig('1日の枠数', 0);
m.throwsWith(() => g.generateSlots(), '「1日の枠数」に1以上',
  '1日の枠数が0なら止める');
setConfig('1日の枠数', 3);

setConfig('面談枠の長さ(分)', 0);
m.throwsWith(() => g.generateSlots(), '「面談枠の長さ(分)」に1以上',
  '枠の長さが0なら止める');
setConfig('面談枠の長さ(分)', 15);

setConfig('枠間の休憩(分)', -5);
m.throwsWith(() => g.generateSlots(), '「枠間の休憩(分)」に0以上',
  '休憩が負の数なら止める');
setConfig('枠間の休憩(分)', 10);

// 休憩0分は「休憩なし」として有効。|| で既定値に化けないこと
setConfig('枠間の休憩(分)', 0);
g.generateSlots();
m.eq(g.daySlotTimes_(g.getConfig()).map((x) => x.start),
  ['13:40', '13:55', '14:10'], '休憩0分は「休憩なし」として扱う（既定値に化けない）');
setConfig('枠間の休憩(分)', 10);
g.generateSlots();

/* ---------------- だめなコマの引き継ぎ ---------------- */

const ngSlot = m.slotIdOf(g, DAY2, CLS, 3);
g.setNgFlag_(ngSlot, true);
res = g.generateSlots();
m.eq(m.statusOf(g, ngSlot), 'ブロック', '再生成のあと、だめなコマの指定が枠へ反映される');
m.ok(res.ngBlocked >= 1, 'ブロックした枠数を報告する');
m.ok(g.readNgSet_()[ngSlot] === true, '再生成でだめなコマシートを作り直しても、チェックは引き継がれる');

// 予約が入っている枠は、だめなコマに指定してもブロックにしない
g.setNgFlag_(s1, true);
res = g.generateSlots();
m.eq(m.statusOf(g, s1), '予約済', '予約のある枠は、だめなコマ指定でも予約済のまま');
m.eq(res.ngConflicts.length, 1, 'ぶつかっている枠を報告する');
m.eq(res.ngConflicts[0].name, A.name, 'ぶつかっている生徒の氏名を返す');
g.setNgFlag_(s1, false);
g.setNgFlag_(ngSlot, false);
g.generateSlots();

/* ---------------- 予備コマ ---------------- */

setConfig(g.RESERVE_COUNT_KEY, 1);
res = g.generateSlots();
m.eq(res.written, 16, '予備を1つ足すと 2クラス×2日×(3+1) = 16');

const reserveId = m.slotIdOf(g, DAY1, CLS, g.RESERVE_INDEX_BASE + 1);
m.eq(m.statusOf(g, reserveId), '予備', '予備の枠は「予備」で作られる');

const rTimes = g.dayReserveTimes_(g.getConfig());
m.eq(rTimes.map((x) => x.start + '-' + x.end), ['14:55-15:10'],
  '予備は最終コマの後ろに、同じ間隔で作られる');
m.eq(rTimes[0].index, 901, '予備のコマ番号は900番台（通常のコマと衝突させない）');

// 保護者の一覧には出さない
const view = m.unwrap(g.apiSlots(B), '予備があるときの apiSlots');
const shown = [];
view.days.forEach((d) => d.slots.forEach((s) => shown.push(s.slotId)));
m.ok(shown.indexOf(reserveId) < 0, '予備の枠は保護者の一覧に出さない');

// 担任が入れたぶんは、再生成でも引き継ぐ
const keptBefore = g.generateSlots().kept;
g.setReserveStudent_(reserveId, '3', 'テスト');
m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '生徒1_3', '予備に生徒が入る');

res = g.generateSlots();
m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '生徒1_3', '再生成しても予備の記入は残る');
m.eq(m.statusOf(g, reserveId), '予備', '再生成しても状態は「予備」のまま');
m.eq(res.kept, keptBefore + 1, '予備に入れたぶんも「引き継いだ」に数える');

// 予備が減る変更も、入っていれば中止する
setConfig(g.RESERVE_COUNT_KEY, 0);
m.throwsWith(() => g.generateSlots(),
  '予約が入っているため再生成を中止しました',
  '予備に生徒が入っているのに予備を0にしたら、再生成を中止する');
m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '生徒1_3', '中止したので予備の記入は残る');

g.setReserveStudent_(reserveId, '', 'テスト');
res = g.generateSlots();
m.eq(res.written, 12, '予備を空にすれば0に戻せる');

/* ---------------- 開始時刻を変えると、予約の時刻が動く ---------------- */
/* 枠IDはコマ番号で作るため、時刻を変えても予約はコマに付いたまま移動する。
   仕様どおりだが、保護者に伝えた時刻とずれる。受付開始後は触らないこと */

setConfig('面談開始時刻', '14:00');
g.generateSlots();
m.eq(m.slotValue(g, s1, g.COL.START), '14:00',
  '開始時刻を変えると、予約はコマに付いたまま時刻だけ動く（保護者への連絡が要る）');
m.eq(m.statusOf(g, s1), '予約済', '時刻が動いても予約そのものは残る');
setConfig('面談開始時刻', '13:40');
g.generateSlots();

/* ---------------- ログ ---------------- */

const log = g.__ss.getSheetByName(g.SH.LOG);
const logRows = log.getRange(2, 1, log.getLastRow() - 1, 7).getValues();
const regen = logRows.filter((r) => String(r[1]) === '枠再生成');
m.ok(regen.length > 0, '再生成がログに残る');
m.ok(String(regen[regen.length - 1][6]).indexOf('引継ぎ') >= 0, 'ログに引き継ぎ件数が残る');

m.report('test_regen');
