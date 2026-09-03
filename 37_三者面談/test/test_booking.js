/**
 * 保護者の予約そのもの — 確定・変更・取消と、名簿照合。
 *
 * ここが壊れると、予約が消えるか二重に入る。どちらも当日まで気づけないので、
 * いちばん先に守る。
 */

'use strict';

const m = require('./gasmock');

const g = m.load();
const S = m.seedSchool(g, { perClass: 3 });

const CLS = '1年1組';
const OTHER = '1年2組';
const DAY1 = '2026-10-29';
const DAY2 = '2026-10-30';

const s1 = m.slotIdOf(g, DAY1, CLS, 1);
const s2 = m.slotIdOf(g, DAY1, CLS, 2);
const s3 = m.slotIdOf(g, DAY2, CLS, 1);
const other1 = m.slotIdOf(g, DAY1, OTHER, 1);

const A = { cls: CLS, no: 1, name: '生徒1_1' };
const B = { cls: CLS, no: 2, name: '生徒1_2' };

/* ---------------- 下ごしらえの確認 ---------------- */

m.eq(g.readSlots_().length, 12, '枠が 2クラス×2日×3コマ = 12 できている');
m.eq(m.statusOf(g, s1), '空き', '作った直後はすべて空き');

/* ---------------- 名簿との照合 ---------------- */

m.failsWith(g.apiSlots({ cls: CLS, no: 1, name: 'ちがう名前' }),
  '名簿と一致しませんでした', '氏名が名簿と違うと弾く');

m.failsWith(g.apiSlots({ cls: CLS, no: 99, name: '生徒1_1' }),
  '名簿と一致しませんでした', '名簿に無い出席番号は弾く');

m.failsWith(g.apiSlots({ cls: '', no: 1, name: '生徒1_1' }),
  'クラスを選んでください', 'クラス未選択を弾く');

// 空白の有無は無視する（案内文でそう約束している）
m.unwrap(g.apiSlots({ cls: CLS, no: 1, name: '生徒1_1 ' }), '前後の空白は無視される');

/* ---------------- 空き枠の一覧 ---------------- */

const slotsView = m.unwrap(g.apiSlots(A), 'apiSlots');
m.eq(slotsView.days.length, 2, '2日ぶんの枠が返る');
m.eq(slotsView.days[0].slots.length, 3, '1日3コマ');
m.eq(slotsView.existing, null, 'まだ予約は無い');
m.ok(slotsView.days[0].slots.every((s) => s.available), '最初はすべて予約できる');

// 他クラスの枠が混ざっていないこと（混ざると他人の枠を取れてしまう）
const allIds = [];
slotsView.days.forEach((d) => d.slots.forEach((s) => allIds.push(s.slotId)));
m.ok(allIds.indexOf(other1) < 0, '他クラスの枠は一覧に出さない');

/* ---------------- 予約の確定 ---------------- */

const booked = m.unwrap(g.apiBook(Object.assign({}, A, {
  slotId: s1, guardian: '保護者1', note: '進路のこと'
})), 'apiBook');

m.ok(/^\d{4}$/.test(booked.code), '予約コードは4桁');
m.eq(m.statusOf(g, s1), '予約済', '枠マスタが予約済になる');
m.eq(booked.booking.ymd, DAY1, 'カレンダー登録用に ymd を返す（年を取り違えないため）');
m.eq(booked.booking.cls, CLS, 'クラスを返す');

const row = m.slotRow(g, s1);
m.eq(String(row.v[g.COL.STUDENT - 1]), A.name, '生徒氏名が入る');
m.eq(String(row.v[g.COL.GUARDIAN - 1]), '保護者1', '保護者氏名が入る');
m.eq(String(row.v[g.COL.NOTE - 1]), '進路のこと', '連絡事項が入る');

/* ---------------- 二重予約を防ぐ ---------------- */

m.failsWith(g.apiBook(Object.assign({}, A, { slotId: s2 })),
  'すでに予約が入っています', '同じ生徒は2つ目を取れない');

m.failsWith(g.apiBook(Object.assign({}, B, { slotId: s1 })),
  'ちょうど埋まりました', '埋まった枠は別の生徒も取れない');

m.failsWith(g.apiBook(Object.assign({}, A, { slotId: other1 })),
  '他のクラスの時間は予約できません', '他クラスの枠は取れない');

m.failsWith(g.apiBook(Object.assign({}, A, { slotId: 'ありえない枠ID' })),
  '見つかりませんでした', '存在しない枠IDを弾く');

/* ---------------- 予約済みの生徒に一覧を出したとき ---------------- */

const after = m.unwrap(g.apiSlots(A), '予約後の apiSlots');
m.ok(after.existing !== null, '自分の予約が existing に入る');
m.eq(after.existing.slotId, s1, 'existing は予約した枠');
m.ok(after.days[0].slots.find((s) => s.slotId === s1).mine, '自分の枠に印が付く');
m.ok(!after.days[0].slots.find((s) => s.slotId === s1).available, '自分の枠は選べない');

/* ---------------- 確認（予約コード） ---------------- */

m.failsWith(g.apiLookup(Object.assign({}, A, { code: '0000' })),
  '予約が見つかりませんでした', '違う予約コードでは引けない');

m.failsWith(g.apiLookup(Object.assign({}, A, { code: '' })),
  '予約コードを入力してください', '空の予約コードを弾く');

const found = m.unwrap(g.apiLookup(Object.assign({}, A, { code: booked.code })), 'apiLookup');
m.eq(found.booking.slotId, s1, '正しいコードなら引ける');
m.eq(found.booking.guardian, '保護者1', '保護者氏名が返る');
m.eq(found.booking.note, '進路のこと', '連絡事項が返る');

/* ---------------- 変更 ---------------- */

const changed = m.unwrap(g.apiChange(Object.assign({}, A, {
  code: booked.code, slotId: s3
})), 'apiChange');

m.eq(m.statusOf(g, s1), '空き', '変更前の枠は空きに戻る');
m.eq(m.statusOf(g, s3), '予約済', '変更先の枠が予約済になる');
m.eq(changed.booking.slotId, s3, '変更後の枠を返す');
m.eq(changed.booking.guardian, '保護者1', '保護者氏名を引き継ぐ');
m.eq(changed.booking.note, '進路のこと', '連絡事項を引き継ぐ');
m.eq(String(m.slotRow(g, s3).v[g.COL.CODE - 1]), booked.code, '予約コードは変わらない');

// 変更前の枠が空いたので、別の生徒が取れるようになる
const bBooked = m.unwrap(g.apiBook(Object.assign({}, B, { slotId: s1 })), '空いた枠を別の生徒が予約');
m.eq(m.statusOf(g, s1), '予約済', '空いた枠は再び使える');

m.failsWith(g.apiChange(Object.assign({}, A, { code: '9999', slotId: s2 })),
  '予約が見つかりませんでした', '違うコードでは変更できない');

m.failsWith(g.apiChange(Object.assign({}, A, { code: booked.code, slotId: s1 })),
  'ちょうど埋まりました', '埋まっている枠へは変更できない');

/* ---------------- 取消 ---------------- */

m.unwrap(g.apiCancel(Object.assign({}, A, { code: booked.code })), 'apiCancel');
m.eq(m.statusOf(g, s3), '空き', '取り消すと空きに戻る');
m.eq(String(m.slotRow(g, s3).v[g.COL.STUDENT - 1]), '', '生徒氏名が消える');
m.eq(String(m.slotRow(g, s3).v[g.COL.CODE - 1]), '', '予約コードが消える');

m.failsWith(g.apiCancel(Object.assign({}, A, { code: booked.code })),
  '予約が見つかりませんでした', '取り消したあとは同じコードで引けない');

/* ---------------- だめなコマに指定された枠の取消 ---------------- */
/* 空きに戻すと、担任が面談を入れないと決めた枠が再び予約されてしまう */

const ngSlot = m.slotIdOf(g, DAY2, CLS, 2);
const C = { cls: CLS, no: 3, name: '生徒1_3' };
const cBooked = m.unwrap(g.apiBook(Object.assign({}, C, { slotId: ngSlot })), 'だめなコマ用に予約');

// 「だめなコマ」シートで、その枠にチェックを入れる
g.setNgFlag_(ngSlot, true);
m.ok(g.readNgSet_()[ngSlot] === true, 'だめなコマの指定が読み取れる');

m.unwrap(g.apiCancel(Object.assign({}, C, { code: cBooked.code })), 'だめなコマの枠を取消');
m.eq(m.statusOf(g, ngSlot), 'ブロック',
  'だめなコマに指定された枠は、取り消しても空きではなくブロックに戻る');

/* ---------------- 受付期間 ---------------- */

g.setConfigValue_('公開', false);
g.dropRefCaches_();
m.failsWith(g.apiSlots(A), '受付を停止しています', '非公開なら一覧を出さない');
m.failsWith(g.apiBook(Object.assign({}, A, { slotId: s2 })), '受付を停止しています', '非公開なら予約できない');

g.setConfigValue_('公開', true);
g.setConfigValue_('予約受付締切', new Date(2020, 0, 1, 12, 0));
g.dropRefCaches_();
m.failsWith(g.apiBook(Object.assign({}, A, { slotId: s2 })), '締め切りました', '締切を過ぎたら予約できない');

g.setConfigValue_('予約受付締切', '');
g.setConfigValue_('予約受付開始', new Date(2099, 0, 1, 9, 0));
g.dropRefCaches_();
m.failsWith(g.apiBook(Object.assign({}, A, { slotId: s2 })), 'から開始します', '受付開始前は予約できない');

g.setConfigValue_('予約受付開始', '');
g.dropRefCaches_();

/* ---------------- 予約ログ ---------------- */

const log = g.__ss.getSheetByName(g.SH.LOG);
const logRows = log.getRange(2, 1, log.getLastRow() - 1, 7).getValues();
const actions = logRows.map((r) => String(r[1]));
m.ok(actions.indexOf('予約') >= 0, '予約がログに残る');
m.ok(actions.indexOf('変更') >= 0, '変更がログに残る');
m.ok(actions.indexOf('取消') >= 0, '取消がログに残る');

// 取り消すと枠マスタの行は空になるので、あとから追えるようコードを残している
const cancelRow = logRows.find((r) => String(r[1]) === '取消');
m.ok(String(cancelRow[6]).indexOf('コード') >= 0, '取消のログに予約コードが残る');

m.report('test_booking');
