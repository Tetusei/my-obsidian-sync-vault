/**
 * きょうだいまとめて予約。
 *
 * 全クラスが同じ時間割で動くので、「同じ日の連続したコマ」を
 * 別々のクラスから1つずつ確保する、という形で扱える。
 *
 * ここが壊れると、**一部の子だけ予約が入って残りが入らない**という
 * 中途半端な状態になる。保護者からは何が起きたのか分からず、
 * 学校側も気づけない。書き込みは全員ぶんまとめて、1人でも欠けたら誰も入れない。
 */

'use strict';

const m = require('./gasmock');

const g = m.load();
m.seedSchool(g, {
  perClass: 3,
  classes: [
    { name: '1年1組', teacher: '山田', email: 'y@example.jp', grade: '1年' },
    { name: '1年2組', teacher: '鈴木', email: 's@example.jp', grade: '1年' },
    { name: '2年1組', teacher: '田中', email: 't@example.jp', grade: '2年' }
  ]
});

const C1 = '1年1組';
const C2 = '1年2組';
const C3 = '2年1組';
const DAY1 = '2026-10-29';
const DAY2 = '2026-10-30';

// 1年1組の1番 と 2年1組の1番 をきょうだいとする
const KID_A = { cls: C1, no: 1, name: '生徒1_1' };
const KID_B = { cls: C3, no: 1, name: '生徒3_1' };
const PAIR = [KID_A, KID_B];

/** 入力の誤りを数えるカウンタを消す（8回続くと15分止まるため） */
function clearFailures() {
  g.__cache._store.clear();
}

/* ---------------- 最初の画面 ---------------- */

const init = m.unwrap(g.apiSiblingInit(), 'apiSiblingInit');
m.eq(init.maxChildren, 6, '一度に申し込めるのは6人まで');
m.eq(init.grades.map((x) => x.grade), ['1年', '2年'], '学年は数字の順に並ぶ');
m.eq(init.grades[0].classes, [C1, C2], '学年ごとにクラスがまとまる');
m.eq(init.grades[1].classes, [C3], '2年のクラス');
m.ok(init.open, '受付中なら open が立つ');

/* ---------------- 人数と入力の確認 ---------------- */

m.failsWith(g.apiSiblingSlots({ children: [KID_A] }),
  '2人以上でご利用ください', '1人ではまとめて予約を使わせない');

const seven = [];
for (let i = 0; i < 7; i++) seven.push({ cls: C1, no: i + 1, name: 'x' });
m.failsWith(g.apiSiblingSlots({ children: seven }),
  '6 人までです', '7人は受け付けない');

clearFailures();
m.failsWith(g.apiSiblingSlots({ children: [KID_A, KID_A] }),
  '同じお子さんが2回入力されています', '同じ子を2回入力したら弾く');

clearFailures();
const badRes = g.apiSiblingSlots({
  children: [
    { cls: C1, no: 1, name: 'ちがう名前' },
    { cls: C3, no: 99, name: 'いない子' }
  ]
});
m.ok(!badRes.ok, '入力が合わなければ失敗する');
m.ok(String(badRes.error).indexOf('1人目') >= 0, '何人目が合わなかったかを示す（1人目）');
m.ok(String(badRes.error).indexOf('2人目') >= 0, '何人目が合わなかったかを示す（2人目）');
m.ok(String(badRes.error).indexOf('氏名が名簿と一致しません') >= 0, '氏名違いの理由が出る');
m.ok(String(badRes.error).indexOf('名簿と一致しません') >= 0, '名簿に無い子の理由が出る');

clearFailures();
m.failsWith(g.apiSiblingSlots({ children: [KID_A, { cls: '', no: '', name: '' }] }),
  'クラスと出席番号を入力してください', '未入力の欄を指摘する');

clearFailures();

/* ---------------- 組み合わせの探索 ---------------- */

const view = m.unwrap(g.apiSiblingSlots({ children: PAIR }), 'apiSiblingSlots');

m.eq(view.children.map((c) => c.cls), [C1, C3], 'お子さんは学年の低い順に並ぶ');
m.eq(view.children[0].teacher, '山田', '担任名を返す');
m.eq(view.already.length, 0, 'まだ誰も予約していない');
m.eq(view.days.length, 2, '2日ぶんの組み合わせが返る');

const d1 = view.days[0];
m.eq(d1.key, DAY1, '1日目のキーは日付');
m.eq(d1.options.length, 2, '3コマに2人なら、続いた組み合わせは2通り（1-2コマ目 と 2-3コマ目）');
m.eq(d1.maxFit, 2, '2人とも入るなら maxFit は2');

const opt = d1.options[0];
m.eq(opt.startTime, '13:40', '1つ目の組み合わせは1コマ目から');
m.eq(opt.endTime, '14:20', '2人目が終わる時刻を返す');
m.eq(opt.picks.length, 2, '人数ぶんの枠が入る');
m.eq(opt.picks.map((p) => p.start), ['13:40', '14:05'], '時間順に並ぶ（続いたコマ）');
m.eq(opt.picks.map((p) => p.cls), [C1, C3], '在籍クラスの枠が割り当てられる');
m.ok(opt.picks.every((p) => p.slotId && p.name && p.teacher !== undefined),
  '画面に出すのに必要な情報がそろっている');

// 割り当てられた枠が、本当にその子のクラスのものであること
opt.picks.forEach((p) => {
  const q = g.parseSlotId_(p.slotId);
  m.eq(q.cls, p.cls, '枠IDのクラスと割り当て先が一致する: ' + p.name);
});

m.eq(view.days[1].options.length, 2, '2日目にも同じだけ組み合わせがある');

/* ---------------- 同じクラスのきょうだい ---------------- */
/* 配布プリントには「同じクラスの場合は1人ずつ」と書いてあるが、
   同じクラスにも複数のコマがあるので、実際には組み合わせが出る */

clearFailures();
const sameCls = m.unwrap(g.apiSiblingSlots({
  children: [{ cls: C1, no: 1, name: '生徒1_1' }, { cls: C1, no: 2, name: '生徒1_2' }]
}), '同じクラスのきょうだい');
m.ok(sameCls.days[0].options.length > 0,
  '同じクラスのきょうだいでも、続いたコマの組み合わせは出る（プリントの記載と食い違う）');

/* ---------------- まとめて予約 ---------------- */

clearFailures();
const picked = view.days[0].options[0];
const booked = m.unwrap(g.apiSiblingBook({
  children: PAIR,
  slotIds: picked.picks.map((p) => p.slotId),
  guardian: '保護者S',
  note: 'きょうだいです'
}), 'apiSiblingBook');

m.ok(/^\d{4}$/.test(booked.code), '予約コードは4桁');
m.eq(booked.bookings.length, 2, '2人ぶんの予約が返る');

picked.picks.forEach((p) => {
  m.eq(m.statusOf(g, p.slotId), '予約済', p.name + ' の枠が予約済になる');
  m.eq(m.slotValue(g, p.slotId, g.COL.CODE), booked.code, p.name + ' も同じ予約コード');
  m.eq(m.slotValue(g, p.slotId, g.COL.GUARDIAN), '保護者S', p.name + ' に保護者氏名が入る');
  m.eq(m.slotValue(g, p.slotId, g.COL.NOTE), 'きょうだいです', p.name + ' に連絡事項が入る');
});

m.eq(booked.bookings[0].ymd, DAY1, 'カレンダー登録用に ymd を返す');
m.ok(booked.bookings.every((b) => b.teacher !== undefined), '担任名を返す');

// 1つのコードで、それぞれの子の予約を引ける
const lookA = m.unwrap(g.apiLookup(Object.assign({}, KID_A, { code: booked.code })), 'Aを引く');
const lookB = m.unwrap(g.apiLookup(Object.assign({}, KID_B, { code: booked.code })), 'Bを引く');
m.ok(lookA.booking.slotId !== lookB.booking.slotId, '2人は別の枠に入っている');
m.eq(lookA.booking.cls, C1, 'Aは1年1組');
m.eq(lookB.booking.cls, C3, 'Bは2年1組');

/* ---------------- すでに予約済みの子がいる場合 ---------------- */

clearFailures();
const again = m.unwrap(g.apiSiblingSlots({ children: PAIR }), '予約後の apiSiblingSlots');
m.eq(again.already.length, 2, '予約済みの子を挙げる');
m.eq(again.days.length, 0, '予約済みがいるときは組み合わせを出さない');
m.eq(again.already[0].ownCls, C1, '在籍クラスを返す（手続き画面へ渡すため）');
m.eq(again.already[0].cls, C1, '予約が入っているクラスを返す');
m.ok(again.already[0].dateLabel.indexOf('10月29日') >= 0, 'いつ予約したかを返す');

m.failsWith(g.apiSiblingBook({
  children: PAIR, slotIds: [m.slotIdOf(g, DAY2, C1, 1), m.slotIdOf(g, DAY2, C3, 1)]
}), 'すでに予約が入っています', '予約済みの子を含むまとめて予約は受け付けない');

// 取り消せば、また使える
m.unwrap(g.apiCancel(Object.assign({}, KID_A, { code: booked.code })), 'Aを取消');
m.unwrap(g.apiCancel(Object.assign({}, KID_B, { code: booked.code })), 'Bを取消');
clearFailures();
m.eq(m.unwrap(g.apiSiblingSlots({ children: PAIR }), '取消後').already.length, 0,
  '取り消せば、またまとめて予約できる');

/* ---------------- 1人でも取れなければ、誰も予約しない ---------------- */
/* ここが本丸。途中まで書いて止まると、片方だけ予約が入った状態になる */

clearFailures();
const fresh = m.unwrap(g.apiSiblingSlots({ children: PAIR }), '再取得');
const target = fresh.days[0].options[0];
const slotA = target.picks[0].slotId;
const slotB = target.picks[1].slotId;

// 組み合わせを見せたあとで、片方の枠が別の家庭に取られる
const OTHER = { cls: C1, no: 2, name: '生徒1_2' };
m.unwrap(g.apiBook(Object.assign({}, OTHER, { slotId: slotA })), '別の家庭が先に押さえる');
m.eq(m.statusOf(g, slotA), '予約済', '片方の枠が埋まった');
m.eq(m.statusOf(g, slotB), '空き', 'もう片方はまだ空き');

clearFailures();
m.failsWith(g.apiSiblingBook({ children: PAIR, slotIds: [slotA, slotB] }),
  'ちょうど埋まりました', '1人ぶんが埋まっていたら、まとめて予約は失敗する');

m.eq(m.statusOf(g, slotB), '空き',
  '**失敗したとき、もう片方の枠は空きのまま**（片方だけ入る中途半端な状態を作らない）');
m.eq(m.slotValue(g, slotB, g.COL.STUDENT), '', '失敗したら生徒氏名も書かれていない');

// 後片付け
m.unwrap(g.apiCancel(Object.assign({}, OTHER, { code: m.slotValue(g, slotA, g.COL.CODE) })), '後片付け');

/* ---------------- 選んだ枠の取り違え ---------------- */

clearFailures();
m.failsWith(g.apiSiblingBook({ children: PAIR, slotIds: [slotA] }),
  'お子さんの人数と合いません', '枠の数が人数と違えば弾く');

clearFailures();
m.failsWith(g.apiSiblingBook({ children: PAIR, slotIds: [slotA, slotA] }),
  '同じ時間が2回選ばれています', '同じ枠を2回選んだら弾く');

clearFailures();
m.failsWith(g.apiSiblingBook({
  children: PAIR, slotIds: [m.slotIdOf(g, DAY1, C3, 1), m.slotIdOf(g, DAY1, C1, 2)]
}), '在籍クラスと合いません', '他の子のクラスの枠を渡したら弾く');

clearFailures();
m.failsWith(g.apiSiblingBook({ children: PAIR, slotIds: ['ありえない枠ID', slotB] }),
  '時間が見つかりませんでした', '存在しない枠IDを弾く');

m.eq(m.statusOf(g, slotB), '空き', 'どの失敗のあとも、枠は書き換わっていない');

/* ---------------- 空いた枠が足りない日 ---------------- */

clearFailures();
// 2年1組の1日目を全部「面談なし」にする
[1, 2, 3].forEach((i) => g.setNgFlag_(m.slotIdOf(g, DAY1, C3, i), true));
g.applyNgSlots();

const tight = m.unwrap(g.apiSiblingSlots({ children: PAIR }), '片方のクラスが埋まった日');
m.eq(tight.days[0].options.length, 0, '片方のクラスに空きが無い日は、組み合わせが出ない');
m.eq(tight.days[0].maxFit, 1, '何人までなら入るかを返す（画面の案内に使う）');
m.ok(tight.days[1].options.length > 0, '別の日には組み合わせが残っている');

[1, 2, 3].forEach((i) => g.setNgFlag_(m.slotIdOf(g, DAY1, C3, i), false));
g.applyNgSlots();

/* ---------------- 受付期間 ---------------- */

clearFailures();
g.setConfigValue_('公開', false);
g.dropRefCaches_();
m.failsWith(g.apiSiblingSlots({ children: PAIR }), '受付を停止しています', '非公開なら使えない');
g.setConfigValue_('公開', true);

// きょうだい専用の受付開始が入っていれば、そちらが優先される
g.setConfigValue_('きょうだい予約の受付開始', new Date(2099, 0, 1, 9, 0));
g.dropRefCaches_();
m.failsWith(g.apiSiblingSlots({ children: PAIR }),
  'きょうだいまとめての予約は', 'きょうだい専用の受付開始前は使えない');
m.unwrap(g.apiSlots(KID_A), 'そのあいだも、ひとりずつの予約は使える');

g.setConfigValue_('きょうだい予約の受付開始', '');
g.dropRefCaches_();

/* ---------------- 3人でも動く ---------------- */

clearFailures();
const TRIO = [KID_A, { cls: C2, no: 1, name: '生徒2_1' }, KID_B];
const trioView = m.unwrap(g.apiSiblingSlots({ children: TRIO }), '3人');
m.eq(trioView.children.length, 3, '3人ぶん返る');
m.eq(trioView.days[0].options.length, 1, '3コマに3人なら、続いた組み合わせは1通り');
m.eq(trioView.days[0].options[0].picks.map((p) => p.start),
  ['13:40', '14:05', '14:30'], '3人が続いたコマに並ぶ');

const trioBooked = m.unwrap(g.apiSiblingBook({
  children: TRIO,
  slotIds: trioView.days[0].options[0].picks.map((p) => p.slotId),
  guardian: '保護者T', note: ''
}), '3人まとめて予約');
m.eq(trioBooked.bookings.length, 3, '3人ぶん予約できる');
m.eq(new Set(trioBooked.bookings.map(() => trioBooked.code)).size, 1, '3人とも同じ予約コード');

/* ---------------- ログ ---------------- */

const log = g.__ss.getSheetByName(g.SH.LOG);
const rows = log.getRange(2, 1, log.getLastRow() - 1, 7).getValues();
const sib = rows.filter((r) => String(r[1]) === 'きょうだい予約');
m.ok(sib.length >= 2, 'きょうだい予約がログに残る');
m.ok(String(sib[sib.length - 1][6]).indexOf('3名') >= 0, 'ログに人数が残る');
m.ok(String(sib[sib.length - 1][6]).indexOf('コード') >= 0, 'ログに予約コードが残る');

m.report('test_sibling');
