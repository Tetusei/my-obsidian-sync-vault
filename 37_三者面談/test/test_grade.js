/**
 * 学年と、交流学級（特別支援学級）の紐づけ。
 *
 * 特別支援学級の児童生徒は、通常学級（交流学級）にも在籍している。
 * 名簿は両方に載るが、面談は**どちらか一方の担任**と行う。
 * 何もしないとシステムからは別人に見えるため、**両方で予約できてしまう。**
 *
 * 二重に入っても、それぞれのクラスの予約表を見ているだけでは気づけない。
 * 当日、同じ時間に別の教室で2つの面談が待っている、という形で表に出る。
 */

'use strict';

const m = require('./gasmock');

const g = m.load();
m.seedSchool(g, {
  perClass: 3,
  classes: [
    { name: '1年1組', teacher: '山田', email: 'y@example.jp', grade: '1年' },
    { name: '1年2組', teacher: '鈴木', email: 's@example.jp', grade: '1年' },
    { name: 'ひまわり学級', teacher: '田中', email: 't@example.jp', grade: '特別支援' }
  ]
});

const NORMAL = '1年1組';
const SPECIAL = 'ひまわり学級';
const DAY1 = '2026-10-29';

/** 名簿を書き換える */
function setRoster(cls, rows) {
  const sh = g.__ss.getSheetByName(g.CLASS_SHEET_PREFIX + cls);
  sh.getRange(2, 1, 10, 2).setValues(
    Array.from({ length: 10 }, (_, i) => rows[i] || ['', '']));
  g.clearRosterCache_();
  g.dropRefCaches_();
}

/** 交流学級シートを書き換える */
function setLinks(rows) {
  const sh = g.__ss.getSheetByName(g.SH.LINK);
  sh.getRange(2, 1, 10, 5).setValues(
    Array.from({ length: 10 }, (_, i) => rows[i] || ['', '', '', '', '']));
  g.dropRefCaches_();
}

/* ---------------- 学年の読み取り ---------------- */

m.eq(g.inferGrade_('1年1組'), '1年', 'クラス名から学年を読む');
m.eq(g.inferGrade_('3年2組'), '3年', '2桁でない学年も読む');
m.eq(g.inferGrade_(' 6 年 1 組'), '6年', '空白が入っていても読む');
m.eq(g.inferGrade_('特別支援学級'), '特別支援', '特別支援学級を判別する');
m.eq(g.inferGrade_('特支1組'), '特別支援', '「特支」でも判別する');
m.eq(g.inferGrade_('ひまわり学級'), '',
  '「ひまわり学級」からは学年を推定できない（クラスシートの学年列に書く必要がある）');
m.eq(g.inferGrade_(''), '', '空欄なら空欄');

m.eq(g.gradeOrder_('1年'), 1, '学年の並び順は数字そのもの');
m.eq(g.gradeOrder_('6年'), 6, '6年');
m.eq(g.gradeOrder_('特別支援'), 900, '特別支援はいちばん後ろ');
m.eq(g.gradeOrder_('その他'), 800, '読めない学年は特別支援の手前');

/* ---------------- クラスシートの学年列 ---------------- */

const classes = g.getClasses();
m.eq(classes.map((c) => c.grade), ['1年', '1年', '特別支援'],
  '学年列に書いた値がそのまま使われる');
m.eq(classes.map((c) => c.special), [false, false, true],
  '学年列が「特別支援」のクラスだけ special になる（クラス名からは推定しない）');
m.ok(g.hasSpecialClass_(), '特別支援学級があることを判定できる');

const clsSh = g.__ss.getSheetByName(g.SH.CLASSES);
m.eq(String(clsSh.getRange(1, 4).getValue()), '学年', 'クラスシートに学年列がある');

/* ---------------- 名簿を、両方に載っている状態にする ---------------- */

setRoster(NORMAL, [[1, '生徒1_1'], [2, '生徒1_2'], [3, '交流 太郎']]);
setRoster(SPECIAL, [[1, '交流 太郎'], [2, '生徒3_2'], [3, '生徒3_3']]);

const KID_SP = { cls: SPECIAL, no: 1, name: '交流 太郎' };
const KID_NM = { cls: NORMAL, no: 3, name: '交流 太郎' };

// 紐づける前は、システムから見て別人
m.eq(g.linkedIdentity_(SPECIAL, 1), null, '紐づけ前は相手が居ない');

setLinks([[SPECIAL, 1, NORMAL, 3, '交流 太郎']]);

m.eq(g.linkedIdentity_(SPECIAL, 1), { cls: NORMAL, no: 3 }, '特別支援側から通常学級を引ける');
m.eq(g.linkedIdentity_(NORMAL, 3), { cls: SPECIAL, no: 1 }, '通常学級側から特別支援を引ける');
m.eq(g.linkedIdentity_(NORMAL, 1), null, '関係のない子には相手が居ない');

/* ---------------- 保護者の画面に、どちらの担任かを出す ---------------- */

const viewSp = m.unwrap(g.apiSlots(KID_SP), '特別支援側から見る');
m.eq(viewSp.linked.cls, NORMAL, 'もう一方のクラスを伝える');
m.eq(viewSp.linked.teacher, '山田', 'もう一方の担任名も伝える');
m.eq(viewSp.linkedBooking, null, 'まだどちらにも予約は無い');
m.eq(viewSp.teacher, '田中', '自分のクラスの担任名');

const viewNm = m.unwrap(g.apiSlots(KID_NM), '通常学級側から見る');
m.eq(viewNm.linked.cls, SPECIAL, '逆向きにも伝える');

/* ---------------- 二重予約を防ぐ ---------------- */

const nmSlot = m.slotIdOf(g, DAY1, NORMAL, 1);
const spSlot = m.slotIdOf(g, DAY1, SPECIAL, 2);

const bookedNm = m.unwrap(g.apiBook(Object.assign({}, KID_NM, { slotId: nmSlot })),
  '通常学級で予約');
m.eq(m.statusOf(g, nmSlot), '予約済', '通常学級の枠が埋まる');

// ここが本丸。紐づけが効いていないと、特別支援側でも予約できてしまう
const blocked = g.apiBook(Object.assign({}, KID_SP, { slotId: spSlot }));
m.failsWith(blocked, 'すでに ' + NORMAL + ' の担任と面談を予約されています',
  '**もう一方のクラスで予約済みなら、こちらでは予約できない**');
m.eq(m.statusOf(g, spSlot), '空き', '弾いたので特別支援側の枠は空きのまま');
m.ok(String(blocked.error).indexOf('10月29日') >= 0, 'いつ予約したかを伝える');
m.ok(String(blocked.error).indexOf('確認・変更・取消') >= 0, 'どこで手続きするかを伝える');

// 予約済みであることも、もう一方の画面に出る
const viewSp2 = m.unwrap(g.apiSlots(KID_SP), '予約後に特別支援側から見る');
m.eq(viewSp2.linkedBooking.cls, NORMAL, 'もう一方で予約済みだと伝える');
m.eq(viewSp2.linkedBooking.teacher, '山田', '相手の担任名も伝える');
m.eq(viewSp2.linkedBooking.start, '13:40', '予約されている時刻を伝える');
m.eq(viewSp2.existing, null, '自分のクラスには予約が無いので existing は空');

/* ---------------- 取り消せば、もう一方で予約できる ---------------- */

m.unwrap(g.apiCancel(Object.assign({}, KID_NM, { code: bookedNm.code })), '通常学級の予約を取消');
const bookedSp = m.unwrap(g.apiBook(Object.assign({}, KID_SP, { slotId: spSlot })),
  '取消後に特別支援側で予約');
m.eq(m.statusOf(g, spSlot), '予約済', '特別支援側で予約できる');

// 逆向きも塞がっていること
m.failsWith(g.apiBook(Object.assign({}, KID_NM, { slotId: nmSlot })),
  'すでに ' + SPECIAL + ' の担任と面談を予約されています',
  '逆向き（特別支援→通常学級）も塞がっている');

/* ---------------- 予備の枠に入れた場合も「予約あり」 ---------------- */

m.unwrap(g.apiCancel(Object.assign({}, KID_SP, { code: bookedSp.code })), '特別支援の予約を取消');

g.setConfigValue_(g.RESERVE_COUNT_KEY, 1);
g.dropRefCaches_();
g.generateSlots();

const reserveId = m.slotIdOf(g, DAY1, NORMAL, g.RESERVE_INDEX_BASE + 1);
g.setReserveStudent_(reserveId, '3', 'テスト');   // 通常学級の予備に「交流 太郎」を入れる
m.eq(m.slotValue(g, reserveId, g.COL.STUDENT), '交流 太郎', '予備に入った');

m.failsWith(g.apiBook(Object.assign({}, KID_SP, { slotId: m.slotIdOf(g, DAY1, SPECIAL, 1) })),
  'すでに ' + NORMAL + ' の担任と面談を予約されています',
  '担任が予備の枠に入れたぶんも「予約あり」として、もう一方を塞ぐ');

g.setReserveStudent_(reserveId, '', 'テスト');
g.setConfigValue_(g.RESERVE_COUNT_KEY, 0);
g.dropRefCaches_();
g.generateSlots();

/* ---------------- 未予約の一覧に、二重に出さない ---------------- */

const again = m.unwrap(g.apiBook(Object.assign({}, KID_NM, { slotId: nmSlot })), '通常学級で予約');

const groups = g.unbookedByClass_();
const spGroup = groups.find((x) => x.cls === SPECIAL);
const nmGroup = groups.find((x) => x.cls === NORMAL);
m.ok(!spGroup.students.some((s) => s.no === 1),
  '交流学級で予約していれば、特別支援学級の未予約一覧には出ない');
m.ok(!nmGroup.students.some((s) => s.no === 3),
  '予約した本人も、通常学級の未予約一覧には出ない');

g.setConfigValue_('管理パスコード', 'test1234');
g.dropRefCaches_();
m.failsWith(g.apiAdminUnbooked('ちがう合言葉'), 'パスコードが違います',
  '管理画面は合言葉が違えば開かない');
const unbooked = m.unwrap(g.apiAdminUnbooked('test1234'), 'apiAdminUnbooked').students || [];
m.ok(!unbooked.some((s) => s.cls === SPECIAL && s.no === 1),
  '担任用の管理画面の未予約一覧にも、二重に出さない');

/* ---------------- きょうだい予約と交流学級 ---------------- */

const SIB = { cls: '1年2組', no: 1, name: '生徒2_1' };
const sibView = m.unwrap(g.apiSiblingSlots({ children: [KID_SP, SIB] }),
  '特別支援側の在籍で、きょうだい予約を試す');
m.eq(sibView.already.length, 1, '交流学級での予約を「予約済み」として拾う');
m.eq(sibView.already[0].ownCls, SPECIAL, '入力された在籍（特別支援）を返す');
m.eq(sibView.already[0].cls, NORMAL, '実際に予約が入っているクラスを返す');
m.eq(sibView.days.length, 0, '予約済みがいるので組み合わせは出さない');

m.unwrap(g.apiCancel(Object.assign({}, KID_NM, { code: again.code })), '後片付け');

/* ---------------- 交流学級シートの点検 ---------------- */

m.eq(g.checkLinks_(), [], '正しい紐づけなら問題なし');

setLinks([[SPECIAL, 1, NORMAL, 99, '交流 太郎']]);
let problems = g.checkLinks_();
m.ok(problems.some((p) => p.indexOf('名簿にありません') >= 0),
  '名簿に無い在籍を指していたら知らせる');

setLinks([[SPECIAL, 1, NORMAL, 3, ''], [SPECIAL, 1, '1年2組', 1, '']]);
problems = g.checkLinks_();
m.ok(problems.some((p) => p.indexOf('複数の行で紐づけられています') >= 0),
  '同じ子が2回紐づけられていたら知らせる');

setLinks([[SPECIAL, 2, NORMAL, 3, '']]);   // 生徒3_2 と 交流 太郎 は別人
problems = g.checkLinks_();
m.ok(problems.some((p) => p.indexOf('別の氏名です') >= 0),
  '氏名が食い違う紐づけを知らせる（入力ミスの可能性）');

// 点検（🩺）からも拾えること
const check = g.checkData();
m.ok(check.errors.some((e) => e.title.indexOf(g.SH.LINK) >= 0),
  '交流学級の誤りは、データ点検で「直したほうがよいもの」に出る');

setLinks([]);
m.eq(g.checkLinks_(), [], '紐づけが無ければ問題なし');
m.eq(g.linkedIdentity_(SPECIAL, 1), null, '紐づけを消せば別人に戻る');

/* ---------------- 全体ビューの学年色 ---------------- */

const gradeList = ['1年', '2年', '3年'];
m.eq(g.gradeHeaderColor_('特別支援', gradeList), '#dfe1e5',
  '特別支援は並び順に関係なく灰色で固定');
m.ok(g.gradeHeaderColor_('1年', gradeList) !== g.gradeHeaderColor_('2年', gradeList),
  '隣り合う学年は違う色になる');
m.eq(g.gradeHeaderColor_('読めない学年', gradeList), '#dfe1e5',
  '一覧に無い学年は灰色にする');

// 本文は「空き」の緑に学年色を薄く重ねる。⚠ の赤には重ねない
const openTint = g.bodyColor_('#e6f4ea', { hex: '#d2e3fc', shade: false });
m.ok(openTint !== '#e6f4ea', '本文に学年の帯が乗る');
m.eq(g.bodyColor_('#fce8e6', null), '#fce8e6', '帯が無ければ色はそのまま');

// 1つおきに少し暗くする（淡い色を重ねるだけでは隣の学年と見分けが付かないため）
const light = g.bodyColor_('#e6f4ea', { hex: '#d2e3fc', shade: false });
const dark = g.bodyColor_('#e6f4ea', { hex: '#d2e3fc', shade: true });
m.ok(light !== dark, '1つおきの学年は明るさを変える');
m.ok(parseInt(dark.slice(1, 3), 16) < parseInt(light.slice(1, 3), 16),
  '暗くするほうが実際に暗い');

/* ---------------- 学年ごとの並び ---------------- */

const init = m.unwrap(g.apiSiblingInit(), 'apiSiblingInit');
m.eq(init.grades.map((x) => x.grade), ['1年', '特別支援'],
  'きょうだい予約の学年一覧で、特別支援はいちばん後ろに並ぶ');

m.report('test_grade');
