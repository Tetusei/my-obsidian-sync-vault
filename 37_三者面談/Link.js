/**
 * 交流学級の紐づけ。
 *
 * 特別支援学級の児童生徒は、通常学級（交流学級）にも在籍している。
 * 名簿は両方に載るが、面談は「どちらか一方の担任」と行う。
 * 何もしないと、システムからは別人に見えるため両方で予約できてしまう。
 *
 * 「交流学級」シートで2つの在籍を結びつけ、
 *   ・保護者はどちらの担任と面談するかを選べる
 *   ・片方を予約したら、もう片方では予約できない
 * ようにする。
 */

/** 交流学級シートの列（1始まり） */
var LINK_COL = {
  SPECIAL_CLASS: 1,
  SPECIAL_NO: 2,
  NORMAL_CLASS: 3,
  NORMAL_NO: 4,
  NAME: 5
};

var LINK_HEADER = ['特別支援学級', '出席番号', '交流学級', '出席番号', '氏名（確認用）'];

/**
 * 紐づけを読む。
 * @return {Array<{a:{cls:string,no:number}, b:{cls:string,no:number}, name:string}>}
 */
function getLinks_() {
  var hit = cacheGet_(CACHE_KEY.LINKS);
  if (hit) return hit;

  var out = readLinks_();
  cachePut_(CACHE_KEY.LINKS, out, CACHE_SEC.LINKS);
  return out;
}

/** 交流学級シートを実際に読む */
function readLinks_() {
  var sh = ss_().getSheetByName(SH.LINK);
  if (!sh) return [];

  var last = sh.getLastRow();
  if (last < 2) return [];

  var vals = sh.getRange(2, 1, last - 1, LINK_HEADER.length).getValues();
  var out = [];

  for (var i = 0; i < vals.length; i++) {
    var sCls = String(vals[i][LINK_COL.SPECIAL_CLASS - 1] || '').trim();
    var sNo = Number(vals[i][LINK_COL.SPECIAL_NO - 1]);
    var nCls = String(vals[i][LINK_COL.NORMAL_CLASS - 1] || '').trim();
    var nNo = Number(vals[i][LINK_COL.NORMAL_NO - 1]);
    if (!sCls || !sNo || !nCls || !nNo) continue;

    out.push({
      a: { cls: sCls, no: sNo },
      b: { cls: nCls, no: nNo },
      name: String(vals[i][LINK_COL.NAME - 1] || '').trim()
    });
  }
  return out;
}

/** (クラス, 出席番号) → 相手側の在籍。無ければ null。 */
function linkedIdentity_(cls, no) {
  var links = getLinks_();
  var c = String(cls), n = Number(no);
  for (var i = 0; i < links.length; i++) {
    if (links[i].a.cls === c && links[i].a.no === n) return links[i].b;
    if (links[i].b.cls === c && links[i].b.no === n) return links[i].a;
  }
  return null;
}

/**
 * その枠に面談が入っているか。
 * 通常の予約に加え、担任が予備の枠へ手で入れたものも数える。
 */
function isTakenSlot_(v) {
  var st = String(v[COL.STATUS - 1]);
  if (st === STATUS.BOOKED) return true;
  return st === STATUS.RESERVE && !!v[COL.NUMBER - 1];
}

/**
 * その児童（と交流先）が、すでに予約を持っているか。
 * @return {Object|null} 予約されている枠の行データ
 */
function findExistingBookingFor_(slots, cls, no, linked) {
  var targets = [{ cls: String(cls), no: Number(no) }];
  if (linked) targets.push({ cls: String(linked.cls), no: Number(linked.no) });

  for (var i = 0; i < slots.length; i++) {
    var v = slots[i].v;
    // 担任が予備の枠に入れたぶんも「予約あり」として扱う
    if (!isTakenSlot_(v)) continue;
    for (var t = 0; t < targets.length; t++) {
      if (String(v[COL.CLASS - 1]) === targets[t].cls &&
        Number(v[COL.NUMBER - 1]) === targets[t].no) {
        return slots[i];
      }
    }
  }
  return null;
}

/**
 * 予約済みの (クラス_出席番号) の集合。交流先も予約済みとして数える。
 * 未予約者の一覧づくりに使う。
 */
function bookedKeySet_(slots) {
  var booked = {};
  for (var i = 0; i < slots.length; i++) {
    var v = slots[i].v;
    if (!isTakenSlot_(v)) continue;
    booked[String(v[COL.CLASS - 1]) + '_' + Number(v[COL.NUMBER - 1])] = true;
  }

  // 交流学級で予約していれば、もう片方も予約済みとして扱う
  var links = getLinks_();
  for (var k = 0; k < links.length; k++) {
    var ka = links[k].a.cls + '_' + links[k].a.no;
    var kb = links[k].b.cls + '_' + links[k].b.no;
    if (booked[ka]) booked[kb] = 'linked';
    else if (booked[kb]) booked[ka] = 'linked';
  }
  return booked;
}

/** 「交流学級」シートを用意する（無ければ作る） */
function ensureLinkSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SH.LINK);
  var created = false;

  if (!sh) {
    sh = ss.insertSheet(SH.LINK);
    created = true;
  }
  ensureSheetSize_(sh, 100, LINK_HEADER.length + 1);

  sh.getRange(1, 1, 1, LINK_HEADER.length).setValues([LINK_HEADER])
    .setFontWeight('bold').setBackground('#e8eaed').setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setColumnWidth(LINK_COL.SPECIAL_CLASS, 140);
  sh.setColumnWidth(LINK_COL.SPECIAL_NO, 90);
  sh.setColumnWidth(LINK_COL.NORMAL_CLASS, 140);
  sh.setColumnWidth(LINK_COL.NORMAL_NO, 90);
  sh.setColumnWidth(LINK_COL.NAME, 160);

  return created;
}

/**
 * 紐づけの内容を点検する。名簿に無い在籍を指していないか、重複していないか。
 * @return {Array<string>} 問題の説明（無ければ空）
 */
function checkLinks_() {
  var links = getLinks_();
  if (!links.length) return [];

  var roster = getRoster();
  var known = {};
  for (var r = 0; r < roster.length; r++) known[roster[r].cls + '_' + roster[r].no] = roster[r].name;

  var problems = [];
  var seen = {};

  for (var i = 0; i < links.length; i++) {
    var L = links[i];
    var ka = L.a.cls + ' ' + L.a.no + '番';
    var kb = L.b.cls + ' ' + L.b.no + '番';

    if (!known[L.a.cls + '_' + L.a.no]) problems.push(ka + ' が名簿にありません');
    if (!known[L.b.cls + '_' + L.b.no]) problems.push(kb + ' が名簿にありません');

    [L.a.cls + '_' + L.a.no, L.b.cls + '_' + L.b.no].forEach(function (key) {
      if (seen[key]) problems.push(key.replace('_', ' ') + '番 が複数の行で紐づけられています');
      seen[key] = true;
    });

    var na = known[L.a.cls + '_' + L.a.no];
    var nb = known[L.b.cls + '_' + L.b.no];
    if (na && nb && norm_(na) !== norm_(nb)) {
      problems.push(ka + '（' + na + '）と ' + kb + '（' + nb + '）は別の氏名です');
    }
  }
  return problems;
}
