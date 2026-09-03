/**
 * QRコードの生成（バイトモード / 誤り訂正レベル L / バージョン1〜10）。
 *
 * 外部サービスに保護者用URLを送らずに済むよう、自前で符号化する。
 * 画像は作らず、白黒の真偽値の並び（マトリクス）を返すだけにしてある。
 * 呼び出し側は、これをセルの背景色として描く。
 *
 * 規格: JIS X 0510 / ISO/IEC 18004
 */

/** 各バージョンの総コードワード数と、レベルLのデータコードワード数・ブロック構成 */
var QR_SPEC = {
  //          総CW, データCW, [ブロック数, ブロック内データCW] の組
  1:  { total: 26,  data: 19,  blocks: [[1, 19]] },
  2:  { total: 44,  data: 34,  blocks: [[1, 34]] },
  3:  { total: 70,  data: 55,  blocks: [[1, 55]] },
  4:  { total: 100, data: 80,  blocks: [[1, 80]] },
  5:  { total: 134, data: 108, blocks: [[1, 108]] },
  6:  { total: 172, data: 136, blocks: [[2, 68]] },
  7:  { total: 196, data: 156, blocks: [[2, 78]] },
  8:  { total: 242, data: 194, blocks: [[2, 97]] },
  9:  { total: 292, data: 232, blocks: [[2, 116]] },
  10: { total: 346, data: 274, blocks: [[2, 68], [2, 69]] }
};

/** バージョン2以降の位置合わせパターンの中心座標 */
var QR_ALIGN = {
  1: [], 2: [6, 18], 3: [6, 22], 4: [6, 26], 5: [6, 30],
  6: [6, 34], 7: [6, 22, 38], 8: [6, 24, 42], 9: [6, 26, 46], 10: [6, 28, 50]
};

/** レベルLの形式情報（マスク0〜7）。規格の固定値 */
var QR_FORMAT_L = [
  0x77C4, 0x72F3, 0x7DAA, 0x789D, 0x662F, 0x6318, 0x6C41, 0x6976
];

/** バージョン7以降の型番情報。規格の固定値 */
var QR_VERSION_INFO = {
  7: 0x07C94, 8: 0x085BC, 9: 0x09A99, 10: 0x0A4D3
};

/* ---------------- GF(256) ---------------- */

var QR_EXP = null;
var QR_LOG = null;

function qrInitGF_() {
  if (QR_EXP) return;
  QR_EXP = new Array(512);
  QR_LOG = new Array(256);
  var x = 1;
  for (var i = 0; i < 255; i++) {
    QR_EXP[i] = x;
    QR_LOG[x] = i;
    x <<= 1;
    if (x & 0x100) x ^= 0x11D;   // 原始多項式 x^8+x^4+x^3+x^2+1
  }
  for (var j = 255; j < 512; j++) QR_EXP[j] = QR_EXP[j - 255];
}

function qrMul_(a, b) {
  if (a === 0 || b === 0) return 0;
  return QR_EXP[QR_LOG[a] + QR_LOG[b]];
}

/** 誤り訂正コードワードを作る */
function qrRS_(data, ecLen) {
  qrInitGF_();

  // 生成多項式 (x-α^0)(x-α^1)...(x-α^(ecLen-1))
  var gen = [1];
  for (var i = 0; i < ecLen; i++) {
    var next = new Array(gen.length + 1);
    for (var k = 0; k < next.length; k++) next[k] = 0;
    for (var j = 0; j < gen.length; j++) {
      next[j] ^= gen[j];
      next[j + 1] ^= qrMul_(gen[j], QR_EXP[i]);
    }
    gen = next;
  }

  var rem = data.slice();
  for (var r = 0; r < ecLen; r++) rem.push(0);

  for (var p = 0; p < data.length; p++) {
    var factor = rem[p];
    if (factor === 0) continue;
    for (var q = 0; q < gen.length; q++) {
      rem[p + q] ^= qrMul_(gen[q], factor);
    }
  }
  return rem.slice(data.length);
}

/* ---------------- 符号化 ---------------- */

/** 文字列をUTF-8のバイト列にする */
function qrUtf8Bytes_(text) {
  var out = [];
  for (var i = 0; i < text.length; i++) {
    var c = text.charCodeAt(i);
    if (c < 0x80) {
      out.push(c);
    } else if (c < 0x800) {
      out.push(0xC0 | (c >> 6), 0x80 | (c & 0x3F));
    } else if (c >= 0xD800 && c <= 0xDBFF && i + 1 < text.length) {
      var lo = text.charCodeAt(i + 1);
      var cp = 0x10000 + ((c - 0xD800) << 10) + (lo - 0xDC00);
      out.push(0xF0 | (cp >> 18), 0x80 | ((cp >> 12) & 0x3F),
        0x80 | ((cp >> 6) & 0x3F), 0x80 | (cp & 0x3F));
      i++;
    } else {
      out.push(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 0x3F), 0x80 | (c & 0x3F));
    }
  }
  return out;
}

/** 収まる最小のバージョンを選ぶ */
function qrPickVersion_(byteLen) {
  for (var v = 1; v <= 10; v++) {
    var lenBits = v < 10 ? 8 : 16;             // バイトモードの文字数指示子
    var need = 4 + lenBits + byteLen * 8;
    if (QR_SPEC[v].data * 8 >= need) return v;
  }
  throw new Error('URLが長すぎてQRコードにできません（' + byteLen + 'バイト）。');
}

/** データ部のコードワードを組み立てる */
function qrDataCodewords_(bytes, version) {
  var spec = QR_SPEC[version];
  var lenBits = version < 10 ? 8 : 16;
  var bits = [];

  function push(value, n) {
    for (var i = n - 1; i >= 0; i--) bits.push((value >> i) & 1);
  }

  push(4, 4);                    // バイトモード
  push(bytes.length, lenBits);
  for (var i = 0; i < bytes.length; i++) push(bytes[i], 8);

  var capacity = spec.data * 8;
  var term = Math.min(4, capacity - bits.length);
  push(0, term);                 // 終端パターン
  while (bits.length % 8 !== 0) bits.push(0);

  var words = [];
  for (var b = 0; b < bits.length; b += 8) {
    var w = 0;
    for (var k = 0; k < 8; k++) w = (w << 1) | bits[b + k];
    words.push(w);
  }

  var pad = [0xEC, 0x11];        // 埋め草
  var p = 0;
  while (words.length < spec.data) words.push(pad[p++ % 2]);
  return words;
}

/** ブロックに分けて誤り訂正を付け、規格どおりに交互配置する */
function qrFinalCodewords_(dataWords, version) {
  var spec = QR_SPEC[version];
  var blocks = [];
  var pos = 0;

  var totalBlocks = 0;
  for (var g = 0; g < spec.blocks.length; g++) totalBlocks += spec.blocks[g][0];
  var ecLen = Math.floor((spec.total - spec.data) / totalBlocks);

  for (var grp = 0; grp < spec.blocks.length; grp++) {
    var count = spec.blocks[grp][0];
    var size = spec.blocks[grp][1];
    for (var n = 0; n < count; n++) {
      var d = dataWords.slice(pos, pos + size);
      pos += size;
      blocks.push({ data: d, ec: qrRS_(d, ecLen) });
    }
  }

  var out = [];
  var maxData = 0;
  for (var i = 0; i < blocks.length; i++) {
    if (blocks[i].data.length > maxData) maxData = blocks[i].data.length;
  }
  for (var c = 0; c < maxData; c++) {
    for (var b = 0; b < blocks.length; b++) {
      if (c < blocks[b].data.length) out.push(blocks[b].data[c]);
    }
  }
  for (var e = 0; e < ecLen; e++) {
    for (var b2 = 0; b2 < blocks.length; b2++) out.push(blocks[b2].ec[e]);
  }
  return out;
}

/* ---------------- 配置 ---------------- */

function qrNewMatrix_(size) {
  var m = [];
  for (var r = 0; r < size; r++) {
    var row = [];
    for (var c = 0; c < size; c++) row.push(null);   // null = 未確定
    m.push(row);
  }
  return m;
}

function qrPlaceFinder_(m, top, left) {
  for (var r = -1; r <= 7; r++) {
    for (var c = -1; c <= 7; c++) {
      var rr = top + r, cc = left + c;
      if (rr < 0 || cc < 0 || rr >= m.length || cc >= m.length) continue;
      var on = (r >= 0 && r <= 6 && (c === 0 || c === 6)) ||
        (c >= 0 && c <= 6 && (r === 0 || r === 6)) ||
        (r >= 2 && r <= 4 && c >= 2 && c <= 4);
      m[rr][cc] = on;
    }
  }
}

function qrPlaceFunctionPatterns_(m, version) {
  var size = m.length;

  qrPlaceFinder_(m, 0, 0);
  qrPlaceFinder_(m, 0, size - 7);
  qrPlaceFinder_(m, size - 7, 0);

  // タイミングパターン
  for (var i = 8; i < size - 8; i++) {
    m[6][i] = (i % 2 === 0);
    m[i][6] = (i % 2 === 0);
  }

  // 位置合わせパターン
  var centers = QR_ALIGN[version];
  for (var a = 0; a < centers.length; a++) {
    for (var b = 0; b < centers.length; b++) {
      var cr = centers[a], cc2 = centers[b];
      if ((cr <= 8 && cc2 <= 8) ||
        (cr <= 8 && cc2 >= size - 9) ||
        (cr >= size - 9 && cc2 <= 8)) continue;
      for (var dr = -2; dr <= 2; dr++) {
        for (var dc = -2; dc <= 2; dc++) {
          m[cr + dr][cc2 + dc] =
            (Math.abs(dr) === 2 || Math.abs(dc) === 2 || (dr === 0 && dc === 0));
        }
      }
    }
  }

  m[size - 8][8] = true;   // 常に暗のモジュール

  // 形式情報の領域を予約（値は後で入れる）
  for (var f = 0; f <= 8; f++) {
    if (m[8][f] === null) m[8][f] = false;
    if (m[f][8] === null) m[f][8] = false;
  }
  for (var g = 0; g < 8; g++) {
    if (m[8][size - 1 - g] === null) m[8][size - 1 - g] = false;
    if (m[size - 1 - g][8] === null) m[size - 1 - g][8] = false;
  }

  // 型番情報の領域を予約
  if (version >= 7) {
    for (var v = 0; v < 18; v++) {
      var row = Math.floor(v / 3);
      var col = size - 11 + (v % 3);
      m[row][col] = false;
      m[col][row] = false;
    }
  }
}

/** 機能パターン以外のマスを、右下から蛇行しながら埋める */
function qrPlaceData_(m, words, reserved) {
  var size = m.length;
  var bitIndex = 0;
  var total = words.length * 8;

  function bitAt(i) {
    if (i >= total) return 0;
    return (words[i >> 3] >> (7 - (i & 7))) & 1;
  }

  var col = size - 1;
  var upward = true;

  while (col > 0) {
    if (col === 6) col--;   // タイミングパターンの列は飛ばす
    for (var i = 0; i < size; i++) {
      var row = upward ? size - 1 - i : i;
      for (var s = 0; s < 2; s++) {
        var c = col - s;
        if (reserved[row][c]) continue;
        m[row][c] = bitAt(bitIndex++) === 1;
      }
    }
    upward = !upward;
    col -= 2;
  }
}

function qrMaskBit_(pattern, r, c) {
  switch (pattern) {
    case 0: return (r + c) % 2 === 0;
    case 1: return r % 2 === 0;
    case 2: return c % 3 === 0;
    case 3: return (r + c) % 3 === 0;
    case 4: return (Math.floor(r / 2) + Math.floor(c / 3)) % 2 === 0;
    case 5: return ((r * c) % 2 + (r * c) % 3) === 0;
    case 6: return (((r * c) % 2 + (r * c) % 3) % 2) === 0;
    default: return (((r + c) % 2 + (r * c) % 3) % 2) === 0;
  }
}

function qrApplyFormat_(m, mask) {
  var size = m.length;
  var bitsVal = QR_FORMAT_L[mask];

  for (var i = 0; i < 15; i++) {
    var bit = ((bitsVal >> i) & 1) === 1;

    // 縦側：左上から左下へ。タイミングパターンの交点は飛ばす。
    if (i < 6) m[i][8] = bit;
    else if (i < 8) m[i + 1][8] = bit;
    else m[size - 15 + i][8] = bit;

    // 横側：右上から左上へ。縦側と同じ15ビットを複写する。
    if (i < 8) m[8][size - 1 - i] = bit;
    else if (i === 8) m[8][7] = bit;
    else m[8][14 - i] = bit;
  }

  // 形式情報の複写と隣接する固定の黒マス。
  m[size - 8][8] = true;
}

function qrApplyVersionInfo_(m, version) {
  if (version < 7) return;
  var size = m.length;
  var info = QR_VERSION_INFO[version];
  for (var i = 0; i < 18; i++) {
    var bit = ((info >> i) & 1) === 1;
    var row = Math.floor(i / 3);
    var col = size - 11 + (i % 3);
    m[row][col] = bit;
    m[col][row] = bit;
  }
}

/** マスクの良し悪しを点数化する（低いほどよい） */
function qrPenalty_(m) {
  var size = m.length;
  var penalty = 0;

  // 規則1: 同じ色が5つ以上並ぶ
  for (var pass = 0; pass < 2; pass++) {
    for (var a = 0; a < size; a++) {
      var run = 1;
      for (var b = 1; b < size; b++) {
        var cur = pass === 0 ? m[a][b] : m[b][a];
        var prev = pass === 0 ? m[a][b - 1] : m[b - 1][a];
        if (cur === prev) {
          run++;
        } else {
          if (run >= 5) penalty += 3 + (run - 5);
          run = 1;
        }
      }
      if (run >= 5) penalty += 3 + (run - 5);
    }
  }

  // 規則2: 2×2の同色ブロック
  for (var r = 0; r < size - 1; r++) {
    for (var c = 0; c < size - 1; c++) {
      if (m[r][c] === m[r][c + 1] && m[r][c] === m[r + 1][c] && m[r][c] === m[r + 1][c + 1]) {
        penalty += 3;
      }
    }
  }

  // 規則3: 1:1:3:1:1 の並び
  var patA = [true, false, true, true, true, false, true, false, false, false, false];
  var patB = [false, false, false, false, true, false, true, true, true, false, true];
  for (var rr = 0; rr < size; rr++) {
    for (var cc = 0; cc + 10 < size; cc++) {
      var okA = true, okB = true, okC = true, okD = true;
      for (var k = 0; k < 11; k++) {
        if (m[rr][cc + k] !== patA[k]) okA = false;
        if (m[rr][cc + k] !== patB[k]) okB = false;
        if (m[cc + k][rr] !== patA[k]) okC = false;
        if (m[cc + k][rr] !== patB[k]) okD = false;
      }
      if (okA) penalty += 40;
      if (okB) penalty += 40;
      if (okC) penalty += 40;
      if (okD) penalty += 40;
    }
  }

  // 規則4: 暗モジュールの割合の偏り
  var dark = 0;
  for (var y = 0; y < size; y++) {
    for (var x = 0; x < size; x++) if (m[y][x]) dark++;
  }
  var ratio = (dark * 100) / (size * size);
  penalty += Math.floor(Math.abs(ratio - 50) / 5) * 10;

  return penalty;
}

/**
 * 文字列からQRコードのマトリクスを作る。
 * @param {string} text 埋め込む文字列（URLなど）
 * @return {Array<Array<boolean>>} true=黒 の二次元配列
 */
function qrMatrix_(text) {
  var bytes = qrUtf8Bytes_(String(text));
  var version = qrPickVersion_(bytes.length);
  var size = version * 4 + 17;

  var dataWords = qrDataCodewords_(bytes, version);
  var words = qrFinalCodewords_(dataWords, version);

  // どこが機能パターンかを覚えておく
  var base = qrNewMatrix_(size);
  qrPlaceFunctionPatterns_(base, version);

  var reserved = [];
  for (var r = 0; r < size; r++) {
    var row = [];
    for (var c = 0; c < size; c++) row.push(base[r][c] !== null);
    reserved.push(row);
  }

  var best = null, bestScore = Infinity;

  for (var mask = 0; mask < 8; mask++) {
    var m = qrNewMatrix_(size);
    for (var y = 0; y < size; y++) {
      for (var x = 0; x < size; x++) m[y][x] = base[y][x];
    }
    qrPlaceData_(m, words, reserved);

    for (var my = 0; my < size; my++) {
      for (var mx = 0; mx < size; mx++) {
        if (reserved[my][mx]) continue;
        if (qrMaskBit_(mask, my, mx)) m[my][mx] = !m[my][mx];
      }
    }

    qrApplyFormat_(m, mask);
    qrApplyVersionInfo_(m, version);

    var score = qrPenalty_(m);
    if (score < bestScore) { bestScore = score; best = m; }
  }

  return best;
}


/* ==================================================================
   QRを画像（PNG）にする

   セルの背景色で描くと、PDFに出力したとき黒マスが1pxはみ出し、
   輪郭が隣のマスに食い込む。読み取り機は倍率によって成否が変わり、
   実機では読めないことがあった。画像として貼れば輪郭が正確に出る。
   ================================================================== */

/** PNG用のCRC32 */
function qrCrc32_(bytes) {
  var table = qrCrc32_.table;
  if (!table) {
    table = qrCrc32_.table = [];
    for (var i = 0; i < 256; i++) {
      var c = i;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[i] = c >>> 0;
    }
  }
  var crc = 0xFFFFFFFF;
  for (var b = 0; b < bytes.length; b++) crc = table[(crc ^ bytes[b]) & 0xFF] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function qrAdler32_(bytes) {
  var a = 1, b = 0;
  for (var i = 0; i < bytes.length; i++) {
    a = (a + bytes[i]) % 65521;
    b = (b + a) % 65521;
  }
  return ((b << 16) | a) >>> 0;
}

function qrU32_(v) {
  return [(v >>> 24) & 0xFF, (v >>> 16) & 0xFF, (v >>> 8) & 0xFF, v & 0xFF];
}

function qrChunk_(type, data) {
  var body = [];
  for (var i = 0; i < type.length; i++) body.push(type.charCodeAt(i));
  for (var j = 0; j < data.length; j++) body.push(data[j]);
  return qrU32_(data.length).concat(body).concat(qrU32_(qrCrc32_(body)));
}

/**
 * QRの行列をPNGのバイト列にする。1画素1ビットの白黒。
 * 圧縮は使わず、そのまま格納するdeflateブロックで組み立てる。
 *
 * @param {Array<Array<boolean>>} m QRの行列
 * @param {number} scale 1マスを何画素で描くか
 * @param {number} quiet 周囲の余白（マス数）
 * @return {Array<number>} PNGのバイト列（0〜255）
 */
function qrPngBytes_(m, scale, quiet) {
  var n = m.length;
  var size = (n + quiet * 2) * scale;
  var raw = [];

  for (var y = 0; y < size; y++) {
    raw.push(0);                                  // フィルタなし
    var my = Math.floor(y / scale) - quiet;
    var acc = 0, bits = 0;
    for (var x = 0; x < size; x++) {
      var mx = Math.floor(x / scale) - quiet;
      var dark = my >= 0 && my < n && mx >= 0 && mx < n && !!m[my][mx];
      acc = ((acc << 1) | (dark ? 0 : 1)) & 0xFF;  // 0=黒 1=白
      bits++;
      if (bits === 8) { raw.push(acc); acc = 0; bits = 0; }
    }
    if (bits) {                                    // 行の余りは白で埋める
      acc = ((acc << (8 - bits)) | ((1 << (8 - bits)) - 1)) & 0xFF;
      raw.push(acc);
    }
  }

  // zlib（無圧縮）
  var z = [0x78, 0x01];
  var pos = 0;
  while (pos < raw.length) {
    var len = Math.min(65535, raw.length - pos);
    var last = (pos + len >= raw.length) ? 1 : 0;
    z.push(last);
    z.push(len & 0xFF, (len >>> 8) & 0xFF);
    var nlen = (~len) & 0xFFFF;
    z.push(nlen & 0xFF, (nlen >>> 8) & 0xFF);
    for (var k = 0; k < len; k++) z.push(raw[pos + k]);
    pos += len;
  }
  z = z.concat(qrU32_(qrAdler32_(raw)));

  var ihdr = qrU32_(size).concat(qrU32_(size)).concat([1, 0, 0, 0, 0]);  // 1bit グレースケール
  return [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]
    .concat(qrChunk_('IHDR', ihdr))
    .concat(qrChunk_('IDAT', z))
    .concat(qrChunk_('IEND', []));
}

/** Apps Script の Blob に渡せる符号付きバイト配列にする */
function qrPngBlob_(m, scale, quiet, name) {
  var u = qrPngBytes_(m, scale, quiet);
  var signed = [];
  for (var i = 0; i < u.length; i++) signed.push(u[i] > 127 ? u[i] - 256 : u[i]);
  return Utilities.newBlob(signed, 'image/png', name || 'qr.png');
}
