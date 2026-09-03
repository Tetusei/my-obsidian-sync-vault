# ローカル検証環境

Apps Script を実際に動かさずに、Node.js だけでロジックを検証するための一式。
`gasmock.js` が SpreadsheetApp などを模した最小の実装で、
1つ上のフォルダの `.js` をそのまま読み込んで動かす。

## 実行

```bash
cd test
node test_booking.js
```

まとめて動かす場合:

```bash
cd test
for t in test_*.js; do printf "%-18s " "$t"; node "$t" | tail -1; done
```

## いまの状態（2026-09-03）

**この一式は一度失われ、作り直している途中。** 元は14スイート・約280項目あったが、
`test/` が git の管理外だったため、2台のPC間の同期が飛んだときに復元できなくなった。
**いまは git の追跡対象にしてある。消える前に必ずコミットすること。**

### できているもの

| ファイル | 内容 | 項目数 |
|---|---|---|
| `gasmock.js` | 土台。実機の制約を再現したモックと、テスト用の道具 | — |
| `test_booking.js` | 予約の確定・変更・取消、名簿照合、受付期間、予約ログ | 53 |

### まだ無いもの（作り直しの残り）

失うと痛い順に並べてある。上から順に足していく。

| ファイル | 内容 | なぜ要るか |
|---|---|---|
| `test_regen.js` | 枠再生成での予約の引き継ぎと、失う枠があるときの中止 | 設定を直した拍子に予約が消えるのを防ぐ |
| `test_sibling.js` | きょうだいまとめて予約 | 新しくて複雑、実運用の経験がゼロ |
| `test_grade.js` | 学年、交流学級（特別支援学級）の紐づけ | 同じ子が両方のクラスで予約できてしまう |
| `test_reserve.js` | 予備コマの記入・取り消し・二重予約の確認 | v4.6.13〜15 で作り替えたばかり |
| `test_ng.js` | だめなコマの指定・反映・再生成での引き継ぎ | |
| `test_warn.js` | 予約とだめなコマがぶつかったときの警告 | |
| `test_views.js` | 予約と表示更新の分離 | |
| `test_guard.js` | 作り直される場所の保護と、未登録の手入力の記録 | v4.6.14〜15 で追加 |
| `test_reset.js` | リセット（予約・ログ・名簿の消去範囲） | |
| `test_classes.js` | クラスの増減・改名 | |
| `test_addclass.js` | 受付中にクラスを増やす通し確認 | |
| `test_order.js` | シートの並び順 | |
| `test_check.js` | データ点検、数値設定の 0 の扱い | |
| `test_auto.js` | 入力エラーの可視化、だめなコマ自動反映、リマインド | |
| `test_qr.js` | QRコードを独立に復号して検証 | |
| `test_handout.js` | 配布プリントのQR描画 | |
| `test_integrity.js` | 入力ミスへの耐性（調査用。⚠は要検討の意味） | |

### 試算・計測（これも失われている）

| ファイル | 内容 |
|---|---|
| `bench_scale.js` | 規模別の「1件の予約にかかるシート操作回数」 |
| `sim_school.js` | 実際の学校規模（400名・17クラス）でのきょうだい成立率 |
| `sim_siblings.js` | 規模を変えた場合のきょうだい成立率 |

## テストの書き方

`gasmock.js` が道具を出している。

```js
const m = require('./gasmock');

const g = m.load();                       // プロジェクトの .js を全部読み込む
m.seedSchool(g, { perClass: 3 });         // 学校を1つ組み立てる

m.eq(actual, expected, '説明');            // 値の一致
m.ok(cond, '説明');                        // 真であること
m.unwrap(g.apiBook({...}), '説明');        // {ok:true} を期待し、中身を返す
m.failsWith(g.apiBook({...}), '文言', '説明'); // {ok:false} で、error に文言を含む
m.throwsWith(() => g.generateSlots(), '文言', '説明');

m.report('test_booking');                 // 最後に必ず呼ぶ（失敗があれば exit 1）
```

`seedSchool()` は**プロジェクト自身の `setupSystem()` と `syncClasses()` を通す**ので、
本番と同じ経路でシートができる。下ごしらえ専用の近道は作らないこと。
近道を作ると、セットアップ側が壊れてもテストが通ってしまう。

## モックについて

**実機の制約はできるだけ再現してある。** 甘くすると、テストが通るのに本番で落ちる。
これまで実際に、次の不具合がモックを厳しくしたことで見つかっている。

- シートは既定 1000行 × 26列しかない → 24クラスで範囲外エラー
- 非表示シートはアクティブにできない
- `clearDataValidations()` は Range のメソッドで、Sheet には無い

いま再現してあるもの:

- 範囲がシートの外に出ると例外（`ensureSheetSize_` を忘れると落ちる）
- `setValues` の行数・列数が範囲と違えば例外
- 行数・列数に 0 以下を渡すと例外（`getLastRow()` が 1 のときの `last - 1` を捕まえる）
- 非表示シートを `setActiveSheet` すると例外
- `CacheService` の有効期限を本当に見る（キャッシュの取り違えを捕まえる）
- 保護（`protect` / `getProtections` / `remove`）

再現していないもの（必要になったら足す）:

- 数式の計算。`=HYPERLINK(...)` は文字列のまま入る
- 書式（色・罫線・フォント）。鎖でつなげるだけの空実装
- 実行時間の上限、同時実行

**モックを緩めたくなったら、まず実機がどうなのかを確認すること。**

## テスト自身の確認

「通った」だけでは、そのテストが本当に噛んでいるか分からない。
新しいスイートを足したら、**わざとコードを壊して落ちることを1回は確かめる**。

```bash
cp ../Booking.js ../Booking.js.bak
sed -i 's/status = STATUS.BLOCKED;/status = STATUS.OPEN;/' ../Booking.js
node test_booking.js          # 落ちることを確認
mv -f ../Booking.js.bak ../Booking.js
```

`clasp push` の対象からは `.claspignore` で除外してある。
