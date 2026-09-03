# ローカル検証環境

Apps Script を実際に動かさずに、Node.js だけでロジックを検証するための一式。
`gasmock.js` が SpreadsheetApp などを模した最小の実装で、
1つ上のフォルダの `.js` をそのまま読み込んで動かす。

## 実行

```bash
cd test
node test_sibling.js
```

まとめて動かす場合:

```bash
cd test
for t in test_*.js; do printf "%-18s " "$t"; node "$t" | tail -1; done
```

## テストの一覧

| ファイル | 内容 |
|---|---|
| `test_ng.js` | だめなコマの指定・反映・再生成での引き継ぎ |
| `test_warn.js` | 予約とだめなコマがぶつかったときの警告 |
| `test_more.js` | 予約表からの枠ID逆引き、連絡事項の保持 |
| `test_reset.js` | リセット（予約・ログ・名簿の消去範囲） |
| `test_classes.js` | クラスの増減・改名 |
| `test_addclass.js` | 受付中にクラスを増やす通し確認 |
| `test_order.js` | シートの並び順 |
| `test_check.js` | データ点検、数値設定の 0 の扱い |
| `test_auto.js` | 入力エラーの可視化、だめなコマ自動反映、リマインド |
| `test_qr.js` | QRコードを独立に復号して検証 |
| `test_handout.js` | 配布プリントのQR描画 |
| `test_views.js` | 予約と表示更新の分離 |
| `test_grade.js` | 学年、交流学級（特別支援学級）の紐づけ |
| `test_sibling.js` | きょうだいまとめて予約 |
| `test_integrity.js` | 入力ミスへの耐性（調査用。⚠は要検討の意味） |

## 試算・計測

| ファイル | 内容 |
|---|---|
| `bench_scale.js` | 規模別の「1件の予約にかかるシート操作回数」 |
| `sim_school.js` | 実際の学校規模（400名・17クラス）でのきょうだい成立率 |
| `sim_siblings.js` | 規模を変えた場合のきょうだい成立率 |

## モックについて

**実機の制約はできるだけ再現してある。** 甘くすると、テストが通るのに本番で落ちる。
これまで実際に、次の不具合がモックを厳しくしたことで見つかった。

- シートは既定 1000行 × 26列しかない → 24クラスで範囲外エラー
- 非表示シートはアクティブにできない
- `clearDataValidations()` は Range のメソッドで、Sheet には無い

モックを緩めたくなったら、まず実機がどうなのかを確認すること。

`clasp push` の対象からは `.claspignore` で除外してある。
