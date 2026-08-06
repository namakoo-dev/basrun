# basrun

平文の Basic ソースを、**マクロを格納できない文書**に対して走らせる。

```
basrun sync  <dir> [lib]                       ソース群 -> ライブラリ
basrun pull  <dir> [lib] [--book <doc>]        ライブラリ -> ソース群（救出）
basrun apply <book> <dir> <lib> <Module.Sub>   同期して、文書に適用して保存
basrun stop                                    起動した LibreOffice を落とす
```

## 何を解いているか

`.xlsx` は仕様としてマクロを格納できない（マクロは `.xlsm` 側）。だから
「この Excel を処理する VBA を書いて」は、`.xlsx` を渡された時点で素直には成立しない。
`.xlsm` に変換する、値を直接書き込む、といった妥協に流れやすい。

**制約は「格納」の側にあって「実行」の側には無い。** ソースを実行時にライブラリへ
流し込めば、文書は普通の `.xlsx` のまま処理できる。

回り道した先のほうが、本来の形として良い:

| | 埋め込みマクロ | basrun |
|---|---|---|
| ソースの可読性 | 文書内の XML に埋まる | 平文。`git diff` が読める |
| レビュー | 実質できない | 通常のコードレビュー |
| 対象の形式 | マクロを持てる形式のみ | `.xlsx` でも `.ods` でも同じソース |
| 真実の在処 | 文書 | ソース。ライブラリは実行時のキャッシュ |

## 使い方

```bash
# .bas を書いて、マクロを持たない .xlsx に適用する
python basrun.py apply book.xlsx src MyLib Amount.FillAmounts --backup

# 文書に埋まってしまっているマクロを平文に救出する
python basrun.py pull rescued Standard --book legacy.ods
```

対象の `Sub` は**文書を引数で受けること**:

```basic
Sub FillAmounts(oDoc As Object)
    Dim oSheet As Object
    oSheet = oDoc.Sheets.getByIndex(0)
    ...
End Sub
```

`ThisComponent` に頼ってはいけない。`--headless` + 非表示で開いた文書は
「現在のコンポーネント」ではないので、**成功したように見えて何も起きない**
（実測: 「読み込んだ / 実行した / 保存した」と全部表示され、セルは 1 つも
変わっていなかった）。

## 2 つの道具を合わせている

同期は **[obasync](https://pypi.org/project/obasync/)**（imacat 作、Apache-2.0）に
任せる。双方向・差分・余剰モジュールの削除まで実装されていて、書き直す理由が無い。
`vendor/obasync/` に**無改変で**同梱している。

実行はこちら側が持つ。obasync にも `--run` はあるが `script.invoke((), (), ())` で
引数を渡しておらず、`ThisComponent` に依存する（上記の穴に落ちる）。

`--book` は**どちらにも無かった**。obasync は文書を開けず（既に開いているものしか
選べない）、こちらの実行部は引き出せなかった。組んで初めて
「文書を開いて、中のマクロを引き出して閉じる」が成立する。

## 確認していること / していないこと

**実測で確認済み:**

- `.xlsx`（`vbaProject.bin` を持たない）に対して計算結果と書式が入ること
- `pull` の往復でソースの内容が保たれること（差は行末 LF/CRLF のみ）
- 埋め込みマクロを持つ `.ods` から平文への救出。**元ファイルは無変更**（SHA-256 で確認）
- `stop` が接続先の LibreOffice だけを終了させること（`taskkill` しない）

**確認していない:**

- **Windows でしか動かしていない。** Linux / macOS のパス解決は書いてあるが未検証
- 検証した表は**1 種類だけ**。結合セル・複数シート・巨大な表は試していない
- 同時に複数の文書を開いた状態での `--target` の曖昧性
- 同梱している obasync は **2017 年公開で以降更新が無い**。現行 LibreOffice で
  動くことは確認したが、将来の互換性は保証できない
- 自動テストが無い（`tests/` にあるのは手動検証用の fixture）

## 前提

- LibreOffice（同梱の Python を使う。素の Python には `uno` が無い）
- 専用プロファイル `~/.nagi/lo-profile` で起動する。
  **既定プロファイルは使わない** —— GUI で開いている LibreOffice を巻き込んで壊すため

環境変数で上書きできる: `BASRUN_OFFICE` / `BASRUN_PROFILE` / `BASRUN_PORT`

## ライセンス

`vendor/obasync/` は imacat 氏による obasync 0.10 で、Apache License 2.0。
無改変で同梱しており、著作権表示とライセンス条項はスクリプト冒頭にそのまま残っている。

それ以外の部分の扱いは未定。
