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

## Word / PowerPoint（実験的）

`apply` の仕組みは Calc（`.xlsx`/`.ods`）専用ではない。文書を引数で受ける `Sub` であれば、
Writer（`.docx`）・Impress（`.pptx`）でも同じ流れで動く。`src-doc/Doc.bas`（段落の走査と
差し替え）・`src-ppt/Deck.bas`（スライド一括の題差し替え・フッタ付与）に例がある。

★ **自動テストがあるのは Calc（`.xlsx`）だけ。** docx/pptx は `tests/word_test.docx`・
`tests/ppt/deck.pptx`・`tests/ppt/fresh.pptx` を使って手動で動作確認したのみで、
pytest には組み込んでいない。回帰しても CI では検知できない。

## 2 つの道具を合わせている

同期は **[obasync](https://pypi.org/project/obasync/)**（imacat 作、Apache-2.0）に
任せる。双方向・差分・余剰モジュールの削除まで実装されていて、書き直す理由が無い。
`vendor/obasync/` に**無改変で**同梱している。

実行はこちら側が持つ。obasync にも `--run` はあるが `script.invoke((), (), ())` で
引数を渡しておらず、`ThisComponent` に依存する（上記の穴に落ちる）。

`--book` は**どちらにも無かった**。obasync は文書を開けず（既に開いているものしか
選べない）、こちらの実行部は引き出せなかった。組んで初めて
「文書を開いて、中のマクロを引き出して閉じる」が成立する。

### 同梱物の危険な挙動（塞いである）

**obasync は接続できないと、自分で LibreOffice を起動する**（`vendor/obasync/obasync`
786-790 行の `Popen`）。そのとき `-env:UserInstallation` も `--headless` も付けないため、
**利用者の既定プロファイルが使われ、窓も出る。**

2026-08-06 に実際に踏んだ。`basrun` 側が obasync に `-p` を渡し忘れていたため
obasync は常に 2002 を見に行き、そこが空だったので既定プロファイルで LibreOffice を
起こし、そこへライブラリを書き込んだ。

`run_obasync()` で 2 段に塞いである:

1. `-p` を必ず渡す（`ensure_office` が用意したポートと一致させる）
2. **呼ぶ直前に接続を確かめ、応答が無ければこちらで中止する** ——
   obasync に「繋がらない」状況を渡さない

隔離が効いていることは実測で確認している（別ポート + 別プロファイルで `sync` を
実行し、ライブラリが隔離側にだけ作られ、既定プロファイルが無変化であること）。

## 確認していること / していないこと

**実測で確認済み:**

- `.xlsx`（`vbaProject.bin` を持たない）に対して計算結果と書式が入ること
- `pull` の往復でソースの内容が保たれること（差は行末 LF/CRLF のみ）
- 埋め込みマクロを持つ `.ods` から平文への救出。**元ファイルは無変更**（SHA-256 で確認）
- `stop` が接続先の LibreOffice だけを終了させること（`taskkill` しない）。
  ★ **戻り値 0 は「もう閉じている」を意味する** —— terminate を投げただけで
  「終了させた」と表示していたのを 2026-08-06 に直した（下記）

**確認していない:**

- **Windows でしか動かしていない。** Linux / macOS のパス解決は書いてあるが未検証
- 検証した表は**1 種類だけ**。結合セル・複数シート・巨大な表は試していない
- 同時に複数の文書を開いた状態での `--target` の曖昧性
- 同梱している obasync は **2017 年公開で以降更新が無い**。現行 LibreOffice で
  動くことは確認したが、将来の互換性は保証できない

## テスト

```
pip install pytest openpyxl          # 依存はこの 2 つだけ

python -m pytest tests -q            # 22 件 (単体 16 / 統合 6)
python -m pytest tests -q -m "not integration"   # LibreOffice が無い環境
```

統合の 6 件は**テスト専用のポート 2199 と使い捨てプロファイル**で本物の
LibreOffice を起こす。利用者の既定プロファイルには触れない。★ うち 1 件は
結合セルと複数シートを扱う（既存の結合が往復で保たれること・2 枚目のシートを
狙えること・新しい結合を作れること）。

### ★ 通しで走らせて初めて出た欠陥（2026-08-06）

単独では 21 件とも通るのに、**通しで走らせると 1 件落ちた。**

    stop_office() が terminate() を投げた直後に「終了させた」と表示して戻る
    → 次のテストの ensure_office() が、まだ開いているポートを見て
      「動いている」と誤認し、起動をやめる
    → 直後に soffice.bin が落ちて、次の接続が拒否される

**この道具が扱っている「設定した ≠ 動く」そのものだった。**
待機はテスト側にあった（テストが自分でポーリングして待っていた）。
**待つ側が道具に無いのが誤り** —— `basrun stop` の直後に `apply` を叩く
利用者にも同じことが起きる。`stop_office()` がポートの閉鎖を確かめてから
戻るように直し、テストからはポーリングを外した（**外さないと、
待機が道具から消えても気づけない**）。通しで 2 回、21/21。

## 前提

- LibreOffice（同梱の Python を使う。素の Python には `uno` が無い）
- 専用プロファイル `~/.nagi/lo-profile` で起動する。
  **既定プロファイルは使わない** —— GUI で開いている LibreOffice を巻き込んで壊すため

環境変数で上書きできる: `BASRUN_OFFICE` / `BASRUN_PROFILE` / `BASRUN_PORT`

## ライセンス

basrun 本体（`basrun.py`・テスト・`.bas` ソース）は **MIT License**。全文は `LICENSE`。

`vendor/obasync/` は imacat 氏による obasync で **Apache License 2.0**。無改変で同梱し、
著作権表示とライセンス条項はスクリプト冒頭にそのまま残してある。ライセンス全文は
`vendor/obasync/LICENSE` にある。MIT と Apache-2.0 は互換で、混在に問題はない。

コードは AI アシスタント（Nagi）に実装させ、作者が動作を確かめて仕上げたもの。commit の
著者名もその立て付けのまま（`git log` は `Nagi <nagi@stg.local>`）。プロジェクトの
著作権者は Namakoo。
