"""basrun.py を実際の LibreOffice に対して走らせる統合検査。

**README の「実測で確認済み」の 4 項目がこのファイルの主題。** 単体テストは
ロジックの分岐だけを見るが、ここは「本当に .xlsx の E 列が埋まるか」
「pull の往復でソースが保たれるか」「埋め込みマクロを本当に無改変で救出
できるか」「stop が本当に接続先だけを終わらせるか」を、宣言ではなく実行で
確認する。

★ PORT/PROFILE は basrun.py の中でモジュール読み込み時に固定される
グローバルで、``ensure_office``/``stop_office`` はさらにそれを *def 時*
のデフォルト引数として 2 重に凍結している。だから
``monkeypatch.setenv("BASRUN_PORT", ...)`` を import 後にやっても
何も変わらない。ここでは env を先に書き換えてから ``importlib.reload``
してモジュール本体を再実行し、グローバルも各関数のデフォルト引数も
まとめて 2199 番ポート / tmp_path 下のプロファイルに束ね直す。テスト終了後は
env を元に戻してからもう一度 reload し、他のテストファイルが import 済みの
``basrun`` を見ても実 (2002 / ~/.nagi/lo-profile) の既定値に戻っている状態
にする — bleed-over を「無害だから許容する」のではなく、実際に消す。
"""
from __future__ import annotations

import hashlib
import importlib
import subprocess
import time
import zipfile
from pathlib import Path

import openpyxl
import pytest
from _pytest.monkeypatch import MonkeyPatch

import basrun

try:
    basrun.office_dir()
    _OFFICE_MISSING_REASON = None
except SystemExit as exc:
    _OFFICE_MISSING_REASON = str(exc)

pytestmark = [
    pytest.mark.integration,
    pytest.mark.skipif(
        _OFFICE_MISSING_REASON is not None,
        reason=f"LibreOffice が見つからない: {_OFFICE_MISSING_REASON}"),
]

TEST_PORT = 2199  # ★ 実運用の 2002 は絶対に使わない

# ★★ 2026-08-06 に実機で発見した、このタスクの絶対制約に直結する実バグ:
#
#   basrun.py の run_obasync() は obasync (vendor/obasync/obasync) を呼ぶとき
#   -p/--port を一切渡していない。obasync 自身の argparse は
#   `-p/--port` の既定値を 2002 にしている (obasync 側のソースで確認済み)。
#   つまり sync/pull (と、apply 内部の同期ステップ) は BASRUN_PORT/basrun.PORT
#   の値に関係なく、常に *obasync の既定である 2002 番* に接続する。
#   ensure_office/stop_office/open_book/close_book/apply の script 実行部は
#   basrun.PORT を正しく使っているので、この不一致は sync/pull の経路だけで起きる。
#
#   実測した実害 (2026-08-06、修正済み):
#   BASRUN_PORT=2199 で走らせたのに、sync が port 2002 に繋ぎに行った。
#   run_obasync() が obasync に -p を渡していなかったため。
#
#   ★ さらに悪いことに、そのとき 2002 は空だった。obasync は繋がらないと
#     自分で LibreOffice を起動する (vendor/obasync/obasync 786-790 の Popen)。
#     -env:UserInstallation を付けないので、起動したのは隔離プロファイルではなく
#     **LibreOffice の既定プロファイル** (%APPDATA%/LibreOffice/4/user)。
#     そこに MyLib が書き込まれ、--headless も無いので窓まで出た。
#     (プロセスのコマンドラインに -env: が無いことを実測して確定させた。
#      当初 ~/.nagi/lo-profile が汚れたと記録されていたが、それは誤り。)
#
#   後片付け: UNO の removeLibrary("MyLib") で外し、正常終了で .xlc を書き直させた。
#   Standard には触れていない。既定プロファイルは事故前の状態に戻っている。
#
#   塞ぎ方: run_obasync() が (1) -p を必ず渡し、(2) 呼ぶ直前に接続を確かめて
#   応答が無ければこちらで中止する。obasync に「繋がらない」状況を渡さない。
_OBASYNC_IGNORES_PORT_REASON = "修正済み (basrun.py run_obasync)。この定数は履歴のため残す。"


@pytest.fixture
def office(tmp_path):
    """テスト専用ポート/プロファイルで basrun を再読込し、LibreOffice を上げる。

    後始末は逆順: まず basrun.stop_office() でテスト自身の LibreOffice だけを
    終了 (この時点ではまだ env は 2199 のまま = 正しい対象を狙える)、それから
    env を元に戻し、最後にもう一度 reload して basrun のグローバルを実の
    既定値に戻す。
    """
    mp = MonkeyPatch()
    mp.setenv("BASRUN_PORT", str(TEST_PORT))
    mp.setenv("BASRUN_PROFILE", str(tmp_path / "lo-profile"))
    importlib.reload(basrun)
    assert basrun.PORT == TEST_PORT  # 束ね直しが効いていることの自己確認
    assert str(tmp_path) in str(basrun.PROFILE)

    basrun.ensure_office()
    try:
        yield basrun
    finally:
        try:
            basrun.stop_office()
        finally:
            mp.undo()
            importlib.reload(basrun)


# ---------------------------------------------------------------------------
# 1. apply: .xlsx に対して実際に計算結果と書式が入ること
# ---------------------------------------------------------------------------

def test_apply_fills_amounts_and_saves_without_embedding_a_macro(
        office, tmp_path):
    """★README の中心の主張: マクロを持てない .xlsx に、値と書式が入る。"""
    book = tmp_path / "book.xlsx"
    book.write_bytes((Path(__file__).parent / "fixture.xlsx").read_bytes())

    src_dir = tmp_path / "src"
    src_dir.mkdir()
    amount_bas = Path(__file__).parent.parent / "src" / "Amount.bas"
    (src_dir / "Amount.bas").write_bytes(amount_bas.read_bytes())

    parser = office.build_parser()
    args = parser.parse_args(
        ["apply", str(book), str(src_dir), "MyLib", "Amount.FillAmounts"])
    rc = args.func(args)
    assert rc == 0

    wb = openpyxl.load_workbook(str(book))
    ws = wb["Sheet"]
    e_values = [ws["E2"].value, ws["E3"].value, ws["E4"].value]
    e_formats = [ws["E2"].number_format, ws["E3"].number_format,
                 ws["E4"].number_format]
    g1 = ws["G1"].value

    assert e_values == [1200, 1350, 2000]
    assert e_formats == ["#,##0", "#,##0", "#,##0"]
    assert g1 == "rows=3"

    with zipfile.ZipFile(book) as z:
        assert "xl/vbaProject.bin" not in z.namelist()


# ---------------------------------------------------------------------------
# 2. pull round-trip: ユーザ側ライブラリに sync したものが pull で戻る
# ---------------------------------------------------------------------------

def test_pull_round_trip_preserves_source_content(office, tmp_path):
    """sync してから pull すると、内容は同じ (行末だけが変わりうる)。"""
    src_dir = tmp_path / "push_src"
    src_dir.mkdir()
    original = (
        "Sub RoundTrip()\n"
        "    ' pull 往復で内容が保たれることを確認するためだけのソース\n"
        "    Dim x As Integer\n"
        "    x = 1\n"
        "End Sub\n"
    )
    (src_dir / "Trip.bas").write_text(original, encoding="utf-8", newline="")

    parser = office.build_parser()

    sync_args = parser.parse_args(["sync", str(src_dir), "RoundTripLib"])
    assert sync_args.func(sync_args) == 0

    pull_dir = tmp_path / "pulled"
    pull_args = parser.parse_args(["pull", str(pull_dir), "RoundTripLib"])
    assert pull_args.func(pull_args) == 0

    pulled_file = pull_dir / "Trip.bas"
    assert pulled_file.exists()
    pulled_raw = pulled_file.read_bytes()
    pulled_text = pulled_raw.decode("utf-8")

    normalized_original = original.replace("\r\n", "\n")
    normalized_pulled = pulled_text.replace("\r\n", "\n")
    assert normalized_pulled == normalized_original

    # ★ README の「差は行末 LF/CRLF のみ」を実測で裏取りする。
    #   一致していればそれはそれで記録する (プラットフォーム依存の可能性は
    #   残るので、差があってもなくても正規化後の一致だけを必須条件にする)。
    if pulled_raw != original.encode("utf-8"):
        assert b"\r\n" in pulled_raw, (
            "正規化後は一致するのに raw が違う場合、原因は LF/CRLF 以外の"
            "はず。これは想定外なので明示的に落とす。")


# ---------------------------------------------------------------------------
# 3. --book: 文書に埋め込まれたマクロを平文へ救出する (元ファイルは無変更)
# ---------------------------------------------------------------------------

_EMBED_MACRO_SRC = r'''
import sys, uno
from com.sun.star.beans import PropertyValue

port, out_path, marker = sys.argv[1], sys.argv[2], sys.argv[3]
port = int(port)
local = uno.getComponentContext()
res = local.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local)
ctx = res.resolve(
    "uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % port)
desktop = ctx.ServiceManager.createInstanceWithContext(
    "com.sun.star.frame.Desktop", ctx)
hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
doc = desktop.loadComponentFromURL(
    "private:factory/scalc", "_blank", 0, (hidden,))
try:
    libs = doc.BasicLibraries
    if not libs.hasByName("Standard"):
        libs.createLibrary("Standard")
    lib = libs.getByName("Standard")
    src = "Sub RescueMe\n    ' %s\nEnd Sub\n" % marker
    if lib.hasByName("RescueModule"):
        lib.replaceByName("RescueModule", src)
    else:
        # ★ 正しい名前は insertByName。insertModuleByName は存在しない
        lib.insertByName("RescueModule", src)

    fmt = PropertyValue(); fmt.Name = "FilterName"; fmt.Value = "calc8"
    doc.storeToURL(uno.systemPathToFileUrl(out_path), (fmt,))
finally:
    doc.close(False)
print("embedded")
'''


def _embed_macro_in_new_ods(office, out_path: Path, marker: str) -> None:
    proc = subprocess.run(
        [str(office.office_python()), "-c", _EMBED_MACRO_SRC,
         str(office.PORT), str(out_path), marker],
        capture_output=True, text=True)
    if proc.returncode != 0 or not out_path.exists():
        pytest.fail(
            "埋め込みマクロ付き .ods の準備に失敗した。\n"
            f"stdout: {proc.stdout}\nstderr: {proc.stderr}")


def test_pull_book_rescues_embedded_macro_without_touching_original(
        office, tmp_path):
    """★埋め込みマクロを平文へ救出し、元の .ods は SHA-256 レベルで無変更。"""
    ods_path = tmp_path / "legacy.ods"
    marker = "RESCUE_MARKER_20260806"

    try:
        _embed_macro_in_new_ods(office, ods_path, marker)
    except Exception as exc:  # pragma: no cover - 環境依存の失敗を明示的に許容
        pytest.skip(
            "埋め込みマクロ付き .ods の準備 (BasicLibraries 経由) が"
            f"この環境で通らなかった: {exc}")

    before_hash = hashlib.sha256(ods_path.read_bytes()).hexdigest()

    rescued_dir = tmp_path / "rescued_out"
    parser = office.build_parser()
    pull_args = parser.parse_args(
        ["pull", str(rescued_dir), "Standard", "--book", str(ods_path)])
    rc = pull_args.func(pull_args)
    assert rc == 0

    bas_files = list(rescued_dir.glob("*.bas"))
    assert bas_files, f"救出された .bas が無い: {list(rescued_dir.iterdir())}"
    matched = [f for f in bas_files if marker in f.read_text(encoding="utf-8")]
    assert matched, (
        f"marker {marker!r} を含む .bas が無い。"
        f"見つかったファイル: {[f.name for f in bas_files]}")

    after_hash = hashlib.sha256(ods_path.read_bytes()).hexdigest()
    assert after_hash == before_hash, "元の .ods が変更されてしまっている"


# ---------------------------------------------------------------------------
# 4. stop のべき等性 (実際に動いているインスタンスに対して)
# ---------------------------------------------------------------------------

def test_stop_office_is_idempotent_against_a_real_instance(office):
    """実際に動いている LibreOffice を stop → 落ちる → もう一度 stop しても 0。

    ★実測 (2026-08-06 以前): stop_office() は d.terminate() を投げてすぐ戻り、
    soffice.bin が実際にポートを手放すまで数百ms〜数秒のずれがあった。
    戻り値=0 が「もう閉じている」ことを保証せず、このテストは自分で
    ポーリングして待っていた —— **待つ側がテストにあるのが誤りだった。**

    実害: 停止の直後に ensure_office() を呼ぶと、まだ開いているポートを見て
    「動いている」と誤認して起動をやめ、直後にプロセスが消えて次の接続が
    拒否される。このファイルのテストが**単独では通り通しでは落ちた**のがこれ。

    ★ 待機は basrun.stop_office() 側へ移した。ここで確かめる契約はこう:
    **戻り値 0 は「もう閉じている」を意味する。** ポーリングはしない ——
    ここで待つと、待機が道具から消えても気づけない。
    """
    assert office.port_open(TEST_PORT) is True

    rc1 = office.stop_office()
    assert rc1 == 0
    assert office.port_open(TEST_PORT) is False  # ★ 待たずに、その場で

    rc2 = office.stop_office()
    assert rc2 == 0


def test_a_macro_that_never_ends_leaves_nothing_behind(office, tmp_path):
    """★★ 2026-09-23: 無限ループの Basic を --timeout で止めた後、soffice も残らず、
    冊のファイルも手放されていること。

    実測した事故: 時間切れの後も soffice が冊を握ったまま生きていて、次の検体が
    unlink で「別のプロセスが使用中」になり落ちた（port が閉じた＝終わった、と読んでいた）。
    さらに terminate が届かない回は soffice が塞がったまま残り、以降が全部詰まった。
    """
    import argparse
    src = tmp_path / "Loop"
    src.mkdir()
    (src / "Gen.bas").write_text(
        "Sub Forever(oDoc As Object)\n    Dim i As Long\n    Do\n        i = i + 1\n"
        "    Loop\nEnd Sub\n", encoding="utf-8")
    book = tmp_path / "loop.xlsx"
    openpyxl.Workbook().save(book)
    owner = office._port_owner(TEST_PORT)
    with pytest.raises(SystemExit):
        office.apply_cmd(argparse.Namespace(
            book=str(book), dir=str(src), library="Loop", entry="Gen.Forever",
            ext=".bas", encoding="utf-8", backup=False, timeout=5.0))
    assert office.port_open(TEST_PORT) is False
    if owner is not None:
        assert not office._pid_alive(owner), f"soffice（pid {owner}）が残っている"
    book.rename(book.with_suffix(".moved"))   # ★ 握られていれば PermissionError


# ---------------------------------------------------------------------------
# 5. 結合セル・複数シート: apply が結合を保ち、2枚目を狙え、新しい結合を作れる
# ---------------------------------------------------------------------------

def test_apply_handles_merged_cells_and_multiple_sheets(office, tmp_path):
    """★結合セルと複数シートに対して apply が正しく動くことを実行で確かめる。

    nagi-site の「確認していない」に挙げていた 2 点を、宣言でなく実行で検証する:
      A 既存の結合 (A1:C1) が open→store の往復で保たれる
      B Sub が getByIndex(1) で 2 枚目を狙って書ける
      C Sub が結合を読み (1枚目)、新しい結合を作れる (2枚目)

    ここは値と結合メタ (openpyxl) を確認する。★ 結合の「見え方」は値では
    測れないので、描画しての目視 (xlview) を別途 1 回行い、両シートの結合が
    正しく表示されることを確認済み (2026-08-09)。
    """
    # 2 シート・結合ヘッダ付きの検体を、この場で組む (バイナリ fixture を増やさない)。
    wb = openpyxl.Workbook()
    data = wb.active
    data.title = "Data"
    data.merge_cells("A1:C1")
    data["A1"] = "月次データ 2026"
    data["A2"] = "単価"; data["B2"] = "数量"; data["C2"] = "金額"
    data["A3"] = 1280; data["B3"] = 40
    summ = wb.create_sheet("Summary")
    summ["A5"] = "（Sub が書き換える前）"
    book = tmp_path / "merged.xlsx"
    wb.save(str(book))

    # 結合を読み/作り、2 枚目を狙う Sub。
    src_dir = tmp_path / "merge_src"
    src_dir.mkdir()
    (src_dir / "MergeTest.bas").write_text(
        "Sub Run(oDoc As Object)\n"
        "    Dim oData As Object, oSum As Object\n"
        "    oData = oDoc.Sheets.getByIndex(0)\n"
        "    oSum  = oDoc.Sheets.getByIndex(1)\n"
        "    Dim sHeader As String\n"
        "    sHeader = oData.getCellByPosition(0, 0).getString()\n"
        "    oSum.getCellByPosition(0, 0).setString(\"元ヘッダ: \" & sHeader)\n"
        "    oSum.getCellByPosition(0, 1).setValue(12345)\n"
        "    Dim oRange As Object\n"
        "    oRange = oSum.getCellRangeByName(\"B3:D3\")\n"
        "    oRange.merge(True)\n"
        "    oSum.getCellByPosition(1, 2).setString(\"2枚目の結合\")\n"
        "End Sub\n",
        encoding="utf-8", newline="")

    parser = office.build_parser()
    args = parser.parse_args(
        ["apply", str(book), str(src_dir), "MergeLib", "MergeTest.Run"])
    assert args.func(args) == 0

    wb2 = openpyxl.load_workbook(str(book))
    d, s = wb2["Data"], wb2["Summary"]

    # A 往復保存: 既存の結合とシート構成が保たれる。
    assert wb2.sheetnames == ["Data", "Summary"]
    assert "A1:C1" in [str(r) for r in d.merged_cells.ranges]
    assert d["A1"].value == "月次データ 2026"

    # B シート指定: 2 枚目に正しく書けている。
    assert s["A1"].value == "元ヘッダ: 月次データ 2026"
    assert s["A2"].value == 12345

    # C 結合操作: 新しい結合ができ、1 枚目の結合を読めていた。
    assert "B3:D3" in [str(r) for r in s.merged_cells.ranges]
    assert s["B3"].value == "2枚目の結合"

    # 副作用なし: 触っていないセルは残る。
    assert s["A5"].value == "（Sub が書き換える前）"


# ---------------------------------------------------------------------------
# bleed-over の確認: 全ての office フィクスチャ利用テストの後始末が終わった後、
# import 済みの basrun が実 (2002 / ~/.nagi/lo-profile) の既定値に戻っている
# ---------------------------------------------------------------------------

def test_no_bleed_into_module_globals_after_all_office_fixtures_torn_down():
    """★office フィクスチャの後始末が、本当に実の既定値まで戻すことの確認。

    このテストは office フィクスチャを使わない。このファイル内で先に走った
    office フィクスチャ利用テストの後始末 (mp.undo + reload) が正しければ、
    ここで見る basrun.PORT / basrun.PROFILE は実の既定値のはず。
    """
    assert basrun.PORT == 2002
    assert basrun.PROFILE == Path.home() / ".nagi" / "lo-profile"
