"""basrun.py のロジック単体の検査。LibreOffice を必要としない。

**LibreOffice が無くても正しく振る舞うことがこのファイルの主題。** ここで
検査するのは「LibreOffice を探す・呼ぶ準備をする」までの純粋なロジック
(パス解決・引数解析・ソケット確認・べき等性) であって、実際に soffice を
起動して検証するのは test_basrun_integration.py の役目。

★ このファイルのどのテストも ``basrun.PORT`` / ``basrun.PROFILE`` の
既定値には依存しない。統合テスト側のフィクスチャは ``importlib.reload``
で基本これらのグローバルを書き換える (そして最終的に実 env へ戻す) ため、
このファイルが「今の値」をあてにすると実行順序次第で結果が変わってしまう。
必要な値は関数の引数として明示的に渡す (例: ``stop_office(port=...)``)。
"""
from __future__ import annotations

import argparse
import socket

import pytest

import basrun


# ---------------------------------------------------------------------------
# office_dir()
# ---------------------------------------------------------------------------

@pytest.mark.office_locator
def test_office_dir_prefers_basrun_office_env_even_over_real_install(
        tmp_path, monkeypatch):
    """BASRUN_OFFICE が最優先。この開発機に本物の LibreOffice があっても勝つ。"""
    fake_dir = tmp_path / "my-lo" / "program"
    fake_dir.mkdir(parents=True)
    (fake_dir / "soffice.exe").touch()
    monkeypatch.setenv("BASRUN_OFFICE", str(fake_dir))

    assert basrun.office_dir() == fake_dir


@pytest.mark.office_locator
def test_office_dir_raises_systemexit_mentioning_env_var_when_nothing_found(
        monkeypatch):
    """★見つからない時は BASRUN_OFFICE という逃げ道を必ず案内する。

    候補を全滅させるため Path.exists を強制的に False にする
    (この開発機に本物の LibreOffice があるかどうかに結果が左右されないため)。
    """
    monkeypatch.delenv("BASRUN_OFFICE", raising=False)
    monkeypatch.setattr(basrun.Path, "exists", lambda self: False)

    with pytest.raises(SystemExit) as exc:
        basrun.office_dir()
    assert "BASRUN_OFFICE" in str(exc.value)


# ---------------------------------------------------------------------------
# office_python()
# ---------------------------------------------------------------------------

@pytest.mark.office_locator
def test_office_python_finds_pythonexe(tmp_path, monkeypatch):
    """Windows 名 (python.exe) が候補の先頭。あればそれを返す。"""
    monkeypatch.setattr(basrun, "office_dir", lambda: tmp_path)
    (tmp_path / "python.exe").touch()

    assert basrun.office_python() == tmp_path / "python.exe"


@pytest.mark.office_locator
def test_office_python_finds_python3_when_no_pythonexe(tmp_path, monkeypatch):
    """python.exe が無ければ python3 (POSIX 系の名前) を試す。"""
    monkeypatch.setattr(basrun, "office_dir", lambda: tmp_path)
    (tmp_path / "python3").touch()

    assert basrun.office_python() == tmp_path / "python3"


@pytest.mark.office_locator
def test_office_python_raises_systemexit_mentioning_dir_when_missing(
        tmp_path, monkeypatch):
    """候補名が 1 つも無ければ、探したディレクトリを示して落ちる。"""
    monkeypatch.setattr(basrun, "office_dir", lambda: tmp_path)

    with pytest.raises(SystemExit) as exc:
        basrun.office_python()
    assert str(tmp_path) in str(exc.value)


# ---------------------------------------------------------------------------
# port_open()
# ---------------------------------------------------------------------------

def test_port_open_false_for_a_definitely_closed_port():
    """誰も listen していないポートは False。LibreOffice 不要。"""
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()  # bind した瞬間に閉じるので、この port はもう閉じている

    assert basrun.port_open(port) is False


def test_port_open_true_for_a_real_listening_socket():
    """自前で listen したポートは True。これも LibreOffice 不要。"""
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    s.listen(1)
    port = s.getsockname()[1]
    try:
        assert basrun.port_open(port) is True
    finally:
        s.close()


# ---------------------------------------------------------------------------
# build_parser() / サブコマンドの配線
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("argv", [
    ["sync", "somedir"],
    ["pull", "somedir"],
    ["apply", "book.xlsx", "somedir", "Module.Sub"],
    ["stop"],
])
def test_subcommands_parse_minimal_argv_without_error(argv):
    """4 つのサブコマンドとも、最小限の argv でエラー無く解析できる。"""
    parser = basrun.build_parser()
    args = parser.parse_args(argv)
    assert callable(args.func)


@pytest.mark.parametrize("cmd", ["sync", "pull"])
def test_library_omitted_on_sync_or_pull_parses_to_none(cmd):
    """library を省略すると argparse の時点では None のまま。

    「ディレクトリ名を既定にする」のは argparse ではなく _obasync_for 側の
    仕事 (下の test_library_defaults_to_source_dir_name で別に検査する)。
    """
    parser = basrun.build_parser()
    args = parser.parse_args([cmd, "somedir"])
    assert args.library is None


# ---------------------------------------------------------------------------
# library はディレクトリ名を既定にする (_obasync_for)
# ---------------------------------------------------------------------------

def test_library_defaults_to_source_dir_name(tmp_path, monkeypatch):
    """library 未指定なら、ソースディレクトリのベース名がそのまま使われる。

    book=None なので open_book/close_book/ensure_office は一切呼ばれない
    (_obasync_for の分岐を読んだ上での前提)。run_obasync だけ差し替えて、
    実際に渡された引数列を検査する純粋な単体テスト。
    """
    recorded = {}

    def fake_run_obasync(args):
        recorded["args"] = args
        return 0

    monkeypatch.setattr(basrun, "run_obasync", fake_run_obasync)

    src_dir = tmp_path / "MyBasSrc"
    src_dir.mkdir()
    ns = argparse.Namespace(
        dir=str(src_dir), library=None, book=None, doc=False,
        ext=".bas", encoding="utf-8")

    rc = basrun.sync_cmd(ns)

    assert rc == 0
    assert recorded["args"] == [
        "-x", ".bas", "-e", "utf-8", "--user", str(src_dir), "MyBasSrc"]


# ---------------------------------------------------------------------------
# apply_cmd の Module.Sub 形式チェック
# ---------------------------------------------------------------------------

def test_apply_cmd_validates_module_dot_sub_before_touching_filesystem(
        tmp_path):
    """★ドットが無ければ、book が存在しなくてもここで先に落ちる。

    basrun.py の apply_cmd はまず ``"." in a.entry`` を見て、その後で
    book.exists() を見る (順序をソースで確認済み)。この順序が壊れると
    「フォーマットミスなのに『文書が無い』という別のエラーになる」という
    紛らわしい壊れ方をする。
    """
    ns = argparse.Namespace(
        entry="NoDotAtAll",
        book=str(tmp_path / "definitely-does-not-exist.xlsx"),
        dir=str(tmp_path), library=None, ext=".bas", encoding="utf-8",
        backup=False)

    with pytest.raises(SystemExit) as exc:
        basrun.apply_cmd(ns)
    assert "Module.Sub" in str(exc.value)


# ---------------------------------------------------------------------------
# stop_office() のべき等性 (何も動いていない場合)
# ---------------------------------------------------------------------------

def test_stop_office_is_idempotent_when_nothing_listens():
    """閉じているポートに対する stop は、例外を出さず 0 を返す。

    port を明示的に渡す (basrun.PORT には触れない) ので、統合テスト側の
    reload で PORT が書き換わっていても結果は変わらない。
    """
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    closed_port = s.getsockname()[1]
    s.close()

    assert basrun.stop_office(port=closed_port) == 0


# ---------------------------------------------------------------------------
# run_obasync(): 上流 obasync の exit 0 バグへの境界ガード
#
# imacat/obasync issue #3 (https://github.com/imacat/obasync/issues/3、
# 投稿済み): 一部のエラー経路が stderr へ "ERROR:" を出しながら bare
# `return` で main() を抜けるため、プロセス全体は exit 0 で戻る。
# ここでは obasync 本体 (vendor/obasync/obasync) は一切呼ばず、
# subprocess.run をモックしてその状態 (returncode=0 かつ stderr に
# "ERROR:") を再現する。ensure_office/uno_ready もモックして
# LibreOffice 不要にする。
# ---------------------------------------------------------------------------

class _FakeCompletedProcess:
    def __init__(self, returncode, stdout="", stderr=""):
        self.returncode = returncode
        self.stdout = stdout
        self.stderr = stderr


def test_run_obasync_converts_exit0_with_error_stderr_to_nonzero(fake_office, monkeypatch, capsys):
    """★上流バグへのガード本体: ERROR: + exit 0 を非ゼロへ変換する。"""
    monkeypatch.setattr(basrun, "ensure_office", lambda *a, **kw: None)
    monkeypatch.setattr(basrun, "uno_ready", lambda *a, **kw: True)
    monkeypatch.setattr(
        basrun, "_run_bounded",
        lambda *a, **kw: _FakeCompletedProcess(
            0, stderr="ERROR: Found no source macros in somedir\n"))

    rc = basrun.run_obasync(["somedir", "MyLib"])

    assert rc != 0
    # ★ 握りつぶさない: ERROR: 本文はそのまま利用者の stderr に出る。
    assert "ERROR: Found no source macros" in capsys.readouterr().err


def test_run_obasync_passes_through_a_real_nonzero_exit_unchanged(fake_office, monkeypatch):
    """obasync 自身が非ゼロで落ちた場合は、ERROR: 検出を経由せずそのまま返す。"""
    monkeypatch.setattr(basrun, "ensure_office", lambda *a, **kw: None)
    monkeypatch.setattr(basrun, "uno_ready", lambda *a, **kw: True)
    monkeypatch.setattr(
        basrun, "_run_bounded",
        lambda *a, **kw: _FakeCompletedProcess(
            2, stderr="some other unrelated failure\n"))

    assert basrun.run_obasync(["somedir", "MyLib"]) == 2


def test_run_obasync_stays_zero_when_exit0_and_no_error_line(fake_office, monkeypatch):
    """正常系: SyntaxWarning が混じっても ERROR: が無ければ 0 のまま。"""
    monkeypatch.setattr(basrun, "ensure_office", lambda *a, **kw: None)
    monkeypatch.setattr(basrun, "uno_ready", lambda *a, **kw: True)
    monkeypatch.setattr(
        basrun, "_run_bounded",
        lambda *a, **kw: _FakeCompletedProcess(
            0, stdout="Done.  00:01 elapsed.\n",
            stderr='obasync:123: SyntaxWarning: "is" with a literal\n'))

    assert basrun.run_obasync(["somedir", "MyLib"]) == 0


# ---------------------------------------------------------------------------
# apply_cmd(): opt-in タイムアウト (BASRUN_APPLY_TIMEOUT / --timeout)
#
# 生成マクロが無限ループすると apply が永久にハングすることを TS 移行の
# 実測で確認した。ここでは実際に重いマクロを走らせず、subprocess.run が
# subprocess.TimeoutExpired を投げる状況を直接モックして再現する
# (軽量なスリープマクロ相当)。stop_office もモックし、LibreOffice 不要。
# ---------------------------------------------------------------------------

def _apply_ns(tmp_path, *, timeout):
    book = tmp_path / "book.xlsx"
    book.write_bytes(b"fake-xlsx-content")
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    return argparse.Namespace(
        book=str(book), dir=str(src_dir), library=None, entry="Mod.Sub",
        ext=".bas", encoding="utf-8", backup=False, timeout=timeout)


def test_apply_cmd_default_has_no_timeout_and_passes_none_through(fake_office, tmp_path, monkeypatch):
    """★既定は今までどおり無制限。timeout=None がそのまま subprocess.run に渡る。"""
    monkeypatch.setattr(basrun, "sync_cmd", lambda ns: 0)
    monkeypatch.setattr(basrun, "ensure_office", lambda *a, **kw: None)
    monkeypatch.setattr(basrun, "APPLY_TIMEOUT", None)

    recorded = {}

    # ★ 2026-09-04: 符号化を明示したので引数が増えた。位置引数を数え上げる形だと
    #   呼び出し側を直すたびに試験が落ちる ── **見たい引数だけ**を名前で受ける。
    def fake_run(cmd, *a, timeout=None, **kw):
        recorded["timeout"] = timeout
        return _FakeCompletedProcess(0, stdout="applied\n")

    monkeypatch.setattr(basrun, "_run_bounded", fake_run)

    rc = basrun.apply_cmd(_apply_ns(tmp_path, timeout=None))

    assert rc == 0
    assert recorded["timeout"] is None


def test_apply_cmd_falls_back_to_module_apply_timeout_when_flag_omitted(fake_office, tmp_path, monkeypatch):
    """--timeout 未指定なら、環境変数由来の basrun.APPLY_TIMEOUT を使う。"""
    monkeypatch.setattr(basrun, "sync_cmd", lambda ns: 0)
    monkeypatch.setattr(basrun, "ensure_office", lambda *a, **kw: None)
    monkeypatch.setattr(basrun, "APPLY_TIMEOUT", 7.5)

    recorded = {}

    # ★ 2026-09-04: 符号化を明示したので引数が増えた。位置引数を数え上げる形だと
    #   呼び出し側を直すたびに試験が落ちる ── **見たい引数だけ**を名前で受ける。
    def fake_run(cmd, *a, timeout=None, **kw):
        recorded["timeout"] = timeout
        return _FakeCompletedProcess(0, stdout="applied\n")

    monkeypatch.setattr(basrun, "_run_bounded", fake_run)

    basrun.apply_cmd(_apply_ns(tmp_path, timeout=None))

    assert recorded["timeout"] == 7.5


def test_apply_cmd_on_hang_stops_office_and_raises_systemexit(fake_office, tmp_path, monkeypatch):
    """★タイムアウト発火時: 接続先だけ stop_office() で終了し、非ゼロ相当で中止する。"""
    monkeypatch.setattr(basrun, "sync_cmd", lambda ns: 0)
    monkeypatch.setattr(basrun, "ensure_office", lambda *a, **kw: None)

    stopped = {"called": False}
    monkeypatch.setattr(
        basrun, "stop_office",
        lambda *a, **kw: stopped.__setitem__("called", True) or 0)

    # ★ 2026-09-04: 符号化を明示したので引数が増えた。位置引数を数え上げる形だと
    #   呼び出し側を直すたびに試験が落ちる ── **見たい引数だけ**を名前で受ける。
    def fake_run(cmd, *a, timeout=None, **kw):
        raise basrun.subprocess.TimeoutExpired(cmd=cmd, timeout=timeout)

    monkeypatch.setattr(basrun, "_run_bounded", fake_run)

    with pytest.raises(SystemExit):
        basrun.apply_cmd(_apply_ns(tmp_path, timeout=5.0))

    assert stopped["called"] is True
