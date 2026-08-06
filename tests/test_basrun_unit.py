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

def test_office_dir_prefers_basrun_office_env_even_over_real_install(
        tmp_path, monkeypatch):
    """BASRUN_OFFICE が最優先。この開発機に本物の LibreOffice があっても勝つ。"""
    fake_dir = tmp_path / "my-lo" / "program"
    fake_dir.mkdir(parents=True)
    (fake_dir / "soffice.exe").touch()
    monkeypatch.setenv("BASRUN_OFFICE", str(fake_dir))

    assert basrun.office_dir() == fake_dir


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

def test_office_python_finds_pythonexe(tmp_path, monkeypatch):
    """Windows 名 (python.exe) が候補の先頭。あればそれを返す。"""
    monkeypatch.setattr(basrun, "office_dir", lambda: tmp_path)
    (tmp_path / "python.exe").touch()

    assert basrun.office_python() == tmp_path / "python.exe"


def test_office_python_finds_python3_when_no_pythonexe(tmp_path, monkeypatch):
    """python.exe が無ければ python3 (POSIX 系の名前) を試す。"""
    monkeypatch.setattr(basrun, "office_dir", lambda: tmp_path)
    (tmp_path / "python3").touch()

    assert basrun.office_python() == tmp_path / "python3"


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
