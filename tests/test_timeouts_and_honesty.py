"""ハングからの復旧が、それ自身ハングしないこと／落ちていないのに落としたと言わないこと。

★★ 2026-09-04 のコードレビューで見つけた実バグ 2 件と、その周辺 3 件。

  ① apply の --timeout は「生成マクロが無限ループしても戻る」ための保険なのに、
     復旧経路の stop_office が **時間を縛らずに** UNO を呼んでいた。
     LibreOffice が Basic の実行で本線を塞いでいると terminate() は返らないので、
     **保険そのものがハングする**形だった。
  ② その stop_office の戻り値を捨てて「接続先の LibreOffice を終了させた」と
     断定していた。stop_office は落とせなければ 1 を返す。
     ★ この関数の docstring 自身が「落ちたことを確かめずに『終了させた』と
       表示していた」と書いている ── 同じ教訓が同じファイルの 2 箇所目に
       配線されていなかった（片配線）。
  ③ close_book が失敗を握りつぶし、保存できなくても同期成功と報告していた。
  ④ 環境変数が数値でないと import 時に traceback で死ぬ（--help も出ない）。
  ⑤ BASRUN_APPLY_TIMEOUT=0 が 0.0 になり、apply が即座に打ち切られていた。

★ ここは実 LibreOffice を使わない（偽の踏み台と差し替えで測る）。
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import basrun  # noqa: E402


# --- ① 復旧経路そのものがハングしないこと --------------------------------

def test_stop_office_bounds_the_terminate_call(fake_office, monkeypatch):
    """★ terminate の呼び出しに**必ず**時間の縛りが渡ること（無制限にしない）。"""
    seen = {}

    def _fake(code, *args, timeout=None):
        seen["timeout"] = timeout
        return subprocess.CompletedProcess([], 0, "", "")

    monkeypatch.setattr(basrun, "_office_py", _fake)
    monkeypatch.setattr(basrun, "port_open", lambda p: False)
    basrun.stop_office(port=1)          # 動いていない場合は即 0
    assert "timeout" not in seen

    calls = iter([True, False])         # 1 回目は開いている、2 回目で閉じた
    monkeypatch.setattr(basrun, "port_open", lambda p: next(calls, False))
    monkeypatch.setattr(basrun, "_office_py", _fake)
    assert basrun.stop_office(port=1, timeout=20.0) == 0
    assert seen["timeout"] is not None, "★ terminate に時間の縛りが渡っていない"
    assert seen["timeout"] <= basrun.TERMINATE_BUDGET


def test_stop_office_survives_a_terminate_that_never_returns(fake_office, monkeypatch):
    """★ 本命 ── terminate が返らなくても、ポートの様子を見て判断まで進むこと。"""
    def _hang(code, *args, timeout=None):
        raise subprocess.TimeoutExpired(cmd="terminate", timeout=timeout or 0)

    monkeypatch.setattr(basrun, "_office_py", _hang)
    monkeypatch.setattr(basrun, "port_open", lambda p: True)   # 落ちない
    rc = basrun.stop_office(port=1, timeout=0.5)
    assert rc == 1, "落ちていないなら 1 を返すこと"


def test_stop_office_reports_success_only_when_the_port_actually_closes(
        fake_office, monkeypatch):
    seq = iter([True, True, False])
    monkeypatch.setattr(basrun, "port_open", lambda p: next(seq, False))
    monkeypatch.setattr(basrun, "_office_py",
                        lambda *a, **k: subprocess.CompletedProcess([], 0, "", ""))
    assert basrun.stop_office(port=1, timeout=5.0) == 0


# --- ② 落としたと嘘をつかないこと ------------------------------------------

def _apply_ns(tmp_path, timeout):
    book = tmp_path / "b.xlsx"; book.write_bytes(b"x")
    d = tmp_path / "src"; d.mkdir()
    import argparse
    return argparse.Namespace(book=str(book), dir=str(d), library="L",
                              entry="M.S", ext=".bas", encoding="utf-8",
                              backup=False, timeout=timeout)


@pytest.mark.parametrize("stop_rc, must_say, must_not_say", [
    (0, "終了させて中止", "終了にも失敗"),
    (1, "終了にも失敗", "終了させて中止"),
])
def test_apply_tells_the_truth_about_whether_office_was_stopped(
        fake_office, tmp_path, monkeypatch, stop_rc, must_say, must_not_say):
    """★ stop_office が落とせなかった回に「終了させた」と言わないこと。"""
    monkeypatch.setattr(basrun, "sync_cmd", lambda a: 0)
    monkeypatch.setattr(basrun, "ensure_office", lambda *a, **k: None)
    monkeypatch.setattr(basrun, "stop_office", lambda *a, **k: stop_rc)

    def _boom(*a, **k):
        raise subprocess.TimeoutExpired(cmd="apply", timeout=1)

    monkeypatch.setattr(basrun.subprocess, "run", _boom)
    with pytest.raises(SystemExit) as e:
        basrun.apply_cmd(_apply_ns(tmp_path, 1.0))
    msg = str(e.value)
    assert must_say in msg, msg
    assert must_not_say not in msg, msg


# --- ③ 保存の失敗を握りつぶさないこと --------------------------------------

def test_close_book_returns_the_failure(fake_office, monkeypatch, tmp_path):
    monkeypatch.setattr(basrun, "_office_py",
                        lambda *a, **k: subprocess.CompletedProcess([], 1, "", "CLOSE: 読み取り専用"))
    assert basrun.close_book(tmp_path / "b.xlsx", save=True) == 1


def test_close_book_reports_when_nothing_matched(fake_office, monkeypatch, tmp_path):
    """★ 「探す場所が空なら必ず通る」形を残さない ── 0 件も黙らせない。"""
    monkeypatch.setattr(basrun, "_office_py",
                        lambda *a, **k: subprocess.CompletedProcess([], 2, "", "CLOSE: 対象の文書が開いていない"))
    assert basrun.close_book(tmp_path / "b.xlsx", save=True) == 2


def test_sync_does_not_report_success_when_saving_failed(fake_office, tmp_path, monkeypatch):
    """★ 本命 ── 同期が通っても、閉じる/保存に失敗したら成功と言わない。"""
    import argparse
    book = tmp_path / "b.xlsx"; book.write_bytes(b"x")
    d = tmp_path / "src"; d.mkdir()
    monkeypatch.setattr(basrun, "open_book", lambda b: None)
    monkeypatch.setattr(basrun, "run_obasync", lambda args: 0)      # 同期は成功
    monkeypatch.setattr(basrun, "close_book", lambda b, save: 1)    # 保存が失敗
    ns = argparse.Namespace(dir=str(d), library="L", ext=".bas",
                            encoding="utf-8", doc=False, book=str(book))
    assert basrun.sync_cmd(ns) != 0, "★ 保存に失敗したのに 0 を返した"


def test_sync_still_returns_zero_when_everything_worked(fake_office, tmp_path, monkeypatch):
    import argparse
    book = tmp_path / "b.xlsx"; book.write_bytes(b"x")
    d = tmp_path / "src"; d.mkdir()
    monkeypatch.setattr(basrun, "open_book", lambda b: None)
    monkeypatch.setattr(basrun, "run_obasync", lambda args: 0)
    monkeypatch.setattr(basrun, "close_book", lambda b, save: 0)
    ns = argparse.Namespace(dir=str(d), library="L", ext=".bas",
                            encoding="utf-8", doc=False, book=str(book))
    assert basrun.sync_cmd(ns) == 0


# --- ④⑤ 環境変数 ------------------------------------------------------------

def test_a_broken_env_var_says_why_instead_of_a_traceback(monkeypatch):
    monkeypatch.setenv("BASRUN_TEST_SECONDS", "abc")
    with pytest.raises(SystemExit) as e:
        basrun._env_seconds("BASRUN_TEST_SECONDS")
    assert "秒数" in str(e.value) and "abc" in str(e.value)


@pytest.mark.parametrize("value, want", [
    ("", None), ("  ", None), ("0", None), ("0.0", None),
    ("30", 30.0), (" 45 ", 45.0),
])
def test_zero_and_blank_both_mean_unlimited(monkeypatch, value, want):
    """★ 0 を 0.0 にしない ── subprocess.run(timeout=0) は即座に打ち切る。"""
    monkeypatch.setenv("BASRUN_TEST_SECONDS", value)
    assert basrun._env_seconds("BASRUN_TEST_SECONDS") == want


def test_a_negative_value_is_refused(monkeypatch):
    monkeypatch.setenv("BASRUN_TEST_SECONDS", "-1")
    with pytest.raises(SystemExit):
        basrun._env_seconds("BASRUN_TEST_SECONDS")


# --- 構造の番人 --------------------------------------------------------------

def test_every_uno_call_goes_through_the_one_place_that_can_bound_it():
    """★ UNO を呼ぶ経路が _office_py に集まっていること（呼び出し側に散らさない）。

    以前は stop_office だけが subprocess.run を直に叩いており、そこにだけ
    時間の縛りが無かった ── 同じ判断が 2 箇所にあって片方が古い形。
    """
    src = Path(basrun.__file__).read_text(encoding="utf-8")
    body = src[src.index("def stop_office"):src.index("\ndef run_obasync")]
    assert "subprocess.run(" not in body, \
        "stop_office が subprocess.run を直に叩いている（_office_py を通すこと）"
    assert "_office_py(" in body


# --- 埋め込みスクリプトの番人 -------------------------------------------------

@pytest.mark.parametrize("name", ["OPEN_SRC", "CLOSE_SRC", "APPLY_SRC", "UNO_READY_SRC"])
def test_every_embedded_script_is_valid_python(name):
    """★ 文字列に埋め込んだスクリプトが、Python として構文が通ること。

    ★★ 2026-09-04 に踏んだ: CLOSE_SRC の中に改行文字を字面で書いてしまい、
      文字列が閉じずに**子プロセスが構文エラーで死んだ**。
      呼び出し側はそれを「保存に失敗した」と読み、実機テストが赤くなって初めて
      気づいた ── **埋め込みは、静かに壊れて別の失敗のふりをする。**
    ★ 構文なら書く前に機械で確かめられる。ここで縛る。
    """
    import ast
    ast.parse(getattr(basrun, name))


def test_the_embedded_scripts_only_emit_ascii():
    """★ 埋め込みスクリプトが**出力する文字列**は ASCII だけ（人向けの文は Python 側で作る）。

    子プロセスの出力の符号化は環境で変わる（Windows では cp932 に落ちる）。
    非 ASCII を子から出すと、それだけで道具が落ちうる ── 実際に踏んだ。
    ★ コメントは対象外 ── 出力されないものまで縛ると、番人が説明を書かせなくなる
      （最初にそう書いて OPEN_SRC の注記で赤くなった）。**出る物だけを見る。**
    """
    import ast as _ast
    for name in ("OPEN_SRC", "CLOSE_SRC", "APPLY_SRC", "UNO_READY_SRC"):
        tree = _ast.parse(getattr(basrun, name))
        for node in _ast.walk(tree):
            if not isinstance(node, _ast.Call):
                continue
            f = node.func
            emits = (isinstance(f, _ast.Name) and f.id == "print") or (
                isinstance(f, _ast.Attribute) and f.attr == "write")
            if not emits:
                continue
            for lit in _ast.walk(node):
                if isinstance(lit, _ast.Constant) and isinstance(lit.value, str):
                    assert lit.value.isascii(), f"{name} が非 ASCII を出力する: {lit.value!r}"


def test_every_subprocess_that_reads_output_declares_its_encoding():
    """★ 出力を読む subprocess.run は、必ず符号化を明示していること。

    text=True だけだと Windows では cp932 で復号され、子が UTF-8 を書いた瞬間に
    UnicodeDecodeError で**道具ごと落ちる**（ファイル名や UNO の例外文が非 ASCII なら
    いつでも起きる）。
    """
    src = Path(basrun.__file__).read_text(encoding="utf-8")
    for i, chunk in enumerate(src.split("capture_output=True")[1:], 1):
        head = chunk[:200]
        assert "encoding=" in head, f"{i} 番目の capture_output に encoding が無い"
