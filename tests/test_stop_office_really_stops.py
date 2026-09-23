"""stop_office は「終わった」を**プロセスが消えたこと**で言い、届かない時は自分の分だけ落とす。

★★ 2026-09-23（ailine の全件が止まった件の広がり方）:
  ① 生成マクロが soffice の中で無限ループすると terminate が届かず、誰も soffice を
    落とさないので、以降の実行が全部詰まった。
  ② port が閉じても soffice はしばらく生きていて冊を握っており、次の検体が
    「別のプロセスが使用中」で落ちた（port が閉じた＝終わった、と読んでいた）。
★ ここは実 LibreOffice を使わない（持ち主・生死・kill を差し替えて測る）。
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import basrun  # noqa: E402


def _stuck(monkeypatch, *, ours: bool, dies_on_kill: bool = True):
    """terminate が届かず port が開いたままの LibreOffice（pid 4242）。"""
    state = {"killed": [], "alive": True}
    monkeypatch.setattr(basrun, "_port_owner", lambda port: 4242)
    monkeypatch.setattr(basrun, "_started_by_us", lambda pid: ours)

    def _boom(*a, **k):
        raise subprocess.TimeoutExpired(cmd="terminate", timeout=1)
    monkeypatch.setattr(basrun, "_office_py", _boom)

    def _kill(pid, group=False):
        state["killed"].append(pid)
        if dies_on_kill:
            state["alive"] = False
    monkeypatch.setattr(basrun, "_kill_pid_tree", _kill)
    monkeypatch.setattr(basrun, "_pid_alive", lambda pid: state["alive"])
    monkeypatch.setattr(basrun, "port_open", lambda port: state["alive"])
    return state


def test_a_stuck_office_that_we_started_is_put_down_by_pid(fake_office, monkeypatch, capsys):
    st = _stuck(monkeypatch, ours=True)
    assert basrun.stop_office(port=1, timeout=0.5) == 0
    assert st["killed"] == [4242]
    assert "PID 指定で落とした" in capsys.readouterr().out


def test_an_office_we_did_not_start_is_never_killed(fake_office, monkeypatch, capsys):
    """★ 利用者が GUI で開いている LibreOffice（専用プロファイルでない）は巻き込まない。"""
    st = _stuck(monkeypatch, ours=False)
    assert basrun.stop_office(port=1, timeout=0.5) == 1
    assert st["killed"] == []
    assert "開いたまま" in capsys.readouterr().out


def test_a_closed_port_is_not_called_stopped_while_the_process_lives(fake_office, monkeypatch,
                                                                     capsys):
    """★ port は閉じたが soffice が生きている（冊を握っている）間は「終了させた」と言わない。"""
    monkeypatch.setattr(basrun, "_port_owner", lambda port: 4242)
    monkeypatch.setattr(basrun, "_started_by_us", lambda pid: False)
    monkeypatch.setattr(basrun, "_office_py",
                        lambda *a, **k: subprocess.CompletedProcess([], 0, "", ""))
    opened = iter([True, False])
    monkeypatch.setattr(basrun, "port_open", lambda port: next(opened, False))
    monkeypatch.setattr(basrun, "_pid_alive", lambda pid: True)
    assert basrun.stop_office(port=1, timeout=0.5) == 1
    out = capsys.readouterr().out
    assert "終了させた" not in out and "まだ残っている" in out, out


def test_it_waits_for_the_process_before_saying_stopped(fake_office, monkeypatch, capsys):
    """port が閉じた後、プロセスが消えるまで待ってから「終了させた」と言う。"""
    monkeypatch.setattr(basrun, "_port_owner", lambda port: 4242)
    monkeypatch.setattr(basrun, "_office_py",
                        lambda *a, **k: subprocess.CompletedProcess([], 0, "", ""))
    opened = iter([True])                # 入口では開いている → その後は閉じている
    monkeypatch.setattr(basrun, "port_open", lambda port: next(opened, False))
    polls: list = []
    life = iter([True, True, True, False])

    def _alive(pid):
        polls.append(pid)
        return next(life, False)
    monkeypatch.setattr(basrun, "_pid_alive", _alive)
    assert basrun.stop_office(port=1, timeout=5.0) == 0
    assert len(polls) >= 4, f"プロセスの生死を見ずに終了と言った（見た回数 {len(polls)}）"
    assert "終了させた" in capsys.readouterr().out
