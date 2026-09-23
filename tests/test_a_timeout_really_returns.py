"""時間切れは**本当に戻ってくる**こと（孫が出力の管を握っていても）。

★★ 2026-09-23（ailine の全件が 74 分止まった件の真因）:
  Windows の LibreOffice 同梱 python.exe は**ランチャー**で、コードは孫の python-core で動く。
  subprocess.run(timeout=...) が時間切れで kill するのはランチャーだけで、孫は生き残り、
  出力の管を握ったまま run() を待たせ続ける。
  ★ 09-04 の試験は「timeout を**渡したか**」を見ていたので緑のまま効いていなかった。
    ここは**時計で**測る。

★ 偽ランチャーは本物と同じ形 ── 孫を起こし（管は継承）、自分は孫を待つだけ。
  実 LibreOffice は要らないので、CI（ubuntu / windows）でも走る。
"""
from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import basrun  # noqa: E402

GRANDCHILD_SLEEP = 8


def _launcher_cmd(sleep: int = GRANDCHILD_SLEEP) -> list:
    child = f"import os, time; print(os.getpid(), flush=True); time.sleep({sleep})"
    launcher = ("import subprocess, sys; "
                f"sys.exit(subprocess.Popen([sys.executable, '-c', {child!r}]).wait())")
    return [sys.executable, "-c", launcher]


def _alive(pid: int) -> bool:
    if os.name == "nt":
        r = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"],
                           capture_output=True, text=True, encoding="utf-8", errors="replace")
        return str(pid) in r.stdout
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    return True


def test_the_specimen_reproduces_the_hang():
    """★ 陽性対照: 素の subprocess.run では**頼んだ時間で戻らない**こと。
       ここが赤なら偽ランチャーが事故を再現しておらず、下の緑は何も言っていない。"""
    if os.name != "nt":
        # POSIX の run() は kill 後に wait() するだけで管を待たないので、この形では詰まらない
        # （★ 詰まるのは Windows の communicate()）。対照は Windows でだけ意味を持つ。
        return
    t = time.monotonic()
    try:
        subprocess.run(_launcher_cmd(), capture_output=True, text=True, timeout=1)
    except subprocess.TimeoutExpired:
        pass
    assert time.monotonic() - t >= GRANDCHILD_SLEEP - 2, \
        "素の run が時間どおり戻った ── 偽ランチャーが事故を再現していない（検体を疑う）"


def test_a_timeout_returns_on_time_and_takes_the_grandchild_with_it():
    """頼んだ時間（+猶予）で戻り、孫も死んでいること。"""
    t = time.monotonic()
    try:
        basrun._run_bounded(_launcher_cmd(sleep=60), timeout=2)
    except subprocess.TimeoutExpired as e:
        elapsed = time.monotonic() - t
        out = e.output or ""
    else:
        raise AssertionError("60 秒眠る孫を 2 秒で縛ったのに、時間切れにならなかった")
    assert elapsed < 2 + basrun.KILL_GRACE, f"時間どおりに戻らなかった: {elapsed:.1f} 秒"
    pid = int(out.split()[0])
    time.sleep(0.5)
    assert not _alive(pid), f"孫（pid {pid}）が生き残っている ── soffice を握ったまま残る形"


def test_a_normal_run_still_returns_its_output():
    """時間内に終わる回は、出力と終了コードをそのまま返す（縛りが普段の形を変えない）。"""
    p = basrun._run_bounded([sys.executable, "-c", "print('日本語も'); raise SystemExit(3)"],
                            timeout=30)
    assert p.returncode == 3 and "日本語も" in p.stdout


def test_every_office_call_goes_through_the_bounded_door():
    """★ 同梱 python を呼ぶのは run_office だけ（4 箇所に散っていた形を戻さない）。
       分母は手で並べず、ソースから数える。"""
    src = Path(basrun.__file__).read_text(encoding="utf-8")
    uses = [ln.strip() for ln in src.splitlines()
            if "office_python()" in ln and not ln.lstrip().startswith("def office_python")]
    assert len(uses) == 1 and "_run_bounded(" in uses[0], \
        f"run_office 以外で同梱 python を直に呼んでいる: {uses}"
