#!/usr/bin/env python
"""平文の Basic ソースを、マクロを格納できない文書に対して走らせる。

    basrun sync  <dir> [lib]                       ソース群 -> ライブラリ
    basrun pull  <dir> [lib]                       ライブラリ -> ソース群
    basrun apply <book> <dir> <lib> <Module.Sub>   同期して、文書に適用して保存
    basrun stop                                    起動した LibreOffice を落とす

## 何を解いているか

`.xlsx` は仕様としてマクロを格納できない (マクロは `.xlsm` 側)。だから
「Excel の VBA を書いてくれ」は、`.xlsx` を渡された時点で素直には成立しない。

**★ だが制約は「格納」の側にあって「実行」の側には無い。** ソースを実行時に
ライブラリへ流し込めば、文書は普通の `.xlsx` のまま処理できる。

回り道した先のほうが、本来の形として良い:

    ソースが平文        git で diff が取れる / レビューできる
    文書に埋め込まない  文書内の XML を壊す事故が起きない
    真実の在処が 1 つ   ソースが原本。ライブラリは実行時のキャッシュにすぎない
    形式が自由          .xlsx でも .ods でも同じソースが走る

## 2 つの道具を合わせている

同期は **obasync** (imacat 作、Apache-2.0、`vendor/obasync/` に無改変で同梱) に
任せる。双方向・差分・余剰モジュールの削除まで実装されていて、こちらで書き直す
理由が無い。

実行はこちらが持つ。obasync にも `--run` はあるが `script.invoke((), (), ())` で
**引数を渡していない**ため、`ThisComponent` に依存する。

    ★ --headless + Hidden で開いた文書は「現在のコンポーネント」ではない。
      2026-08-06 に実測: 「読み込んだ/実行した/保存した」と全部表示されて、
      セルは 1 つも変わっていなかった。成功を報告して何もしていない状態。

だから `apply` は**文書を引数で明示的に渡す**。対象の Sub は次の形で受けること:

    Sub FillAmounts(oDoc As Object)

## LibreOffice の起動管理

obasync も自作の実行部も「LibreOffice が既に動いていること」を前提にしていて、
落ちていると分かりにくい失敗をする (実測: 接続不能を `illegal object given!` と
いう別の場所のエラーとして観測した)。**ここで面倒を見る。**

★ 専用プロファイルで起動する。既定プロファイルを使うと、GUI で開いている
LibreOffice を巻き込んで壊す (2026-08-04 に実際に壊した)。
"""
from __future__ import annotations

import argparse
import os
import shutil
import socket
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
OBASYNC = HERE / "vendor" / "obasync" / "obasync"

PORT = int(os.environ.get("BASRUN_PORT", "2002"))
# ★ 既定プロファイルを使わない。GUI 側を巻き込まないため。
PROFILE = Path(os.environ.get(
    "BASRUN_PROFILE", str(Path.home() / ".nagi" / "lo-profile")))


def _env_seconds(name: str) -> float | None:
    """秒指定の環境変数を読む。未設定/空/0 なら None (=無制限)。

    ★ 2026-09-04 のレビューで見つけた 2 件:
      ① 数値でない値を入れると **import 時に traceback で死ぬ** ──
         `basrun stop` も `--help` も出せなくなる。この repo の他の失敗は
         すべて理由つきの SystemExit なので、ここだけ作法が違っていた。
      ② `0` が `0.0` になっていた。`subprocess.run(timeout=0.0)` は即座に
         打ち切るので、「0＝無制限」と読んだ利用者の apply は必ず失敗する。
         **0 は無制限**として扱う（無制限を意図した書き方を裏切らない）。
    """
    v = (os.environ.get(name) or "").strip()
    if not v:
        return None
    try:
        seconds = float(v)
    except ValueError:
        raise SystemExit(
            f"環境変数 {name} は秒数で指定すること（今の値: {v!r}）。"
            "無制限にしたいなら空にするか 0 を入れる。")
    if seconds < 0:
        raise SystemExit(f"環境変数 {name} に負の秒数は指定できない（今の値: {v!r}）")
    return seconds or None


# ★ opt-in。既定は今までどおり無制限。生成マクロが無限ループすると apply が
# 永久にハングすることを TS 移行の実測で確認した (2026-08-14) が、無条件の
# タイムアウトは「重いが正常に終わる」処理まで巻き込むので既定にはしない。
APPLY_TIMEOUT = _env_seconds("BASRUN_APPLY_TIMEOUT")

# ★ 2026-09-04: apply 以外の UNO 呼び出し（開く・閉じる・終了要求）も時間を縛れる
# ようにした。既定は無制限 ── 重いが正常に終わる読み込みを巻き込まないため、
# APPLY_TIMEOUT と同じ判断に揃える。
UNO_TIMEOUT = _env_seconds("BASRUN_UNO_TIMEOUT")
#: ★ 同期（obasync）の既定の縛り。UNO_TIMEOUT が無い（無制限）時でも同期だけは必ず縛る（2026-09-23）。
SYNC_TIMEOUT = 120.0

# ★ 終了要求だけは**必ず縛る**。ここは「ハングからの復旧」経路で、
#   縛らないと保険そのものがハングする（下の stop_office のコメント参照）。
TERMINATE_BUDGET = 10.0


# ---------------------------------------------------------------------------
# LibreOffice の在処と起動
# ---------------------------------------------------------------------------

def office_dir() -> Path:
    """LibreOffice の program ディレクトリ。見つからなければ理由つきで落とす。"""
    env = os.environ.get("BASRUN_OFFICE")
    cands = [Path(env)] if env else []
    cands += [
        Path(r"C:\Program Files\LibreOffice\program"),
        Path(r"C:\Program Files (x86)\LibreOffice\program"),
        Path("/usr/lib/libreoffice/program"),
        Path("/Applications/LibreOffice.app/Contents/MacOS"),
    ]
    for c in cands:
        if (c / "soffice.exe").exists() or (c / "soffice").exists():
            return c
    raise SystemExit(
        "LibreOffice が見つからない。BASRUN_OFFICE に program ディレクトリを指定するか、"
        "標準の場所にインストールすること。")


def office_python() -> Path:
    """同梱 python。★ 素の python には uno が無いので、これでなければ動かない。"""
    d = office_dir()
    for name in ("python.exe", "python3", "python"):
        p = d / name
        if p.exists():
            return p
    raise SystemExit(
        f"LibreOffice 同梱の python が {d} に無い。Linux では別パッケージ "
        "(python3-uno 等) になっていることがある。")


# ★ 木ごと kill した後、管が閉じるのを待つ上限（それでも孫が残る環境でも戻るため）。
KILL_GRACE = 10.0


def _kill_tree(proc: subprocess.Popen) -> None:
    """proc を根に、子孫ごと落とす。★ proc が**まだ生きているうちに**呼ぶこと
    （根が先に死ぬと、孫は親を失って木から外れ、Windows の /T では辿れない）。
    ★ 名前一括（taskkill /IM）はしない ── 無関係なプロセスを巻き込む。"""
    _kill_pid_tree(proc.pid, group=True)


def _kill_pid_tree(pid: int, group: bool = False) -> None:
    """PID を根に子孫ごと落とす。group=True は POSIX で新しいセッションを作った相手。"""
    if os.name == "nt":
        subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=KILL_GRACE)
    else:
        import signal
        try:
            if group:
                os.killpg(pid, signal.SIGKILL)
            else:
                os.kill(pid, signal.SIGKILL)
        except (ProcessLookupError, PermissionError):
            pass


def _port_owner(port: int) -> int | None:
    """port を LISTEN しているプロセスの PID（Windows のみ。取れなければ None）。"""
    if os.name != "nt":
        return None
    try:
        r = subprocess.run(["netstat", "-ano", "-p", "TCP"], capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=KILL_GRACE)
    except (OSError, subprocess.TimeoutExpired):
        return None
    for line in r.stdout.splitlines():
        parts = line.split()
        if (len(parts) == 5 and parts[0] == "TCP" and parts[1].endswith(f":{port}")
                and parts[3] == "LISTENING" and parts[4].isdigit()):
            return int(parts[4])
    return None


def _started_by_us(pid: int) -> bool:
    """★ その PID が basrun の専用プロファイルで起こした LibreOffice か。
    違えば落とさない ── 利用者が GUI で開いている LibreOffice を巻き込まない。"""
    if os.name != "nt":
        return False
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             f"(Get-CimInstance Win32_Process -Filter 'ProcessId={pid}').CommandLine"],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=KILL_GRACE * 3)
    except (OSError, subprocess.TimeoutExpired):
        return False
    return f"-env:UserInstallation={PROFILE.resolve().as_uri()}" in (r.stdout or "")


def _pid_alive(pid: int) -> bool:
    if os.name == "nt":
        r = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"], capture_output=True,
                           text=True, encoding="utf-8", errors="replace", timeout=KILL_GRACE)
        return str(pid) in (r.stdout or "")
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def _wait_gone(pid: int, deadline: float) -> bool:
    while time.monotonic() < deadline:
        if not _pid_alive(pid):
            return True
        time.sleep(0.2)
    return not _pid_alive(pid)


def _run_bounded(cmd: list, timeout: float | None = None) -> subprocess.CompletedProcess:
    """cmd を走らせ、timeout を過ぎたら**本当に戻る**（子孫ごと落としてから）。

    ★★ 2026-09-23（実測）: Windows の LibreOffice 同梱 python.exe は**ランチャー**で、
      コードは孫の python-core で動く（Popen の pid 17992 に対し、中の pid 4644 の親が 17992）。
      subprocess.run(timeout=...) が時間切れで kill するのは**ランチャーだけ**で、
      孫は生き残り、出力の管を握ったまま run() を待たせ続ける。
      実測: timeout=3 で 25 秒眠るスクリプトを呼ぶと、例外まで **25.0 秒**かかった。
      soffice が塞がっていれば、これは永遠になる（ailine の全件が 74 分止まった）。
    ★★ 2026-09-04 に付けた timeout は**すべて**この理由で効いていなかった。
      試験は「timeout を渡したか」を見ていて、「時間どおりに戻るか」を見ていなかった。
    """
    kw = {} if os.name == "nt" else {"start_new_session": True}
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            text=True, encoding="utf-8", errors="replace", **kw)
    try:
        out, err = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        _kill_tree(proc)
        try:
            out, err = proc.communicate(timeout=KILL_GRACE)
        except subprocess.TimeoutExpired:
            out, err = "", ""
        raise subprocess.TimeoutExpired(cmd, timeout, output=out, stderr=err) from None
    return subprocess.CompletedProcess(cmd, proc.returncode, out, err)


def run_office(args: list, timeout: float | None = None) -> subprocess.CompletedProcess:
    """LibreOffice 同梱 python を走らせる**唯一の口**。★ 呼び出し側で subprocess を叩かない
    （4 箇所が同じ形で叩いていて、4 箇所とも時間切れが効いていなかった）。"""
    return _run_bounded([str(office_python()), *args], timeout=timeout)


def port_open(port: int) -> bool:
    with socket.socket() as s:
        s.settimeout(0.5)
        return s.connect_ex(("127.0.0.1", port)) == 0


# ★ TCP が開いたことと、UNO で resolve できることは別。
#   ポートが開いた直後の数秒は resolve が NoConnectException で失敗する。
#   そこで obasync を呼ぶと、obasync は「繋がらない」と判断して
#   **既定プロファイルで LibreOffice を起こす**。2026-08-06 に実際に起きた。
UNO_READY_SRC = r'''
import sys, time, uno
from com.sun.star.connection import NoConnectException
port, timeout = int(sys.argv[1]), float(sys.argv[2])
local = uno.getComponentContext()
res = local.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local)
url = "uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % port
deadline = time.time() + timeout
while time.time() < deadline:
    try:
        res.resolve(url)
        sys.exit(0)
    except NoConnectException:
        time.sleep(0.5)
    except Exception:
        time.sleep(0.5)
sys.exit(1)
'''


def uno_ready(port: int = PORT, timeout: float = 5.0) -> bool:
    """★ 実際に resolve できるかを確かめる。TCP が開いているだけでは足りない。"""
    # ★ 中のスクリプトは自分で timeout 秒まで試す。外側はそれに猶予を足した上限で縛る。
    try:
        p = run_office(["-c", UNO_READY_SRC, str(port), str(timeout)],
                       timeout=timeout + KILL_GRACE)
    except subprocess.TimeoutExpired:
        return False
    return p.returncode == 0


def ensure_office(port: int = PORT, timeout: float = 90.0) -> None:
    """UNO で resolve できるまで待つ。動いていなければ起こす。

    ★ 判定を `port_open` にすると足りない。上の UNO_READY_SRC のコメント参照。
    """
    if port_open(port) and uno_ready(port, timeout=3.0):
        return
    if not port_open(port):
        PROFILE.mkdir(parents=True, exist_ok=True)
        soffice = office_dir() / ("soffice.exe" if os.name == "nt" else "soffice")
        url = PROFILE.resolve().as_uri()
        subprocess.Popen(
            [str(soffice), "--headless", "--norestore", "--nologo",
             f"--accept=socket,host=127.0.0.1,port={port};urp;",
             f"-env:UserInstallation={url}"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if uno_ready(port, timeout=timeout):
        return
    raise SystemExit(
        f"LibreOffice が {timeout:.0f} 秒で UNO 応答しなかった (port={port})")


def stop_office(port: int = PORT, timeout: float = 20.0) -> int:
    """★ 接続先だけを terminate する。名前一括の taskkill / pkill はしない。

    taskkill / pkill は利用者が GUI で開いている LibreOffice も巻き込む。

    ★★ 2026-09-23: terminate が**届かない**回がある（生成マクロが soffice の中で無限ループ
      して本線を塞いでいる）。その時は誰も soffice を落とさず、以降の実行が全部詰まっていた
      （ailine の全件が 74 分止まった件の広がり方）。そこで、**port の持ち主の PID** が
      basrun の専用プロファイルで起こしたものだと確かめられた時だけ、PID 指定で落とす。
    ★★ 2026-09-23: port が閉じても、soffice はしばらく生きていて**冊のファイルを握っている**
      （実測: 戻った直後に soffice.bin が残っていた・次の検体の unlink が「使用中」で落ちた）。
      port でなく**プロセスが消えたこと**で「終了させた」と言う。

    ★ terminate() は投げたら即座に戻る。soffice.bin が実際に落ちてポートを
    手放すまでには間がある。**落ちたことを確かめずに「終了させた」と表示
    していた** —— この道具が扱っている「設定した ≠ 動く」そのものの形。

    2026-08-06 に実害が出た: 停止の直後に `ensure_office()` を呼ぶと、まだ
    開いたままのポートを見て「動いている」と誤認し、起動をやめる。その直後に
    プロセスが消えるので、次の接続が拒否される。テストが**単独では通り、
    通しで走らせると落ちた**のがこれ。`basrun stop` の直後に `apply` を
    叩く利用者にも同じことが起きる。

    そこで**ポートが実際に閉じるまで待ち、閉じなければそう言う。**
    """
    if not port_open(port):
        print("LibreOffice は動いていない")
        return 0
    # ★ 持ち主は**頼む前に**取る ── port が閉じた後では辿れない。
    owner = _port_owner(port)
    code = (
        "import uno\n"
        "l = uno.getComponentContext()\n"
        "r = l.ServiceManager.createInstanceWithContext("
        "'com.sun.star.bridge.UnoUrlResolver', l)\n"
        f"c = r.resolve('uno:socket,host=127.0.0.1,port={port};urp;"
        "StarOffice.ComponentContext')\n"
        "d = c.ServiceManager.createInstanceWithContext("
        "'com.sun.star.frame.Desktop', c)\n"
        "try:\n"
        "    d.terminate()\n"
        "except Exception:\n"
        "    pass\n"
    )
    # ★★ 2026-09-04 のレビューで見つけた穴: ここに時間の縛りが無かった。
    #   LibreOffice が Basic の実行で本線を塞いでいると、Desktop の生成も
    #   terminate() も**返ってこない**。stop_office は apply のタイムアウト
    #   復旧から呼ばれるので、**保険そのものがハングする**形だった。
    #   ★ 返らなくても諦めない ── ポートが閉じるかどうかは下で別に確かめる
    #     （terminate が届いていて、応答だけが返らない場合がある）。
    try:
        _office_py(code, timeout=min(TERMINATE_BUDGET, timeout))
    except subprocess.TimeoutExpired:
        print(f"終了要求が {min(TERMINATE_BUDGET, timeout):.0f} 秒で返らなかった"
               "（実行中のマクロが本線を塞いでいる可能性がある）。ポートの様子を見る")

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not port_open(port):
            if owner is None or _wait_gone(owner, deadline):
                print("接続先の LibreOffice を終了させた")
                return 0
            break               # port は閉じたが、プロセスがまだ残っている
        time.sleep(0.2)
    if owner is not None and _started_by_us(owner):
        _kill_pid_tree(owner)
        if _wait_gone(owner, time.monotonic() + KILL_GRACE) and not port_open(port):
            print(f"終了要求が届かなかったので、basrun が起こした LibreOffice（pid {owner}）を"
                  "PID 指定で落とした")
            return 0
    if port_open(port):
        print(f"終了を要求したが、{timeout:.0f} 秒たってもポート {port} が開いたままだ")
    else:
        print(f"ポート {port} は閉じたが、LibreOffice（pid {owner}）がまだ残っている")
    return 1


# ---------------------------------------------------------------------------
# 同期 (obasync に委譲)
# ---------------------------------------------------------------------------

def run_obasync(args: list[str]) -> int:
    """obasync を回す。★ ポートを必ず渡し、繋がることを確かめてから呼ぶ。

    obasync は接続できないと **自分で LibreOffice を起動する** (vendor 側
    786-790 行の Popen)。そのとき `-env:UserInstallation` も `--headless` も
    付けないので、**利用者の既定プロファイルが使われる。**

    2026-08-06 に実際に起きた: `-p` を渡し忘れていたため obasync は常に 2002 を
    見に行き、そこが空だったので自分で既定プロファイルの LibreOffice を起こし、
    そこへライブラリを書き込んだ。窓も出た。

    だから 2 段で塞ぐ:

    1. `-p` を必ず渡す         -> ensure_office が用意したポートと一致させる
    2. 呼ぶ直前に接続を確かめる -> 万一空なら **こちらで落とす**。
                                  obasync に「繋がらない」状況を渡さない

    ★ 3 段目: obasync は失敗時に stderr へ "ERROR:" を出しながら **exit 0**
    で返ることがある (`main()` 内の 2 箇所が他の 8 箇所と違って `sys.exit(1)`
    を挟まず bare `return` しているため。imacat/obasync HEAD で現存確認済み、
    issue 投稿済み: https://github.com/imacat/obasync/issues/3)。exit code
    だけを見て成功と判断すると、この失敗を握りつぶす。ここで stderr の
    "ERROR:" も見て非ゼロへ変換する。vendor 本体には手を入れない。
    """
    if not OBASYNC.exists():
        raise SystemExit(f"同梱の obasync が無い: {OBASYNC}")
    ensure_office()
    # ★ port_open では足りない。obasync がやるのと同じ resolve で確かめる。
    if not uno_ready(PORT, timeout=10.0):
        raise SystemExit(
            f"LibreOffice が port={PORT} で UNO 応答しない。ここで中止する。\n"
            "★ このまま obasync を呼ぶと、obasync が resolve に失敗し、"
            "**既定プロファイルで** LibreOffice を起動して、"
            "利用者の環境にライブラリを書き込む。")
    args = ["-p", str(PORT), *args]
    # ★★ 2026-09-23: 同期の段だけ時間の縛りが無かった（1d81e1d で他の 3 か所を縛った時の取り残し）。
    #   soffice が塞がると、ここが無制限に待ち、呼ぶ側（ailine）の外側の時間切れが basrun を殺しても
    #   **soffice は basrun の子でないので残り**、次の実行も同じ所で詰まった
    #   （ailine の pre-push の素の環境で、断りの盤の歩きが 1 件ずつ 210 秒の時間切れを繰り返した）。
    #   ★ 縛りは BASRUN_UNO_TIMEOUT（既定は無制限なので、同期だけは既定 SYNC_TIMEOUT=120 秒で必ず縛る ──
    #     同期は数秒で終わる段。時間切れなら stop_office で塞がりを取り除く）。
    limit = UNO_TIMEOUT or SYNC_TIMEOUT
    try:
        proc = run_office([str(OBASYNC), *args], timeout=limit)
    except subprocess.TimeoutExpired:
        rc = stop_office()
        raise SystemExit(
            f"同期（obasync）が {limit:.0f} 秒応答しなかった (BASRUN_UNO_TIMEOUT)。"
            + ("接続先の LibreOffice を終了させて中止した。" if rc == 0 else
               f"★ LibreOffice の終了にも失敗した (port={PORT})。手で落とすこと。"))
    # obasync は現行 python で SyntaxWarning を出す (`is` と文字列リテラル)。
    # 動作には影響しないので、利用者の目からは落とす。★ それ以外は必ず出す。
    saw_error = False
    for line in (proc.stderr or "").splitlines():
        if "SyntaxWarning" in line or line.strip().startswith("if storage.type"):
            continue
        print(line, file=sys.stderr)
        if line.strip().startswith("ERROR:"):
            saw_error = True
    sys.stdout.write(proc.stdout or "")
    if proc.returncode != 0:
        return proc.returncode
    return 1 if saw_error else 0


OPEN_SRC = r'''
import sys, uno
from com.sun.star.beans import PropertyValue
book, port = sys.argv[1], int(sys.argv[2])
local = uno.getComponentContext()
res = local.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local)
ctx = res.resolve(
    "uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % port)
desktop = ctx.ServiceManager.createInstanceWithContext(
    "com.sun.star.frame.Desktop", ctx)
hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
desktop.loadComponentFromURL(uno.systemPathToFileUrl(book), "_blank", 0, (hidden,))
# ★ この処理が終わっても文書は soffice 側に残る。UNO の参照を手放すだけ。
'''

CLOSE_SRC = r'''
import sys, uno
book, port, save = sys.argv[1], int(sys.argv[2]), sys.argv[3] == "1"
local = uno.getComponentContext()
res = local.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local)
ctx = res.resolve(
    "uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % port)
desktop = ctx.ServiceManager.createInstanceWithContext(
    "com.sun.star.frame.Desktop", ctx)
want = uno.systemPathToFileUrl(book)
matched, failed = 0, []
comps = desktop.getComponents().createEnumeration()
while comps.hasMoreElements():
    c = comps.nextElement()
    try:
        url = c.getURL()
    except Exception:
        continue
    if url != want:
        continue
    matched += 1
    # ★ 2026-09-04: ここは以前 except Exception: pass で**全部握りつぶして**いた。
    #   store() が落ちても（読み取り専用・容量不足・排他）呼び出し側は成功と読む。
    #   この道具が扱っている「成功を報告して何もしていない」そのものの形だった。
    try:
        if save:
            c.store()
        c.close(False)
    except Exception as e:
        failed.append("%s: %s" % (type(e).__name__, e))
if matched == 0:
    sys.stderr.write("CLOSE:no-match" + chr(10))
    sys.exit(2)
if failed:
    sys.stderr.write("CLOSE:failed " + "; ".join(failed) + chr(10))
    sys.exit(1)
sys.exit(0)
'''


def _office_py(code: str, *args: str,
                timeout: float | None = None) -> subprocess.CompletedProcess:
    """LibreOffice 同梱 python で短いコードを走らせる。

    ★ 2026-09-04: timeout を通せるようにした。UNO 呼び出しは 5 本あるのに、
      時間が縛られていたのは apply の 1 本だけで、開く・閉じる・終了要求は
      LibreOffice が塞がると無言で永久に待っていた。
      ★ 呼び出し側に散らさず**この 1 箇所**で受ける（同じ判断を写さない）。
    """
    # ★ 2026-09-04: 符号化を明示する。text=True だけだと Windows では cp932 で
    #   復号され、子側が UTF-8 を書いた瞬間に **UnicodeDecodeError で道具ごと落ちる**。
    #   実際にこのレビューの直しで踏んだ（埋め込みスクリプトに日本語を書いた回）。
    #   ★ 元から在った穴でもある ── ファイル名や UNO の例外文が非 ASCII なら同じことが起きる。
    return run_office(["-c", code, *args], timeout=timeout)


def open_book(book: Path) -> None:
    """★ obasync は文書を開けない。開ける側がこちらなので、ここで開く。

    短命なプロセスで開いても、文書は soffice 側に残る (UNO の参照を手放すだけ)。
    """
    ensure_office()
    p = _office_py(OPEN_SRC, str(book), str(PORT), timeout=UNO_TIMEOUT)
    if p.returncode != 0:
        sys.stderr.write(p.stderr or "")
        raise SystemExit(f"文書を開けなかった: {book}")


def close_book(book: Path, save: bool) -> int:
    """開いた文書を閉じる（sync の回は保存する）。★ 失敗を戻り値で返す。

    ★★ 2026-09-04 のレビューで見つけた穴: 以前は戻り値を捨てており、しかも
      CLOSE_SRC 側が全例外を握りつぶしていた。つまり
      **保存できなくても「同期できた」と報告していた**。
    ★ 0=閉じた / 1=保存か終了に失敗 / 2=対象が開いていなかった。
      2 も黙らない ── 「探す場所が空なら必ず通る」形を残さない。
    """
    p = _office_py(CLOSE_SRC, str(book), str(PORT), "1" if save else "0",
                   timeout=UNO_TIMEOUT)
    if p.returncode == 2:
        print(f"閉じようとした文書が開いていない: {book}", file=sys.stderr)
    elif p.returncode != 0:
        detail = (p.stderr or "").replace("CLOSE:failed ", "").strip()
        print(f"文書を閉じる/保存する側で失敗した: {book}"
              + (f" ── {detail}" if detail else ""), file=sys.stderr)
    return p.returncode


def _obasync_for(a: argparse.Namespace, extra: list[str]) -> int:
    """--book が指定されていれば、文書を開いてから obasync を回す。"""
    lib = a.library or Path(a.dir).resolve().name
    book = Path(a.book).resolve() if getattr(a, "book", None) else None
    use_doc = bool(book) or a.doc
    if book:
        if not book.exists():
            raise SystemExit(f"文書が無い: {book}")
        open_book(book)
    rc, crc = 1, 0
    try:
        args = [*extra, "-x", a.ext, "-e", a.encoding,
                "--doc" if use_doc else "--user"]
        if book:
            # obasync は開いている文書を部分パスで選ぶ
            args += ["--target", book.name]
        rc = run_obasync([*args, a.dir, lib])
    finally:
        if book:
            # ★ pull は読むだけ。sync は文書側を書き換えるので保存する。
            crc = close_book(book, save=("--get" not in extra))
    # ★ 2026-09-04: 同期が通っても**閉じる/保存に失敗していたら成功と言わない**。
    #   finally の中で return を書き換えることはできない（Python は finally の
    #   代入で戻り値を変えない）ので、閉じた後にここで判定する。
    if rc == 0 and crc != 0:
        print("同期はできたが、文書を閉じる/保存する側で失敗した", file=sys.stderr)
        return crc
    return rc


def sync_cmd(a: argparse.Namespace) -> int:
    return _obasync_for(a, [])


def pull_cmd(a: argparse.Namespace) -> int:
    Path(a.dir).mkdir(parents=True, exist_ok=True)
    return _obasync_for(a, ["--get"])


# ---------------------------------------------------------------------------
# 適用 (こちら側)
# ---------------------------------------------------------------------------

APPLY_SRC = r'''
import sys, uno
from com.sun.star.beans import PropertyValue

book, lib, module, sub, port = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], int(sys.argv[5])
local = uno.getComponentContext()
res = local.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local)
ctx = res.resolve(
    "uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % port)
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(book), "_blank", 0, (hidden,))
try:
    factory = smgr.createInstanceWithContext(
        "com.sun.star.script.provider.MasterScriptProviderFactory", ctx)
    provider = factory.createScriptProvider("")
    url = ("vnd.sun.star.script:%s.%s.%s?language=Basic&location=application"
           % (lib, module, sub))
    script = provider.getScript(url)
    # ★ 文書を引数で渡す。ThisComponent に頼ると headless で空振りする。
    script.invoke((doc,), (), ())
    doc.store()          # ★ 元の形式のまま上書き
finally:
    doc.close(False)
print("applied %s.%s.%s -> %s" % (lib, module, sub, book))
'''


def apply_cmd(a: argparse.Namespace) -> int:
    if "." not in a.entry:
        raise SystemExit("Module.Sub の形で指定すること (例: Amount.FillAmounts)")
    module, sub = a.entry.rsplit(".", 1)
    book = Path(a.book).resolve()
    if not book.exists():
        raise SystemExit(f"文書が無い: {book}")

    lib = a.library or Path(a.dir).resolve().name
    rc = sync_cmd(argparse.Namespace(dir=a.dir, library=lib, ext=a.ext,
                                     encoding=a.encoding, doc=False, book=None))
    if rc != 0:
        return rc

    if a.backup:
        bak = book.with_suffix(book.suffix + ".bak")
        shutil.copy2(book, bak)
        print(f"控えを作った: {bak}")

    ensure_office()
    timeout = a.timeout if a.timeout is not None else APPLY_TIMEOUT
    try:
        proc = run_office(["-c", APPLY_SRC, str(book), lib, module, sub, str(PORT)],
                          timeout=timeout)
    except subprocess.TimeoutExpired:
        # ★ opt-in (既定は無制限、上の APPLY_TIMEOUT のコメント参照)。生成
        # マクロが無限ループすると、この呼び出しは永久にハングする (TS 移行
        # での実測)。timeout に達したら、ハングしている接続先の LibreOffice
        # だけを終了する —— 既存の stop_office() をそのまま再利用するので、
        # taskkill のようなプロセス名一括終了はしない。
        # ★★ 2026-09-04: 戻り値を見ずに「終了させた」と断定していた。
        #   stop_office は落とせなければ 1 を返す ── まさにこの関数の docstring が
        #   「落ちたことを確かめずに『終了させた』と表示していた」と書いている形が、
        #   **同じファイルの 2 箇所目で再発していた**（片配線）。
        rc = stop_office()
        head = (f"apply が {timeout:.0f} 秒応答しなかった (BASRUN_APPLY_TIMEOUT/"
                "--timeout)。")
        if rc == 0:
            raise SystemExit(head + "接続先の LibreOffice を終了させて中止した。")
        raise SystemExit(
            head + f"★ LibreOffice の終了にも失敗した (port={PORT})。"
            "手で落とすこと。マクロが無限ループしている可能性がある。")
    sys.stdout.write(proc.stdout or "")
    sys.stderr.write(proc.stderr or "")
    return proc.returncode


# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = p.add_subparsers(dest="cmd", required=True)

    def common(sp: argparse.ArgumentParser) -> argparse.ArgumentParser:
        sp.add_argument("dir", help="ソースを置くディレクトリ")
        sp.add_argument("library", nargs="?", default=None,
                        help="ライブラリ名 (既定: ディレクトリ名)")
        sp.add_argument("-x", "--ext", default=".bas", help="拡張子 (既定 .bas)")
        sp.add_argument("-e", "--encoding", default="utf-8",
                        help="ソースの符号化 (既定 utf-8)")
        sp.add_argument("--doc", action="store_true",
                        help="文書側の格納先を対象にする (既定はユーザ側)")
        sp.add_argument("--book", default=None,
                        help="この文書を開いてから、その文書側の格納先を対象にする。"
                             "★ obasync 単体では開いている文書しか選べない")
        return sp

    s = common(sub.add_parser("sync", help="ソース群 -> ライブラリ"))
    s.set_defaults(func=sync_cmd)

    g = common(sub.add_parser("pull", help="ライブラリ -> ソース群 (救出)"))
    g.set_defaults(func=pull_cmd)

    ap = sub.add_parser("apply", help="同期して文書に適用し保存する")
    ap.add_argument("book", help="対象の文書 (.xlsx / .ods など)")
    ap.add_argument("dir", help="ソースを置くディレクトリ")
    ap.add_argument("library", nargs="?", default=None)
    ap.add_argument("entry", help="Module.Sub (Sub は文書を引数で受けること)")
    ap.add_argument("-x", "--ext", default=".bas")
    ap.add_argument("-e", "--encoding", default="utf-8")
    ap.add_argument("--backup", action="store_true", help="上書き前に .bak を作る")
    ap.add_argument("--timeout", type=float, default=None,
                     help="apply がハングしたとみなす秒数 (opt-in、既定は無制限。"
                          "BASRUN_APPLY_TIMEOUT でも指定できる)")
    ap.set_defaults(func=apply_cmd)

    st = sub.add_parser("stop", help="起動した LibreOffice を落とす")
    st.set_defaults(func=lambda a: stop_office())

    return p


def main() -> int:
    a = build_parser().parse_args()
    return a.func(a)


if __name__ == "__main__":
    raise SystemExit(main())
