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
    p = subprocess.run(
        [str(office_python()), "-c", UNO_READY_SRC, str(port), str(timeout)],
        capture_output=True, text=True)
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


def stop_office(port: int = PORT) -> int:
    """★ 接続先だけを terminate する。taskkill しない。

    taskkill / pkill は利用者が GUI で開いている LibreOffice も巻き込む。
    """
    if not port_open(port):
        print("LibreOffice は動いていない")
        return 0
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
    subprocess.run([str(office_python()), "-c", code], check=False)
    print("接続先の LibreOffice を終了させた")
    return 0


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
    cmd = [str(office_python()), str(OBASYNC), *args]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    # obasync は現行 python で SyntaxWarning を出す (`is` と文字列リテラル)。
    # 動作には影響しないので、利用者の目からは落とす。★ それ以外は必ず出す。
    for line in (proc.stderr or "").splitlines():
        if "SyntaxWarning" in line or line.strip().startswith("if storage.type"):
            continue
        print(line, file=sys.stderr)
    sys.stdout.write(proc.stdout or "")
    return proc.returncode


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
comps = desktop.getComponents().createEnumeration()
while comps.hasMoreElements():
    c = comps.nextElement()
    try:
        if c.getURL() == want:
            if save:
                c.store()
            c.close(False)
    except Exception:
        pass
'''


def _office_py(code: str, *args: str) -> subprocess.CompletedProcess:
    return subprocess.run([str(office_python()), "-c", code, *args],
                          capture_output=True, text=True)


def open_book(book: Path) -> None:
    """★ obasync は文書を開けない。開ける側がこちらなので、ここで開く。

    短命なプロセスで開いても、文書は soffice 側に残る (UNO の参照を手放すだけ)。
    """
    ensure_office()
    p = _office_py(OPEN_SRC, str(book), str(PORT))
    if p.returncode != 0:
        sys.stderr.write(p.stderr or "")
        raise SystemExit(f"文書を開けなかった: {book}")


def close_book(book: Path, save: bool) -> None:
    _office_py(CLOSE_SRC, str(book), str(PORT), "1" if save else "0")


def _obasync_for(a, extra: list[str]) -> int:
    """--book が指定されていれば、文書を開いてから obasync を回す。"""
    lib = a.library or Path(a.dir).resolve().name
    book = Path(a.book).resolve() if getattr(a, "book", None) else None
    use_doc = bool(book) or a.doc
    if book:
        if not book.exists():
            raise SystemExit(f"文書が無い: {book}")
        open_book(book)
    try:
        args = [*extra, "-x", a.ext, "-e", a.encoding,
                "--doc" if use_doc else "--user"]
        if book:
            # obasync は開いている文書を部分パスで選ぶ
            args += ["--target", book.name]
        return run_obasync([*args, a.dir, lib])
    finally:
        if book:
            # ★ pull は読むだけ。sync は文書側を書き換えるので保存する。
            close_book(book, save=("--get" not in extra))


def sync_cmd(a) -> int:
    return _obasync_for(a, [])


def pull_cmd(a) -> int:
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


def apply_cmd(a) -> int:
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
    proc = subprocess.run(
        [str(office_python()), "-c", APPLY_SRC,
         str(book), lib, module, sub, str(PORT)],
        capture_output=True, text=True)
    sys.stdout.write(proc.stdout or "")
    sys.stderr.write(proc.stderr or "")
    return proc.returncode


# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = p.add_subparsers(dest="cmd", required=True)

    def common(sp):
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
    ap.set_defaults(func=apply_cmd)

    st = sub.add_parser("stop", help="起動した LibreOffice を落とす")
    st.set_defaults(func=lambda a: stop_office())

    return p


def main() -> int:
    a = build_parser().parse_args()
    return a.func(a)


if __name__ == "__main__":
    raise SystemExit(main())
