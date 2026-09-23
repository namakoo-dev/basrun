"""pytest 共有セットアップ。

basrun.py はパッケージ化されていない、リポジトリ直下の単体スクリプト。
tests/ から見て一つ上のディレクトリ (basrun.py が置かれている場所) を
import できるようにする。テスト対象を tests/ 配下に複製せず、既存ファイルを
そのまま import するための橋渡し。

``integration`` マーカーもここで登録する。LibreOffice を実際に起動する
統合テストにだけ付ける。登録しておかないと ``-m "not integration"`` が
未知マーカーとして警告を出す。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import basrun  # noqa: E402  （sys.path を通した後でなければ import できない）


def pytest_configure(config):
    config.addinivalue_line(
        "markers",
        "integration: LibreOffice を実際に起動する統合テスト (遅い・環境依存)")
    config.addinivalue_line(
        "markers",
        "office_locator: LO の在処を引く関数そのものを（入力を制御して）試すテスト。"
        "_no_real_libreoffice の stub 対象から外す ── 実 LO には依存しない")


@pytest.fixture(autouse=True)
def _no_real_libreoffice(request, monkeypatch):
    """★ 2026-08-23: integration マーカーが嘘の分割だった件の構造的な直し。

    非 integration のテストが実 LibreOffice を要求していても、開発機には LO が
    入っているのでローカルでは緑になり、CI（LO 無し）で初めて落ちた。
    「居るから見えない」── 居ない側を常に再現して、その場で鳴らす。

    非 integration のテストでは LO の在処を引く関数を必ず失敗させる。実 LO が要る
    テストは @pytest.mark.integration を付ける（そちらは免除）。
    """
    if "integration" in request.keywords or "office_locator" in request.keywords:
        yield
        return

    def _boom(*a, **kw):
        raise AssertionError(
            "非 integration のテストが実 LibreOffice を要求した"
            "（CI には存在しない ── tests/conftest.py の _no_real_libreoffice）")

    for name in ("office_dir", "office_python", "soffice_bin"):
        if hasattr(basrun, name):
            monkeypatch.setattr(basrun, name, _boom, raising=False)
    yield


@pytest.fixture
def fake_office(tmp_path, monkeypatch):
    """LO の在処だけを偽の踏み台に差し替える（実 LO は起動しない）。

    ★ 「LO が在ることを前提にした上の層」を試すテストが、この踏み台を**明示的に**
    要求する形にした ── 以前は開発機の実 LO に暗黙に寄りかかっており、CI で初めて
    落ちた（22 件中 11 件が実は LO 依存だった・2026-08-23 実測）。
    """
    prog = tmp_path / "program"
    prog.mkdir()
    (prog / "soffice.exe").write_text("", encoding="utf-8")
    (prog / "python.exe").write_text("", encoding="utf-8")
    monkeypatch.setattr(basrun, "office_dir", lambda *a, **kw: prog, raising=False)
    monkeypatch.setattr(basrun, "office_python", lambda *a, **kw: prog / "python.exe",
                         raising=False)
    # ★ 偽の LibreOffice には port の持ち主がいない ── 単体試験が開発機の本物の soffice の
    #   PID を拾って落とすことを、既定で起こさない。
    monkeypatch.setattr(basrun, "_port_owner", lambda *a, **kw: None, raising=False)
    return prog
