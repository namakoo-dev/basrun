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

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


def pytest_configure(config):
    config.addinivalue_line(
        "markers",
        "integration: LibreOffice を実際に起動する統合テスト (遅い・環境依存)")
