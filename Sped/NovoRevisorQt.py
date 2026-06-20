from __future__ import annotations

import sys
from pathlib import Path

# ── Bootstrap: garante que estamos rodando com o Python do venv ───────────────
# No Windows, o alias do sistema pode interceptar o comando "python" mesmo com
# o venv ativado. Se pandas/PySide6 não estiver disponível, procura o python
# do venv (.venv-dev ou .venv-prod) e reinicia o processo com ele.
def _bootstrap_venv() -> None:
    try:
        import pandas  # noqa: F401
        import PySide6  # noqa: F401
        return  # já no ambiente correto
    except ImportError:
        pass

    script = Path(__file__).resolve()
    project_root = script.parent.parent

    env = "prod" if "prod" in sys.argv else "dev"
    candidates = [
        project_root / f".venv-{env}" / "Scripts" / "python.exe",
        project_root / ".venv-dev" / "Scripts" / "python.exe",
        project_root / ".venv-prod" / "Scripts" / "python.exe",
    ]
    for venv_py in candidates:
        if venv_py.exists():
            import os, subprocess
            result = subprocess.run([str(venv_py), str(script)] + sys.argv[1:])
            sys.exit(result.returncode)

    print(
        "ERRO: pandas/PySide6 nao encontrado e nenhum venv localizado.\n"
        "Execute: .\\scripts\\run-dev.ps1",
        file=sys.stderr,
    )
    sys.exit(1)


_bootstrap_venv()

# ── Imports normais (somente após o bootstrap) ────────────────────────────────
import faulthandler

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from app.ui_qt.app import QtSpedApp


def main() -> None:
    crash_log = Path(__file__).resolve().parent / "qt_crash.log"
    try:
        crash_file = crash_log.open("a", encoding="utf-8")
        faulthandler.enable(crash_file)
    except Exception:
        crash_file = None
    app = QApplication(sys.argv)
    app.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    window = QtSpedApp()
    window.showMaximized()
    try:
        sys.exit(app.exec())
    finally:
        if crash_file is not None:
            crash_file.close()


if __name__ == "__main__":
    main()
