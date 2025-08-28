# SPDX-FileCopyrightText: 2024–2025 Mattia Rubino
# SPDX-License-Identifier: AGPL-3.0-or-later

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from wizarddocx.utils.errors.errors import (
    AntiwordNotFoundError,
    ExtractionError,
    FileNotFoundCustomError,
)


class DocReader:
    """Extract text from .doc using, in order: antiword → catdoc → soffice (LibreOffice)."""

    def __init__(self) -> None:
        pass

    @staticmethod
    def _which(cmd: str) -> bool:
        return shutil.which(cmd) is not None

    @staticmethod
    def _run(cmd: list[str]) -> tuple[int, str, str]:
        p = subprocess.run(cmd, capture_output=True, text=True)
        return p.returncode, p.stdout, p.stderr

    @staticmethod
    def _ensure_exists(path: str) -> None:
        if not Path(path).exists():
            raise FileNotFoundCustomError(f"The file '{path}' does not exist or is inaccessible.")

    def _try_antiword(self, doc: str) -> str | None:
        if not self._which("antiword"):
            return None
        code, out, _ = self._run(["antiword", doc])
        return out if code == 0 and out else None

    def _try_catdoc(self, doc: str) -> str | None:
        if not self._which("catdoc"):
            return None
        code, out, _ = self._run(["catdoc", "-w", doc])
        return out if code == 0 and out else None

    def _try_soffice_txt(self, doc: str) -> str | None:
        if not self._which("soffice"):
            return None
        with tempfile.TemporaryDirectory() as tmp:
            code, _, _ = self._run(
                ["soffice", "--headless", "--convert-to", "txt:Text", "--outdir", tmp, doc]
            )
            if code != 0:
                return None
            txt_path = Path(tmp) / (Path(doc).stem + ".txt")
            if not txt_path.exists():
                return None
            try:
                return txt_path.read_text(encoding="utf-8", errors="replace")
            except Exception:
                return txt_path.read_text(errors="replace")

    def doc_reader(self, doc: str) -> str:
        """
        Extract text from a .doc file.

        Raises:
            FileNotFoundCustomError: File missing or inaccessible.
            AntiwordNotFoundError: No available extractor (antiword/catdoc/soffice).
            ExtractionError: Extractors present but all failed.
        """
        self._ensure_exists(doc)

        out = self._try_antiword(doc) or self._try_catdoc(doc) or self._try_soffice_txt(doc)
        if out:
            return out

        if not (self._which("antiword") or self._which("catdoc") or self._which("soffice")):
            raise AntiwordNotFoundError(
                "No available .doc extractor. Install one of: antiword, catdoc, or LibreOffice (soffice)."
            )
        raise ExtractionError("Failed to extract text from .doc with antiword/catdoc/soffice.")
