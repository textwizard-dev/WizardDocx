# SPDX-FileCopyrightText: 2024–2025 Mattia Rubino
# SPDX-License-Identifier: AGPL-3.0-or-later

from __future__ import annotations

import io, re, zipfile, xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Sequence, Union, Optional

from wizarddocx.utils.errors.errors import (
    DocxFileError,
    ImageProcessingError,
    OCRNotConfiguredError,
)
from wizarddocx.wizard_docx.utils.selector import  normalize_pages_selector

_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
}
def qn(p: str, t: str) -> str:
    return f"{{{_NS[p]}}}{t}"

_RELS_RX   = re.compile(r"word/_rels/[^/]+\.xml\.rels$")
_HEADER_RX = re.compile(r"word/header\d*\.xml$")
_FOOTER_RX = re.compile(r"word/footer\d*\.xml$")

# ── registry tag → handler ─────────────────────────────────────────
class _Registry(dict):
    def register(self, tag: str):
        def deco(f): self[tag] = f; return f
        return deco
_handlers: _Registry = _Registry()

@dataclass
class _Extractor:
    zf: zipfile.ZipFile
    lang: str
    do_ocr: bool

    _rels: Dict[str, str]  = field(init=False, default_factory=dict)
    _ocr:  Dict[str, str]  = field(init=False, default_factory=dict)
    _pages: List[str]      = field(init=False, default_factory=list)
    _buf:   List[str]      = field(init=False, default_factory=list)

    def run(self, xml_parts: list[tuple[str, bytes]]) -> List[str]:
        self._load_relationships()
        if self.do_ocr:
            self._build_ocr_cache()

        emit = self._buf.append
        for _name, data in xml_parts:
            for _ev, el in ET.iterparse(io.BytesIO(data), events=("end",)):
                if h := _handlers.get(el.tag):
                    h(self, el, emit)
                el.clear()

        self._pages.append("".join(self._buf).strip())
        return self._pages

    def _load_relationships(self) -> None:
        for fname in self.zf.namelist():
            if not _RELS_RX.fullmatch(fname):
                continue
            for rel in ET.fromstring(self.zf.read(fname)):
                rid, tgt = rel.get("Id"), rel.get("Target")
                if rid and tgt:
                    tgt = tgt.replace("\\", "/").lstrip("/")
                    while tgt.startswith("../"):
                        tgt = tgt[3:]
                    self._rels[rid] = tgt

    def _build_ocr_cache(self) -> None:
        try:
            import pytesseract
            from PIL import Image as PILImage
        except ModuleNotFoundError as exc:
            raise OCRNotConfiguredError("OCR requested but pytesseract/Pillow not installed.") from exc

        try:
            if pytesseract.get_tesseract_version() is None:  # type: ignore[attr-defined]
                raise OCRNotConfiguredError("Tesseract not found in PATH")
        except Exception as exc:
            raise OCRNotConfiguredError("Tesseract not available") from exc

        SUP = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"}
        for tgt in set(self._rels.values()):
            if Path(tgt).suffix.lower() not in SUP:
                continue
            member = f"word/{tgt}" if not tgt.startswith("word/") else tgt
            try:
                data = self.zf.read(member)
                with PILImage.open(io.BytesIO(data)) as im:
                    txt = pytesseract.image_to_string(im, lang=self.lang).strip() or "[Empty OCR]"
            except Exception as exc:
                raise ImageProcessingError(f"OCR failed on {member}: {exc}") from exc
            self._ocr[tgt] = txt

    def _ocr_for_rid(self, rid: str | None) -> str | None:
        if rid and self.do_ocr and rid in self._rels:
            return self._ocr.get(self._rels[rid])
        return None

# ── tag handlers ──────────────────────────────────────────────────
@_handlers.register(qn("w", "t"))
def _(ex, el, emit): emit(el.text or "")

@_handlers.register(qn("w", "tab"))
def _(ex, el, emit): emit("\t")

@_handlers.register(qn("w", "br"))
@_handlers.register(qn("w", "cr"))
def _(ex, el, emit): emit("\n")

@_handlers.register(qn("w", "p"))
def _(ex, el, emit): emit("\n\n")

@_handlers.register(qn("w", "lastRenderedPageBreak"))
def _(ex, el, emit):
    ex._pages.append("".join(ex._buf).strip())
    ex._buf.clear()

@_handlers.register(qn("a", "blip"))          # DrawingML image
def _(ex, el, emit):
    if txt := ex._ocr_for_rid(el.get(qn("r", "embed"))):
        emit("\n" + txt + "\n")

@_handlers.register(qn("v", "imagedata"))     # VML image
def _(ex, el, emit):
    if txt := ex._ocr_for_rid(el.get(qn("r", "id"))):
        emit("\n" + txt + "\n")

# ── public API ────────────────────────────────────────────────────
class DocxReader:
    __slots__ = ()

    def docx_reader(
        self,
        docx_source: Union[bytes, str, Path, io.BytesIO],
        pages_list: Optional[Sequence[Union[int, str]]] = None,  # 1, "3-5", "1,3,5-7", [...]
        ocr: bool = False,
        language_ocr: str = "eng",
    ) -> str:
        zf = self._open_zip(docx_source)
        try:
            xml_parts = [
                (n, zf.read(n))
                for n in sorted(zf.namelist())
                if n.endswith(".xml") and (
                    n == "word/document.xml"
                    or _HEADER_RX.fullmatch(n)
                    or _FOOTER_RX.fullmatch(n)
                )
            ]

            ext = _Extractor(zf, language_ocr, ocr)
            pages = ext.run(xml_parts) 

            idx = normalize_pages_selector(pages_list, len(pages))
            return "\n".join(pages[i] for i in idx)
        finally:
            zf.close()

    @staticmethod
    def _open_zip(src):
        try:
            if isinstance(src, (bytes, bytearray)):
                return zipfile.ZipFile(io.BytesIO(src))
            if hasattr(src, "read"):
                return zipfile.ZipFile(src)
            return zipfile.ZipFile(src)
        except zipfile.BadZipFile as exc:
            raise DocxFileError("Invalid .docx file") from exc
