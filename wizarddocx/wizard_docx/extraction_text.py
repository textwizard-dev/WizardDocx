# SPDX-FileCopyrightText: 2024–2025 Mattia Rubino
# SPDX-License-Identifier: AGPL-3.0-or-later

import io
from pathlib import Path
from typing import Union, Optional, List

from wizarddocx.utils.errors.errors import (
    FileNotFoundCustomError,
    UnsupportedExtensionError,
    DocFileAsBytesError,
    InvalidPagesError,
)
from wizarddocx.wizard_docx.tool.docx_reader import DocxReader
from wizarddocx.wizard_docx.tool.doc_reader import DocReader



class TextExtractor:
    """
    Handles text extraction from various formats, with optional OCR and page selection.
    """

    def __init__(self) -> None:
        self._doc = DocReader()
        self._docx = DocxReader()


    @staticmethod
    def _validate_selector(sel):
        if sel is None:
            return None
        if isinstance(sel, (int, str)):
            return [sel]
        if isinstance(sel, list) and all(isinstance(x, (int, str)) for x in sel):
            return sel
        raise InvalidPagesError(sel)

    def data_extractor(
            self,
            input_data: Union[str, bytes, Path],
            extension: Optional[str] = None,
            pages_or_sheets: Optional[Union[int, str,
            List[Union[int, str]]]] = None,
            ocr: bool = False,
            language_ocr: str = "eng",
    ) -> str:

        # 1) Validate & canonicalize pages
        selector = self._validate_selector(pages_or_sheets)

        # 2) Normalize input_data → always bytes (or path for .doc)
        if isinstance(input_data, (str, Path)):
            path = Path(input_data)
            if not path.exists():
                raise FileNotFoundCustomError(path)
            ext = path.suffix.lower().lstrip(".")
            if ext == "doc":
                # .doc reader needs a file path
                raw = str(path)
            else:
                raw = path.read_bytes()
            extension = ext
        else:
            # input is bytes
            if extension is None:
                raise UnsupportedExtensionError(None)
            extension = extension.lower()
            raw = input_data
            if extension == "doc":
                raise DocFileAsBytesError()

        # 3) Dispatch
        reader_map = {
            "doc": lambda: self._doc.doc_reader(raw),
            "docx": lambda: self._docx.docx_reader(io.BytesIO(raw), pages_list=selector, ocr=ocr, language_ocr=language_ocr),
        }

        if extension not in reader_map:
            raise UnsupportedExtensionError(extension)

        return reader_map[extension]()

