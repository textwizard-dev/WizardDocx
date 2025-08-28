# SPDX-FileCopyrightText: 2024–2025 Mattia Rubino
# SPDX-License-Identifier: AGPL-3.0-or-later


from pathlib import Path
from typing import Union, Optional,List
from wizarddocx.wizard_docx.extraction_text import TextExtractor
from wizarddocx.utils.errors.errors_handle import handle_errors


class WizardDocx:
    def __init__(self):
        self._text_extractor = TextExtractor()

    
    # ----------------------------------------------------------------------
    # Text extraction
    # ----------------------------------------------------------------------

    @handle_errors
    def extract_text(
            self,
            input_data: Union[str, bytes, Path],
            extension: Optional[str] = None,
            pages: Optional[Union[int, str, List[Union[int, str]]]] = None,
            ocr: bool = False,
            language_ocr: str = "eng",    
    ) -> str:
        """
        Extracts text from the provided input data based on its format and type.

        Args:
            input_data (Union[str, bytes, Path]):
                The input for extraction: a filesystem path, raw bytes, or string content.
            extension (Optional[str]):
                File extension to use when `input_data` is bytes (e.g. 'docx', 'doc').
            pages (Optional[int | str | list[int | str]]):
                • For paged formats: one-based page numbers to extract.
                • If None (default), all pages are extracted.
            ocr (bool):
                Enables OCR for text extraction using Tesseract OCR. Applicable for formats
                like DOCX
            language_ocr (str):
                Tesseract language code (default: 'eng').

        Returns:
            str: The extracted text content.

        Raises:
            InvalidInputError: If the input data is invalid or unsupported.

        Supported formats:
            'doc', 'docx'
        """
        
        
        return self._text_extractor.data_extractor(
            input_data,
            extension,
            pages,
            ocr,
            language_ocr,
        )

