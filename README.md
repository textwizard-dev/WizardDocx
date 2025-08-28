<img src="https://raw.githubusercontent.com/textwizard-dev/wizarddocx/main/asset/wizarddocx%20Banner.png"
     alt="wizarddocx Banner" width="800" height="300">

---

# Wizard Docx
[![PyPI - Version](https://img.shields.io/pypi/v/wizarddocx)](https://pypi.org/project/wizarddocx/)
[![PyPI - Downloads/month](https://img.shields.io/pypi/dm/wizarddocx?label=PyPI%20downloads)](https://pypistats.org/packages/wizarddocx)
[![License](https://img.shields.io/pypi/l/wizarddocx)](https://github.com/textwizard-dev/wizarddocx/blob/main/LICENSE)



**WizardDocx** is a Python library focused on text extraction from Microsoft Word documents.  
It parses Word documents natively and can apply local OCR with Tesseract for embedded images or scanned pages inside 'docx'.  
Legacy `.doc` is supported in read-only mode without OCR.

---

## Contents

- [Installation](#installation)
- [Quick start](#quick-start)
- [Text extraction](#text-extraction)
- [License](#license)
- [Resources](#resources)


---
## Installation

Requires Python 3.9+.

~~~bash
pip install wizarddocx
~~~


> For OCR capabilities, ensure you have [Tesseract](https://github.com/tesseract-ocr/tesseract) installed on your system.  

---

## Quick start

~~~python
import wizarddocx as wd

text = wd.extract_text("example.docx")
print(text)
~~~

---

## Text extraction

### Parameters

- `input_data`: `[str, bytes, Path]`  
- `extension`: The file extension, required only if `input_data` is `bytes`.  
- `pages`: page selection for .docx.  
  • Examples: `1`, `"1-3"`, `[1, 3, "5-8"]`  
- `ocr`: Enables OCR using Tesseract. Applies to DOCX and image-based files no for doc.  
- `language_ocr`: Language code for OCR. Defaults to `'eng'`.

### Examples

Basic:

~~~python
import wizarddocx as wd

txt = wd.extract_text("docs/report.docx")
~~~

From bytes:

~~~python
from pathlib import Path
import wizarddocx as wd

raw = Path("img.docx").read_bytes()
txt_img = wd.extract_text(raw, extension="docx")
~~~

Paged selection and OCR:

~~~python
import wizarddocx as wd

sel = wd.extract_text("docs/big.docx", pages=[1, 3, "5-7"])
ocr_txt = wd.extract_text("scan.docx", ocr=True, language_ocr="ita")
~~~

#### **Supported Formats**

| Format | OCR Option |
|---|---|
| DOC | Not available |
| DOCX | Optional |

---

## License

[AGPL-3.0-or-later](LICENSE).

## RESOURCES

- [GitHub Repository](https://github.com/textwizard-dev/wizarddocx)
- [Documentation](https://wizarddocx.readthedocs.io/en/latest/)
- [PyPI Package](https://pypi.org/project/wizarddocx/)
---

## Contact & Author

**Author:** Mattia Rubino  
**Email:** <textwizard.dev@gmail.com>
