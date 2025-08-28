===========
Wizard Docx
===========

.. figure:: _static/img/wizarddocxBanner.png
   :alt: wizarddocx Banner
   :width: 800
   :height: 300
   :align: center

.. image:: https://img.shields.io/pypi/v/wizarddocx.svg
   :target: https://pypi.org/project/wizarddocx/
   :alt: PyPI - Version

.. image:: https://img.shields.io/pypi/dm/wizarddocx.svg?label=PyPI%20downloads
   :target: https://pypistats.org/packages/wizarddocx
   :alt: PyPI - Downloads/month

.. image:: https://img.shields.io/pypi/l/wizarddocx.svg
   :target: https://github.com/textwizard-dev/wizarddocx/blob/main/LICENSE
   :alt: License



**WizardDocx** is a Python library focused on text extraction from Microsoft Word documents.  
It parses Word documents natively and can apply local OCR with Tesseract for embedded images or scanned pages inside 'docx'.  
Legacy `.doc` is supported in read-only mode without OCR.

Installation
============

Requires Python 3.9+.

.. code-block:: bash

   pip install wizarddocx


.. note::
   For OCR, install `Tesseract <https://github.com/tesseract-ocr/tesseract>`_.  

Quick start
===========

.. code-block:: python

   import wizarddocx as wd

   text = wd.extract_text("example.docx")
   print(text)


Supported formats
=================

+-----------+--------------+
| Format    | OCR option   |
+===========+==============+
| DOC       | Not available|
+-----------+--------------+
| DOCX      | Optional     |
+-----------+--------------+


Parameters
==========

+---------------------------+--------------------------------------------------------------------------+
| **Parameter**             | **Description**                                                          |
+===========================+==========================================================================+
| ``input_data``            | (*str | bytes | Path*) Source to extract from: path string, bytes, or    |
|                           | ``pathlib.Path``.                                                        |
+---------------------------+--------------------------------------------------------------------------+
| ``extension``             | (*str, optional*) File extension when ``input_data`` is bytes            |
|                           | (e.g., ``"doc"``, ``"docx"``).                                           |
+---------------------------+--------------------------------------------------------------------------+
| ``pages``                 | (*int | str | list[int|str] | None*) Page selection only docx. For paged |
|                           | formats use numbers and ranges (``1``, ``"2-5"``, ``[1, "5-7"]``).       |
+---------------------------+--------------------------------------------------------------------------+
| ``ocr``                   | (*bool, optional*) Enable Tesseract OCR for images and scanned DOCX.     |
|                           | Defaults to ``False``.                                                   |
+---------------------------+--------------------------------------------------------------------------+
| ``language_ocr``          | (*str, optional*) Tesseract language code. Defaults to ``"eng"``.        |
+---------------------------+--------------------------------------------------------------------------+


Detailed parameters and examples
================================

``input_data``
--------------

Accepts a filesystem path, a ``pathlib.Path``, or raw ``bytes``.

**Path string**

.. code-block:: python

   import wizarddocx as wd
   text = wd.extract_text("docs/report.docx")

**pathlib.Path**

.. code-block:: python

   from pathlib import Path
   import wizarddocx as wd
   text = wd.extract_text(Path("docs/report.docx"))

**Bytes (must set ``extension``)**

.. code-block:: python

   from pathlib import Path
   import wizarddocx as wd
   raw = Path("img.doc").read_bytes()
   text = wd.extract_text(raw, extension="doc")

**BytesIO (streams)**

.. code-block:: python

   import io, wizarddocx as wd
   buf = io.BytesIO(open("img.docx", "rb").read())
   text = wd.extract_text(buf.getvalue(), extension="docx")

``extension``
-------------

Required only when passing ``bytes``. Indicates the file type.

**Example**

.. code-block:: python

   import wizarddocx as wd
   doc_bytes = open("img.doc", "rb").read()
   text = wd.extract_text(doc_bytes, extension="doc")

.. warning::
   Passing bytes without ``extension`` raises a validation error.

``pages``
---------

Select pages.

Accepted forms by format:

- **Paged** — 1-based:
  - single int: ``1``
  - range string: ``"1-3"``
  - CSV string: ``"1,3,5-7"``
  - mixed list: ``[1, 3, "5-7"]``
  Invalid tokens and out-of-range pages are silently skipped.


**Examples — paged**

.. code-block:: python

   import wizarddocx as wd
   page1  = wd.extract_text("docs/big.docx", pages=1)
   subset = wd.extract_text("docs/big.docx", pages=[1, 3, "5-7"])


----------------------------

Enable OCR for raster content and scanned documents. ``language_ocr`` controls the recognition language.

**Images**

.. code-block:: python

   import wizarddocx as wd
   img_txt = wd.extract_text("scan.docx", ocr=True)               # default 'eng'

**Scanned DOCX**

.. code-block:: python

   import wizarddocx as wd
   docx_txt = wd.extract_text("contract_scanned.docx", ocr=True, language_ocr="ita")

Returns
=======

``str`` — concatenated Unicode text from the selected pages.

Errors
======

- Bytes without ``extension`` → validation error.
- Unsupported or invalid input → domain-specific error.
- Missing or unreadable file → I/O error.


License
=======

`AGPL-3.0-or-later <_static/LICENSE>`_.

Resources
=========

- `PyPI Package <https://pypi.org/project/wizarddocx/>`_
- `Documentation <https://wizarddocx.readthedocs.io/en/latest/>`_
- `GitHub Repository <https://github.com/textwizard-dev/wizarddocx>`_

.. _contact_author:

Contact & Author
================

:Author: Mattia Rubino
:Email: `textwizard.dev@gmail.com <mailto:textwizard.dev@gmail.com>`_