# SPDX-FileCopyrightText: 2024–2025 Mattia Rubino
# SPDX-License-Identifier: AGPL-3.0-or-later

from .wizard_extract import WizardDocx

_wizard = WizardDocx()

extract_text       = _wizard.extract_text


__all__ = [
    "WizardDocx",
    "extract_text",

]
