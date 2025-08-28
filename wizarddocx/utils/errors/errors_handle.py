# SPDX-FileCopyrightText: 2024–2025 Mattia Rubino
# SPDX-License-Identifier: AGPL-3.0-or-later

import traceback
import logging
from functools import wraps
from wizarddocx.utils.errors.errors import TextWizardError

logger = logging.getLogger(__name__)

def handle_errors(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except TextWizardError:
            raise
        except Exception as e:
            original_traceback = "".join(
                traceback.format_exception(type(e), e, e.__traceback__)
            )
            logger.error(f"Unexpected error:\n{original_traceback}")
            raise
    return wrapper
