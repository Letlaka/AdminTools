from __future__ import annotations

import re
from typing import Any


FORMULA_PREFIX = re.compile(r"^\s*[=+\-@]")
LEADING_CONTROL_PREFIXES = ("\t", "\r", "\n")


def safe_cell_value(value: Any) -> Any:
    if isinstance(value, str) and (
        FORMULA_PREFIX.match(value) or value.startswith(LEADING_CONTROL_PREFIXES)
    ):
        return "'" + value
    return value