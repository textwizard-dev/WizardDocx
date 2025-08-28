# SPDX-FileCopyrightText: 2024–2025 Mattia Rubino
# SPDX-License-Identifier: AGPL-3.0-or-later

from __future__ import annotations
import re
from typing import Sequence, Union, List, Set

_rng = re.compile(r"^\s*(\d+)\s*-\s*(\d+)\s*$")

def normalize_pages_selector(sel: Sequence[Union[int, str]] | None, total: int) -> List[int]:
    if sel is None:
        return list(range(max(total, 0)))
    got: Set[int] = set()

    def add1(p: int) -> None:
        if 1 <= p <= total:
            got.add(p - 1)

    for item in sel:
        if isinstance(item, int):
            add1(item)
        elif isinstance(item, str):
            for chunk in item.split(","):
                s = chunk.strip()
                if not s:
                    continue
                if s.isdigit():
                    add1(int(s))
                    continue
                m = _rng.match(s)
                if m:
                    a, b = int(m.group(1)), int(m.group(2))
                    lo, hi = (a, b) if a <= b else (b, a)
                    for p in range(lo, hi + 1):
                        add1(p)
    return sorted(got)

def normalize_sheets_selector(sel: Sequence[Union[int, str]] | None, sheet_names: List[str]) -> List[int]:
    if sel is None:
        return list(range(len(sheet_names)))
    name2idx = {n: i for i, n in enumerate(sheet_names)}
    got: Set[int] = set()
    for item in sel:
        if isinstance(item, int):
            if 0 <= item < len(sheet_names):
                got.add(item)
        elif isinstance(item, str):
            if item in name2idx:  
                got.add(name2idx[item])
            else:
                for chunk in item.split(","):  
                    s = chunk.strip()
                    if not s:
                        continue
                    if s.isdigit():
                        i = int(s)
                        if 0 <= i < len(sheet_names):
                            got.add(i)
                        continue
                    m = _rng.match(s)
                    if m:
                        a, b = int(m.group(1)), int(m.group(2))
                        lo, hi = (a, b) if a <= b else (b, a)
                        for i in range(lo, hi + 1):
                            if 0 <= i < len(sheet_names):
                                got.add(i)
    return sorted(got)
