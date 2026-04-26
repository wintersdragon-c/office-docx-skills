#!/usr/bin/env python3
from __future__ import annotations


def has_chinese(text: str) -> bool:
    return any("\u4e00" <= char <= "\u9fff" for char in text)
