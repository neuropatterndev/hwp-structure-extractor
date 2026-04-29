"""Structured HWP parser package with CLI, API, MCP, and Claude Skill support."""

from __future__ import annotations

__version__ = "0.1.0"

__all__ = [
    "ParseMode",
    "parse_hwp_document",
    "load_parsed_json",
    "compute_json_summary",
    "extract_plain_text_from_json",
    "extract_tables_from_json",
    "list_media_from_json",
    "get_block_by_order",
    "run_verification_ui",
]


def __getattr__(name: str):
    if name in __all__:
        from . import api

        return getattr(api, name)
    raise AttributeError(name)
