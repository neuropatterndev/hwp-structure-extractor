from __future__ import annotations


def test_import_package() -> None:
    import hwp_full_parser

    assert hwp_full_parser.__version__ == "0.1.0"


def test_core_symbols_available() -> None:
    from hwp_full_parser.core import FullHwpParser, summarize, write_json

    assert FullHwpParser is not None
    assert summarize is not None
    assert write_json is not None


def test_api_symbols_available() -> None:
    from hwp_full_parser.api import parse_hwp_document, extract_plain_text_from_json

    assert parse_hwp_document is not None
    assert extract_plain_text_from_json is not None
