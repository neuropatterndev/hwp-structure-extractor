"""Stable programmatic API for the HWP parser.

The original parser is preserved in :mod:`hwp_full_parser.core`. This module is
an intentionally thin compatibility layer used by both the MCP server and Claude
Skill scripts. Keep the heavy parsing heuristics in ``core.py``; add integration
logic here.
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path
from typing import Any, Literal

from .core import FullHwpParser, ParsedDocument, start_web_ui, summarize, write_json

ParseMode = Literal["auto", "pyhwp", "binary", "xml"]


def parse_hwp_document(
    input_path: str | Path,
    output_dir: str | Path | None = None,
    mode: ParseMode = "auto",
    keep_intermediate: bool = False,
    hwp5proc_path: str = "hwp5proc",
    json_name: str = "result.json",
    compact: bool = False,
) -> dict[str, Any]:
    """Parse a HWP/XML document and write the structured JSON result.

    Returns a compact integration response rather than the full document JSON.
    This avoids excessive MCP tool output while preserving the full evidence in
    ``result_json_path``.
    """
    input_path = Path(input_path).expanduser().resolve()

    if output_dir is None:
        output_dir = Path(tempfile.mkdtemp(prefix="hwp_parsed_"))
    else:
        output_dir = Path(output_dir).expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

    doc: ParsedDocument = FullHwpParser(
        input_path=input_path,
        output_dir=output_dir,
        mode=mode,
        keep_intermediate=keep_intermediate,
        hwp5proc_path=hwp5proc_path,
    ).parse()

    json_path = write_json(doc, output_dir / json_name, pretty=not compact)
    summary = summarize(doc)

    return {
        "input_path": str(input_path),
        "output_dir": str(output_dir),
        "result_json_path": str(json_path),
        "method": doc.method,
        "summary": summary,
        "warnings": doc.warnings,
        "errors": doc.errors,
    }


def load_parsed_json(json_path: str | Path) -> dict[str, Any]:
    """Load a parser ``result.json`` file."""
    p = Path(json_path).expanduser().resolve()
    return json.loads(p.read_text(encoding="utf-8"))


def _iter_blocks(blocks: list[dict[str, Any]]):
    for block in blocks or []:
        if not isinstance(block, dict):
            continue
        yield block
        for child in block.get("children", []) or []:
            if isinstance(child, dict):
                yield from _iter_blocks([child])
        if block.get("type") == "table":
            for row in block.get("rows", []) or []:
                for cell in row or []:
                    if isinstance(cell, dict):
                        for item in cell.get("content", []) or []:
                            if isinstance(item, dict):
                                yield from _iter_blocks([item])


def compute_json_summary(json_path: str | Path) -> dict[str, Any]:
    """Compute a compact summary from a parser JSON file."""
    data = load_parsed_json(json_path)
    counts: dict[str, int] = {}
    caption_count = 0
    table_cells = 0
    image_linked = 0

    for block in _iter_blocks(data.get("blocks", []) or []):
        btype = str(block.get("type", "unknown"))
        counts[btype] = counts.get(btype, 0) + 1
        if block.get("caption"):
            caption_count += 1
        if btype == "image" and block.get("image_path"):
            image_linked += 1
        if btype == "table":
            for row in block.get("rows", []) or []:
                table_cells += len(row or [])

    return {
        "source_path": data.get("source_path"),
        "method": data.get("method"),
        "block_counts": counts,
        "caption_count": caption_count,
        "table_cell_count": table_cells,
        "media_count": len(data.get("media_files", []) or []),
        "image_linked_count": image_linked,
        "warning_count": len(data.get("warnings", []) or []),
        "error_count": len(data.get("errors", []) or []),
        "warnings": data.get("warnings", []) or [],
        "errors": data.get("errors", []) or [],
    }


def extract_plain_text_from_json(json_path: str | Path, include_captions: bool = True) -> str:
    """Extract LLM/RAG-friendly plain text from a parser result JSON."""
    data = load_parsed_json(json_path)
    lines: list[str] = []

    for block in _iter_blocks(data.get("blocks", []) or []):
        btype = block.get("type")
        if btype == "paragraph" and block.get("text"):
            lines.append(str(block["text"]).strip())
        elif btype == "caption" and block.get("text") and include_captions:
            lines.append(f"[CAPTION] {str(block['text']).strip()}")
        elif btype == "image" and include_captions:
            caption = block.get("caption") or {}
            if isinstance(caption, dict) and caption.get("text"):
                lines.append(f"[IMAGE CAPTION] {str(caption['text']).strip()}")
        elif btype == "table":
            caption = block.get("caption") or {}
            if include_captions and isinstance(caption, dict) and caption.get("text"):
                lines.append(f"[TABLE CAPTION] {str(caption['text']).strip()}")
            for row in block.get("rows", []) or []:
                row_texts: list[str] = []
                for cell in row or []:
                    if isinstance(cell, dict) and cell.get("text"):
                        row_texts.append(str(cell["text"]).strip())
                if row_texts:
                    lines.append(" | ".join(row_texts))

    return "\n\n".join(x for x in lines if x).strip()


def extract_tables_from_json(json_path: str | Path) -> list[dict[str, Any]]:
    """Return parsed top-level and inline table blocks from a parser JSON file."""
    data = load_parsed_json(json_path)
    tables = [block for block in _iter_blocks(data.get("blocks", []) or []) if block.get("type") == "table"]
    return tables


def list_media_from_json(json_path: str | Path) -> list[dict[str, Any]]:
    """Return extracted media metadata from a parser JSON file."""
    data = load_parsed_json(json_path)
    return list(data.get("media_files", []) or [])


def get_block_by_order(json_path: str | Path, order: int) -> dict[str, Any]:
    """Return a parsed block by its ``order`` field."""
    data = load_parsed_json(json_path)
    for block in _iter_blocks(data.get("blocks", []) or []):
        if block.get("order") == order:
            return block
    raise KeyError(f"No parsed block with order={order}")


def run_verification_ui(
    input_path: str | Path | None = None,
    output_dir: str | Path = "hwp_parsed_output",
    mode: ParseMode = "auto",
    keep_intermediate: bool = True,
    hwp5proc_path: str = "hwp5proc",
    host: str = "127.0.0.1",
    port: int = 7860,
    open_browser: bool = True,
) -> int:
    """Start the built-in parser verification web UI."""
    return start_web_ui(
        input_path=input_path,
        output_dir=Path(output_dir).expanduser().resolve(),
        mode=mode,
        keep_intermediate=keep_intermediate,
        hwp5proc_path=hwp5proc_path,
        host=host,
        port=port,
        open_browser=open_browser,
    )
