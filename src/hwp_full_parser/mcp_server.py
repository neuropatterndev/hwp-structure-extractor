"""MCP server exposing the HWP parser as Claude-callable tools."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from typing import Any, Literal

try:  # MCP Python SDK
    from mcp.server.fastmcp import FastMCP
except Exception:  # pragma: no cover - optional fallback for standalone fastmcp
    try:
        from fastmcp import FastMCP  # type: ignore
    except Exception as exc:  # pragma: no cover
        raise RuntimeError(
            "MCP support is not installed. Install with: pip install -e '.[mcp]'"
        ) from exc

from .api import (
    compute_json_summary,
    extract_plain_text_from_json,
    extract_tables_from_json,
    get_block_by_order,
    list_media_from_json,
    parse_hwp_document,
)

ParseMode = Literal["auto", "pyhwp", "binary", "xml"]

mcp = FastMCP("hwp-parser")


def _allowed_root() -> Path | None:
    value = os.environ.get("HWP_PARSER_ALLOWED_ROOT", "").strip()
    if not value:
        return None
    return Path(value).expanduser().resolve()


def _resolve_existing_path(path: str | Path) -> Path:
    resolved = Path(path).expanduser().resolve()
    root = _allowed_root()
    if root is not None:
        try:
            resolved.relative_to(root)
        except ValueError as exc:
            raise PermissionError(f"Path is outside HWP_PARSER_ALLOWED_ROOT: {resolved}") from exc
    return resolved


def _resolve_output_dir(path: str | Path) -> Path:
    resolved = Path(path).expanduser().resolve()
    root = _allowed_root()
    if root is not None:
        try:
            resolved.relative_to(root)
        except ValueError as exc:
            raise PermissionError(f"Output directory is outside HWP_PARSER_ALLOWED_ROOT: {resolved}") from exc
    resolved.mkdir(parents=True, exist_ok=True)
    return resolved


@mcp.tool()
def parse_hwp_to_json(
    input_path: str,
    output_dir: str = "hwp_parsed_output",
    mode: ParseMode = "auto",
    keep_intermediate: bool = False,
    hwp5proc_path: str = "hwp5proc",
    json_name: str = "result.json",
) -> dict[str, Any]:
    """Parse a .hwp or intermediate .xml file into structured JSON.

    Use this when the user asks to extract HWP paragraphs, tables, images,
    captions, media files, or document-AI verification evidence.
    """
    safe_input = _resolve_existing_path(input_path)
    safe_output = _resolve_output_dir(output_dir)
    return parse_hwp_document(
        input_path=safe_input,
        output_dir=safe_output,
        mode=mode,
        keep_intermediate=keep_intermediate,
        hwp5proc_path=hwp5proc_path,
        json_name=json_name,
    )


@mcp.tool()
def summarize_hwp_result(json_path: str) -> dict[str, Any]:
    """Read a parser result JSON and return compact extraction statistics."""
    safe_json = _resolve_existing_path(json_path)
    return compute_json_summary(safe_json)


@mcp.tool()
def extract_hwp_plain_text(json_path: str, include_captions: bool = True) -> str:
    """Extract LLM/RAG-friendly plain text from parser result JSON."""
    safe_json = _resolve_existing_path(json_path)
    return extract_plain_text_from_json(safe_json, include_captions=include_captions)


@mcp.tool()
def list_hwp_tables(json_path: str, max_tables: int = 20) -> list[dict[str, Any]]:
    """Return parsed table blocks from result JSON, capped for tool-output safety."""
    safe_json = _resolve_existing_path(json_path)
    tables = extract_tables_from_json(safe_json)
    return tables[: max(0, max_tables)]


@mcp.tool()
def list_hwp_media(json_path: str, max_items: int = 100) -> list[dict[str, Any]]:
    """Return extracted media metadata from result JSON."""
    safe_json = _resolve_existing_path(json_path)
    media = list_media_from_json(safe_json)
    return media[: max(0, max_items)]


@mcp.tool()
def get_hwp_block(json_path: str, order: int) -> dict[str, Any]:
    """Return one parsed block by its order number."""
    safe_json = _resolve_existing_path(json_path)
    return get_block_by_order(safe_json, order=order)


@mcp.tool()
def start_hwp_verification_ui(
    input_path: str = "",
    output_dir: str = "hwp_parsed_output",
    mode: ParseMode = "auto",
    host: str = "127.0.0.1",
    port: int = 7860,
) -> dict[str, str]:
    """Start the original built-in verification web UI in a detached process.

    Prefer parse_hwp_to_json for automated workflows. This tool is for manual
    visual inspection of extraction quality.
    """
    safe_output = _resolve_output_dir(output_dir)
    cmd = [
        sys.executable,
        "-m",
        "hwp_full_parser.cli",
        "--web",
        "--output-dir",
        str(safe_output),
        "--mode",
        mode,
        "--host",
        host,
        "--port",
        str(port),
    ]
    if input_path:
        safe_input = _resolve_existing_path(input_path)
        cmd.insert(3, str(safe_input))

    subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
    )
    return {"url": f"http://{host}:{port}/", "note": "HWP verification UI started."}


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
