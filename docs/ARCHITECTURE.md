# Architecture

## Goal

This repository preserves the original single-file HWP parser while making it usable as:

1. a Python package,
2. a CLI tool,
3. a Claude MCP server,
4. a Claude Skill-backed workflow.

## Layering

```text
Claude / user
   │
   ├─ CLI: hwp-parse
   │
   ├─ MCP: hwp-parser-mcp
   │
   ├─ Skill: .claude/skills/hwp-parser/SKILL.md
   │
   ▼
src/hwp_full_parser/api.py
   │
   ▼
src/hwp_full_parser/core.py
```

## `core.py`

`core.py` contains the original parser logic. It intentionally remains monolithic because the versioned patches and trial-and-error corrections are order-sensitive.

Do not aggressively refactor `core.py` until you have a regression corpus.

## `api.py`

`api.py` is the stable integration boundary.

Main functions:

- `parse_hwp_document`
- `load_parsed_json`
- `compute_json_summary`
- `extract_plain_text_from_json`
- `extract_tables_from_json`
- `list_media_from_json`
- `get_block_by_order`
- `run_verification_ui`

## `mcp_server.py`

`mcp_server.py` exposes safe, compact parser operations as MCP tools.

It does not return the full `result.json` by default because tool output can become too large. It returns the path to the full JSON and exposes additional inspection tools.

## Claude Skill

The Skill is not a parser implementation. It is an operating procedure that tells Claude how to use the parser correctly:

- prefer MCP tools when available,
- fall back to script execution when MCP is unavailable,
- choose parser modes conservatively,
- summarize output without pasting huge JSON,
- use the web UI only for visual verification.

## Extension points

Future work can add:

- regression fixtures,
- `.hwpx` support,
- additional plain-text chunking for RAG,
- CSV/Markdown table export,
- batch parser CLI,
- Streamlit/Gradio wrapper,
- packaged Claude Desktop extension.
