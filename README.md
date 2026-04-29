# HWP Structure Extractor

Structured HWP extraction for document-AI, LLM/RAG preprocessing, and verification workflows.

This project extracts text blocks, tables, table cells, images, captions, embedded media, and diagnostic metadata from Korean HWP documents. It is designed for document understanding, dataset construction, extraction verification, and LLM/RAG preprocessing rather than visual page rendering.

## Features

- Parse `.hwp` documents into structured JSON
- Extract paragraphs, tables, table cells, images, captions, and embedded media
- Preserve document order as much as possible through ordered block objects
- Detect and recover image streams from HWP `BinData`
- Handle HWP preview image (`PrvImage`) when available
- Use `pyhwp/hwp5proc` XML extraction when available
- Fall back to direct OLE/Binary parsing when XML extraction is unavailable
- Provide table/image/caption diagnostics for verification workflows
- Export media files into a dedicated output directory
- Provide a built-in local web verification UI
- Provide Python API, CLI, Claude MCP server, and Claude Skill support

## What this parser is for

HWP Structure Extractor is useful when `.hwp` files need to be converted into machine-readable structures for downstream processing.

Typical use cases include:

- Building document-AI datasets from HWP files
- Extracting paragraphs and tables for NLP pipelines
- Preparing HWP documents for LLM or RAG systems
- Verifying whether captions, images, and tables were extracted correctly
- Inspecting embedded HWP media streams
- Creating structured JSON from Korean administrative, research, or report-style documents
- Connecting HWP parsing functionality to Claude Code or Claude Desktop through MCP

## What this parser is not

This project is **not** a full HWP layout renderer.

It does not aim to reproduce exact visual page layout, typography, pagination, line wrapping, or complete rendering fidelity. For exact visual rendering, a dedicated HWP rendering engine is required.

The built-in web UI is a verification interface for extracted structure, not a full replacement for Hancom Office or a browser-based HWP viewer.

## Extraction strategy

The parser uses a layered extraction strategy.

| Mode | Description |
|---|---|
| `auto` | Try XML extraction first when available, then use binary fallback if needed |
| `pyhwp` | Use `hwp5proc xml --embedbin` and parse the generated XML |
| `binary` | Read HWP OLE/CFB streams directly and extract records/media |
| `xml` | Parse an already generated intermediate XML file |

Recommended default:

```bash
hwp-parse sample.hwp --mode auto
```

If `pyhwp/hwp5proc` is not installed, use binary mode or install the optional `pyhwp` extra after reviewing the dependency notice below.

## Installation

Clone the repository:

```bash
git clone https://github.com/<your-username>/hwp-structure-extractor.git
cd hwp-structure-extractor
```

Create a virtual environment:

```bash
python -m venv .venv
```

Activate it.

macOS/Linux:

```bash
source .venv/bin/activate
```

Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Install the package:

```bash
pip install -e .
```

Install with MCP support:

```bash
pip install -e ".[mcp]"
```

Install with optional `pyhwp/hwp5proc` support:

```bash
pip install -e ".[pyhwp]"
```

Install all optional development dependencies:

```bash
pip install -e ".[mcp,pyhwp,dev]"
```

## Requirements

Core dependencies:

- Python 3.10+
- `olefile`
- `lxml`

Optional dependencies:

- `pyhwp` for XML-based extraction through `hwp5proc`
- `mcp` for Claude MCP server integration
- `pytest`, `ruff`, and `build` for development workflows

The core parser can run without `pyhwp` by using the binary fallback path. Install `pyhwp` only when XML-based extraction is needed and its AGPL license terms are acceptable for your use case.

## Quick start

Parse an HWP document:

```bash
hwp-parse sample.hwp --output-dir parsed --print-summary
```

Output directory:

```text
parsed/
├── result.json
├── media/
└── intermediate.xml   # only when intermediate XML is preserved
```

Launch the verification UI:

```bash
hwp-parse sample.hwp --web --output-dir parsed
```

The UI opens a local web interface that shows:

- internal first-page preview when available
- extracted block list
- paragraphs
- tables
- images
- captions
- warnings and errors
- extracted media gallery
- raw diagnostic metadata

## CLI usage

### Parse with default mode

```bash
hwp-parse sample.hwp --output-dir parsed
```

### Print extraction summary

```bash
hwp-parse sample.hwp --output-dir parsed --print-summary
```

### Keep intermediate XML

```bash
hwp-parse sample.hwp --output-dir parsed --keep-intermediate
```

### Force XML extraction

```bash
hwp-parse sample.hwp --output-dir parsed --mode pyhwp
```

### Force binary fallback

```bash
hwp-parse sample.hwp --output-dir parsed --mode binary
```

### Parse an intermediate XML file

```bash
hwp-parse intermediate.xml --output-dir parsed --mode xml
```

### Start verification UI

```bash
hwp-parse sample.hwp --web --output-dir parsed
```

### Start UI without opening browser automatically

```bash
hwp-parse sample.hwp --web --no-browser --host 127.0.0.1 --port 7860
```

## Python API

```python
from hwp_full_parser.api import (
    parse_hwp_document,
    extract_plain_text_from_json,
    summarize_result_json,
)

result = parse_hwp_document(
    input_path="sample.hwp",
    output_dir="parsed",
    mode="auto",
    keep_intermediate=True,
)

print(result["summary"])
print(result["result_json_path"])

plain_text = extract_plain_text_from_json(result["result_json_path"])
print(plain_text[:1000])

summary = summarize_result_json(result["result_json_path"])
print(summary)
```

## Output JSON

The main output file is:

```text
parsed/result.json
```

Top-level structure:

```json
{
  "source_path": "sample.hwp",
  "method": "pyhwp_xml+ole_media",
  "metadata": {},
  "warnings": [],
  "errors": [],
  "media_files": [],
  "blocks": []
}
```

Each item in `blocks` represents a document-level structure such as:

- paragraph
- table
- image
- caption
- placeholder or diagnostic block when recovery is partial

Example paragraph block:

```json
{
  "type": "paragraph",
  "id": "p-0001",
  "order": 1,
  "section": 0,
  "text": "Example paragraph text"
}
```

Example table block:

```json
{
  "type": "table",
  "id": "tbl-0001",
  "order": 2,
  "rows": [
    [
      {
        "text": "Cell text",
        "row": 0,
        "col": 0,
        "row_span": 1,
        "col_span": 1
      }
    ]
  ],
  "caption": {
    "text": "표 1. Example table caption",
    "method": "pattern"
  }
}
```

Example image block:

```json
{
  "type": "image",
  "id": "img-0001",
  "order": 3,
  "image_path": "parsed/media/BIN0001.png",
  "media_type": "image/png",
  "caption": {
    "text": "그림 1. Example figure caption",
    "method": "pattern"
  }
}
```

See:

```text
examples/expected_output_schema.json
```

for a more complete schema example.

## Media extraction

Embedded media files are written to:

```text
parsed/media/
```

The parser attempts to detect common media formats by binary signature, including:

- PNG
- JPEG
- GIF
- BMP
- TIFF
- WEBP
- WMF/EMF
- SVG
- PDF
- ZIP-like embedded payloads

Extracted media metadata is also recorded in `result.json` under `media_files`.

## Verification UI

The verification UI is intended for checking extraction quality.

Run:

```bash
hwp-parse sample.hwp --web --output-dir parsed
```

The UI is useful for answering questions such as:

- Were all paragraphs extracted?
- Did table cells preserve row/column structure?
- Were image files recovered?
- Were captions attached to the correct table or image?
- Are there unlinked media files?
- Did the parser produce warnings or partial-recovery diagnostics?

By default, the UI runs locally.

```text
http://127.0.0.1:7860
```

## Claude MCP integration

This repository includes a Claude MCP server.

Register it with Claude Code:

```bash
claude mcp add --transport stdio hwp-parser -- hwp-parser-mcp
```

Project-scoped registration:

```bash
claude mcp add --scope project --transport stdio hwp-parser -- hwp-parser-mcp
```

Provided MCP tools:

| Tool | Description |
|---|---|
| `parse_hwp_to_json` | Parse `.hwp` or intermediate `.xml` into structured JSON |
| `summarize_hwp_result` | Summarize parsed block/media/caption statistics |
| `extract_hwp_plain_text` | Extract plain text for LLM/RAG workflows |
| `list_hwp_tables` | List parsed table blocks |
| `list_hwp_media` | List extracted media metadata |
| `get_hwp_block` | Retrieve a specific block by `order` |
| `start_hwp_verification_ui` | Start the local verification UI |

Security recommendation:

```bash
export HWP_PARSER_ALLOWED_ROOT=/absolute/path/to/workspace
```

When this environment variable is set, the MCP server restricts file access to the specified workspace root.

## Claude Skill integration

This repository includes a Claude Code project skill:

```text
.claude/skills/hwp-parser/SKILL.md
```

Use it in Claude Code with natural language:

```text
Parse sample.hwp and summarize extracted paragraphs, tables, images, and captions.
```

Or explicitly:

```text
/hwp-parser sample.hwp 파일을 파싱하고 추출 결과를 요약해줘.
```

The skill is designed to:

- choose an appropriate parser mode
- prefer MCP tools when available
- fall back to the local parser script when MCP is unavailable
- summarize extraction results without dumping the entire JSON
- guide inspection of tables, images, captions, and warnings

## Repository structure

```text
hwp-structure-extractor/
├── src/
│   └── hwp_full_parser/
│       ├── api.py
│       ├── cli.py
│       ├── core.py
│       └── mcp_server.py
├── .claude/
│   └── skills/
│       └── hwp-parser/
├── docs/
├── examples/
├── tests/
├── legacy/
├── pyproject.toml
├── requirements.txt
├── requirements-mcp.txt
└── README.md
```

## Testing

Run smoke tests:

```bash
pytest
```

The included tests verify importability and basic package wiring. For production use, it is recommended to maintain a private regression set of representative HWP files and compare extraction summaries across versions.

## Known limitations

HWP is a complex binary document format. This parser is designed for robust structured extraction, but some documents may still require manual inspection.

Known limitations:

- It does not provide exact visual page rendering.
- Some layout-specific information may be incomplete.
- Complex nested tables may require verification.
- Caption attachment can be document-dependent.
- Damaged or encrypted HWP files may not be parsable.
- Exact rendering of fonts, pagination, and object positioning is outside the project scope.
- Some image formats such as WMF/EMF may be extracted but not previewable in all environments.

## Troubleshooting

### `hwp5proc` is not found

Install optional `pyhwp` support only if its AGPL license terms are acceptable for your project:

```bash
pip install -e ".[pyhwp]"
```

Or use binary fallback:

```bash
hwp-parse sample.hwp --mode binary --output-dir parsed
```

### No images are shown in the UI

Check:

```text
parsed/media/
```

and inspect `media_files` in:

```text
parsed/result.json
```

Some files may be extracted as media but not previewable by the browser.

### Captions are not attached correctly

Run the verification UI:

```bash
hwp-parse sample.hwp --web --output-dir parsed
```

Then inspect nearby blocks, unlinked media, and raw diagnostics.

### Output JSON is too large for LLM use

Use plain text extraction:

```python
from hwp_full_parser.api import extract_plain_text_from_json

text = extract_plain_text_from_json("parsed/result.json")
```

Or use the MCP tools that return summaries, table lists, media lists, or individual blocks instead of the full JSON.

## Security and privacy

Do not commit private HWP documents, extracted media files, generated JSON files, or intermediate XML files to a public repository.

Recommended `.gitignore` targets include:

```text
parsed/
hwp_parsed_output/
*.hwp
*.hwpx
*.pdf
.env
```

When using MCP, restrict file access with:

```bash
export HWP_PARSER_ALLOWED_ROOT=/absolute/path/to/workspace
```

## Third-party dependencies and license notes

This project is licensed under the Apache License 2.0. Third-party dependencies remain under their own licenses.

| Dependency | Required? | Used for | License note |
|---|---:|---|---|
| `olefile` | Core | Reading OLE/CFB streams and extracting HWP `BinData` | BSD-2-Clause according to conda-forge metadata |
| `lxml` | Core | XML parsing for intermediate HWP XML output | BSD-3-Clause |
| `pyhwp` | Optional | `hwp5proc` XML extraction path | GNU AGPL v3 or later |
| `mcp` | Optional | Claude MCP server integration | MIT |
| `pytest`, `ruff`, `build` | Development only | Tests, linting, packaging | See each package's license metadata |

### Important note about `pyhwp`

`pyhwp` is an optional dependency used only for the XML extraction path through `hwp5proc`. The parser can still run without `pyhwp` by using the binary fallback mode.

`pyhwp` is licensed under the GNU Affero General Public License v3 or later. AGPL is a strong copyleft license with additional obligations for software made available over a network. If you install, modify, redistribute, bundle, or provide a service based on `pyhwp`, review the AGPL terms before distribution or deployment.

Relevant upstream information:

- pyhwp repository: https://github.com/mete0r/pyhwp
- pyhwp PyPI: https://pypi.org/project/pyhwp/
- GNU AGPL v3: https://www.gnu.org/licenses/agpl-3.0.en.html
- Apache License 2.0: https://www.apache.org/licenses/LICENSE-2.0

## License

This project is licensed under the Apache License 2.0. See the `LICENSE` file for the full license text.

This license applies to the source code in this repository. It does not modify the license terms of third-party dependencies. In particular, optional `pyhwp` support remains subject to `pyhwp`'s AGPL license terms when that optional path is installed or used.

SPDX-License-Identifier: Apache-2.0

## Citation

If this parser is used in a research workflow, report, dataset construction pipeline, or document-AI benchmark, please cite the repository URL and version or commit hash used for extraction.

Example:

```text
HWP Structure Extractor, version <commit-hash>, https://github.com/<your-username>/hwp-structure-extractor
```
