# Python API Reference

## `parse_hwp_document`

```python
from hwp_full_parser.api import parse_hwp_document

result = parse_hwp_document(
    input_path="sample.hwp",
    output_dir="parsed",
    mode="auto",
    keep_intermediate=True,
)
```

Returns:

```python
{
    "input_path": "...",
    "output_dir": "...",
    "result_json_path": ".../result.json",
    "method": "pyhwp_xml+ole_media",
    "summary": {...},
    "warnings": [...],
    "errors": [...],
}
```

## `compute_json_summary`

```python
from hwp_full_parser.api import compute_json_summary

summary = compute_json_summary("parsed/result.json")
```

## `extract_plain_text_from_json`

```python
from hwp_full_parser.api import extract_plain_text_from_json

text = extract_plain_text_from_json("parsed/result.json")
```

## `extract_tables_from_json`

```python
from hwp_full_parser.api import extract_tables_from_json

tables = extract_tables_from_json("parsed/result.json")
```

## `list_media_from_json`

```python
from hwp_full_parser.api import list_media_from_json

media = list_media_from_json("parsed/result.json")
```

## `get_block_by_order`

```python
from hwp_full_parser.api import get_block_by_order

block = get_block_by_order("parsed/result.json", order=12)
```

## `run_verification_ui`

```python
from hwp_full_parser.api import run_verification_ui

run_verification_ui("sample.hwp", output_dir="parsed", port=7860)
```
