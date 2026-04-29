# Sample Usage

```bash
hwp-parse sample.hwp --output-dir parsed --print-summary
```

```python
from hwp_full_parser.api import parse_hwp_document, compute_json_summary

result = parse_hwp_document("sample.hwp", output_dir="parsed")
summary = compute_json_summary(result["result_json_path"])
print(summary)
```

Claude Code example:

```text
/hwp-parser ./sample.hwp 를 파싱하고 표/이미지/캡션 추출 결과를 요약해줘.
```
