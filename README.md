# HWP Full Parser

HWP Full Parser는 `.hwp` 문서를 문서 AI 검증 워크플로우에서 사용할 수 있도록 구조화된 JSON으로 추출하는 Python 패키지입니다. 기존 단일 파일 파서의 trial-and-error 로직을 `src/hwp_full_parser/core.py`에 그대로 보존하고, 그 위에 CLI, Python API, Claude MCP Server, Claude Skill을 얇은 wrapper로 추가했습니다.

## 핵심 기능

- HWP 문서에서 문단, 표, 셀, 이미지, 캡션, media stream 추출
- `pyhwp/hwp5proc` 기반 XML 추출 경로 우선 사용
- `olefile` 기반 OLE/Binary fallback 경로 제공
- `PrvImage`, `BinData`, caption-image hint, table/image placeholder 진단 유지
- parser 결과를 `result.json`으로 저장
- 내장 웹 검증 UI 제공
- Claude Code / Claude Desktop용 MCP server 제공
- Claude Skill 제공

## 저장소 구조

```text
hwp-full-parser/
├── src/hwp_full_parser/
│   ├── core.py          # 원본 HWP parser 로직 보존
│   ├── api.py           # MCP/Skill 공통 API wrapper
│   ├── cli.py           # hwp-parse CLI entrypoint
│   └── mcp_server.py    # Claude MCP server
├── legacy/
│   └── hwp_full_parser_v30_original.py
├── .claude/skills/hwp-parser/
│   ├── SKILL.md
│   ├── scripts/parse_hwp.py
│   └── references/
├── docs/
│   ├── GITHUB_UPLOAD_GUIDE.md
│   ├── MCP_SETUP.md
│   └── SKILL_SETUP.md
├── examples/
│   └── expected_output_schema.json
└── tests/
```

## 설치

```bash
git clone https://github.com/YOUR_ID/hwp-full-parser.git
cd hwp-full-parser
python -m venv .venv
```

macOS/Linux:

```bash
source .venv/bin/activate
```

Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

기본 설치:

```bash
pip install -e .
```

MCP까지 사용할 경우:

```bash
pip install -e ".[mcp]"
```

`pyhwp/hwp5proc` 기반 XML 변환까지 사용할 경우:

```bash
pip install -e ".[pyhwp]"
```

전체 개발 환경:

```bash
pip install -e ".[mcp,pyhwp,dev]"
```

## CLI 사용법

문서 파싱:

```bash
hwp-parse sample.hwp --output-dir parsed --print-summary
```

중간 XML 보존:

```bash
hwp-parse sample.hwp --output-dir parsed --keep-intermediate
```

Binary fallback만 강제:

```bash
hwp-parse sample.hwp --output-dir parsed --mode binary
```

웹 검증 UI 실행:

```bash
hwp-parse sample.hwp --web --output-dir parsed
```

브라우저 자동 실행 없이 localhost 서버만 실행:

```bash
hwp-parse sample.hwp --web --no-browser --host 127.0.0.1 --port 7860
```

## Python API 사용법

```python
from hwp_full_parser.api import parse_hwp_document, extract_plain_text_from_json

result = parse_hwp_document(
    input_path="sample.hwp",
    output_dir="parsed",
    mode="auto",
    keep_intermediate=True,
)

print(result["summary"])

plain_text = extract_plain_text_from_json(result["result_json_path"])
print(plain_text[:1000])
```

## MCP 사용법

Claude Code에 local stdio MCP server로 등록합니다.

```bash
claude mcp add --transport stdio hwp-parser -- hwp-parser-mcp
```

프로젝트 단위로 공유하려면:

```bash
claude mcp add --scope project --transport stdio hwp-parser -- hwp-parser-mcp
```

수동 설정 예시는 `.mcp.json.example`에 포함되어 있습니다.

제공 MCP tools:

| Tool | 역할 |
|---|---|
| `parse_hwp_to_json` | `.hwp` 또는 중간 `.xml`을 구조화 JSON으로 파싱 |
| `summarize_hwp_result` | `result.json`의 block/media/caption 통계 요약 |
| `extract_hwp_plain_text` | LLM/RAG용 plain text 추출 |
| `list_hwp_tables` | 파싱된 table block 목록 반환 |
| `list_hwp_media` | 추출 media metadata 반환 |
| `get_hwp_block` | 특정 `order` block 조회 |
| `start_hwp_verification_ui` | 내장 검증 UI를 detached process로 실행 |

보안을 위해 다음 환경변수를 설정하는 것을 권장합니다.

```bash
export HWP_PARSER_ALLOWED_ROOT=/absolute/path/to/your/workspace
```

이 값을 설정하면 MCP server는 해당 root 바깥의 파일을 읽거나 출력하지 않습니다.

## Claude Skill 사용법

이 저장소는 Claude Code project skill을 포함합니다.

```text
.claude/skills/hwp-parser/SKILL.md
```

Claude Code에서 자연어로 요청하거나, 명시적으로 skill을 호출합니다.

```text
/hwp-parser sample.hwp 파일을 파싱하고 표, 이미지, 캡션 추출 상태를 요약해줘.
```

Skill은 MCP server가 있으면 MCP tool 사용을 우선하고, 없으면 `scripts/parse_hwp.py`를 CLI fallback으로 사용하도록 설계되어 있습니다.

## 출력 JSON 개요

기본 출력은 다음 위치에 생성됩니다.

```text
parsed/result.json
parsed/media/
parsed/intermediate.xml  # --keep-intermediate 사용 시
```

`result.json`의 상위 구조:

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

상세 schema 예시는 `examples/expected_output_schema.json`을 참고하십시오.

## 설계 원칙

이 패키지는 원본 파서의 trial-and-error 로직을 보존하는 것을 최우선으로 합니다. 따라서 핵심 파서 로직은 `core.py`에 그대로 유지하고, MCP/Skill/CLI/API는 모두 wrapper로만 동작합니다. 향후 refactoring을 하더라도 먼저 regression sample set을 만든 뒤, `core.py`의 기능 동등성을 검증한 후 단계적으로 분리하는 것이 안전합니다.

## 주의 사항

- 이 파서는 HWP layout renderer가 아닙니다. 문서 내부 구조를 추출하고 검증하기 위한 parser입니다.
- 전체 페이지 렌더링은 별도의 HWP rendering engine이 필요합니다.
- 비공개 HWP 문서, 과제 자료, 개인정보 포함 파일, 추출 media는 GitHub에 올리지 마십시오.
- `LICENSE`는 기본적으로 all-rights-reserved placeholder입니다. 공개 오픈소스로 배포하려면 의도에 맞는 라이선스로 교체하십시오.
