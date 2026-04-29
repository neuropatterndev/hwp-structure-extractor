# MCP 설정 가이드

## 개념

MCP server는 Claude가 외부 도구를 호출할 수 있게 해주는 연결 계층입니다. 이 저장소의 MCP server는 HWP parser를 다음 tool로 노출합니다.

- `parse_hwp_to_json`
- `summarize_hwp_result`
- `extract_hwp_plain_text`
- `list_hwp_tables`
- `list_hwp_media`
- `get_hwp_block`
- `start_hwp_verification_ui`

## 설치

```bash
pip install -e ".[mcp,pyhwp]"
```

## Claude Code 등록

```bash
claude mcp add --transport stdio hwp-parser -- hwp-parser-mcp
```

프로젝트 공유용:

```bash
claude mcp add --scope project --transport stdio hwp-parser -- hwp-parser-mcp
```

상태 확인:

```bash
claude mcp list
claude mcp get hwp-parser
```

Claude Code 내부에서는 다음 명령으로 MCP 연결 상태를 확인할 수 있습니다.

```text
/mcp
```

## 수동 `.mcp.json` 예시

```json
{
  "mcpServers": {
    "hwp-parser": {
      "type": "stdio",
      "command": "python",
      "args": ["-m", "hwp_full_parser.mcp_server"],
      "env": {
        "HWP_PARSER_ALLOWED_ROOT": "."
      }
    }
  }
}
```

## 보안 권장사항

MCP server는 로컬 파일을 읽을 수 있습니다. 다음을 권장합니다.

```bash
export HWP_PARSER_ALLOWED_ROOT=/absolute/path/to/your/workspace
```

이 값을 설정하면 parser는 해당 root 바깥의 파일을 읽거나 출력하지 않습니다.

## 사용 예시

Claude에게 다음처럼 요청할 수 있습니다.

```text
이 HWP 파일을 parse_hwp_to_json으로 파싱하고, 표/이미지/캡션 추출 상태를 요약해줘: ./samples/report.hwp
```

```text
방금 생성한 result.json에서 plain text만 추출해서 RAG 입력용으로 정리해줘.
```

```text
order=12 block을 열어서 표인지 이미지인지 확인해줘.
```
