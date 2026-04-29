# Claude Skill 설정 가이드

## 개념

Skill은 Claude에게 특정 작업 절차를 알려주는 패키지입니다. 이 저장소의 Skill은 HWP parser를 사용할 때 다음 판단을 돕습니다.

- `auto`, `pyhwp`, `binary`, `xml` 모드 선택
- MCP server 사용 우선순위
- MCP가 없을 때 CLI/script fallback
- `result.json` 요약 방식
- 웹 검증 UI를 언제 사용할지

## 위치

Project skill 위치:

```text
.claude/skills/hwp-parser/SKILL.md
```

Claude Code는 프로젝트 루트에서 이 파일을 인식합니다.

## 사용 예시

```text
/hwp-parser sample.hwp 파일을 파싱하고 result.json과 media 추출 상태를 요약해줘.
```

또는 자연어로:

```text
이 HWP 문서에서 표와 그림 캡션을 추출해줘. 필요하면 검증 UI를 실행해줘.
```

## Skill fallback script

MCP가 없거나 연결되지 않은 경우 Skill은 다음 script를 사용할 수 있습니다.

```text
.claude/skills/hwp-parser/scripts/parse_hwp.py
```

직접 실행:

```bash
python .claude/skills/hwp-parser/scripts/parse_hwp.py sample.hwp --output-dir parsed --mode auto
```

## MCP와 Skill의 역할 분담

| 구성 | 역할 |
|---|---|
| MCP | 실제 parser 함수를 Claude tool로 호출 |
| Skill | parser 사용 절차, 모드 선택, 결과 해석 기준 제공 |

가장 안정적인 운용 방식은 둘을 함께 사용하는 것입니다.
