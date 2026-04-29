# GitHub 업로드 가이드

이 문서는 HWP Full Parser 저장소를 GitHub에 올리는 절차를 설명합니다.

## 1. GitHub 저장소 생성

GitHub 웹에서 새 repository를 만듭니다.

권장 이름:

```text
hwp-full-parser
```

처음에는 `Public`보다 `Private`을 권장합니다. 샘플 HWP, 추출 이미지, 프로젝트 문서가 실수로 포함될 가능성이 있기 때문입니다.

## 2. 로컬에서 압축 해제

제공된 zip을 원하는 위치에 풉니다.

```bash
unzip hwp-full-parser-package.zip
cd hwp-full-parser
```

## 3. 개인정보 및 비공개 데이터 확인

다음 파일은 GitHub에 올리지 마십시오.

```text
*.hwp
*.hwpx
parsed/
hwp_parsed_output/
media/
result.json
intermediate.xml
```

이미 `.gitignore`에 포함되어 있지만, `git status`로 반드시 확인하십시오.

## 4. 작성자 정보 수정

`pyproject.toml`에서 다음 값을 수정합니다.

```toml
authors = [
  {name = "YOUR NAME", email = "YOUR_EMAIL@example.com"}
]

[project.urls]
Homepage = "https://github.com/YOUR_ID/hwp-full-parser"
Repository = "https://github.com/YOUR_ID/hwp-full-parser"
Issues = "https://github.com/YOUR_ID/hwp-full-parser/issues"
```

`LICENSE`의 `YOUR NAME`도 수정하십시오.

## 5. 라이선스 결정

현재 `LICENSE`는 all-rights-reserved placeholder입니다.

오픈소스로 배포하려면 다음 중 하나로 교체하십시오.

- MIT: 가장 단순하고 허용적
- Apache-2.0: 특허 조항 포함
- BSD-3-Clause: 허용적이면서 비교적 엄격
- GPL-3.0: 파생물 공개 의무가 강함

연구실/과제 결과물이라면 지도교수, 기관, 기업 협약 조건을 먼저 확인하십시오.

## 6. 설치 및 smoke test

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\Activate.ps1
pip install -e ".[mcp,pyhwp,dev]"
python -m pytest
```

CLI 확인:

```bash
hwp-parse --help
```

MCP server import 확인:

```bash
python -m hwp_full_parser.mcp_server --help
```

일부 MCP server는 `--help`를 처리하지 않고 stdio server를 시작할 수 있습니다. 이 경우 `Ctrl+C`로 종료하면 됩니다.

## 7. Git 초기화 및 첫 커밋

```bash
git init
git add .
git status
git commit -m "Initial HWP parser package with MCP and Claude Skill support"
```

## 8. 원격 저장소 연결

```bash
git branch -M main
git remote add origin https://github.com/YOUR_ID/hwp-full-parser.git
git push -u origin main
```

## 9. GitHub 업로드 후 확인

GitHub 페이지에서 다음을 확인합니다.

- `README.md`가 정상 표시되는지
- `.hwp`, `result.json`, `media/`가 올라가지 않았는지
- `.claude/skills/hwp-parser/SKILL.md`가 포함되어 있는지
- `.mcp.json.example`만 포함되고 개인 `.mcp.json`이 올라가지 않았는지

## 10. Release 생성 선택 사항

GitHub Releases에서 `v0.1.0` 태그를 만들 수 있습니다.

```bash
git tag v0.1.0
git push origin v0.1.0
```

Release note 예시:

```text
Initial release:
- Preserves original HWP v30 parser logic
- Adds Python package layout
- Adds CLI entrypoint
- Adds MCP server
- Adds Claude Code Skill
- Adds API wrappers for JSON, text, table, and media extraction
```
