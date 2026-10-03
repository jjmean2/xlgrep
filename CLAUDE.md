# xlgrep

Excel 워크북(.xlsx/.xlsm/.xltx/.xltm) 셀을 grep처럼 검색하는 Python CLI. 설계와 로드맵은 [DESIGN.md](DESIGN.md).

## 명령

```bash
uv sync                 # 환경 구성 (프로젝트는 editable 설치)
uv run xlgrep --help
uv run pytest           # 전체 테스트
uv build                # dist/ 에 sdist + wheel
```

폴더를 옮기거나 이름을 바꾸면 `.venv`에 절대 경로가 남아 있으므로 `rm -rf .venv && uv sync`.

## 구조 (src/xlgrep/)

- `cli.py` — argparse, `Searcher`(셀 → 검색 텍스트 선택, 매칭), 파일 루프, 종료 코드(0/1/2)
- `workbook.py` — openpyxl read_only로 시트별 `Cell` 격자 생성. 계산값이 필요할 때만 `data_only=True`로 두 번째 패스
- `objects.py` — 셀 밖 대상(이름 정의, 조건부 서식, 유효성 검사, 메모). openpyxl 대신 xlsx XML을 직접 파싱
- `matcher.py` — 일반 패턴 + `-f` 함수 패턴(문자열 리터럴 마스킹 후 매칭)
- `text.py` — `_xlfn.` 등 접두사 정규화, 값 문자열화, 이스케이프, 동아시아 문자 표시 폭
- `output.py` — Line / Pretty / Json / Csv 포매터
- `address.py`, `files.py` — 셀 주소·범위, 파일 탐색

## 규칙과 주의점

- 출력 형식과 옵션 이름은 grep/ripgrep 관례를 따른다. 새 옵션도 rg에 같은 개념이 있으면 그 이름을 쓴다.
- 포매터는 `sys.stdout`을 생성 시점에 읽는다(기본 인자로 바인딩하면 pytest capsys가 출력을 못 잡는다).
- Match처럼 비교 불가능한 객체를 튜플에 넣어 정렬할 때는 반드시 `key=`를 준다(같은 행 다중 매칭 시 TypeError 이력).
- openpyxl은 계산값(캐시)을 쓰지 않는다. 테스트에서 계산값이 필요하면 `tests/conftest.py`의 `make_workbook(cached=...)`가 시트 XML에 `<v>`를 패치한다.
- openpyxl read_only 시트는 행을 순회할 때 파싱되므로 경고도 그때 난다. 경고 억제(`_quiet()`)는 순회 구간까지 감싸야 한다.
- `.xls`는 지원하지 않는다(사용자 결정).
- `data/`는 사용자의 실험용 파일 폴더로 gitignore 대상이다. 커밋하지 않는다.
