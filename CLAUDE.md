# xlgrep

Excel 워크북(.xlsx/.xlsm/.xltx/.xltm) 셀을 grep처럼 검색하는 Python CLI. 설계와 로드맵은 [DESIGN.md](DESIGN.md),
코드 구조와 읽는 순서는 [ARCHITECTURE.md](ARCHITECTURE.md). 구조를 바꾸면 ARCHITECTURE.md도 함께 고친다.

사용자는 이 코드를 직접 공부해서 이해하고 설명할 수 있기를 원한다. 불필요한 복잡성을 줄이고, 필요한 복잡성(성능 등)은
격리하고 이유를 적는다.

## 명령

```bash
uv sync                 # 환경 구성 (프로젝트는 editable 설치)
uv run xlgrep --help
uv run pytest           # 전체 테스트
uv build                # dist/ 에 sdist + wheel
```

폴더를 옮기거나 이름을 바꾸면 `.venv`에 절대 경로가 남아 있으므로 `rm -rf .venv && uv sync`.

## 구조 (src/xlgrep/)

- `cli.py` — argparse, 옵션 검증(`_search_config`, `_func_config`), 실행(`run_search`, `run_list_funcs`), 종료 코드(0/1/2)
- `scope.py` — 범위 규칙(시트 글롭, 숨김, --range, --objects, 통합문서 이름)을 한 곳에. `ScopedWorkbook.parts()`가 시트별 범위 안 셀·대상을 준다
- `workbook.py` — 시트 XML을 직접 읽어 `Cell` 격자 생성(수식+계산값 한 번에). 정규식 경로가 기본, 예외적 형식은 `ElementTree` 경로
- `package.py` — xlsx 패키지 구조: workbook.xml, rels, 시트 목록 (objects.py와 workbook.py가 공유)
- `objects.py` — 셀 밖 대상(이름 정의, 조건부 서식, 유효성 검사, 메모). openpyxl 대신 xlsx XML을 직접 파싱
- `funcs.py` — Tokenizer로 함수 추출(`CallScanner`, 수식 모양별 캐시), 내장/lambda/custom 분류, `VOLATILE`, `--list-funcs` 집계와 출력
- `stats.py` — `--stats`: 파일별 지표 수집(`stats_file`), 합계(`combine`), 표·카드·JSON·CSV 출력
- `refs.py` — `SharedFormula`(공유 수식 템플릿, 셀 위치로 상대 참조 이동), `relative_form`(R1C1 정규화, 고유 수식 수), `--ref`: 정규식으로 참조 추출(문자열·대괄호 마스킹), 대상 범위 겹침 판정, 이름 정의 연쇄 추적
- `search.py` — 파일 하나 처리(`Searcher`, `search_file`, `count_file`)와 병렬 실행(`Runner`, `choose_jobs`). 범위 판단은 하지 않는다(scope.py)
- `matcher.py` — 일반 패턴 + `-f` 함수 패턴(문자열 리터럴 마스킹 후 매칭)
- `text.py` — `_xlfn.` 등 접두사 정규화, 값 문자열화, 이스케이프, 동아시아 문자 표시 폭
- `output.py` — Line / Pretty / Json / Csv 포매터(하나가 파일 하나를 렌더링), `ResultStream`(파일 사이 구분자·CSV 헤더)
- `address.py`, `files.py` — 셀 주소·범위, 파일 탐색

## 규칙과 주의점

- 출력 형식과 옵션 이름은 grep/ripgrep 관례를 따른다. 새 옵션도 rg에 같은 개념이 있으면 그 이름을 쓴다.
- 포매터는 `sys.stdout`을 생성 시점에 읽는다(기본 인자로 바인딩하면 pytest capsys가 출력을 못 잡는다).
- Match처럼 비교 불가능한 객체를 튜플에 넣어 정렬할 때는 반드시 `key=`를 준다(같은 행 다중 매칭 시 TypeError 이력).
- openpyxl은 계산값(캐시)을 파일에 쓰지 않는다. 테스트에서 계산값이 필요하면 `tests/conftest.py`의 `make_workbook(cached=...)`가 시트 XML에 `<v>`를 패치한다.
- 셀 읽기를 바꾸면 `tests/test_reader.py`로 정규식 경로 = 표준 파서 경로 = openpyxl 결과인지 확인한다. 테스트용 xlsx는 XML을 직접 써서 만든다(공유 문자열·공유 수식 등 openpyxl이 쓰지 않는 구조 포함).
- 작업자 프로세스로 넘기는 설정(`SearchConfig`, `FuncConfig`, `OutputOptions`)은 pickle 가능해야 한다. 클로저·람다 필드 금지.
- 파일 사이에 걸친 출력(`--` 구분자, pretty 파일 간 빈 줄, CSV 헤더)은 포매터가 아니라 `output.ResultStream`이 붙인다. `-j1`과 `-j2` 출력이 같아야 한다(`tests/test_parallel.py`).
- `.xls`는 지원하지 않는다(사용자 결정).
- `data/`는 사용자의 실험용 파일 폴더로 gitignore 대상이다. 커밋하지 않는다.
