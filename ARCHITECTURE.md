# xlgrep 코드 읽기 안내

이 문서는 xlgrep의 코드를 처음부터 이해하고 싶은 사람을 위한 길잡이다. 무엇을 하는지는 [README.md](README.md),
왜 그렇게 정했는지는 [DESIGN.md](DESIGN.md)에 있고, 여기서는 **코드가 어떻게 짜여 있고 어떤 순서로 읽으면 되는지**를 다룬다.

## 1. 한 장 요약

명령 하나가 실행되는 흐름:

```
cli.main
 ├─ 옵션 검증 → SearchConfig (무엇을 어떻게 찾을지, 프로세스 간 전달 가능)
 ├─ files.iter_files → 검색할 .xlsx 목록
 └─ search.Runner → 파일마다 search_file(path, cfg)   (여러 파일이면 병렬)
        │
        ├─ scope.ScopedWorkbook   이 파일에서 "범위 안"인 것만 시트 단위로 꺼내 준다
        │     ├─ package.open_package      xlsx(zip) 구조: workbook.xml, 시트 목록
        │     ├─ objects.read_package_objects  이름 정의·조건부 서식·유효성 검사·메모
        │     └─ workbook.CellReader       시트 XML → Cell 들 (수식 + 계산값)
        │
        ├─ search.Searcher        셀/대상마다 매칭 여부와 강조 구간 결정
        │     ├─ matcher.Matcher           정규식, -f 함수 패턴
        │     └─ refs.RefFinder            --ref 참조 판정
        │
        └─ output.Formatter       그 파일의 결과를 문자열로 렌더링
 ← 부모 프로세스: output.ResultStream 이 파일 순서대로 출력 (파일 사이 구분자 담당)
```

요약 모드도 같은 길을 간다. `--list-funcs`는 `search_file` 대신 `count_file`이 함수를 세고 `funcs.write_func_stats`가
표를 그린다. `--stats`는 `stats.stats_file`이 지표를 모으고 `stats.write_stats`가 그린다. `--deps`는 `deps.deps_file`이 의존을 모으고,
부모가 경로를 해석한 뒤(`deps.resolve_target`, 검색한 파일 전체를 알아야 해서) `deps.write_deps`가 그린다. 요약 기능은 수집과 출력을
**자기 모듈 안에** 두고(`funcs.py`, `stats.py`), `output.py`는 검색 결과 출력만 맡는다.

**핵심 원칙 세 가지**만 기억하면 대부분의 구조가 설명된다.

1. **범위 규칙은 한 곳(`scope.py`)**. 검색과 집계가 같은 규칙으로 시트·셀·대상을 고른다.
2. **포매터 하나 = 파일 하나**. 파일 사이에 걸치는 출력은 `ResultStream`만 안다. 그래서 병렬로 돌려도 출력이 같다.
3. **복잡함은 이유가 있는 곳에만**. 빠른 경로(`workbook.py`), 정규식 참조 스캐너(`refs.py`)는 측정이나
   실제 문제 때문에 생긴 것이고, 각각 이유가 docstring에 적혀 있으며 교차 검증 테스트가 붙어 있다.

## 2. xlsx 파일에 대해 알아야 할 것

코드를 읽기 전에 xlsx 내부 구조를 알면 절반은 이해한 셈이다. `unzip -l 파일.xlsx`로 직접 열어 보길 권한다.

```
[Content_Types].xml
_rels/.rels                       → 통합문서 본체가 어디 있는지
xl/workbook.xml                   → 시트 목록(이름, 숨김 여부), 이름 정의(<definedNames>)
xl/_rels/workbook.xml.rels        → 시트 id → 실제 파일 경로, sharedStrings/styles 위치
xl/worksheets/sheet1.xml          → 셀 (<sheetData>), 조건부 서식, 유효성 검사, 확장(<extLst>)
xl/worksheets/_rels/sheet1.xml.rels → 그 시트의 메모 파일 위치
xl/sharedStrings.xml              → 문자열 셀의 실제 텍스트 (셀에는 번호만 저장)
xl/styles.xml                     → 서식. 날짜인지는 여기 숫자 서식을 봐야 안다
xl/comments1.xml                  → 셀 메모
```

셀 하나는 대략 이렇게 생겼다:

```xml
<c r="B2" t="s"><v>5</v></c>                 <!-- 공유 문자열 5번 -->
<c r="C2" s="3"><v>45000</v></c>             <!-- 숫자. 스타일 3이 날짜 서식이면 날짜 -->
<c r="D2"><f>SUM(A1:A3)</f><v>6</v></c>      <!-- 수식(앞에 = 없음)과 마지막 계산값 -->
<c r="D3"><f t="shared" si="0"/><v>9</v></c> <!-- 위쪽 수식을 채우기로 복사한 셀: 수식은 마스터에만 -->
```

알아 둘 개념:
- **공유 수식(shared formula)**: 채우기 핸들로 복사한 수식은 첫 셀(마스터)에만 텍스트가 있다. 나머지 셀의 수식은
  마스터 수식의 상대 참조를 위치만큼 옮겨 만든다 → `refs.SharedFormula`.
- **`_xlfn.` 접두사**: Excel 2007 이후 생긴 함수는 파일에 `_xlfn.XLOOKUP`처럼 저장된다. 화면에는 안 보인다
  → `text.normalize_formula`가 지운다. `--list-funcs`는 이 접두사로 내장 함수를 판별하므로 원문을 쓴다.
- **캐시된 계산값**: `<v>`는 Excel이 마지막으로 저장할 때의 결과다. 다른 프로그램이 만든 파일엔 없을 수 있다.

## 3. 읽는 순서

작은 것부터, 의존하는 쪽으로 올라간다. 각 단계에서 "이것만 잡으면 된다"를 적었다.

| 순서 | 파일 | 줄 | 잡을 것 |
|---|---|---|---|
| 1 | `address.py` | 90 | `A1`↔(행,열), 시트명 인용, `CellRange`(빈 쪽은 열린 범위). 나머지 전부가 쓰는 기초 |
| 2 | `text.py` | 85 | 수식 정규화, 문자열 리터럴 가리기(위치는 유지), 값→문자열, 한 줄 이스케이프, 한글 폭 |
| 3 | `matcher.py` | 75 | 패턴 → 정규식 컴파일, `spans()`가 강조 구간 반환. `-f`는 문자열을 가린 뒤 `NAME(` 매칭 |
| 4 | `files.py` | 60 | 경로 → .xlsx 목록. 지원 안 하는 파일은 예외 객체로 *순서 안에* 섞여 나온다(오류 순서 유지용) |
| 5 | `package.py` | 85 | 2절의 xlsx 구조를 코드로. `open_package`가 시트 목록과 rels를 준다 |
| 6 | `workbook.py` | 405 | 셀 읽기. **아래 "workbook.py 읽는 법" 참고** |
| 7 | `objects.py` | 205 | 셀 밖 대상. 시트 XML에서 `<sheetData>`를 잘라내고 뒷부분만 파싱하는 요령(`_sheet_tail`) |
| 8 | `scope.py` | 110 | `ScopedWorkbook.parts()`: 시트 순서대로 (그 시트의 범위 안 셀, 대상), 마지막에 통합문서 이름 |
| 9 | `refs.py` | 390 | 참조 스캔(`scan_refs`) → `RefContext`로 시트·이름 해석 → `--ref`의 `RefFinder`가 겹침 판정. `SharedFormula`, `relative_form` |
| 10 | `funcs.py` | 200 | 함수 추출(Tokenizer + 수식 모양 캐시), 내장/lambda/custom 분류, 휘발성 목록, `--list-funcs` |
| 11 | `search.py` | 260 | `Searcher._select`가 매칭의 핵심(패턴 AND --ref, -v). `search_file`은 parts를 돌며 찾고 렌더링 |
| 12 | `output.py` | 470 | 포매터 4종 + `ResultStream`. 길지만 각 클래스는 독립적 |
| 13 | `stats.py` | 300 | `--stats`. 수집(`stats_file`, `_count_cells`) → 합계(`combine`) → 출력. `refs.relative_form`으로 고유 수식 |
| 14 | `deps.py` | 330 | `--deps`. `_Targets`가 수식 → (통합문서, 시트) 집합(이름 추적). 경로 해석, Mermaid는 `_Mermaid`가 노드를 모은 뒤 그림 |
| 15 | `cli.py` | 450 | 옵션 정의, 검증(`_search_config`, `_func_config`, `_stats_config`, `_deps_config`), 실행(`run_search` 등) |

### workbook.py 읽는 법

가장 어려운 파일이다. 세 층으로 나눠 읽는다.

1. **바깥 (위쪽)**: `Cell`, `Sheet`, `CellReader`, `read_sheets`. 여기까지만 알아도 나머지 코드를 읽을 수 있다.
2. **값 해석 (가운데)**: `_context`(공유 문자열·날짜 서식을 한 번 로드), `_convert`(문자열 `<v>` → 파이썬 값),
   `_Formulas`(수식 텍스트, 공유 수식 펼치기), `_make_cell`. 두 스캐너가 공유한다.
3. **스캐너 (아래)**: 같은 일을 하는 두 구현.
   - `_scan_regex`: 기본. 정규식 한 번(`_CELL_RE.findall`)으로 모든 셀의 (열, 행, 타입, 스타일, 본문)을 뽑는다.
     셀마다 `_fast_cell`(흔한 모양: 숫자, 공유 문자열, 공유 수식 종속 셀)을 먼저 시도하고, 아니면 `_general_cell`.
   - `_scan_etree`: 표준 XML 파서. 접두사·좌표 없는 셀 등 정규식이 다루지 않는 형식이면 `_Irregular`가 던져지고 이쪽으로 온다.
   - **왜 둘인가**: 정규식 쪽이 2~3배 빠르다(DESIGN.md "셀 읽기" 측정). 대신 두 경로가 같은 결과를 내는지
     `tests/test_reader.py`가 매번 확인한다. 의도된 중복 + 기계적 검증.

## 4. 명령 하나 따라가기: `xlgrep -f VLOOKUP data`

1. `cli.main` → `_search_config`: `-f`가 있으니 위치 인자 `data`는 경로. `build_matcher(funcs=["VLOOKUP"])`.
   `SearchConfig.formulas_only`는 참(값 셀은 절대 매칭 안 됨).
2. `run_search` → `_each_file`: `iter_files(["data"])`가 `data/Test.xlsx`를 찾는다. 파일이 하나라
   `choose_jobs`는 1 → 같은 프로세스에서 `search_file` 실행.
3. `search_file`: 주변 셀을 보여주지 않으므로 `formulas_only=True`로 `ScopedWorkbook`을 연다
   → `CellReader`가 수식 셀만 만든다.
4. `book.parts()`가 시트마다 `Part`를 준다. `Searcher.search_cells`가 셀마다 `_searched_text`(수식 텍스트) →
   `_select` → `matcher.spans()`로 `VLOOKUP` 위치를 찾는다.
5. 찾으면 `LineFormatter.write_sheet`가 `data/Test.xlsx:Sheet1!B4:=…` 줄을 버퍼에 쓴다.
6. 부모: `ResultStream.add`가 버퍼를 출력. 매칭이 있었으면 종료 코드 0.

## 5. 테스트 지도

| 테스트 | 지키는 것 |
|---|---|
| `test_units.py` | 주소·정규화·매처·폭 계산 같은 작은 함수 |
| `test_cli.py` | 기본 검색의 출력 형식 (줄 단위로 정확히 비교) |
| `test_objects.py` | 셀 밖 대상, x14 확장 블록 |
| `test_funcs.py` | `--list-funcs` 분류·집계·캐시 |
| `test_refs.py` | `--ref` 스캔·해석·이름 추적, `SharedFormula` = openpyxl `Translator` |
| `test_reader.py` | 정규식 경로 = 표준 파서 경로 = openpyxl. 테스트 xlsx를 XML로 직접 써서 만든다 |
| `test_stats.py` | `--stats` 지표(셀·패키지), 합계, 표·JSON·CSV·카드 |
| `test_deps.py` | `--deps` 파일·시트·데이터 의존, 경로 해석(found / matched / missing), Mermaid, 연결 문자열 비노출 |
| `test_parallel.py` | `-j1`과 `-j2` 출력이 모든 형식에서 같음 |

`tests/conftest.py`의 `make_workbook`은 openpyxl로 파일을 만든 뒤 계산값을 XML에 패치한다(openpyxl은 계산값을 쓰지 않음).

## 6. 이해했는지 확인하는 질문

답은 코드에 있다. 막히면 괄호 안 위치를 보면 된다.

1. `xlgrep -p -f VLOOKUP`은 왜 값 셀도 읽는가? (`search.search_file`, `SearchConfig.shows_neighbours`)
2. `--sheet Data`를 주면 통합문서 범위 이름이 결과에서 빠진다. 그 규칙은 어디에 있나? (`scope.ScopedWorkbook.parts`)
3. 병렬로 돌릴 때 line 모드의 `--` 구분자가 파일 사이에서 어떻게 유지되나? (`FileResult.grouped`, `ResultStream.add`)
4. `=SUM(2:$5)`의 `$5`에서 `$`는 행에 붙는다. 이걸 처리하는 함수는? 이게 틀리면 어떤 기능 두 개가 영향을 받나? (`refs._split_endpoint`)
5. 정규식 스캐너가 포기하고 표준 파서로 넘기는 경우 다섯 가지는? (`workbook._sheet_data`)
6. `--list-funcs`가 1만 개의 공유 수식을 왜 한 번만 토큰화하나? (`funcs._SHAPE_RE`)
7. `--stats`에서 D2의 `=B2*C2`와 D3의 `=B3*C3`이 고유 수식 1개로 세어지는 과정은? 공유 수식 그룹이면 어떤 계산이 생략되나? (`refs.relative_form`, `stats._count_cells`)
8. 수식의 `[2]Rates!A1`이 `C:\공유\rates.xlsx`로 바뀌기까지 어떤 파일들을 거치나? 경로 해석은 왜 워커가 아니라 부모에서 하나? (`package.external_books`, `deps.resolve_target`, `cli.run_deps`)
9. `--ref`와 `--deps`가 이름 정의를 같은 방식으로 해석한다는 것은 어디서 보장되나? (`refs.RefContext`)
