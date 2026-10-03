# xlgrep 설계

여러 Excel 파일(.xlsx/.xlsm/.xltx/.xltm)의 셀과 셀 밖의 수식(이름 정의, 조건부 서식, 데이터 유효성 검사, 메모)을
grep과 비슷한 인터페이스로 검색하는 CLI 도구.
구형 바이너리 형식(.xls)은 지원하지 않는다.

## 사용 예

```
xlgrep VLOOKUP ./reports            # 수식/값에서 VLOOKUP 검색
xlgrep -f VLOOKUP,XLOOKUP ./reports # 함수 호출만 검색 (문자열 리터럴 안은 제외)
xlgrep -p -C 1 SUMIFS a.xlsx         # pretty 모드, 위아래 1행 컨텍스트
xlgrep --json -i 'total' .           # JSON Lines 출력
```

## 출력 형식

기본: 결과 하나당 한 줄, `경로:시트!셀:내용`

```
reports/sales.xlsx:Summary!B12:=VLOOKUP(A12,Data!A:D,4,FALSE)
reports/sales.xlsx:'Raw Data'!F3:=IFERROR(VLOOKUP(E3,Map,2,0),"")
```

- 셀 주소는 Excel 표기(`Sheet!A1`). 시트명에 공백·특수문자가 있으면 `'Raw Data'!F3`처럼 따옴표로 감싼다.
- 셀 내용의 줄바꿈/탭/캐리지리턴은 `\n`, `\t`, `\r`로 이스케이프해 항상 한 줄로 출력한다.
- 컨텍스트 줄은 grep처럼 구분자로 `-`를 쓰고, 연속되지 않는 그룹 사이에 `--`를 출력한다.
- `--show-value`: 수식 셀 뒤에 캐시된 계산값을 붙인다. `=SUM(B2:B9) → 1240`
- `--header ROW`: 셀 주소 뒤에 해당 열의 헤더를 붙인다. `Summary!B12[매출액]:...`
- 색상: 경로(보라), 셀 주소(초록), 매칭(빨강 굵게). `--color=auto|always|never`, `NO_COLOR` 환경 변수 존중.

## 검색 대상 (`--in`)

| 값 | 동작 |
|---|---|
| `auto` (기본) | 수식 셀은 수식 텍스트, 일반 셀은 값 |
| `formula` | 수식 셀만, 수식 텍스트 |
| `value` | 모든 셀의 값 (수식 셀은 Excel이 마지막으로 저장한 계산값) |

openpyxl은 수식과 계산값을 한 번에 읽지 못하므로, 계산값이 필요할 때(`--in value`, `--show-value`, pretty 모드)만
`data_only=True`로 파일을 한 번 더 읽는다.

최신 함수는 파일에 `_xlfn.XLOOKUP`, `_xlfn._xlws.FILTER`, `_xlpm.x`(LAMBDA 인수)처럼 저장된다.
수식 텍스트는 이 접두사를 제거해 Excel에서 보이는 그대로 정규화한 뒤 검색·출력한다 (`--raw-formula`로 비활성화).

## 셀 밖의 대상 (`--objects`)

함수 사용처를 빠짐없이 찾으려면 셀 밖도 봐야 한다. 기본으로 모두 검색하고 `--objects`로 좁힌다.

| `--objects` 값 | 대상 | 종류 | 출력 위치 |
|---|---|---|---|
| `cells` | 셀 | 수식/값 | `Summary!B4` |
| `names` | 이름 정의 | 수식 | `TaxRate#name`, 시트 범위는 `Summary!Local#name` |
| `cf` | 조건부 서식 규칙의 수식 | 수식 | `Summary!B2:B50,F2:F50#cf` |
| `dv` | 데이터 유효성 검사 formula1/formula2 | 수식 | `Summary!A1:A100#dv` |
| `notes` | 셀 메모 텍스트 | 값 | `Summary!C3#note` |

```
book.xlsx:Summary!B4:=SUM(B2:B3)
book.xlsx:TaxRate#name:=Summary!$B$1
book.xlsx:Summary!B2:B50#cf:=VLOOKUP($A2,Data!A:B,2,0)>100
book.xlsx:Summary!C3#note:check VLOOKUP here
```

- 종류(수식/값)로 기존 옵션과 맞물린다: `--in formula`와 `-f`는 메모 제외, `--in value`는 이름/조건부 서식/유효성 검사 제외.
- 조건부 서식·유효성 검사의 수식은 파일에 `=` 없이 저장되므로 앞에 `=`를 붙여 셀 수식과 같은 모양으로 출력한다.
  규칙 하나에 수식이 여럿이면(`between`의 두 경계, formula1/formula2) 수식마다 한 줄.
- 건너뛰는 이름: `_xlnm.*`(인쇄 영역, 필터 등 내장 이름), `_xlfn.*`(최신 함수용으로 Excel이 만드는 숨김 자리표시자).
- `-C`/`--row`/`--header`/`--show-value`는 셀에만 적용.
- `--range`: 조건부 서식·유효성 검사는 적용 범위가 겹치면, 메모는 셀이 범위 안이면 포함. 이름은 위치가 없으므로 제외.
- `--sheet`: 시트 범위 이름에는 적용, 통합문서 범위 이름은 시트에 속하지 않으므로 `--sheet`/`--range` 지정 시 제외.
- pretty 모드: 시트 격자 아래에 `#cf  B2:B50  =...` 목록. 통합문서 범위 이름은 `(workbook)` 제목 아래.
- JSON: 모든 레코드에 `"object": "cell|name|cf|dv|note"`. 셀 밖 대상은 `"ref"`와 부가 정보
  (`rule_type`, `dv_type`, `part`, 이름의 `hidden`)를 갖는다.

### 읽는 방식

openpyxl read_only는 조건부 서식·유효성 검사·메모를 읽지 않고, 일반 모드는 큰 시트에서 느리며, 둘 다 Excel 2010+가
다른 시트를 참조하는 규칙을 저장하는 `x14` 확장 블록(`<extLst>`)을 버린다. 그래서 `objects.py`가 패키지 XML을 직접 읽는다.

- 시트 XML에서 `<sheetData>`를 잘라내고 루트 시작 태그 + 나머지만 파싱한다(규칙들은 모두 sheetData 뒤에 온다).
  10만 행 × 10열 파일에서 추가 비용이 측정 오차 수준(6.76s → 6.75s).
- 요소는 네임스페이스 대신 로컬 이름으로 찾는다(x14/xm 확장, Strict OOXML 대응).
- 메모는 시트의 관계 파일(`_rels/sheetN.xml.rels`)이 가리키는 comments 파트에서 읽는다. 최신 "스레드 댓글"은 Excel이
  호환용 메모도 함께 저장하므로 그쪽으로 잡힌다.
- 범위 밖: 차트 수식, 피벗, 외부 연결, VBA.

## 함수 사용 집계 (`--list-funcs`)

검색 결과 대신, 수식에서 쓰인 함수를 집계한다.

```
$ xlgrep --list-funcs reports/
FUNCTION    CALLS  PLACES  FILES
VLOOKUP       142     130     12
SUMIFS         88      80      9
XLOOKUP         7       7      2
MyLambda        3       3      1  lambda
GetRate         2       2      1  custom
```

- CALLS: 호출 횟수(중첩 포함). PLACES: 그 함수가 든 수식 위치 수. FILES: 파일 수.
- 마지막 열: 내장 함수는 비움, `lambda`(같은 통합문서의 이름 정의와 일치), `custom`(VBA/추가 기능).
  - 판별 규칙: `openpyxl.utils.formulas.FORMULAE`(Excel 2007 내장 355개)에 있거나, 파일에 `_xlfn.`/`_xlws.` 접두사로
    저장된 함수(2007 이후 추가 함수는 항상 이 접두사로 저장됨)는 내장. `_xll.` 접두사는 XLL 추가 기능 → custom.
    별도 함수 목록을 유지할 필요가 없다.
- 함수 추출은 `openpyxl.formula.Tokenizer`의 FUNC/OPEN 토큰(문자열 리터럴 안은 자동 제외). 토크나이저가 실패하는
  깨진 수식은 문자열 리터럴을 가린 뒤 정규식으로 대체 추출.
- 이 모드에서 위치 인자는 모두 경로이고, `-e PATTERN`(정규식, `-i`/`-F`/`-w`/`-S` 적용)과 `-f NAMES`(정확히 일치)는
  **함수 이름**을 거른다(사용자 결정). 둘 다 주면 합집합.
- `--by file|sheet`: 파일별/시트별로 나눠 표시. 시트별에서 통합문서 범위 이름은 `(workbook)`.
- `--sort calls|name`: 기본은 CALLS 내림차순.
- 범위 옵션(`--objects`, `--sheet`, `--range`, `-g`, `--no-hidden`)은 그대로 적용. 메모는 수식이 아니므로 집계 대상 아님.
- `--json`(집계 행마다 한 줄), `--csv` 지원. `-p`는 기본 표와 같으므로 무시.
- 셀 단위 출력 옵션(`-l -c -q -o -v -m -A -B -C --row --header --show-value`, `--in value`)은 함께 쓰면 오류.
  반대로 `--by`/`--sort`를 `--list-funcs` 없이 쓰면 오류.

## 옵션

### 패턴
| 옵션 | 의미 |
|---|---|
| `PATTERN` | 정규식 (Python `re`) |
| `-e PAT` (반복) | 여러 패턴. 지정 시 위치 인자는 모두 경로 |
| `-F` | 고정 문자열 |
| `-i` / `-S` | 대소문자 무시 / 스마트 케이스(패턴에 대문자가 없으면 무시) |
| `-w` | 단어 단위 |
| `-f NAMES` | 함수 호출 검색. 쉼표 구분. `NAME(` 형태만, 대소문자 무시, 문자열 리터럴 내부 제외 |
| `-v` | 반전: 매칭되지 않는 (비어있지 않은) 셀/대상 |

### 출력
| 옵션 | 의미 |
|---|---|
| `-p`, `--pretty` | 파일/시트별로 묶고 주변을 격자로 표시 |
| `--json` | 매칭당 JSON 한 줄 |
| `--csv` | CSV (`file,sheet,object,location,kind,content,value`) |
| `-l` / `-c` | 매칭된 파일 목록 / 파일별 매칭 수 |
| `-o` | 매칭된 부분만 |
| `-m N` | 파일당 최대 매칭 수 |
| `-q` | 출력 없이 종료 코드만 |
| `--max-width N` | pretty 모드 셀 최대 표시 폭 (기본 40) |

### 컨텍스트
| 옵션 | 의미 |
|---|---|
| `-A N` / `-B N` / `-C N` | 같은 열의 아래/위/위아래 N행 |
| `--row` | 매칭 셀이 있는 행의 다른 셀들 |
| `--header [ROW]` | 열 헤더 표시 (기본 1행) |
| `--col-context N` | pretty 모드의 좌우 열 수 (기본 1) |

### 범위
| 옵션 | 의미 |
|---|---|
| `PATH...` | 파일 또는 디렉터리(항상 재귀). 생략 시 현재 디렉터리 |
| `-g GLOB` (반복) | 파일 이름 글롭. `!`로 시작하면 제외 |
| `--sheet GLOB` (반복) | 시트 이름 글롭 |
| `--range A1:F100` | 셀 범위 제한 |
| `--no-hidden` | 숨김 시트 제외 (기본은 포함) |
| `--objects LIST` | 검색 대상: `cells,names,cf,dv,notes` 또는 `all`(기본) |

Excel 잠금 파일(`~$*.xlsx`)과 숨김 디렉터리(`.`으로 시작)는 탐색 시 건너뛴다.

### 종료 코드
grep과 동일: 매칭 있음 0, 없음 1, 오류 발생 2.

## pretty 모드

```
reports/sales.xlsx
  Summary
       │ A    │ B                              │ C
    11 │ 서울 │ =VLOOKUP(A11,Data!A:D,4,FALSE) │ 1200
  ▶ 12 │ 부산 │ =VLOOKUP(A12,Data!A:D,4,FALSE) │ 980
```

- 매칭 행 ± `-C`행(pretty 기본 0), 매칭 열 ± `--col-context`열을 표시.
- 겹치는 행 구간은 하나의 블록으로 합친다.
- 표시 폭은 동아시아 문자(한글 등)를 2칸으로 계산하고, 긴 셀은 `…`로 자른다.

## 구조

```
src/xlgrep/
  cli.py       인자 파싱, 실행 흐름, 종료 코드
  files.py     경로 탐색, 글롭 필터
  workbook.py  워크북 읽기 → 시트별 셀 격자 (수식/값)
  objects.py   셀 밖 대상 (이름 정의, 조건부 서식, 유효성 검사, 메모) — XML 직접 파싱
  funcs.py     --list-funcs: 함수 추출, 내장/lambda/custom 분류, 집계
  text.py      수식 정규화, 값 문자열화, 이스케이프, 표시 폭
  matcher.py   패턴 컴파일, 매칭 구간 계산
  address.py   셀 주소/범위 변환, 시트명 인용
  output.py    line / pretty / json / csv 포매터, 색상
```

## 결정 사항

| 결정 | 이유 |
|---|---|
| Python + openpyxl, uv(uv_build 백엔드), src 레이아웃 | 배포 가능한 정석 구조. 수식 텍스트를 읽을 수 있는 라이브러리 |
| `.xls` 미지원 | openpyxl이 읽지 못함. 필요해지면 xlrd(값만 가능) 검토 |
| 숨김 시트 기본 포함 | 주 용도가 "함수 사용처를 빠짐없이 찾기" |
| 디렉터리는 항상 재귀 | rg와 동일. `-r` 옵션 불필요 |
| 빈 셀은 컨텍스트 줄에서 생략 (line 모드) | 노이즈 감소. pretty 모드는 격자이므로 빈칸으로 표시 |
| 시트 전체를 메모리에 적재 | 2차원 컨텍스트 계산이 단순. 대용량 파일에서 문제되면 재검토 |
| 셀 밖 대상 기본 포함 (사용자 결정) | 함수 사용처를 빠짐없이 찾는 주 용도. 추가 비용이 거의 없음 |
| 셀 밖 대상은 위치 뒤 `#종류` 표기 (사용자 결정) | `경로:위치:내용` 3필드 유지. `[...]`는 `--header`가 사용 중 |
| 집계는 하위 명령 대신 `--list-funcs` 플래그 (사용자 결정) | 단일 명령 유지. rg의 `--files`, `--type-list`와 같은 방식 |
| 집계 모드의 `-e`/`-f`는 함수 이름을 거름 (사용자 결정) | "LOOKUP 계열 함수가 얼마나 쓰이나" 같은 질문에 바로 답함 |

## 로드맵

### 1단계 — 완료
위 "옵션" 절의 모든 기능. 테스트 `tests/` (단위 + CLI).

### 2단계 (우선순위 순)

1. ~~**셀 밖의 수식 검색**~~ — 완료. 위 "셀 밖의 대상" 절.
2. ~~**`--list-funcs`**~~ — 완료. 위 "함수 사용 집계" 절. (`--count-by`는 `--by file|sheet`로 구현)
3. **`--ref 'Data!A:D'`** — 특정 범위를 참조하는 수식 검색.
   - Tokenizer의 OPERAND/RANGE 토큰으로 참조 추출 후 범위 겹침 판정.
   - 어려운 경우: 시트명 생략 참조(같은 시트), 이름 정의 경유, INDIRECT/OFFSET, 구조적 참조(`Table[Col]`), 외부 통합문서 `[1]Sheet!A1`.
4. **여러 파일 병렬 처리** — ProcessPoolExecutor, 출력 순서는 파일 순서 유지.
   - 참고: 10만 행 × 10열(100만 셀) 파일 하나에 약 6.7초. 대부분 openpyxl의 셀 읽기 시간이다.
     병렬화로 부족하면 셀도 XML을 직접 읽거나 `python-calamine`(Rust, 값 위주) 검토.

### 기타 후보
- 숨김 행/열 제외 옵션
- PyPI 배포 (2026-10-03 기준 이름 `xlgrep` 비어 있음), CI, 린터(ruff) 설정
- 메모 작성자 표시 (Excel은 메모 본문 앞에 "작성자:" 굵은 글씨 런을 넣는다 — 실제 파일에서 어떻게 보이는지 확인 필요)

## 검증되지 않은 부분

- 대용량 파일 성능 (두 번 읽기 + 시트 전체 메모리 적재).
- `--list-funcs`는 실제 Excel 파일로 아직 확인하지 않았다.

사용자가 실제 업무 파일로 v0.2.0(셀, 계산값, 이름/조건부 서식/유효성 검사/메모)을 써 보고 잘 동작한다고 확인함 (2026-10-03).

확인된 것: **공유 수식**(채우기 핸들로 복사한 수식. 마스터 셀에만 텍스트가 있고 나머지는 `<f t="shared" si="0"/>`)은
openpyxl read_only가 셀별 수식(`=A2*2`, `=A3*2`)으로 풀어준다. XML을 직접 작성해 확인함 (2026-10-03).
