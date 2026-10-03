# xlgrep 설계

여러 Excel 파일(.xlsx/.xlsm/.xltx/.xltm)의 셀을 grep과 비슷한 인터페이스로 검색하는 CLI 도구.
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
| `-v` | 반전: 매칭되지 않는 (비어있지 않은) 셀 |

### 출력
| 옵션 | 의미 |
|---|---|
| `-p`, `--pretty` | 파일/시트별로 묶고 주변을 격자로 표시 |
| `--json` | 매칭당 JSON 한 줄 |
| `--csv` | CSV (`file,sheet,cell,kind,content,value`) |
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
  text.py      수식 정규화, 값 문자열화, 이스케이프, 표시 폭
  matcher.py   패턴 컴파일, 매칭 구간 계산
  address.py   셀 주소/범위 변환, 시트명 인용
  output.py    line / pretty / json / csv 포매터, 색상
```

## 2단계 (예정)

- 셀 밖의 수식 검색: 이름 정의, 조건부 서식, 데이터 유효성 검사, 셀 메모
- `--list-funcs` / `--count-by`: 사용된 함수 집계
- `--ref 'Data!A:D'`: 특정 범위를 참조하는 수식 검색
- 여러 파일 병렬 처리
