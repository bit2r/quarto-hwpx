# Quarto HWPX Extension

Quarto 마크다운(.qmd)을 한글과컴퓨터 HWPX 문서(.hwpx)로 변환하는 Quarto 확장.

## 프로젝트 구조

```
quarto-hwpx/
├── _extensions/hwpx/
│   ├── _extension.yml          # Quarto 확장 설정 (base: docx)
│   ├── hwpx-filter.lua         # Lua 필터: AST → JSON → Python 호출
│   ├── hwpx_writer.py          # Python: JSON AST → HWPX 파일 생성
│   ├── Skeleton.hwpx           # 번들 HWPX 템플릿 (header.xml 포함)
│   ├── cleanup-docx.sh         # 중간 .docx 삭제 스크립트
│   └── assets/
│       ├── extension-diagram.svg   # 파이프라인 다이어그램 (Tufte 스타일)
│       ├── brother_template.hwpx   # 참조 템플릿 (공공기관 보고서 양식)
│       └── fonts/                  # 번들 폰트
│           ├── NanumSquare{R,B}.otf          # 한글 본문/제목
│           ├── NimbusSanL-{Reg,Bol,*Ita}.otf # 영문 산세리프
│           ├── STIXTwo{Text,Math}-*.otf      # 수식
│           ├── D2Coding{,Bold}-*.ttf         # 코드블록
│           └── NotoSansCJKkr-Hanja.otf       # 한자 (KS X 1001 서브셋)
├── _quarto.yml                 # 프로젝트 설정 (post-render hook)
├── example.qmd                 # 테스트 문서 (공공기관 보고서 양식)
├── README.md                   # 프로젝트 소개
└── CLAUDE.md                   # 이 파일
```

## 빌드 & 사용법

```bash
quarto render example.qmd --to hwpx-docx
# → example.hwpx 생성됨 (.docx는 post-render에서 삭제)
```

## 동작 원리

1. Quarto가 `docx` 베이스 포맷으로 Pandoc 호출
2. `hwpx-filter.lua`가 Pandoc AST를 JSON으로 직렬화
3. `pandoc.pipe('python3', ...)` 로 `hwpx_writer.py` 호출
4. Python이 `Skeleton.hwpx`를 열어 `section0.xml` 새로 생성, `header.xml` 수정, `content.hpf` 메타데이터 업데이트
5. 새 ZIP으로 `.hwpx` 저장
6. `cleanup-docx.sh`가 중간 `.docx` 삭제

## 지원 요소 (Pandoc AST → HWPX)

### Block 요소

| Pandoc Block | HWPX 처리 |
|---|---|
| `Para`, `Plain` | `<hp:p>` 본문 스타일 (10pt, 160% 줄간격) |
| `Header(1~6)` | `<hp:p>` 개요 스타일 (H1=22pt bold, H2=16pt bold, H3=13pt) |
| `CodeBlock` | 줄별 분리하여 각각 `<hp:p>` (D2Coding 고정폭 폰트) |
| `BulletList` | "• " 글머리 기호 + 자식 블록 재귀 처리 |
| `OrderedList` | "1. " 접두사 자동 추가 |
| `BlockQuote` | 전각 공백으로 들여쓰기 |
| `Table` | `<hp:tbl>` — 가운데정렬, 실선 테두리, 캡션 지원 |
| `HorizontalRule` | "━━━" 구분선 |
| `Div` | 자식 블록 투명 패스스루 |
| `DefinitionList` | 용어 + 들여쓰기 정의 |

### Inline 요소

`Str`, `Space`, `SoftBreak`, `LineBreak`, `Strong`, `Emph`, `Code`, `Link`, `Quoted`, `Math`, `Span`

### 수식 (Math → hp:equation)

단독 Math 요소를 `<hp:equation>` + `<hp:script>` 태그로 변환:
- `latex_to_hwp_script()`: LaTeX → 한글 수식 스크립트 변환
- `\frac{a}{b}` → `{a} over {b}`, `\sum_{i}^{n}` → `sum from{i} to{n}` 등
- 인라인 Math(텍스트 혼합)는 텍스트 추출 유지

### 메타데이터 (Title Block)

YAML 헤더의 `title`, `subtitle`, `author`, `date`를 문서 본문 상단에 렌더링:
- title: 본문 스타일 + 22pt bold (charPrIDRef=7)
- subtitle: 본문 스타일 + 16pt bold (charPrIDRef=8)
- author | date: 본문 스타일 (10pt)

## Skeleton.hwpx 스타일 매핑

- 출처: [airmang/python-hwpx](https://github.com/airmang/python-hwpx) (MIT)
- 페이지: A4 (59528×84186 HWPUNIT), 여백: 좌우=8504, 상=5668, 하=4252
- 텍스트 영역 폭: 42520 HWPUNIT

### 폰트 매핑 (언어별)

| fontface lang | id=0 (primary) | 용도 |
|---|---|---|
| HANGUL | NanumSquareOTF | 한글 본문/제목 |
| LATIN | NimbusSanL | 영문 산세리프 |
| HANJA | Noto Sans CJK KR | 한자 (KS X 1001 서브셋 4,620자) |
| JAPANESE | Noto Sans CJK KR | 일본어 (한자 공유) |
| SYMBOL | STIX Two Text | 수식/기호 |
| OTHER, USER | NimbusSanL | 기타 |

모든 블록 공통: id=1 = id=0과 동일, id=2 = D2Coding (fontCnt=3)

### 런타임 주입 (hwpx_writer.py가 header.xml에 추가)

| 항목 | ID | 내용 |
|---|---|---|
| fontface 재작성 | 7블록 전체 | 언어별 폰트 매핑 (LANG_FONT_MAP) |
| charPr 7 | H1용 | 22pt, bold, fontRef=0 |
| charPr 8 | H2용 | 16pt, bold, fontRef=0 |
| charPr 9 | H3용 | 13pt, fontRef=0 |
| charPr 10 | 코드블록용 | 10pt, D2Coding (fontRef=2) |
| borderFill 3 | 표 셀용 | 사방 실선 0.12mm |
| paraPr 2 prev=800 | H1 상단 간격 | 8pt |
| paraPr 3 prev=600 | H2 상단 간격 | 6pt |
| paraPr 4 prev=400 | H3 상단 간격 | 4pt |

## 기술 노트

### 한글 Mac namespace 호환성

- 한글 Mac은 XML namespace prefix를 하드코딩으로 인식 (`hp:`, `hs:`, `hc:`, `hh:`)
- ElementTree가 생성하는 `ns0`, `ns1` prefix → 빈 문서로 표시됨
- **해결**: raw XML 문자열 방식으로 원본 namespace prefix 유지
- 참고: [airmang/python-hwpx#8](https://github.com/airmang/python-hwpx/pull/8)

### linesegarray (레이아웃 캐시)

- 한글 Mac은 linesegarray 없으면 레이아웃을 재계산하지 않음
- `compute_lineseg_xml()`: 텍스트 길이 기반 다중 lineseg 항목 생성
- CJK 문자 ≈ char_height 폭, Latin ≈ char_height/2 폭
- flags: first=0x20000, last=0x40000, firstlast=0x60000

### HWPUNIT 변환

- A4: 59528 × 84186 = 210mm × 297mm
- 10pt ≈ 1000 HWPUNIT (높이), 한글 문자 폭 ≈ 1000 HWPUNIT
- 한 줄 약 42자 (42520 / 1000)

### LaTeX → 한글 수식 스크립트 변환

| LaTeX | 한글 수식 스크립트 |
|---|---|
| `\frac{a}{b}` | `{a} over {b}` |
| `\sum_{i}^{n}` | `sum from{i} to{n}` |
| `\int_{a}^{b}` | `int from{a} to{b}` |
| `\sqrt{x}` | `sqrt{x}` |
| `\alpha` 등 | `alpha` (백슬래시 제거) |
| `\geq`, `\leq` | `>=`, `<=` |
| `\times`, `\cdot` | `times`, `cdot` |
| `\infty` | `inf` |

### 외부 의존성

- Python 표준 라이브러리만 사용: `zipfile`, `json`, `re`, `xml.sax.saxutils`
- 외부 패키지 없음
