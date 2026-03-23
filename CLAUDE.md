# IO 문서 자동 생성기

## 개요
엑셀(xlsx) 또는 CSV 데이터로부터 광고게재신청서(IO) docx + PDF 문서를 자동 생성한다.

## 폴더 구조
- `data/` — 데이터 파일(xlsx 또는 csv)을 여기에 넣으면 됨 (sample.xlsx 참고)
- `output/` — 생성된 docx, PDF 문서가 여기에 저장됨
- `src/` — 스크립트 및 템플릿 (수정 불필요)

## 사용법

### 1. 데이터 파일 준비
`data/sample.xlsx`를 참고하여 입력 데이터를 준비하고 `data/` 폴더에 넣는다. xlsx와 csv 모두 지원한다.

필수 컬럼:
- `client_name` — 매체사명 (같은 값의 행이 하나의 문서로 그룹핑됨)
- `client_address` — 매체사 주소
- `client_email` — 매체사 이메일
- `client_manager` — 담당자명
- `gross_rate` — 매체 수수료율
- `service` — 매체명
- `service_name` — Service Name (Domain)
- `widget_name` — 위젯명
- `value` — 계약 단가
- `date_start` — 계약 시작일

### 2. 문서 생성 실행
```bash
python3 src/generate.py --input data/<파일명>
```

옵션:
- `--template`, `-t` : 템플릿 파일 (기본: `src/io-sample.docx`)
- `--output`, `-o` : 출력 디렉토리 (기본: `output/`)
- `--no-pdf` : PDF 변환 건너뛰기

### 3. 출력
- `output/io-{매체사명}.docx`
- `output/io-{매체사명}.pdf` (Microsoft Word 필요)

## 실행 전 필수 사항
문서 생성 전에 반드시 의존성이 설치되어 있는지 확인하고, 없으면 먼저 설치한다:
```bash
pip3 install -r src/requirements.txt
```

## Commands
- 문서 생성: `python3 src/generate.py -i data/<파일명>`
- 문서 생성 (PDF 제외): `python3 src/generate.py -i data/<파일명> --no-pdf`
- 테스트: `python3 -m pytest src/tests/ -v`
- PDF 변환에 Microsoft Word 필요 (macOS/Windows 모두 지원)
