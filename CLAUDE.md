# IO 문서 자동 생성기

## 개요
엑셀 데이터로부터 광고게재신청서(IO) docx + PDF 문서를 자동 생성한다.

## 사용법

### 1. 엑셀 파일 준비
`sample-data.xlsx`를 참고하여 입력 데이터를 준비한다.

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
python3 generate.py --input <엑셀파일> --output <출력디렉토리>
```

옵션:
- `--template`, `-t` : 템플릿 파일 (기본: `io-sample.docx`)
- `--no-pdf` : PDF 변환 건너뛰기

### 3. 출력
- `output/io-{매체사명}.docx`
- `output/io-{매체사명}.pdf` (LibreOffice 필요)

## Commands
- 의존성 설치: `pip3 install -r requirements.txt`
- 문서 생성: `python3 generate.py -i data.xlsx`
- 테스트: `python3 -m pytest tests/ -v`
- PDF 변환에 LibreOffice 필요: `brew install --cask libreoffice`
