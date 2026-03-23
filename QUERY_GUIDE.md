# IO 문서 생성용 로우데이터 추출 가이드

## 개요
IO 문서 자동 생성에 필요한 엑셀 데이터를 DB에서 추출하는 방법을 안내한다.

## 추출 쿼리

아래 쿼리를 실행하여 위젯별 계약 정보를 추출한다.

```sql
SELECT
    C.client_name,
    NULL                                  AS client_address,
    NULL                                  AS client_email,
    NULL                                  AS client_manager,
    NULL                                  AS gross_rate,
    S.service_name_human                  AS service,
    S.service_name                        AS service_name,
    W.widget_name,
    CASE
        WHEN G.income_type = 'cpm'        THEN FLOOR(G.cpm_value)
        WHEN G.income_type = 'grt'        THEN FLOOR(G.grt_value)
        WHEN G.income_type = 'fixed_cpc'  THEN FLOOR(G.fixed_cpc_value)
        WHEN G.income_type = 'cpm_pv'     THEN FLOOR(G.cpm_value)
        ELSE 0
    END                                   AS value,
    DATE_FORMAT(G.start_time, '%Y-%m-%d') AS date_start
FROM dable_media_income.GUARANTEE_SETTING G
INNER JOIN dable.WIDGET W
    ON G.widget_id = W.widget_id
LEFT JOIN dable.SERVICE S
    ON W.service_id = S.service_id
LEFT JOIN dable.CLIENT C
    ON S.client_id = C.client_id
WHERE G.deleted = 0
  AND G.start_time >= '2026-01-01'
  AND C.country = 'KR'
ORDER BY G.start_time DESC, W.widget_id
```

### WHERE 조건 수정
- `G.start_time >= '2026-01-01'` — 추출 기간을 필요에 따라 변경
- `C.country = 'KR'` — 한국 매체만 추출. 다른 국가가 필요하면 변경

## NULL 컬럼 안내

아래 컬럼은 DB에서 자동 추출되지 않으므로, 쿼리 결과를 엑셀에 저장한 후 **수동으로 입력**해야 한다.

| 컬럼 | 설명 | 입력 예시 |
|---|---|---|
| `client_address` | 매체사 주소 | 서울시 중구 세종대로 21길 30 |
| `client_email` | 매체사 이메일 | ad@example.com |
| `client_manager` | 매체사 담당자명 | 홍길동 |
| `gross_rate` | 매체 수수료율 | 55% |

> 같은 `client_name`의 행은 하나의 문서로 그룹핑되므로, 해당 매체사의 모든 행에 동일한 값을 입력한다.

## 엑셀 저장 및 문서 생성

1. 쿼리 결과를 `.xlsx` 파일로 저장하여 `data/` 폴더에 넣는다
2. NULL 컬럼을 수동으로 채운다
3. 문서 생성 실행:
   ```bash
   python3 src/generate.py -i data/<파일명>.xlsx
   ```
