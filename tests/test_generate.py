import pytest
from pathlib import Path

def test_read_and_group_excel():
    """엑셀 파일을 읽어 client_name 기준으로 그룹핑한다."""
    from generate import read_excel_data

    groups = read_excel_data(Path("sample-data.xlsx"))

    assert len(groups) == 2  # 테스트매체A, 테스트매체B

    group_a = groups["테스트매체A"]
    assert group_a["client_name"] == "테스트매체A"
    assert group_a["client_address"] == "서울시 강남구 역삼동 123"
    assert group_a["client_email"] == "a@test.com"
    assert group_a["client_manager"] == "김철수"
    assert group_a["gross_rate"] == "50%"
    assert len(group_a["widgets"]) == 3

    widget0 = group_a["widgets"][0]
    assert widget0["service"] == "매체A-1"
    assert widget0["service_name"] == "a1.example.com"
    assert widget0["widget_name"] == "위젯Alpha"
    assert widget0["value"] == "CPM 1,000원"
    assert widget0["date_start"] == "2026-04-01"

    group_b = groups["테스트매체B"]
    assert len(group_b["widgets"]) == 1
