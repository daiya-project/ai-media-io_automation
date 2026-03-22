from pathlib import Path
from collections import OrderedDict
import openpyxl

COMMON_FIELDS = ["client_name", "client_address", "client_email", "client_manager", "gross_rate"]
WIDGET_FIELDS = ["service", "service_name", "widget_name", "value", "date_start"]


def read_excel_data(excel_path: Path) -> OrderedDict:
    """엑셀 파일을 읽어 client_name 기준으로 그룹핑하여 반환한다."""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    groups = OrderedDict()

    for row in ws.iter_rows(min_row=2, values_only=True):
        row_dict = dict(zip(headers, row))
        name = str(row_dict["client_name"])

        if name not in groups:
            groups[name] = {
                field: str(row_dict[field] or "") for field in COMMON_FIELDS
            }
            groups[name]["widgets"] = []

        widget = {field: str(row_dict[field] or "") for field in WIDGET_FIELDS}
        groups[name]["widgets"].append(widget)

    return groups
