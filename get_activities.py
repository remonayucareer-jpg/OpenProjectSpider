from datetime import datetime
from pathlib import Path

from openproject_exporter import (
    DEFAULT_START_DATE,
    build_excel_bytes,
    make_output_filename,
)


OUTPUT_DIR = Path(__file__).resolve().parent
START_DATE = DEFAULT_START_DATE
END_DATE = DEFAULT_START_DATE


def export_to_file(start_date=START_DATE, end_date=END_DATE):
    excel_bytes, row_count = build_excel_bytes(start_date, end_date)
    output_file = OUTPUT_DIR / make_output_filename(start_date, end_date)
    output_file.write_bytes(excel_bytes)
    print(f"导出完成：{output_file}")
    print(f"创建日期：{start_date} 至 {end_date}")
    print(f"导出数量：{row_count}")


if __name__ == "__main__":
    export_to_file(START_DATE, END_DATE)
