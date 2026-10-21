"""Create a fresh QA report from the template design and one qa-result.json.

The source template is validated and copied first. Only the output copy is
loaded for editing. Historical cell values and evidence images are removed
from the copy while styles, dimensions, merged ranges, and sheet structure are
retained. The validated Excel writer is then reused for data mapping, image
insertion, KPI calculation, and verification.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Sequence

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


WRITER_DIR = Path(r"C:\QA-Agent\.agents\skills\excel-report\scripts")
sys.path.insert(0, str(WRITER_DIR))
import write_results as wr  # noqa: E402


def clear_values(ws: object) -> None:
    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None


def populate_fresh_design(workbook: object, document: dict) -> None:
    testcases = document["testcases"]
    bugs = document["bugs"]
    run = document["run"]
    summary = document["summary"]

    for sheet_name in (wr.TESTCASE_SHEET, wr.BUG_SHEET, wr.SUMMARY_SHEET):
        if sheet_name not in workbook.sheetnames:
            raise wr.ReportError(f"Required sheet not found: {sheet_name}")

    testcase_ws = workbook[wr.TESTCASE_SHEET]
    bug_ws = workbook[wr.BUG_SHEET]
    summary_ws = workbook[wr.SUMMARY_SHEET]

    # Capture existing result styles before all historical values are cleared.
    pass_style_row = 27
    fail_style_row = 38

    for ws in (testcase_ws, bug_ws, summary_ws):
        clear_values(ws)
        ws._images = []

    finished_date = datetime.fromisoformat(run["finished_at"]).replace(tzinfo=None)

    testcase_ws["A20"] = finished_date
    testcase_headers = {
        "A26": "TC_ID",
        "B26": "Module",
        "C26": "Test Title",
        "D26": "Description/Test Summary",
        "E26": "Test Steps",
        "F26": "Test Data",
        "G26": "Expected Result",
        "H26": "Actual Result",
        "I26": "Log/Screenshot/Video",
        "L26": "Result",
        "M26": "Phản hồi",
        "N26": "FEEDBACK",
    }
    for coordinate, value in testcase_headers.items():
        testcase_ws[coordinate] = value

    bug_ws["A8"] = finished_date
    bug_headers = {
        "A14": "Bug ID",
        "B14": "Người report",
        "C14": "Thiết bị",
        "D14": "OS/App version",
        "E14": "Module",
        "F14": "Testcase liên quan",
        "G14": "Tiêu đề bug",
        "H14": "Bước tái hiện",
        "I14": "Kết quả thực tế",
        "J14": "Kết quả mong muốn",
        "K14": "Log/Screenshot/Video",
    }
    for coordinate, value in bug_headers.items():
        bug_ws[coordinate] = value

    testcase_summary = "; ".join(
        f"{case['tc_id']} — {case['test_title']} ({case['status']})"
        for case in testcases
    )
    bug_ids = ", ".join(bug["bug_id"] for bug in bugs) if bugs else "Không có"
    linked_ids = ", ".join(bug["linked_tc_id"] for bug in bugs) if bugs else "Không có"
    bug_titles = "; ".join(bug["title"] for bug in bugs) if bugs else "Không ghi nhận bug."

    summary_values = {
        "A35": "MIDESK PHP 8.5 - EVENT TEST SUMMARY",
        "A37": "Hạng mục",
        "B37": "Số lượng",
        "D37": "Nhóm kiểm thử",
        "A38": "Tổng số testcase",
        "A39": "Pass",
        "A40": "Fail",
        "A41": "Bug riêng ghi nhận",
        "A42": "Tỷ lệ Pass",
        "D38": "1. Sự kiện",
        "E38": summary["total"],
        "F38": testcase_summary,
        "A45": "BUG SUMMARY",
        "A46": "Nhóm lỗi",
        "B46": "Số lượng",
        "D46": "Bug ID",
        "E46": "Testcase liên quan",
        "F46": "Mô tả ngắn",
        "A47": "Sự kiện",
        "B47": summary["bugs"],
        "D47": bug_ids,
        "E47": linked_ids,
        "F47": bug_titles,
        "A50": "Phạm vi bao phủ",
        "B50": (
            f"{run['target']}. Chế độ {run['mode']}; trình duyệt {run['browser']}. "
            "Phạm vi chỉ gồm module Sự kiện và bằng chứng của lần chạy hiện tại."
        ),
        "A53": "Kết quả tổng quan",
        "B53": (
            f"Đã thực thi {summary['total']} testcase: {summary['pass']} PASS, "
            f"{summary['fail']} FAIL, {summary['unverified']} UNVERIFIED."
        ),
        "A56": "Ghi chú Bug",
        "B56": f"Ghi nhận {summary['bugs']} bug: {bug_titles}",
    }
    for coordinate, value in summary_values.items():
        summary_ws[coordinate] = value

    # The normal writer requires an existing visual source for PASS/FAIL cells.
    # Historical values are gone, so provide the preserved template style rows
    # directly without restoring any historical identifiers or results.
    original_find_status_style_row = wr.find_status_style_row

    def fresh_status_style_row(
        ws: object,
        header_row: int,
        target_row: int,
        tc_id_column: int,
        result_column: int,
        status: str,
    ) -> int | None:
        normalized = wr.normalized(status)
        if normalized == "pass":
            return pass_style_row
        if normalized == "fail":
            return fail_style_row
        return original_find_status_style_row(
            ws, header_row, target_row, tc_id_column, result_column, status
        )

    wr.find_status_style_row = fresh_status_style_row


def verify_no_historical_content(workbook: object, current_ids: set[str], current_bugs: set[str]) -> list[str]:
    failures: list[str] = []
    testcase_ws = workbook[wr.TESTCASE_SHEET]
    bug_ws = workbook[wr.BUG_SHEET]

    found_tc_ids = {
        str(cell.value)
        for cell in testcase_ws["A"]
        if cell.value not in (None, "", "TC_ID") and not isinstance(cell.value, datetime)
    }
    found_bug_ids = {
        str(cell.value)
        for cell in bug_ws["A"]
        if cell.value not in (None, "", "Bug ID") and not isinstance(cell.value, datetime)
    }
    if found_tc_ids != current_ids:
        failures.append(
            f"Testcase sheet contains unexpected/missing IDs: expected={sorted(current_ids)}, "
            f"actual={sorted(found_tc_ids)}"
        )
    if found_bug_ids != current_bugs:
        failures.append(
            f"NoteBug sheet contains unexpected/missing IDs: expected={sorted(current_bugs)}, "
            f"actual={sorted(found_bug_ids)}"
        )
    return failures


def run(input_path: Path, template_path: Path, output_path: Path) -> bool:
    # Full contract validation occurs before any copy or workbook write.
    validated = wr.load_and_validate_json(input_path)
    template, output = wr.validate_paths(template_path, output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    template_hash = wr.file_sha256(template)
    shutil.copy2(template, output)
    if wr.file_sha256(output) != template_hash:
        raise wr.ReportError("Copied output does not match the source template")

    workbook = load_workbook(output, read_only=False, data_only=False)
    try:
        populate_fresh_design(workbook, validated.document)
        before = wr.snapshot(workbook)
        testcase_records, testcase_header_row, testcase_headers = wr.write_testcases(
            workbook, validated
        )
        bug_records, bug_header_row, bug_headers = wr.write_bugs(workbook, validated)
        summary_record = wr.write_summary(workbook)
        workbook.save(output)
    finally:
        workbook.close()

    if wr.file_sha256(template) != template_hash:
        raise wr.ReportError("Template content changed during report generation")

    reopened = load_workbook(output, read_only=False, data_only=False)
    try:
        after, failures, summary_preserved, unexpected_formulas = wr.verify_output(
            reopened,
            before,
            testcase_records,
            testcase_header_row,
            testcase_headers,
            bug_records,
            bug_header_row,
            bug_headers,
            summary_record,
        )
        failures.extend(
            verify_no_historical_content(
                reopened,
                {case["tc_id"] for case in validated.document["testcases"]},
                {bug["bug_id"] for bug in validated.document["bugs"]},
            )
        )
    finally:
        reopened.close()

    wr.print_report(
        validated.path,
        template,
        output,
        testcase_records,
        bug_records,
        before,
        after,
        summary_record,
        summary_preserved,
        unexpected_formulas,
        failures,
    )
    print("- historical template content removed from output: " + ("NO" if failures else "YES"))
    return not failures


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--template", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        return 0 if run(args.input, args.template, args.output) else 1
    except wr.ReportError as exc:
        print("FRESH REPORT WRITE", file=sys.stderr)
        print("- verification: FAIL", file=sys.stderr)
        print(f"- error: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print("FRESH REPORT WRITE", file=sys.stderr)
        print("- verification: FAIL", file=sys.stderr)
        print(f"- error: unexpected failure: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
