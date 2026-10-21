---
name: excel-report
description: Inspect and safely write project QA Excel reports based on the existing template, with read-only analysis and mandatory copy-only output workflows.
---

# Excel Report

Use this skill for QA reports based on `C:\QA-Agent\templates\QA_Report_Template.xlsx`.

Before any workbook work, read [references/template-rules.md](references/template-rules.md).

Before consuming QA results, read and follow [references/qa-result-schema.md](references/qa-result-schema.md). Validate the entire `qa-result.json` before any workbook write. Stop and report validation errors for missing or invalid required data; never guess or repair QA content.

For read-only template analysis, use [scripts/inspect_template.py](scripts/inspect_template.py) with `C:\QA-Agent\.venv\Scripts\python.exe`. The inspection workflow must never save a workbook.

For an explicitly authorized write smoke test, use [scripts/write_smoke_test.py](scripts/write_smoke_test.py). Every write workflow must first copy the source template into `C:\QA-Agent\output\` and must load, edit, and save only that output copy. Never load the source template for a write workflow, never save back to the template, and never place a generated workbook under `C:\QA-Agent\templates\`.

For the primary JSON-to-report workflow, use [scripts/write_results.py](scripts/write_results.py). It validates the complete `qa-result.json` contract before copying the template, upserts records in the current `Testcase` and `NoteBug` blocks, computes current-block KPI values in `Summary`, and reopens the result for semantic preservation and value verification. Follow the contract's authoritative `NoteBug` mapping; never claim a row or KPI was updated unless output verification confirms it.

Do not run a write script merely because it exists. Run it only when the user explicitly requests execution. Before treating a write as successful, reopen the output and verify workbook structure, images, merged ranges, formulas, required sheets, and the requested written values.

## QA ownership boundary

- Treat `qa-tester` output as the source of truth. Preserve testcase data, Expected Result, Actual Result, `PASS | FAIL | UNVERIFIED` verdicts, bugs, and evidence meaning exactly as supplied.
- Do not create missing Expected Result, Actual Result, bug, or evidence content. A `FAIL` supplied by `qa-tester` must remain `FAIL` in the Excel report.
- Evidence references must be local file paths and must not contain base64 data.
- For `NoteBug`, require complete `run.environment` metadata whenever bugs exist, apply the documented column mapping, render reproduction steps as a numbered list, and never guess environment values or repurpose another column for bug severity.
- Do not write testcase columns `Phản hồi` and `FEEDBACK` unless a separate workflow explicitly supplies them. New rows keep them blank; existing rows preserve their current values.
- When a testcase already exists, preserve its definition fields and update only `Actual Result`, `Log/Screenshot/Video`, and `Result`. Reject an Expected Result mismatch instead of overwriting the workbook definition.
- Treat only `None` and `""` as equivalent blank values during verification. Do not normalize other values.
- Compute Summary totals from the current workbook blocks, locate KPI value cells by their labels, and preserve non-KPI Summary content and formatting semantically.
- This skill may run independently when the user supplies an existing `qa-result.json` and asks only for an Excel update.
- Do not perform browser interaction or testing. Browser work remains in `qa-tester` through `browser-skill`.
