# QA Excel template rules

## Non-negotiable safety rules

- Treat `C:\QA-Agent\templates\QA_Report_Template.xlsx` as read-only. Never overwrite or save back to it.
- Inspection is read-only. Do not save a workbook, create an output workbook, change styles, edit formulas, or modify any byte of the template during inspection.
- For every explicitly authorized write workflow, first use `shutil.copy2` to copy the template into `C:\QA-Agent\output\`, then load, edit, and save only that output copy.
- Refuse a write when the resolved input and output paths are the same or when the output is inside a directory named `templates`, including `C:\QA-Agent\templates\`.
- Never load the source template as the editable workbook in a write workflow. Never call `save()` with the source template path.
- Load with `openpyxl.load_workbook(..., read_only=False, data_only=False)` so formulas, merged cells, styles, drawings, and embedded images remain inspectable. Never call `wb.save()` during inspection.

## Block detection

- Detect worksheets by their content and headers; do not rely only on fixed sheet positions or fixed row numbers.
- On the testcase worksheet, the current test block begins at the last occurrence of the `TC_ID` header. Never hard-code its row.
- On the bug worksheet, the current bug block begins at the last occurrence of the `Bug ID` header. Never hard-code its row.
- Detect the latest summary block from its content and repeated summary/KPI labels. Do not assume a fixed summary row.
- Determine current block data bounds from populated identifier cells beneath the detected header rather than from a preselected row range.

## Preservation and write requirements

- Use `scripts/write_results.py` as the primary writer for a validated `qa-result.json`. It upserts the current `Testcase` and `NoteBug` blocks, then updates only label-identified KPI value cells in the current `Summary` block.

- Preserve all embedded screenshot/image evidence and existing drawings.
- Preserve the template's current styles, merged cells, conditional formatting, row/column visibility, and freeze panes unless a future request explicitly changes them.
- Do not insert or delete rows and do not add, remove, or alter merged ranges unless the user explicitly authorizes that structural change.
- Do not write evidence into `Log/Screenshot/Video` unless the user explicitly requests it.
- For an existing testcase ID, update only `Actual Result`, `Log/Screenshot/Video`, and `Result`. Preserve testcase definition fields plus `Phản hồi` and `FEEDBACK`; reject an Expected Result mismatch.
- For an existing Bug ID, update that row using the authoritative `NoteBug` mapping. Reject duplicate identifiers in either current block.
- Recompute Total, Pass, Fail, Unverified, bug count, and pass rate from the current workbook blocks. Do not depend on cached formula results or old formula ranges.
- Find Summary KPI value cells by label. A KPI formula may be replaced by its computed static value only in the identified KPI cell; preserve its existing style and number format.
- New Pass/Fail values must inherit the corresponding Pass/Fail style already present in the template; do not invent replacement formatting.
- When a destination row already has styles, preserve those styles and assign values only. When it is completely unstyled, use `copy.copy` to copy `font`, `fill`, `border`, `alignment`, `number_format`, and `protection` from the preceding testcase row, and copy row height when needed.
- Preserve existing formulas except for Summary KPI formulas intentionally replaced by computed static values.
- Write only to a copied workbook under `C:\QA-Agent\output\`; never write to the source template.
- After saving, reopen the output and compare sheet names, existing images, merged ranges, formulas, requested values, and the semantic Summary state. Permit formula removal only at intentionally replaced Summary KPI cells. Any other difference makes verification fail.
- The Summary semantic snapshot covers cell coordinates and values/formulas, font, fill, border, alignment, number format, protection, merged ranges, row heights, column widths, hidden rows/columns, image count, and freeze panes. Report the exact cell/property for every unexpected difference.
- During text verification, treat only `None` and `""` as the same blank state; compare every other value without normalization.
- A write is successful only after requested values are verified in the reopened output and required sheets remain present.

## Dry-run report expectations

Report workbook sheet names and dimensions; current testcase and bug header/data bounds and column maps; testcase, Pass, Fail, and bug counts when readable; latest summary block and KPI formulas; suspicious references to older blocks; and per-sheet merged cells, formulas, images, conditional formatting, freeze panes, styled cells, and hidden rows/columns. Explicitly state whether embedded images were detected.
