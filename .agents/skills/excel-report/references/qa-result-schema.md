# QA Result Contract

This document is the handoff contract between `qa-tester` and `excel-report`. Both skills must use this logical structure without changing its meaning.

## Ownership boundary

- `qa-tester` is the source of truth for test cases, Expected Result, Actual Result, `PASS | FAIL | UNVERIFIED` verdicts, bugs, and the meaning of evidence.
- `excel-report` is a presentation consumer. It may validate and format a QA result, but it must not change a verdict or invent an Expected Result, Actual Result, bug, or evidence meaning.
- Browser interaction belongs to `qa-tester` through `browser-skill`. `excel-report` must contain no browser-testing logic.

## Logical data shape

```json
{
  "run": {
    "run_id": "...",
    "target": "...",
    "mode": "EXECUTE | DISCOVER",
    "browser": "...",
    "started_at": "...",
    "finished_at": "...",
    "environment": {
      "reporter": "... | NOT_RECORDED",
      "device": "... | NOT_RECORDED",
      "os_app_version": "... | NOT_RECORDED"
    }
  },
  "testcases": [
    {
      "tc_id": "...",
      "module": "...",
      "test_title": "...",
      "description": "...",
      "test_steps": ["..."],
      "test_data": "...",
      "expected_result": "...",
      "actual_result": "...",
      "severity": "...",
      "status": "PASS | FAIL | UNVERIFIED",
      "evidence": [
        {
          "type": "screenshot | network | console | url | observation",
          "path": "...",
          "description": "..."
        }
      ],
      "notes": "..."
    }
  ],
  "bugs": [
    {
      "bug_id": "...",
      "linked_tc_id": "...",
      "module": "...",
      "title": "...",
      "severity": "...",
      "reproduction_steps": ["..."],
      "actual_result": "...",
      "expected_result": "...",
      "evidence": []
    }
  ],
  "summary": {
    "total": 0,
    "pass": 0,
    "fail": 0,
    "unverified": 0,
    "bugs": 0
  }
}
```

## Required fields and validation

- Every property shown in the logical data shape is required except that `run.environment` is conditionally required as described below. Arrays may be empty only when the run legitimately has no corresponding items. `notes` and `test_data` may be empty strings when not applicable, but the fields must still be present.
- When `bugs.length > 0`, `run.environment` must exist and must contain all three fields: `reporter`, `device`, and `os_app_version`. A value may be `NOT_RECORDED`, but none of these fields may be absent.
- When `bugs.length === 0`, `run.environment` may be omitted. If it is present, it must still contain all three fields.
- `reporter`, `device`, and `os_app_version` must not be invented or guessed. `qa-tester` may record a value only when the user supplied it or when it can be observed or determined with certainty from the test environment. Otherwise it must use the exact string `NOT_RECORDED`.
- `run.mode` must be exactly `EXECUTE` or `DISCOVER`.
- `testcases[*].status` must be exactly `PASS`, `FAIL`, or `UNVERIFIED`.
- `evidence[*].type` must be exactly `screenshot`, `network`, `console`, `url`, or `observation`.
- Every evidence `path` must be a local filesystem path to an evidence file. Do not embed base64 data in JSON. A page address belongs in the evidence description or in a local evidence artifact referenced by `path`; the `url` type does not permit a remote URL in `path`.
- Each `bugs[*].linked_tc_id` must match an existing `testcases[*].tc_id`.
- Summary counts must agree with the testcase statuses and the number of bug records: `total = testcases.length`, `pass + fail + unverified = total`, and `bugs = bugs.length`.
- If a required field is absent, an enum is invalid, a link is unresolved, an evidence path is not local, or summary counts are inconsistent, `excel-report` must stop and report a validation error. It must not guess, repair, synthesize, or reinterpret QA data.

## Producer requirements (`qa-tester`)

After a completed run, `qa-tester` must be able to create:

```text
C:\QA-Agent\runs\<run_id>\qa-result.json
```

The saved JSON must conform to this contract and contain the verdicts and evidence meanings decided by `qa-tester`.

For a run containing bugs, `qa-tester` must populate the complete `run.environment` object before handoff. It must not infer a reporter name, device, operating system, or application version from uncertain context.

## Consumer requirements (`excel-report`)

- Treat a conforming `qa-result.json` as immutable QA input.
- Preserve all verdicts verbatim. For example, a testcase supplied as `FAIL` must remain `FAIL` in the report.
- Do not create Expected Result, Actual Result, bug, or evidence content absent from the input.
- Report validation errors before any workbook write. Validation failure produces no report update.
- Upsert the current workbook blocks by exact `tc_id` and `bug_id`. Reject duplicate identifiers already present in a current block.
- For an existing testcase, preserve the testcase definition and update only execution fields. If JSON `expected_result` differs from the workbook value (except `None`/empty-string blank equivalence), stop with a definition mismatch.
- After testcase and bug upserts, compute Summary KPIs from all records in the current workbook blocks rather than from the JSON subset.

## `NoteBug` mapping

For each item in `bugs`, `excel-report` uses the following mapping and no substitute columns:

| Contract value | `NoteBug` column |
| --- | --- |
| `bug.bug_id` | `Bug ID` |
| `run.environment.reporter` | `Người report` |
| `run.environment.device` | `Thiết bị` |
| `run.environment.os_app_version` | `OS/App version` |
| `bug.module` | `Module` |
| `bug.linked_tc_id` | `Testcase liên quan` |
| `bug.title` | `Tiêu đề bug` |
| `bug.reproduction_steps` | `Bước tái hiện` |
| `bug.actual_result` | `Kết quả thực tế` |
| `bug.expected_result` | `Kết quả mong muốn` |
| `bug.evidence` | `Log/Screenshot/Video` |

- Render `bug.reproduction_steps` as a one-based numbered list with one step per line: `1. ...`, `2. ...`, `3. ...`.
- Preserve `bug.severity` in JSON. The current `NoteBug` template has no `Severity` column, so `excel-report` must not place severity in any other column.
- The testcase fields `Phản hồi` and `FEEDBACK` are not part of the QA verdict contract. A new row leaves them blank and an updated row preserves their existing values unless a separate workflow explicitly supplies content; neither skill may synthesize it.

## Orchestration

When the user requests both testing and report creation or update, `qa-tester` coordinates this exact sequence:

```text
PLAN
-> browser-skill
-> EXECUTE TEST
-> QA VERDICT
-> SAVE qa-result.json
-> excel-report
-> VERIFY REPORT
```

- If the user requests testing without a report, `qa-tester` does not have to invoke `excel-report`.
- If the user requests an Excel update from an existing `qa-result.json`, `excel-report` may run independently after validating the file.

