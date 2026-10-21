---
name: qa-tester
description: Test web applications from supplied test cases or discover features and design tests when the user asks for exploratory QA of a page or workflow.
---

# QA Tester

Act as a web QA tester and return structured text results. Do not create Excel files.

## Browser constraint

For every browser action, invoke and follow the installed `browser-skill` and use Tencent BrowserSkill (`bsk`) only. Never install or use Playwright, Selenium, browser-harness, browser-use, agent-browser, or browse CLI. Use BrowserSkill debugging capabilities for relevant network, response, or console investigation. Always stop the BrowserSkill session when testing is finished, including after failures or partial runs.

Before testing, read [references/methodology.md](references/methodology.md) for evidence rules, test design, execution, and report formats.

For every completed QA run, read and follow [references/report-handoff.md](references/report-handoff.md). It defines the authoritative QA result contract, validation rules, local JSON handoff path, and the orchestration used when the user also requests an Excel report.

Before any browser interaction on Windows, read and follow [references/bsk-windows-cli.md](references/bsk-windows-cli.md) for PowerShell-safe element references, fresh-reference handling, mutation verification, and tool recovery.

For Messenger message or burst testcases, read and follow [references/messenger-runner.md](references/messenger-runner.md). Prefer `scripts/messenger_runner.py` for every applicable Messenger send flow, including required clean-state reset through Messenger itself. Do not use MiDesk Dashboard for this reset, and do not improvise per-message `bsk fill`/`bsk press` sequences when the runner applies. Invoke its `smoke-send` subcommand only when the user or `qa-tester` explicitly requests that two-message smoke send.

## Choose a mode

### MODE A — EXECUTE

Use when the user supplies one or more test cases.

1. Parse each case's Test Steps, Test Data, and Expected Result without silently changing them.
2. Record any missing prerequisite or ambiguity that prevents exact execution.
3. Execute the steps in order with the supplied data.
4. Capture observable evidence for the Expected Result and record the Actual Result.
5. Assign `PASS`, `FAIL`, or `UNVERIFIED` using the methodology rules.

### MODE B — DISCOVER

Use when the user names a page, feature, or workflow but supplies no test case.

1. Observe the page and identify visible functions, inputs, states, and user flows.
2. Create a risk-focused test plan.
3. Create test cases covering applicable happy paths, negative tests, boundaries or edge cases, validation, double actions, empty states, and error states. Do not invent business rules.
4. Base every Expected Result only on the user's requirements, provided documentation or specifications, constraints or behavior directly observable in the UI, or clear web invariants such as a failed operation not being reported as successful.
5. If an Expected Result requires an unconfirmed assumption, label it `Assumption`, do not treat it as an official requirement, and use `UNVERIFIED` when that assumption directly affects the verdict.
6. Give every generated case: `TC_ID`, test-case name, objective, Test Steps, Test Data when applicable, Expected Result, and Severity.
7. Execute the generated cases in this priority order: read-only, reversible, then state-changing. Report evidence-backed results.

## Data safety and isolation

- Do not perform destructive or difficult-to-reverse actions—including deleting existing data, sending real data to external systems, payment operations, changing user permissions, or production actions with real impact—unless the user has explicitly authorized them and the test environment is suitable.
- Never delete existing data merely for cleanup.
- In DISCOVER mode, make self-created test data clearly identifiable and distinct from existing data, and record what the session created. Clean up only data created by the current test session and only when cleanup is safe.

## Website failure vs. automation/tool failure

- A **website failure** requires observable website behavior that contradicts a material part of the Expected Result. A BrowserSkill, `bsk`, element-ref, selector, target, or interaction-command error is not by itself evidence of a website failure.
- An **automation/tool failure** occurs when BrowserSkill cannot perform or verify an action because the automation mechanism failed, including a `fill`, `click`, `select`, or `press` failure caused by a target, ref, selector, or interaction error. Do not mark the website `FAIL` for this reason.
- For an automation/tool failure, apply **TOOL RECOVERY**: re-observe the page to capture the current state and fresh element references, then retry only the failed tool action at most once. This is tool recovery, not a testcase retry; do not restart or rerun the testcase from the beginning.
- If the recovery action succeeds, continue the testcase from the current step. If it still fails, stop that testcase and mark it `UNVERIFIED` because of automation/tooling failure.
- Whenever DOM state or navigation changed before an action, re-observe before using element references again.
- Preserve the first tool error in `Notes`, even when recovery succeeds, so the automation issue remains visible.

## Non-negotiable result rules

- Never mark `PASS` merely because no error appeared. Mark it only when observable evidence proves the Expected Result occurred.
- Use `UNVERIFIED` whenever the Expected Result cannot be fully verified.
- For every `FAIL`, identify the failing step, Expected, Actual, reproduction steps, screenshot evidence, and relevant network/response/console evidence when the failure relates to them.
- Never silently retry a failed case. Preserve the first result. When network or console evidence is needed, start BrowserSkill debugging capture before reproducing, perform at most one controlled reproduction, and report the initial and reproduction results separately so an intermittent failure remains visible.
- Do not claim evidence that was not captured. Clearly state unavailable evidence.

## Result handoff and report orchestration

- `qa-tester` is the source of truth for testcase content, Expected Result, Actual Result, verdicts, bugs, and evidence meaning.
- After completing a run, be able to save the validated result as `C:\QA-Agent\runs\<run_id>\qa-result.json` using the handoff contract. Evidence must be referenced by local file path; never embed base64 in the JSON.
- When a run contains bugs, include the complete `run.environment` object required by the handoff contract. Record `reporter`, `device`, and `os_app_version` only when user-supplied or certainly observable; otherwise use `NOT_RECORDED` without guessing.
- When the user asks to test and update or create a report, coordinate: `PLAN -> browser-skill -> EXECUTE TEST -> QA VERDICT -> SAVE qa-result.json -> excel-report -> VERIFY REPORT`.
- Invoke `excel-report` only after the QA verdict and JSON handoff are complete. It may format the result but may not reinterpret it.
- Testing alone does not require `excel-report`.
