# Web QA Methodology

Use this reference for both supplied and generated test cases.

## 1. Prepare and observe

- Confirm the target page, starting state, account or role, prerequisites, and available test data from the user's request and current browser state.
- Do not invent credentials, permissions, business rules, hidden state, or data. If a missing item blocks a case, record the case as `UNVERIFIED` with the reason.
- Establish the observable precondition before executing a case. Avoid letting one case contaminate another; restore the required state when possible.
- Use `browser-skill` with Tencent BrowserSkill (`bsk`) for all browser interaction. Follow its own session and debugging instructions.

## 2. Design tests in DISCOVER mode

First inventory visible functionality: navigation, actions, forms, field constraints, tables or lists, search and filters, pagination, dialogs, notifications, permissions, loading indicators, and empty or error presentations. Prioritize by user impact and likelihood of failure.

### Source of truth for Expected Results

Do not invent business rules. Derive each Expected Result only from one or more of these sources:

- requirements supplied by the user;
- provided documentation or specifications;
- constraints or behavior directly observable in the UI;
- clear web invariants, such as a failed operation not being reported as successful.

When an Expected Result needs an unconfirmed assumption, label it explicitly as `Assumption` and do not present it as an official requirement. If the assumption directly affects the verdict, use `UNVERIFIED` rather than resolving the uncertainty as `PASS` or `FAIL`.

Apply the techniques that fit the feature:

- **Happy path:** valid common input completes the main user goal.
- **Negative testing:** invalid formats, forbidden actions, incompatible values, or rejected operations produce safe and useful behavior.
- **Boundary and edge cases:** test documented or observable minima, maxima, just-inside and just-outside values, long text, special characters, whitespace, and relevant date or numeric limits.
- **Validation:** required fields, format checks, inline messages, focus behavior, preservation of entered values, and prevention of invalid submission.
- **Double action:** where duplicate effects are risky, try a rapid double click or repeated submission and verify that only the intended number of effects occurs. Do not use this on destructive or costly actions unless the user's authorization and a safe test environment make it appropriate.
- **Empty state:** no records, no search results, or optional values absent should remain understandable and actionable.
- **Error state:** when a safe way to induce or observe an error exists, verify the message, recovery path, data integrity, and absence of misleading success feedback. Do not disrupt external systems merely to force an error.

Each generated case must include:

```text
TC_ID:
Tên testcase:
Mục tiêu:
Test Steps:
Test Data: (omit only when genuinely not applicable)
Expected Result:
Severity: Critical | High | Medium | Low
```

Severity reflects impact, not execution status:

- `Critical`: security, irreversible loss, broad outage, or the primary service is unusable.
- `High`: a core workflow is blocked or produces materially wrong data with no practical workaround.
- `Medium`: an important behavior is wrong but a workaround exists or impact is limited.
- `Low`: minor usability, copy, or visual problem with little functional impact.

### Safe test selection and test-data isolation

Prioritize cases in this order: read-only, reversible, then state-changing.

Do not perform destructive or difficult-to-reverse actions unless the user has explicitly authorized them and the test environment is suitable. This includes deleting existing data, sending real data to external systems, payment operations, changing user permissions, and production actions with real impact. Do not delete existing data merely to clean up after testing.

When creating test data in DISCOVER mode:

- use values that are clearly identifiable as test data;
- avoid collisions with existing data;
- record exactly what the current test session created;
- clean up only data created by the current test session, and only when cleanup is safe.

## 3. Execute precisely

- In EXECUTE mode, preserve the supplied order, values, and Expected Result. Do not silently repair a questionable case or substitute data.
- Record the outcome of each material step. If execution diverges, identify where and why.
- In DISCOVER mode, execute the generated cases in a safe order: read-only, reversible, then state-changing; execute state-changing cases only within the user's scope and the safety constraints above.
- Do not silently retry a failed case. Always preserve the first result.
- If network or console evidence is needed, start BrowserSkill debugging capture before reproducing the behavior. Perform at most one controlled reproduction when needed, report the first-run and reproduction results separately, and do not let the reproduction hide an intermittent failure.

### Website failure vs. automation/tool failure

Keep failures of the application under test separate from failures of the automation mechanism:

- **Website failure:** observable website behavior contradicts a material part of the Expected Result. This may support a `FAIL` verdict under the existing verdict rules.
- **Automation/tool failure:** BrowserSkill cannot perform or verify an action because `bsk`, a target, element ref, selector, or interaction command fails. Examples include `fill`, `click`, `select`, or `press` failing at the tooling/interaction layer. This is not evidence that the website failed and must not, by itself, produce a `FAIL` verdict.

Apply this **TOOL RECOVERY** procedure to an automation/tool failure:

1. Preserve the first tool error for the case's `Notes`.
2. Re-observe the page to capture its current state and obtain fresh element references.
3. Retry only the specific tool action that failed, at most once. Do not restart or rerun the testcase from the beginning.
4. This retry is tool recovery, not a testcase retry or the controlled reproduction of a website result described above.
5. If the recovered action succeeds, continue the testcase from the current step.
6. If the recovered action still fails, stop the testcase and assign `UNVERIFIED`, explicitly stating that an automation/tooling failure prevented verification.

If navigation or a DOM change occurred before the next action, always re-observe before using element references again. Never attribute a stale-ref or interaction-command failure to the website without independent observable website evidence.

## 4. Evidence and verdicts

Evidence must be observable and tied to the Expected Result. Depending on the case, acceptable evidence includes a changed UI state, exact visible message, persisted record after refresh or navigation, correct list or detail data, URL or navigation change, enabled or disabled control, download presence, or relevant request and response details.

Use these verdicts:

- `PASS`: every material part of the Expected Result is directly supported by captured observations.
- `FAIL`: an observed result contradicts at least one material part of the Expected Result.
- `UNVERIFIED`: execution or observation cannot establish the full Expected Result, including blocked prerequisites, inaccessible downstream state, missing permissions, ambiguous expectations, or insufficient evidence.

No exception or visible error is required for a failure, and absence of an error is never sufficient for a pass. Partial confirmation remains `UNVERIFIED` unless an observed contradiction makes it `FAIL`.

Capture screenshots at decisive states, especially failures. When behavior suggests an API, loading, JavaScript, or server problem, use BrowserSkill debugging to inspect relevant network requests, response status or payload, and console entries. Report only evidence actually observed. Redact secrets and sensitive values from written results.

## 5. Structured text output

Start with a compact scope and environment summary, then report each case:

```text
TC_ID: <id>
Tên testcase: <name>
Severity: <severity>
Status: PASS | FAIL | UNVERIFIED
Actual Result: <what was observed>
Evidence: <observable proof and screenshot reference>
Notes: <limitations or relevant diagnostics>
```

For `FAIL`, include all of the following:

```text
Bước gây lỗi: <step number and action>
Expected: <expected behavior>
Actual: <observed behavior>
Bước tái hiện:
1. ...
2. ...
Screenshot evidence: <reference or explicit reason unavailable>
Network/response/console evidence: <relevant findings, not applicable, or explicit reason unavailable>
```

For `UNVERIFIED`, state exactly what could not be verified, why, which parts were observed, and what access or evidence would complete verification.

When tool recovery was used, keep the initial BrowserSkill or `bsk` error in `Notes` whether recovery succeeded or failed. If recovery failed, identify the blocked step and the retried action, and state that the verdict reflects automation/tooling failure rather than an observed website contradiction.

End with totals by verdict, a concise defect list, and remaining risks or untested areas. In DISCOVER mode, include the test plan and generated case definitions before or alongside execution results.

## 6. Close the browser session

After the final case—or when testing stops early—stop the BrowserSkill session. Mention session cleanup only if it failed or affected the results.
