# Deterministic Messenger runner

Use `scripts/messenger_runner.py` for Messenger burst testcases. The runner owns
the entire message mutation flow and invokes `bsk.exe` directly with Python
`subprocess` and `shell=False`.

## When to use it

- Prefer this runner whenever a Messenger testcase sends one or more messages,
  especially burst/timing cases.
- Do not improvise a manual sequence of `bsk fill` and `bsk press` calls when
  the runner applies.
- Invoke `smoke-send` only when the user or `qa-tester` explicitly requests the
  two-message smoke action. Merely preparing or reviewing a plan does not
  authorize a send.
- Conversation reset runs directly in Messenger through the same safe message
  flow. Do not use MiDesk Dashboard for this reset.

## Plan JSON

`--plan` is a UTF-8 JSON file containing either one case object or an object
with a `cases` array. A case has this shape:

```json
{
  "case_id": "MSG-BURST-001",
  "reset": {
    "method": "message",
    "text": "hello",
    "required": true,
    "success_indicators": ["Welcome", "Main menu"]
  },
  "messages": [
    {"text": "first message", "delay_after_previous_ms": 0},
    {"text": "second message", "delay_after_previous_ms": 1000}
  ]
}
```

The first message must have a zero delay. Later delays are measured against the
previous message's actual successful Enter timestamp. The execution log records
requested interval, actual interval, and deviation.

The reset defaults are `method: "message"` and `text: "hello"`. The legacy
`clean_conversation: true` flag maps to a required reset with those defaults.
Use `reset.required: false` when clean state is not required.

For a required reset, the runner obtains a fresh observation, sends the reset
text through the same safe SEND_MESSAGE flow, and polls fresh observations for
a newly rendered initial-flow indicator. It does not equate a successful
`hello` send with a successful reset. `success_indicators` should contain
stable, observable text from the chatbot's welcome/menu/start response. If the
plan omits them, conservative English and Vietnamese welcome/menu defaults are
used. Optional positive `response_timeout_ms` and `poll_interval_ms` values
default to 15000 and 500.

If no new indicator appears, the runner returns `PRECONDITION_BLOCKED`, captures
the decision state, and sends no testcase data or later-case data.

## Commands

Run with the project interpreter:

```text
C:\QA-Agent\.venv\Scripts\python.exe C:\QA-Agent\.agents\skills\qa-tester\scripts\messenger_runner.py run-case --session <SESSION> --tab-id <TAB_ID> --plan <CASE_JSON> --evidence-dir <DIR>
```

The smoke command uses the same arguments, but its plan must contain exactly
one case with exactly two messages:

```text
C:\QA-Agent\.venv\Scripts\python.exe C:\QA-Agent\.agents\skills\qa-tester\scripts\messenger_runner.py smoke-send --session <SESSION> --tab-id <TAB_ID> --plan <CASE_JSON> --evidence-dir <DIR>
```

The runner assumes the supplied BrowserSkill session and tab already exist. It
does not start or stop a session. The caller remains responsible for normal
BrowserSkill session cleanup after the QA run.

## Safety and result contract

The runner only permits these BrowserSkill commands: `snapshot`, `observe`,
`focus`, `fill`, `press`, and `screenshot`. Its current send flow uses
`observe`, `focus`, `fill`, `press`, and `screenshot`.

Before Enter, the current composer draft must equal the expected text exactly.
On mismatch the runner clears with `Ctrl+A`, `Backspace`, obtains a fresh ref,
and retries the fill once. A second mismatch returns `TOOL_FAILURE`, stops the
case, and sends no later case messages. Empty/whitespace-only text and the exact
string `.` are rejected before browser mutation.

After Enter, the runner requires both a cleared composer and an increase in the
count of exact matching message bubbles. Failure to prove either condition is a
`TOOL_FAILURE` and no later message is sent.

Each started case writes `precondition.png`, `decision.png`, and
`execution-log.json` beneath its own numbered directory in `--evidence-dir`.
For a required reset, `precondition.png` is captured only after the initial-flow
state has been confirmed.
