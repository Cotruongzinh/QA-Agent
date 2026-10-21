#!/usr/bin/env python3
"""Deterministic Messenger message runner backed only by Tencent BrowserSkill.

This module intentionally owns the complete focus/fill/verify/press flow.  It
does not start or stop BrowserSkill sessions. Required clean-state reset is
performed directly in Messenger through the same verified message flow.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Sequence


ALLOWED_BSK_COMMANDS = {"snapshot", "observe", "focus", "fill", "press", "screenshot"}
REF_RE = re.compile(r"(?<![A-Za-z0-9_])@?(e\d+)\b", re.IGNORECASE)
TEXTBOX_RE = re.compile(r"\b(textbox|contenteditable)\b", re.IGNORECASE)
VALUE_RE = re.compile(
    r"\b(?:value|input[_ -]?value|text[_ -]?value)\s*[:=]\s*(?:\"((?:\\.|[^\"])*)\"|'([^']*)')",
    re.IGNORECASE,
)
QUOTED_RE = re.compile(r'\"((?:\\.|[^\"])*)\"')
COMPOSER_POSITIVE = (
    "message",
    "messenger",
    "tin nhắn",
    "nhắn tin",
    "write a message",
    "type a message",
    "nhập tin nhắn",
    "composer",
    "chat message",
)
COMPOSER_NEGATIVE = ("search", "tìm kiếm", "find", "filter")
DEFAULT_RESET_INDICATORS = (
    "welcome",
    "main menu",
    "get started",
    "start over",
    "how can i help",
    "xin chào",
    "chào mừng",
    "menu chính",
    "bắt đầu",
    "tôi có thể giúp",
    "vui lòng chọn",
    "hãy chọn",
)


class RunnerError(RuntimeError):
    pass


class ToolFailure(RunnerError):
    pass


class PreconditionBlocked(RunnerError):
    pass


@dataclass(frozen=True)
class CommandResult:
    command: list[str]
    returncode: int
    stdout: str
    stderr: str
    started_monotonic: float
    finished_monotonic: float


@dataclass(frozen=True)
class Composer:
    ref: str
    descriptor: str
    draft_known: bool
    draft: str | None


@dataclass(frozen=True)
class Observation:
    raw: str
    payload: Any | None
    semantic_text: str


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _decode_quoted(value: str) -> str:
    try:
        return json.loads(f'"{value}"')
    except (json.JSONDecodeError, UnicodeDecodeError):
        return value


def _json_payload(stdout: str) -> Any | None:
    stripped = stdout.strip()
    if not stripped:
        return None
    try:
        return json.loads(stripped)
    except json.JSONDecodeError:
        # Some bsk builds prefix a short informational line even with --json.
        for index, char in enumerate(stripped):
            if char not in "[{":
                continue
            try:
                return json.loads(stripped[index:])
            except json.JSONDecodeError:
                continue
    return None


def _collect_strings(value: Any) -> Iterable[str]:
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for child in value.values():
            yield from _collect_strings(child)
    elif isinstance(value, list):
        for child in value:
            yield from _collect_strings(child)


def _semantic_text(stdout: str, payload: Any | None) -> str:
    if payload is None:
        return stdout
    strings = [item for item in _collect_strings(payload) if "\n" in item or REF_RE.search(item)]
    return "\n".join(strings) if strings else stdout


def _dict_ref(node: dict[str, Any]) -> str | None:
    for key in ("ref", "element_ref", "elementRef", "snapshot_ref", "snapshotRef"):
        value = node.get(key)
        if isinstance(value, str):
            match = REF_RE.search(value)
            if match:
                return match.group(1).lower()
    return None


def _walk_dicts(value: Any) -> Iterable[dict[str, Any]]:
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from _walk_dicts(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk_dicts(child)


def _node_descriptor(node: dict[str, Any]) -> str:
    fields: list[str] = []
    for key in (
        "role",
        "type",
        "name",
        "label",
        "placeholder",
        "description",
        "value",
        "text",
    ):
        value = node.get(key)
        if isinstance(value, (str, int, float, bool)):
            fields.append(f"{key}={value}")
    return " ".join(fields)


def _node_is_textbox(node: dict[str, Any]) -> bool:
    descriptor = _node_descriptor(node)
    return bool(TEXTBOX_RE.search(descriptor)) or node.get("contenteditable") is True


def _node_draft(node: dict[str, Any]) -> tuple[bool, str | None]:
    for key in ("value", "input_value", "inputValue", "text_value", "textValue"):
        if key in node and isinstance(node[key], str):
            return True, node[key]
    for key in ("children", "child", "content", "nodes"):
        child = node.get(key)
        if child is None:
            continue
        fragments: list[str] = []
        for descendant in _walk_dicts(child):
            for text_key in ("text", "value"):
                value = descendant.get(text_key)
                if isinstance(value, str) and value:
                    fragments.append(value)
                    break
        if fragments:
            return True, "\n".join(fragments)
    return False, None


def _line_indent(line: str) -> int:
    return len(line) - len(line.lstrip(" \t"))


def _text_draft(lines: list[str], line_index: int) -> tuple[bool, str | None]:
    line = lines[line_index]
    match = VALUE_RE.search(line)
    if match:
        return True, _decode_quoted(match.group(1) if match.group(1) is not None else match.group(2))

    base_indent = _line_indent(line)
    child_text: list[str] = []
    for child in lines[line_index + 1 :]:
        if not child.strip():
            continue
        if _line_indent(child) <= base_indent:
            break
        lower = child.lower()
        if any(token in lower for token in ("statictext", "text ", "paragraph", "generic")):
            matches = QUOTED_RE.findall(child)
            if matches:
                child_text.append(_decode_quoted(matches[-1]))
    if child_text:
        return True, "\n".join(child_text)
    # An empty textbox is represented without a value or text child in VOM.
    return True, ""


def _composer_score(descriptor: str) -> int:
    lowered = descriptor.casefold()
    score = sum(4 for token in COMPOSER_POSITIVE if token in lowered)
    score -= sum(10 for token in COMPOSER_NEGATIVE if token in lowered)
    if "multiline" in lowered:
        score += 2
    return score


def find_composer(observation: Observation) -> Composer:
    candidates: dict[str, Composer] = {}

    if observation.payload is not None:
        for node in _walk_dicts(observation.payload):
            ref = _dict_ref(node)
            if ref and _node_is_textbox(node):
                known, draft = _node_draft(node)
                descriptor = _node_descriptor(node)
                candidates[ref] = Composer(ref, descriptor, known, draft)

    lines = observation.semantic_text.splitlines()
    for index, line in enumerate(lines):
        ref_match = REF_RE.search(line)
        if not ref_match or not TEXTBOX_RE.search(line):
            continue
        ref = ref_match.group(1).lower()
        known, draft = _text_draft(lines, index)
        previous = candidates.get(ref)
        if previous and previous.draft_known:
            known, draft = previous.draft_known, previous.draft
        candidates[ref] = Composer(ref, line.strip(), known, draft)

    if not candidates:
        raise ToolFailure("No Messenger composer textbox was found in the fresh observation")

    ranked = sorted(
        candidates.values(),
        key=lambda item: (_composer_score(item.descriptor), item.ref),
        reverse=True,
    )
    best_score = _composer_score(ranked[0].descriptor)
    best = [item for item in ranked if _composer_score(item.descriptor) == best_score]
    if best_score < 0 or len(best) != 1:
        refs = ", ".join(item.ref for item in best)
        raise ToolFailure(f"Messenger composer is ambiguous in the fresh observation: {refs}")
    return best[0]


def exact_message_count(observation: Observation, expected: str) -> int:
    """Count exact non-textbox message renderings in the semantic observation."""
    structured_count = 0
    if observation.payload is not None:
        for node in _walk_dicts(observation.payload):
            if _node_is_textbox(node):
                continue
            values = {
                value
                for key in ("name", "label", "text", "value", "description")
                if isinstance((value := node.get(key)), str)
            }
            if expected in values:
                structured_count += 1

    line_count = 0
    for line in observation.semantic_text.splitlines():
        if TEXTBOX_RE.search(line):
            continue
        quoted = [_decode_quoted(item) for item in QUOTED_RE.findall(line)]
        if expected in quoted or line.strip() == expected:
            line_count += 1
    # JSON builds may expose either structured VOM nodes or one rendered VOM
    # string. Using the larger representation avoids double-counting both.
    return max(structured_count, line_count)


def indicator_count(observation: Observation, indicator: str) -> int:
    """Count a visible, non-composer initial-flow indicator."""
    needle = indicator.casefold()
    structured_count = 0
    if observation.payload is not None:
        for node in _walk_dicts(observation.payload):
            if _node_is_textbox(node):
                continue
            values = {
                value.casefold()
                for key in ("name", "label", "text", "value", "description")
                if isinstance((value := node.get(key)), str)
            }
            if any(needle in value for value in values):
                structured_count += 1

    line_count = sum(
        1
        for line in observation.semantic_text.splitlines()
        if not TEXTBOX_RE.search(line) and needle in line.casefold()
    )
    return max(structured_count, line_count)


class MessengerRunner:
    def __init__(self, session: str, tab_id: str, evidence_dir: Path) -> None:
        self.session = session
        self.tab_id = tab_id
        self.evidence_dir = evidence_dir.resolve()
        self.bsk = shutil.which("bsk.exe")
        if not self.bsk:
            raise PreconditionBlocked("bsk.exe is not available on PATH")
        self.command_log: list[dict[str, Any]] = []

    def _run_bsk(self, command: str, *arguments: str, timeout: float = 45.0) -> CommandResult:
        if command not in ALLOWED_BSK_COMMANDS:
            raise ToolFailure(f"Blocked non-allowlisted bsk command: {command}")
        argv = [self.bsk, command, *arguments, "--session", self.session, "--tab-id", self.tab_id]
        started = time.monotonic()
        try:
            completed = subprocess.run(
                argv,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                shell=False,
                check=False,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            finished = time.monotonic()
            self.command_log.append(
                {
                    "command": argv,
                    "started_monotonic": started,
                    "finished_monotonic": finished,
                    "exception": repr(exc),
                }
            )
            raise ToolFailure(f"bsk {command} could not complete: {exc}") from exc

        finished = time.monotonic()
        result = CommandResult(argv, completed.returncode, completed.stdout, completed.stderr, started, finished)
        self.command_log.append(
            {
                "command": argv,
                "returncode": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
                "started_monotonic": started,
                "finished_monotonic": finished,
            }
        )
        if result.returncode != 0:
            detail = result.stderr.strip() or result.stdout.strip() or f"exit {result.returncode}"
            raise ToolFailure(f"bsk {command} failed: {detail}")
        return result

    def observe(self) -> Observation:
        result = self._run_bsk("observe", "--json")
        payload = _json_payload(result.stdout)
        return Observation(result.stdout, payload, _semantic_text(result.stdout, payload))

    def screenshot(self, output: Path) -> None:
        output.parent.mkdir(parents=True, exist_ok=True)
        self._run_bsk("screenshot", "--out", str(output), timeout=150.0)
        if not output.is_file() or output.stat().st_size == 0:
            raise ToolFailure(f"bsk screenshot did not create a non-empty file: {output}")

    def focus(self, ref: str) -> None:
        self._run_bsk("focus", "--ref", ref)

    def fill(self, ref: str, text: str) -> None:
        self._run_bsk("fill", "--ref", ref, "--value", text)

    def press(self, key: str, ref: str) -> None:
        self._run_bsk("press", key, "--ref", ref)

    def send_message(self, expected: str) -> dict[str, Any]:
        validate_message_text(expected)
        initial = self.observe()
        composer = find_composer(initial)
        baseline_count = exact_message_count(initial, expected)

        self.focus(composer.ref)
        self.fill(composer.ref, expected)

        verified = self.observe()
        verified_composer = find_composer(verified)
        if not verified_composer.draft_known or verified_composer.draft != expected:
            mismatch = verified_composer.draft
            # Clear only the freshly observed composer, then obtain another fresh
            # ref before the single permitted recovery fill.
            self.press("Ctrl+A", verified_composer.ref)
            self.press("Backspace", verified_composer.ref)
            recovery_observation = self.observe()
            recovery_composer = find_composer(recovery_observation)
            self.focus(recovery_composer.ref)
            self.fill(recovery_composer.ref, expected)
            verified = self.observe()
            verified_composer = find_composer(verified)
            if not verified_composer.draft_known or verified_composer.draft != expected:
                raise ToolFailure(
                    "Draft mismatch after one recovery fill; Enter was not pressed "
                    f"(initial mismatch={mismatch!r}, final={verified_composer.draft!r})"
                )

        # The ref comes from the observation that proved exact draft equality.
        enter_started = time.monotonic()
        self.press("Enter", verified_composer.ref)
        send_timestamp = time.monotonic()

        after = self.observe()
        after_composer = find_composer(after)
        if not after_composer.draft_known or after_composer.draft != "":
            raise ToolFailure(f"Composer did not clear after Enter: {after_composer.draft!r}")
        after_count = exact_message_count(after, expected)
        if after_count <= baseline_count:
            raise ToolFailure("Could not verify a newly rendered bubble with the exact sent text")

        return {
            "text": expected,
            "enter_started_monotonic": enter_started,
            "send_timestamp_monotonic": send_timestamp,
            "baseline_exact_bubble_count": baseline_count,
            "after_exact_bubble_count": after_count,
            "composer_cleared": True,
            "exact_bubble_verified": True,
        }

    def reset_conversation(self, reset: dict[str, Any]) -> dict[str, Any]:
        if reset["method"] != "message":
            raise PreconditionBlocked(f"Unsupported reset method: {reset['method']}")

        try:
            before = self.observe()
        except ToolFailure as exc:
            raise PreconditionBlocked(f"Could not observe the reset precondition: {exc}") from exc
        baseline = {
            indicator: indicator_count(before, indicator)
            for indicator in reset["success_indicators"]
        }
        try:
            reset_send = self.send_message(reset["text"])
        except ToolFailure as exc:
            raise PreconditionBlocked(f"Safe reset message could not be sent and verified: {exc}") from exc

        deadline = time.monotonic() + reset["response_timeout_ms"] / 1000.0
        last_counts = dict(baseline)
        while True:
            try:
                current = self.observe()
            except ToolFailure as exc:
                raise PreconditionBlocked(
                    f"Could not observe and confirm the initial flow after reset: {exc}"
                ) from exc
            last_counts = {
                indicator: indicator_count(current, indicator)
                for indicator in reset["success_indicators"]
            }
            matched = [
                indicator
                for indicator in reset["success_indicators"]
                if last_counts[indicator] > baseline[indicator]
            ]
            if matched:
                return {
                    "method": "message",
                    "text": reset["text"],
                    "send": reset_send,
                    "matched_initial_state_indicators": matched,
                    "baseline_indicator_counts": baseline,
                    "confirmed_indicator_counts": last_counts,
                    "confirmed_at": utc_now(),
                    "confirmed_monotonic": time.monotonic(),
                }
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            time.sleep(min(reset["poll_interval_ms"] / 1000.0, remaining))

        raise PreconditionBlocked(
            "Reset message was sent, but no new observable initial-flow indicator "
            f"appeared before timeout (baseline={baseline}, final={last_counts})"
        )


def validate_message_text(text: Any) -> None:
    if not isinstance(text, str):
        raise ValueError("Each message text must be a string")
    if text == "" or text.strip() == "":
        raise ValueError("Empty or whitespace-only messages are forbidden")
    if text == ".":
        raise ValueError('The message "." is forbidden')


def validate_case(case: Any, index: int) -> dict[str, Any]:
    if not isinstance(case, dict):
        raise ValueError(f"Case {index} must be a JSON object")
    case_id = case.get("case_id", case.get("id", f"case-{index + 1}"))
    if not isinstance(case_id, str) or not case_id.strip():
        raise ValueError(f"Case {index} has an invalid case_id")
    messages = case.get("messages")
    if not isinstance(messages, list) or not messages:
        raise ValueError(f"Case {case_id} must contain a non-empty messages array")
    normalized_messages: list[dict[str, Any]] = []
    for message_index, message in enumerate(messages):
        if not isinstance(message, dict):
            raise ValueError(f"Case {case_id} message {message_index} must be an object")
        text = message.get("text")
        validate_message_text(text)
        delay = message.get("delay_after_previous_ms", 0)
        if isinstance(delay, bool) or not isinstance(delay, (int, float)) or delay < 0:
            raise ValueError(f"Case {case_id} message {message_index} has an invalid delay")
        if message_index == 0 and delay != 0:
            raise ValueError(f"Case {case_id} first message must use delay_after_previous_ms=0")
        normalized_messages.append({"text": text, "delay_after_previous_ms": float(delay)})
    clean = case.get("clean_conversation", False)
    if not isinstance(clean, bool):
        raise ValueError(f"Case {case_id} clean_conversation must be boolean")
    raw_reset = case.get("reset")
    if raw_reset is None:
        raw_reset = {"method": "message", "text": "hello", "required": clean}
    if not isinstance(raw_reset, dict):
        raise ValueError(f"Case {case_id} reset must be an object")
    method = raw_reset.get("method", "message")
    if method != "message":
        raise ValueError(f"Case {case_id} supports only reset.method=message")
    reset_text = raw_reset.get("text", "hello")
    validate_message_text(reset_text)
    required = raw_reset.get("required", clean)
    if not isinstance(required, bool):
        raise ValueError(f"Case {case_id} reset.required must be boolean")
    raw_indicators = raw_reset.get("success_indicators", list(DEFAULT_RESET_INDICATORS))
    if not isinstance(raw_indicators, list) or not raw_indicators:
        raise ValueError(f"Case {case_id} reset.success_indicators must be a non-empty array")
    indicators: list[str] = []
    for indicator in raw_indicators:
        if not isinstance(indicator, str) or not indicator.strip():
            raise ValueError(f"Case {case_id} has an invalid reset success indicator")
        indicators.append(indicator)
    response_timeout_ms = raw_reset.get("response_timeout_ms", 15000)
    poll_interval_ms = raw_reset.get("poll_interval_ms", 500)
    for field_name, value in (
        ("response_timeout_ms", response_timeout_ms),
        ("poll_interval_ms", poll_interval_ms),
    ):
        if isinstance(value, bool) or not isinstance(value, (int, float)) or value <= 0:
            raise ValueError(f"Case {case_id} reset.{field_name} must be positive")
    reset = {
        "method": method,
        "text": reset_text,
        "required": required,
        "success_indicators": indicators,
        "response_timeout_ms": float(response_timeout_ms),
        "poll_interval_ms": float(poll_interval_ms),
    }
    return {
        "case_id": case_id,
        "clean_conversation": required,
        "reset": reset,
        "messages": normalized_messages,
    }


def load_cases(path: Path) -> list[dict[str, Any]]:
    with path.open("r", encoding="utf-8-sig") as handle:
        plan = json.load(handle)
    raw_cases = plan.get("cases") if isinstance(plan, dict) and "cases" in plan else [plan]
    if not isinstance(raw_cases, list) or not raw_cases:
        raise ValueError("Plan cases must be a non-empty array")
    return [validate_case(case, index) for index, case in enumerate(raw_cases)]


def wait_until(deadline: float) -> None:
    while True:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return
        time.sleep(min(remaining, 0.050))


def safe_case_dir_name(case_id: str, index: int) -> str:
    safe = re.sub(r"[^A-Za-z0-9._-]+", "-", case_id).strip(".-") or f"case-{index + 1}"
    return f"{index + 1:03d}-{safe}"


def write_log(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    with temporary.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    temporary.replace(path)


def execute_cases(
    runner: MessengerRunner,
    cases: Sequence[dict[str, Any]],
) -> tuple[str, list[dict[str, Any]]]:
    results: list[dict[str, Any]] = []
    overall = "PASS"

    for case_index, case in enumerate(cases):
        case_dir = runner.evidence_dir / safe_case_dir_name(case["case_id"], case_index)
        case_dir.mkdir(parents=True, exist_ok=True)
        case_log: dict[str, Any] = {
            "case_id": case["case_id"],
            "status": "RUNNING",
            "started_at": utc_now(),
            "started_monotonic": time.monotonic(),
            "clean_conversation": case["clean_conversation"],
            "reset_plan": case["reset"],
            "messages": [],
            "evidence": {
                "precondition": str(case_dir / "precondition.png"),
                "decision": str(case_dir / "decision.png"),
            },
        }
        command_start = len(runner.command_log)
        stop_remaining = False

        try:
            if case["reset"]["required"]:
                case_log["reset"] = runner.reset_conversation(case["reset"])

            runner.screenshot(case_dir / "precondition.png")
            previous_send: float | None = None

            for message_index, message in enumerate(case["messages"]):
                requested_ms = message["delay_after_previous_ms"]
                if previous_send is not None:
                    wait_until(previous_send + requested_ms / 1000.0)

                sent = runner.send_message(message["text"])
                actual_ms = None if previous_send is None else (
                    sent["send_timestamp_monotonic"] - previous_send
                ) * 1000.0
                sent.update(
                    {
                        "index": message_index,
                        "requested_interval_ms": None if previous_send is None else requested_ms,
                        "actual_interval_ms": actual_ms,
                        "deviation_ms": None if actual_ms is None else actual_ms - requested_ms,
                    }
                )
                case_log["messages"].append(sent)
                previous_send = sent["send_timestamp_monotonic"]

            case_log["status"] = "PASS"
        except PreconditionBlocked as exc:
            case_log["status"] = "PRECONDITION_BLOCKED"
            case_log["error"] = str(exc)
            overall = "PRECONDITION_BLOCKED"
            stop_remaining = True
        except (ToolFailure, ValueError) as exc:
            case_log["status"] = "TOOL_FAILURE"
            case_log["error"] = str(exc)
            overall = "TOOL_FAILURE"
            stop_remaining = True
        finally:
            try:
                runner.screenshot(case_dir / "decision.png")
            except ToolFailure as screenshot_error:
                case_log.setdefault("evidence_errors", []).append(str(screenshot_error))
                if case_log["status"] == "PASS":
                    case_log["status"] = "TOOL_FAILURE"
                    overall = "TOOL_FAILURE"
                    stop_remaining = True
            case_log["finished_at"] = utc_now()
            case_log["finished_monotonic"] = time.monotonic()
            case_log["bsk_commands"] = runner.command_log[command_start:]
            write_log(case_dir / "execution-log.json", case_log)
            results.append(case_log)

        if stop_remaining:
            break

    return overall, results


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    def add_common(subparser: argparse.ArgumentParser) -> None:
        subparser.add_argument("--session", required=True)
        subparser.add_argument("--tab-id", required=True)
        subparser.add_argument("--evidence-dir", type=Path, required=True)

    run_case = subparsers.add_parser("run-case", help="Run one case or an explicit cases array")
    add_common(run_case)
    run_case.add_argument("--plan", type=Path, required=True)

    smoke = subparsers.add_parser("smoke-send", help="Send exactly two plan messages")
    add_common(smoke)
    smoke.add_argument("--plan", type=Path, required=True)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        cases = load_cases(args.plan)
        if args.subcommand == "smoke-send":
            if len(cases) != 1 or len(cases[0]["messages"]) != 2:
                raise ValueError("smoke-send requires exactly one case containing exactly two messages")
        runner = MessengerRunner(args.session, args.tab_id, args.evidence_dir)
        overall, case_results = execute_cases(runner, cases)
        summary = {
            "status": overall,
            "subcommand": args.subcommand,
            "plan": str(args.plan.resolve()),
            "evidence_dir": str(args.evidence_dir.resolve()),
            "case_results": [
                {
                    "case_id": result["case_id"],
                    "status": result["status"],
                    "log": str(
                        args.evidence_dir.resolve()
                        / safe_case_dir_name(result["case_id"], index)
                        / "execution-log.json"
                    ),
                }
                for index, result in enumerate(case_results)
            ],
        }
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return {"PASS": 0, "TOOL_FAILURE": 2, "PRECONDITION_BLOCKED": 3}[overall]
    except (OSError, json.JSONDecodeError, ValueError, PreconditionBlocked) as exc:
        print(json.dumps({"status": "PRECONDITION_BLOCKED", "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
