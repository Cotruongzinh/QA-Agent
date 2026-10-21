# BrowserSkill on Windows PowerShell

Apply these rules before performing browser interactions with Tencent BrowserSkill (`bsk`) on Windows PowerShell.

## PowerShell-safe element refs

Do not pass an element ref as a bare token in Windows PowerShell:

```powershell
bsk fill @e1 ...
bsk click @e14 ...
```

PowerShell can interpret a leading `@`. Prefer the explicit option form:

```powershell
bsk fill --ref e1 --value "..." --session <SESSION>
```

Apply the same `--ref eN` principle to every interaction command that supports `--ref`. Use `@eN` only when it is quoted as a literal and the command actually requires a positional target.

## Fill

After `observe`, if the element is:

```text
@e1 textbox "Text input"
```

prefer:

```powershell
bsk fill --ref e1 --value "<VALUE>" --session <SESSION>
```

Never generate either of these forms:

```powershell
bsk fill --ref --value ...
bsk fill @e1 ...
```

## Fresh references

After navigation or any DOM change:

- Run `observe` again.
- Obtain a fresh element ref.
- Do not continue using an old ref by default.

## Verify mutations

After `fill`, `select`, or `click`, when the result needs verification:

- Run `observe` again.
- Confirm that the UI, value, or state actually changed.
- Do not treat a successful exit code as sufficient evidence for `PASS`.

## Tool recovery

If an interaction command fails because of its target or ref:

- Re-observe the page.
- Obtain a fresh ref.
- Retry the exact failed action at most once, following the existing TOOL RECOVERY rules in `SKILL.md`.

Do not modify Tencent BrowserSkill itself. Do not add Playwright, Selenium, or another browser engine. These Windows CLI rules do not change the existing `PASS`, `FAIL`, or `UNVERIFIED` rules.
