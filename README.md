# SAP GUI Scripting Automation Patterns

A small, self-contained toolkit extracted from a production RPA project I built to automate recurring SAP GUI workflows (financial extracts, reconciliation, reporting) for a finance team. Company-specific details (account codes, transaction names, business logic) have been removed or genericized — what's left is the *engineering patterns* that made the automation reliable enough to run unattended.

This isn't a tutorial repo. It's a record of real problems that came up automating SAP GUI Scripting + native Windows dialogs, and how they were actually solved.

## Why this exists

SAP GUI Scripting (the COM API that lets you drive the SAP GUI client from code) is well documented for *scripted field entry*. It's much less documented for what happens when SAP GUI triggers a **native Windows dialog** mid-script — a "Save As" window, a security prompt, an Excel "file already exists" popup. Those aren't part of the SAP GUI object model at all; you have to automate the actual Windows UI (via `pywinauto`), *while* a blocking SAP GUI call is still waiting for that dialog to close.

That combination — one blocking call, racing against a native dialog you also have to drive — is where most of the real engineering in this repo lives.

## The case study: a bug that looked like three different bugs

This is worth walking through because the debugging process is the actual demonstration of skill, more than the final fix.

A recurring "grant access to this program/file" security popup kept appearing mid-extract and had to be dismissed automatically. Over several iterations, three plausible-looking fixes were tried and discarded:

1. **"It's a timing issue."** The popup seemed to take a while to appear, so the watch window was extended from 45 seconds to 5 minutes. It didn't help — the popup was appearing instantly; something else was eating the time.
2. **"The automation backend is wrong."** `pywinauto` supports two backends for driving Windows UI (`win32` and `uia`). A fallback to `uia` was added on the theory that `win32` alone wasn't reliable. A direct diagnostic later proved `uia` **never found this particular window at all** — its failed attempts were the actual source of the slowdown from hypothesis 1. Removed entirely.
3. **The real cause**, found only by dumping the popup's actual control tree (`dialog.descendants()`) instead of guessing further: the buttons were internally labeled `'&Allow'` and `'&Remember My Decision'` — the `&` is Windows' keyboard-accelerator marker, invisible on screen but present in the real `title` property. The code was matching on an *exact* title (`"Allow"`), which can never match text containing that prefix.

The fix was a one-line change — match button titles with a regular expression (`title_re=".*Allow.*"`) instead of an exact string — but it only became obvious after verifying against the real control tree rather than reasoning about what the fix "should" be.

**The general lesson, applied throughout this repo:** when automating a native UI you don't control, verify your assumptions about it directly (dump the real control tree) before writing code against it. Two of three fix attempts above were wasted effort because they were guesses rather than diagnoses.

## What's in here

- **`sap_connection.py`** — a reusable session wrapper: attaches to an existing SAP GUI session where possible, exposes the popup-handling utilities below, and provides `export_current_report()`, which demonstrates the core concurrency pattern.
- **`diagnose_popup.py`** — the actual diagnostic tool used to find the `&` bug above: dumps every control of a named native window (class name, automation ID, text) via both `pywinauto` backends. Point this at *any* mystery popup before writing matching code against it.
- **`example_extract.py`** — a minimal, fictional example (fake transaction, fake account codes) showing how the pieces compose end to end.

## The core concurrency pattern

```python
# A blocking SAP GUI call can leave a native Windows dialog open and
# wait for it to close before returning control to Python. Dialog
# handling must therefore start in a background thread BEFORE the
# blocking call, not after it.

watcher_thread = threading.Thread(target=handle_save_as_dialog, kwargs={...}, daemon=True)
watcher_thread.start()

session.FindById(export_menu_path).select()  # blocks until the dialog chain resolves

watcher_thread.join(timeout=40)
```

Getting the order right here — thread started *before* the blocking call, not after — was itself a source of an earlier bug: starting the watcher after `.select()` meant it never got a chance to run until the dialog had already timed out unhandled.

## Native Windows button matching — the general rule

```python
# Fragile: breaks silently on any accelerator character, localization,
# or minor SAP GUI version difference.
dialog.child_window(title="Allow", class_name="Button").click_input()

# Robust: matches regardless of '&', surrounding whitespace, or case.
dialog.child_window(title_re=".*Allow.*", class_name="Button").click_input()
```

## What this repo intentionally leaves out

No real transaction codes, account structures, company identifiers, or business logic from the production system. The login/session-establishment step (specific to the employer's internal SSO portal) is replaced here with a stub — in production this is swapped for whatever your organization's login flow is. The value here is the automation *pattern*, not the business process it was built for.

## Stack

Python 3.13, `pywinauto`, `pywin32` (`win32com`), standard library `threading`.
