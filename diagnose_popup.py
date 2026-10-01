"""
diagnose_popup.py

The actual tool used to find the '&' button-title bug described in the
README: dumps every control (class name / control type, automation ID, text)
of a named native window, via both pywinauto backends.

Usage: trigger the popup you're trying to automate, leave it open on screen,
then run this script with the popup's window title (or part of it).

    python diagnose_popup.py "Security Warning"

Don't guess a button's exact text before writing matching code against a
native dialog you don't control — dump its real control tree first.
"""

import sys
from pywinauto import Application, Desktop


def diagnose(title_substring: str, timeout: int = 5) -> None:
    for backend in ["win32", "uia"]:
        print(f"\n{'#' * 70}\n# BACKEND: {backend}\n{'#' * 70}")
        try:
            if backend == "win32":
                app = Application(backend="win32").connect(title_re=f".*{title_substring}.*", timeout=timeout)
                dialog = app.window(title_re=f".*{title_substring}.*")
            else:
                dialog = Desktop(backend="uia").window(title_re=f".*{title_substring}.*")
                if not dialog.exists(timeout=timeout):
                    raise Exception("window not found")

            for ctrl in dialog.descendants():
                try:
                    if backend == "uia":
                        print(f"  control_type={ctrl.element_info.control_type!r:20} "
                              f"auto_id={ctrl.automation_id()!r:15} texte={ctrl.window_text()!r}")
                    else:
                        print(f"  class_name={ctrl.class_name()!r:25} "
                              f"auto_id={ctrl.automation_id()!r:15} texte={ctrl.window_text()!r}")
                except Exception as e:
                    print(f"  (could not read a control: {e})")

        except Exception as e:
            print(f"❌ Could not connect via {backend}: {e}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python diagnose_popup.py \"<window title substring>\"")
        sys.exit(1)
    diagnose(sys.argv[1])
