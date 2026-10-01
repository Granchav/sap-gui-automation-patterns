"""
sap_connection.py — reusable SAP GUI Scripting session wrapper.

Demonstrates two patterns that matter once SAP GUI Scripting triggers native
Windows dialogs mid-script:

1. A blocking SAP GUI call (session.FindById(...).select()) can leave a
   native dialog open and wait for it to close before returning control to
   Python. Dialog-handling must start in a background thread BEFORE the
   blocking call — see export_current_report().

2. Native Windows buttons should be matched with a regex (title_re), not an
   exact title string. Internal button titles often carry a leading '&'
   (the keyboard-accelerator marker), which an exact match will never find.
   See dismiss_security_popup() — this is the fix for the bug documented in
   the README.

This file is extracted from a production project and genericized: no real
transaction names, account codes, or company-specific login flow. The
session-establishment step below is a stub — swap in your organization's
actual login/SSO flow.
"""

import logging
import threading
import time
from typing import Optional

import win32com.client
from pywinauto import Application

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


def dismiss_security_popup(timeout: int = 10, context_label: str = "") -> bool:
    """Closes a native 'grant access' style security popup, if present.

    Matches buttons by regex rather than exact title — real SAP GUI security
    popups label their buttons '&Allow' / '&Remember My Decision' (the '&' is
    Windows' keyboard-accelerator marker). An exact title match silently
    never finds these controls; see the README case study for how this was
    diagnosed.
    """
    try:
        app = Application(backend="win32").connect(title="Security Warning", timeout=timeout)
        dialog = app.window(title="Security Warning")

        try:
            checkbox = dialog.child_window(title_re=".*Remember.*", class_name="Button")
            if not checkbox.get_toggle_state():
                checkbox.click_input()
            logging.info("Checkbox checked in security popup.")
        except Exception as checkbox_error:
            # Log the real exception rather than swallowing it — this is what
            # eventually surfaces a mismatch like the '&' issue instead of
            # hiding it behind a generic "not found" message.
            logging.warning(f"Checkbox not checked ({type(checkbox_error).__name__}): {checkbox_error}")

        dialog.child_window(title_re=".*Allow.*", class_name="Button").click_input()
        logging.info(f"Security popup dismissed ({context_label}).")
        return True
    except Exception:
        logging.info(f"No security popup detected ({context_label}).")
        return False


def watch_security_popup(timeout: int = 60, context_label: str = "") -> bool:
    """Polls for the security popup for up to `timeout` seconds. Meant to run
    in a thread started BEFORE a blocking SAP GUI call — see
    export_current_report() below for why the ordering matters."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if dismiss_security_popup(timeout=1, context_label=context_label):
            return True
        time.sleep(0.25)
    logging.info(f"No security popup appeared within {timeout}s ({context_label}).")
    return False


def handle_save_as_dialog(filename: str, timeout: int = 30) -> None:
    """Automates the native Windows 'Save As' dialog: types the FULL path
    (not just the filename) — the dialog remembers the last folder used by
    SAP, which drifts over time, so a bare filename can silently save to the
    wrong directory."""
    try:
        app = Application(backend="win32").connect(title_re=".*Save As.*", timeout=timeout)
        dialog = app.window(title_re=".*Save As.*")
        dialog.wait("visible", timeout=timeout)

        field = dialog.child_window(class_name="Edit")
        field.set_edit_text(filename)
        field.type_keys("{ENTER}")
        logging.info(f"Save As: path submitted ({filename}).")
    except Exception as e:
        logging.error(f"Save As dialog handling failed: {e}")


class SAPConnection:
    """Wraps a SAP GUI Scripting session. Attaches to an existing session
    where possible; otherwise establishes one (stubbed here — plug in your
    organization's own login flow)."""

    def __init__(self):
        self.session: Optional["win32com.client.CDispatch"] = None
        self._attach_existing_session()

    def _attach_existing_session(self) -> bool:
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine
            if application.Connections.Count > 0:
                connection = application.Children(0)
                if connection.Sessions.Count > 0:
                    self.session = connection.Children(0)
                    logging.info("Attached to existing SAP GUI session.")
                    return True
        except Exception:
            pass

        logging.info("No existing session found — establish one via your organization's login flow here.")
        return False

    def export_current_report(self, filename: str, export_menu_path: str) -> None:
        """Triggers an export, watching for BOTH the Save As dialog and a
        security popup in parallel, in threads started BEFORE the blocking
        SAP GUI call. This ordering is the core pattern in this repo — see
        the README for the bug that motivated it."""
        save_as_thread = threading.Thread(
            target=handle_save_as_dialog, kwargs={"filename": filename}, daemon=True
        )
        security_thread = threading.Thread(
            target=watch_security_popup, kwargs={"context_label": f"export of {filename}"}, daemon=True
        )

        save_as_thread.start()
        security_thread.start()

        try:
            self.session.FindById(export_menu_path).select()  # blocks until dialogs resolve
            logging.info("Export triggered.")
        finally:
            save_as_thread.join(timeout=40)
            security_thread.join(timeout=65)
