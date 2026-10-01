"""
example_extract.py

A minimal, entirely fictional example of how the pieces in this repo compose
end to end — fake transaction code, fake account range, no real business
logic. Shows the shape of a real extract script, not its content.
"""

import logging
from sap_connection import SAPConnection

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


def run_demo_extract(account_range_low: str, account_range_high: str, output_filename: str):
    sap = SAPConnection()
    if not sap.session:
        logging.error("No SAP session available.")
        return

    session = sap.session

    # --- Fictional transaction navigation and field entry ---
    session.FindById("wnd[0]/tbar[0]/okcd").text = "/nZDEMO_REPORT"
    session.FindById("wnd[0]").sendVKey(0)
    session.FindById("wnd[0]").maximize()

    session.FindById("wnd[0]/usr/ctxtSO_ACCT-LOW").text = account_range_low
    session.FindById("wnd[0]/usr/ctxtSO_ACCT-HIGH").text = account_range_high

    session.FindById("wnd[0]/tbar[1]/btn[8]").press()  # Execute
    logging.info("Report executed.")

    # --- Export, using the concurrency pattern from sap_connection.py ---
    sap.export_current_report(
        filename=output_filename,
        export_menu_path="wnd[0]/mbar/menu[0]/menu[3]/menu[1]",  # List > Export > Spreadsheet
    )


if __name__ == "__main__":
    run_demo_extract(account_range_low="1000", account_range_high="1999", output_filename="Demo Extract")
