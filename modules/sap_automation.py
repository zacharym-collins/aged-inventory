import logging
import os
import time
import win32com.client
import pyautogui

logger = logging.getLogger(__name__)  # Initialize the module-specific logger

def get_sap_stock() -> None:
    """
    Automates the GUI process in SAP to retrieve stock data for Plant 8. 
    Assumes the user is actively logged into SAP and has required permissions.
    """
    stock_transaction=os.getenv("STOCK_TRANSACTION")

    try:
        logger.info("Starting SAP stock retrieval for Plant 8.")

        # Connect to SAP GUI
        sap_gui = win32com.client.GetObject("SAPGUI")
        application = sap_gui.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        logger.info("Connected to SAP GUI successfully.")

        # Start the transaction
        session.StartTransaction(Transaction=stock_transaction)
        logger.info("Transaction LX02 started.")

        # Set the variant for Plant 8 aged inventory
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        variant_input = session.findById("wnd[1]/usr/txtV-LOW")
        variant_input.text = "P8AGEDINV"
        logger.info("Variant 'P8AGEDINV' selected.")

        # Clear the "User" field and press Execute
        ename_field = session.findById("wnd[1]/usr/txtENAME-LOW")
        ename_field.text = ""
        ename_field.setFocus()
        ename_field.caretPosition = 0
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        logger.info("Executed the transaction.")

        # Proceed with the report export
        session.findById("wnd[0]/tbar[1]/btn[8]").press()  # Execute report
        session.findById("wnd[0]/tbar[1]/btn[16]").press()  # Export report
        session.findById("wnd[1]/tbar[0]/btn[0]").press()   # Confirm export type
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "paint_inventory.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 20
        session.findById("wnd[1]/tbar[0]/btn[11]").press()  # Save file

        logger.info("Report exported to 'paint_inventory.XLSX'.")

    except Exception as e:
        logger.error(f"Error occurred during SAP stock retrieval: {e}", exc_info=True)

    # Wait and close the Excel document
    try:
        time.sleep(5)  # Wait for Excel to open
        pyautogui.click(865, 557)  # Focus on the Excel window
        pyautogui.hotkey('alt', 'f4')  # Close Excel
        logger.info("Closed Excel document.")
    except Exception as e:
        logger.error(f"Error while closing Excel: {e}", exc_info=True)