import logging
import os
from modules.tableau_automation import get_tableau
from modules.sap_automation import get_sap_stock
from modules.file_utils import copy_and_rename_files, save_to_excel, move_to_sharepoint
from modules.inventory_processing import process_inventory_data

# Setup logging

if not os.path.exists('logs'):
    os.mkdir('logs')            # Create logs directory if it does not already exist
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', handlers=[
    logging.FileHandler("logs/app.log"),
    logging.StreamHandler()
])

def main():
    logging.info("Starting inventory processing")

    # Download Tableau report via GUI automation if enabled
    if os.getenv("DOWNLOAD_TABLEAU")  == "False":
        pass
    else:
        get_tableau()

    # Download SAP stock data. Assumes the user is already logged into the SAP application
    get_sap_stock()
    # Retrieve tableau and SAP stock data
    copy_and_rename_files()

    # Initialize the paint load and unload DataFrames
    aged_load_inv, aged_unload_inv, aged_8qi_inv = process_inventory_data()
    # Save each filtered DataFrame to an Excel file
    save_to_excel(aged_load_inv, r'data\aged_load_inv.xlsx')
    save_to_excel(aged_unload_inv, r'data\aged_unload_inv.xlsx')
    save_to_excel(aged_8qi_inv, r'data\aged_8qi_inv.xlsx')

    move_to_sharepoint(dest_dir=os.getenv("SHAREPOINT_DIRECTORY"))

if __name__ == "__main__":
    main()