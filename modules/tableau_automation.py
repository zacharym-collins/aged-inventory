import logging
import os
import time
import pyautogui
from dotenv import load_dotenv

# Get a logger for this module
logger = logging.getLogger(__name__)

def get_tableau(browser_open: bool = False) -> None:
    """
    Automates GUI interaction to download data from Tableau. Ensure that all screen coordinates match the 
    user's resolution. Requires environment variables for Tableau credentials and URLs.
    
    Parameters:
        browser_open (bool): If True, assumes that the browser is already open to Tableau.
    """
    try:
        logger.info("Attempting to open Chrome for Tableau data download...")
        # Click the chrome icon on the taskbar. If user is logged in to tableau, window should be minimized
        # before running this program.
        pyautogui.click(460, 1061)
        time.sleep(5)

        # Retrieve credentials and URLs from environment variables
        tableau_url = os.getenv("TABLEAU_URL")
        username = os.getenv("TABLEAU_USERNAME")
        password = os.getenv("TABLEAU_PASSWORD")
        workbook_url = os.getenv("WORKBOOK_URL")

        # Stop the tableau process if environment variables are not properly set.
        if not all([tableau_url, username, password, workbook_url]):
            logger.error("One or more Tableau environment variables are missing.")
            return

        # If tableau is not already logged in, log in.
        if not browser_open:
            logger.info("Navigating to Tableau login page.")
            pyautogui.hotkey('ctrl', 'l')
            pyautogui.write(tableau_url)
            pyautogui.press('enter')
            time.sleep(15)

            pyautogui.write(username)
            pyautogui.press('tab')
            pyautogui.write(password)
            pyautogui.press('enter')
            time.sleep(15)

            logger.info("Logging into Tableau and accessing the workbook.")
            pyautogui.hotkey('ctrl', 'l')
            pyautogui.write(workbook_url)
            pyautogui.press('enter')
            time.sleep(30)

        # Begin data download process on Tableau dashboard
        logger.info("Initiating Tableau data download.")
        pyautogui.click(165, 145)  # Open download menu
        time.sleep(30)
        pyautogui.click(1740, 145)  # Select 'Data' tab
        time.sleep(0.5)
        pyautogui.click(1766, 249)  # Click 'Download'
        time.sleep(2)
        pyautogui.press('tab')
        pyautogui.press('right')
        pyautogui.press('tab')
        pyautogui.press('enter')
        time.sleep(30)

        if not browser_open:
            pyautogui.hotkey('alt', 'f4')  # Close browser if it was opened by this script
            logger.info("Closed browser after download.")
        else:
            pyautogui.click(460, 1061)  # Refocus on Tableau if browser_open was True, minimizing window
        logger.info("Tableau data download completed successfully.")
        
    except Exception as e:
        logger.error(f"An error occurred during Tableau data download: {e}", exc_info=True)