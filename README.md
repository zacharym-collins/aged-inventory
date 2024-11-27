# Inventory Managemnt and Reporting Automation
This project automates the retrieval, processing, and reporting of paint inventory data for manufacturing operations. It integrates data from Tableau, SAP, and local file systems to streamline workflows, improve accuracy, and save time. This repository is incomplete and is a work in progress.

## Table of Contents

- [Features](#features)
- [Technologies Used](#technologies-used)
- [Getting Started](#getting-started)
- [Environment Variables](#environment-variables)
- [Usage](#usage)
- [Project Structure](#project-structure)
- [Future Enhancements](#future-enhancements)
- [License](#license)
- [Customization and Setup](#customization-and-setup)

## Features
- Automates Tableau data downloads using `pyautogui`.
- Integrates with SAP to retrieve inventory data.
- Produces "Aged Inventory" report, that is customized to show inventory that is located in a given location for `x` amount of working hours.
- Dynamically adjusts to working shifts and days based on configuration.
- Applies custom formatting to Excel files for reporting.
- Supports error handling and logging for reliable operations.

## Technologies Used
- **Python**: Core programming language.
- **Pandas**: Data manipulation and analysis.
- **PyAutoGUI**: GUI automation.
- **OpenPyXL**: Excel file processing and formatting.
- **Logging**: Comprehensive logging for debuggin and monitoring.
- **Python-dotenv**: Secure handling of environment variables.

## Getting Started
1. Clone the repository:
```
git clone https://github.com/zacharym-collins/aged-inventory.git
cd aged-inventory
```
2. Create a virtual environment and install dependencies:
```
python -m venv venv
source venv/bin/activate # Windows: venv/Scripts/activate
pip install -r requirements.txt
```
3. Set up environment variables (see [Environment Variables](#environment-variables))

## Environment Variables
The project requires the following environment variables, stored in a .env file in the root directory (not tracked by Git for security). Create a `.env` file as follows:
```
# .env
DOWNLOAD_TABLEAU="False" # "True" will enable main.py to attempt GUI browser automation, not reccemonded
TABLEAU_URL="your-tableau-url-here"
TABLEAU_USERNAME="your-tableau-username-here"
TABLEAU_PASSWORD="your-tableau-password-here"
WORKBOOK_URL="specific-url-for-workbook-here"
DOWNLOADS_PATH="C:\Users\example\Downloads"
SAP_GUI_PATH="C:\Users\example\Documents\SAP\SAP GUI"
BASE_DIRECTORY = "C:\Users\example\aged-inventory"
STOCK_TRANSACTION="AB12" # Specific SAP transaction for stock data, AB12 is not a real transaction
SHAREPOINT_DIRECTORY="C:\Users\example\my-linked-sharepoint-directory"
```
Ensure sensitive information like credentials is kept secure. Automating the Tableau download will require changes to the `tableau_automation.py` file, as the program was written to run on the author's machine. You can set `TABLEAU_DOWNLOAD="False"` in the `.env` file to skip this part of the code if you choose to download the Tableau workbook data manually. Ensure that ALL paths are set correctly before proceeding. 

## Usage
1. Create a `.env` file before running any portion of the program (see [Environment Variables](#environment-variables)).
2. Navigate to the `aged-inventory` directory via your terminal or shell of choice.
3. Run `python main.py` to begin the automation. Assumes that all directories are correctly set in the `.env` file.

## Project Structure
```
aged-inventory/
|
|-data/
|-logs/
|-modules/
|    |-file_utils.py
|    |-inventory_processing.py
|    |-sap_automation.py
|    |-tableau_automation.py
|-.env
|-.gitignore
|-.README
|-LICENSE.txt
|-main.py
|-requirements.txt
|-shift_schedules.txt
```

## Future Enhancements
- Explore predictive analysis for inventory optimization
- Extend integration to other manufacturing systems.

## License
This project is licensed under the [MIT License](./LICENSE.txt).

## Customization and Setup

### Environment-Specific Automation
This project automates specific tasks tailored to the author's machine. Although many of the functions could be replaced with a more reliable method than those implemented with `pyautogui`, certain security restrictions and safety measure dictate that this is the most efficient way to retrieve the data on the author's machine. Scripts such as `tableau_automation.py` depend on:
- GUI automation with hardcoded screen coordinates that match the author's screen resolution.
- Specific file paths and directory structures unique to the author's environment.

To adapt these scripts for your use, you may need to:
- Modify screen coordinates in `pyautogui` commands.
- Update paths to match your local file structure.
- Ensure required software is installed and accessible.

### Sensitive Information Exclusion
Sensitive information, such as credentials and internal URLs, has been removed and replaced with environment variable placeholders. You must create a `.env` file as described in the **Setup** section to provide necessary values.

### Limitations
This project is an example of automating specific workflows and may not work "as-is" in other environments without customization.
