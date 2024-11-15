import pandas as pd
from datetime import datetime, timedelta
import logging
from typing import Tuple
from .file_utils import read_shift_parameters, get_work_shifts, get_work_days

# Set up logging for this module
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

def calculate_elapsed_hours(start_time: datetime, end_time: datetime, work_shifts: list, work_days: list) -> float:
    """
    Function to determine the overall time that material has been located in a storage bin
    based on working hours for each shift.

    Args:
        start_time (datetime): The time when the material was placed in storage.
        end_time (datetime): The current time or the time being evaluated.
        work_shifts (list): A list of tuples defining work shifts as (start_hour, start_minute, end_hour, end_minute).
        work_days (list): A list of weekdays (0-6, where 0 is Monday) that represent workdays.

    Returns:
        float: The total hours that the material has been in the storage bin during working hours.
    """
    total_hours = 0
    current_time = start_time
    while current_time < end_time:
        if current_time.weekday() in work_days:  # Check if it's a workday
            for start_hour, start_minute, end_hour, end_minute in work_shifts:
                shift_start = current_time.replace(hour=start_hour, minute=start_minute, second=0, microsecond=0)
                shift_end = current_time.replace(hour=end_hour, minute=end_minute, second=0, microsecond=0)
                # Handle overnight shifts
                if shift_end < shift_start:
                    shift_end += timedelta(days=1)
                # Calculate the overlap of the shift with the time period
                if end_time > shift_start:
                    overlap_start = max(start_time, shift_start)
                    overlap_end = min(end_time, shift_end)
                    if overlap_start < overlap_end:
                        total_hours += (overlap_end - overlap_start).total_seconds() / 3600
        # Move to the next day
        current_time += timedelta(days=1)
        current_time = current_time.replace(hour=0, minute=0, second=0, microsecond=0)

    return total_hours

def compute_elapsed_hours(row: pd.Series) -> float:
    """
    Computes the elapsed hours based on working hours that material has
    been located in a storage bin for each entry.

    Args:
        row (pd.Series): A single row of the DataFrame containing data for a specific material.

    Returns:
        float: The number of hours the material has been in storage.
    """
    try:
        logger.debug(f"Computing elapsed hours for material: {row['Material']}")

        # Get work shifts and work days based on parameters
        parameters = read_shift_parameters()
        work_shifts = get_work_shifts(parameters)
        work_days = get_work_days(parameters)

        # Prepare calculation for now and last stock placement
        last_placement = pd.to_datetime(row['last_stock_placement'])
        now = datetime.now()

        # Calculate elapsed hours
        elapsed_hours = calculate_elapsed_hours(last_placement, now, work_shifts, work_days)
        logger.debug(f"Elapsed hours for material {row['Material']}: {elapsed_hours:.2f} hours")

        return round(elapsed_hours, 2)

    except Exception as e:
        logger.error(f"Error computing elapsed hours for material {row['Material']}: {e}", exc_info=True)
        raise

def process_inventory_data() -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Performs the bulk of the processing work. Returns two dataframes: aged_load_inv and aged_unload_inv.

    Returns:
        Tuple[pd.DataFrame, pd.DataFrame]: DataFrames for aged load and unload inventory.
    """
    try:
        logger.info("Starting inventory data processing.")

        # Load the paint inventory Excel file and drop rows with missing 'Last stock placement'
        df = pd.read_excel(r'data\paint_inventory.xlsx').dropna(subset=['Last stock placement'])
        df['last_stock_placement'] = pd.to_datetime(df['Last stock placement'].astype(str) + ' ' + df['Time'].astype(str))
        df = df[['Material', 'Material Description', 'Storage Type', 'Storage Bin', 'Total Stock', 'Storage Unit',
                 'last_stock_placement', 'Last addtn to stock']]

        # Compute elapsed hours for each entry
        df['hours_elapsed'] = df.apply(compute_elapsed_hours, axis=1)

        # Filter data for aged load inventory
        # Only show data that is in location for more than 4 working hours
        aged_load_inv = df[(df['Storage Type'].isin([800, 801, 802])) & (df['hours_elapsed'] >= 4)]

        # Filter data for aged unload inventory
        unload_bins = ["UNLOAD01", "UNLOAD02", "UNLOAD03", "UNLOAD04", "UNLOAD05", "LGUNLOAD"]
        # Only show data that is in location for more than 4 working hours
        aged_unload_inv = df[(df['Storage Bin'].isin(unload_bins)) & (df['hours_elapsed'] >= 4)]

        # Load the Paint Processed CSV file
        paint_df = pd.read_csv(r'data\paint_processed.csv', encoding='utf-16', sep='\t')

        # Filter and group the paint_df based on conditions
        paint_filtered = paint_df[
            (paint_df['GOOD'] == 0) & 
            (paint_df['NON-CONFIRMED'] == 0) & 
            (paint_df['SCRAP'] == 0)
        ]

        # Group by PART_NO and aggregate LOADED column
        paint_grouped = paint_filtered.groupby('PART_NO', as_index=False).agg({'LOADED': 'sum'})

        # Merge the aggregated data with the aged_load_inv DataFrame
        aged_load_inv = pd.merge(aged_load_inv, paint_grouped, left_on='Material', right_on='PART_NO', how='left')
        aged_load_inv['In Station'] = aged_load_inv['Total Stock'] - aged_load_inv['LOADED']
        aged_load_inv = aged_load_inv[['Material', 'Material Description', 'Storage Type', 'Storage Bin', 'Total Stock', 'LOADED',
                                       'In Station', 'Storage Unit', 'last_stock_placement', 'Last addtn to stock', 'hours_elapsed']]
        aged_load_inv = aged_load_inv[(aged_load_inv['In Station'] != 0)]

        logger.info("Inventory data processing complete.")

        return aged_load_inv, aged_unload_inv

    except Exception as e:
        logger.error(f"Error processing inventory data: {e}", exc_info=True)
        raise