"""
Functions to extract data from inspection files.
"""
import os
from datetime import datetime
from typing import Dict, Any, Optional

import pandas as pd

# Import logging
from logging_config import get_logger

logger = get_logger()


def extract_data_from_inspection_file(file_path: str) -> Dict[str, Any]:
    """Extract relevant data from an inspection Excel file.

    Args:
        file_path: Path to the inspection Excel file

    Returns:
        Dict containing extracted data or error information
    """
    filename = os.path.basename(file_path)
    logger.info(f"Extracting data from inspection file: {filename}")

    try:
        # Read the Excel file
        logger.debug(f"Reading Excel file: {file_path}")
        df = pd.read_excel(file_path, sheet_name=0, header=None)
        logger.debug(f"Excel file loaded. Shape: {df.shape}")

        # Find inspection type
        logger.debug("Extracting inspection type")
        tarkastus_type_row = None
        for i in range(3):
            if pd.notna(df.iloc[i, 18]):
                tarkastus_type_row = i
                break

        tarkastus_type = df.iloc[tarkastus_type_row, 18] if tarkastus_type_row is not None else "Määräaikaistarkastus"
        logger.debug(f"Inspection type found: {tarkastus_type}")

        # Extract inspection date and place
        logger.debug("Extracting inspection date")
        inspection_date = _extract_inspection_date(df)
        logger.debug(f"Inspection date: {inspection_date}")

        # Find the next inspection date
        logger.debug("Extracting next inspection date")
        next_inspection_date = _extract_next_inspection_date(df)
        logger.debug(f"Next inspection date: {next_inspection_date}")

        # Extract device information
        logger.debug("Extracting device information")
        device_info = _extract_device_info(df)
        logger.debug(f"Device info extracted: {device_info}")

        # Extract inspection result
        logger.debug("Extracting inspection result")
        inspection_result = _extract_inspection_result(df)
        logger.debug(f"Inspection result: {inspection_result}")

        # Compile the extracted data
        result = {
            "inspection_type": tarkastus_type,
            "inspection_date": inspection_date,
            "next_inspection_date": next_inspection_date,
            "manufacturer": device_info.get("manufacturer"),
            "model": device_info.get("model"),
            "serial_number": device_info.get("serial_number"),
            "owner": device_info.get("owner"),
            "owner_address": device_info.get("owner_address"),
            "inspection_result": inspection_result,
            "filename": filename
        }

        logger.info(f"Data extraction completed for file: {filename}")
        return result
    except Exception as e:
        logger.error(f"Error extracting data from file {file_path}: {e}", exc_info=True)
        return {
            "error": str(e),
            "filename": filename
        }


def _extract_inspection_date(df: pd.DataFrame) -> Optional[datetime]:
    """Extract the inspection date from the DataFrame.

    Args:
        df: DataFrame containing the inspection data

    Returns:
        datetime: Inspection date or None if not found
    """
    logger.debug("Extracting inspection date from DataFrame")
    inspection_date_cell = None
    for row_idx in range(5, 10):
        if (row_idx < df.shape[0] and pd.notna(df.iloc[row_idx, 16])
                and isinstance(df.iloc[row_idx, 16], str)
                and "Paikka ja pvm" in df.iloc[row_idx, 16]):
            if 18 < df.shape[1] and pd.notna(df.iloc[row_idx, 18]):
                inspection_date_cell = df.iloc[row_idx, 18]
                logger.debug(f"Found date cell at row {row_idx}: {inspection_date_cell}")
            break

    inspection_date = None
    if inspection_date_cell:
        try:
            if isinstance(inspection_date_cell, datetime):
                inspection_date = inspection_date_cell
                logger.debug(f"Date is already a datetime object: {inspection_date}")
            elif isinstance(inspection_date_cell, str):
                logger.debug(f"Date is a string, parsing: '{inspection_date_cell}'")
                # Try to parse date from string
                parts = inspection_date_cell.split('/')
                if len(parts) >= 2:
                    date_part = parts[0].strip()
                    logger.debug(f"Extracted date part: '{date_part}'")
                    # Convert Finnish format (dd.mm.yyyy) to datetime
                    if '.' in date_part:
                        date_parts = date_part.split('.')
                        logger.debug(f"Date parts: {date_parts}")
                        if len(date_parts) >= 3:
                            day, month, year = map(int, date_parts)
                            inspection_date = datetime(year, month, day)
                            logger.debug(f"Converted date to datetime: {inspection_date}")
        except Exception as e:
            logger.error(f"Error parsing date: {e}", exc_info=True)
    else:
        logger.debug("No inspection date cell found")

    return inspection_date


def _extract_next_inspection_date(df: pd.DataFrame) -> Optional[datetime]:
    """Extract the next inspection date from the DataFrame.

    Args:
        df: DataFrame containing the inspection data

    Returns:
        datetime: Next inspection date or None if not found
    """
    logger.debug("Extracting next inspection date from DataFrame")
    next_inspection_date = None
    found_row = False

    for row_idx in range(df.shape[0] - 10, df.shape[0]):
        if (row_idx >= 0 and row_idx < df.shape[0] and 15 < df.shape[1]
                and pd.notna(df.iloc[row_idx, 15])):
            cell_value = df.iloc[row_idx, 15]
            if isinstance(cell_value, str) and "Seuraava määräaikaistarkastus" in cell_value:
                logger.debug(f"Found next inspection row at index {row_idx}")
                found_row = True
                if 18 < df.shape[1] and pd.notna(df.iloc[row_idx, 18]):
                    next_date_cell = df.iloc[row_idx, 18]
                    logger.debug(f"Next inspection date cell value: {next_date_cell}")
                    if isinstance(next_date_cell, datetime):
                        next_inspection_date = next_date_cell
                        logger.debug(f"Next inspection date: {next_inspection_date}")
                break

    if not found_row:
        logger.debug("Next inspection date row not found")
    elif not next_inspection_date:
        logger.debug("Next inspection date row found, but date could not be extracted")

    return next_inspection_date


def _extract_device_info(df: pd.DataFrame) -> Dict[str, Any]:
    """Extract device information (manufacturer, model, serial number, owner, address).

    Args:
        df: DataFrame containing the inspection data

    Returns:
        Dict containing device information
    """
    logger.debug("Extracting device information from DataFrame")
    device_info = {
        "manufacturer": None,
        "model": None,
        "serial_number": None,
        "owner": None,
        "owner_address": None
    }

    # Find the device info section
    device_section_found = False
    for row_idx in range(10, 20):
        if (row_idx < df.shape[0] and pd.notna(df.iloc[row_idx, 0])
                and isinstance(df.iloc[row_idx, 0], str)
                and "NOSTIMEN PERUSTIEDOT" in df.iloc[row_idx, 0]):
            logger.debug(f"Found device info section at row {row_idx}")
            device_section_found = True

            # Extract manufacturer info
            manufacturer_row = row_idx + 1
            if manufacturer_row < df.shape[0] and 18 < df.shape[1] and pd.notna(df.iloc[manufacturer_row, 18]):
                device_info["manufacturer"] = df.iloc[manufacturer_row, 18]
                logger.debug(f"Manufacturer: {device_info['manufacturer']}")

            # Extract model info
            model_row = row_idx + 2
            if model_row < df.shape[0] and 2 < df.shape[1] and pd.notna(df.iloc[model_row, 2]):
                device_info["model"] = df.iloc[model_row, 2]
                logger.debug(f"Model: {device_info['model']}")

            # Extract serial number
            serial_row = row_idx + 3
            if serial_row < df.shape[0] and 2 < df.shape[1] and pd.notna(df.iloc[serial_row, 2]):
                device_info["serial_number"] = df.iloc[serial_row, 2]
                logger.debug(f"Serial number: {device_info['serial_number']}")

            # Extract owner info
            owner_row = row_idx + 2
            if owner_row < df.shape[0] and 18 < df.shape[1] and pd.notna(df.iloc[owner_row, 18]):
                device_info["owner"] = df.iloc[owner_row, 18]
                logger.debug(f"Owner: {device_info['owner']}")

            # Extract owner address
            address_row = row_idx + 3
            if address_row < df.shape[0] and 18 < df.shape[1] and pd.notna(df.iloc[address_row, 18]):
                device_info["owner_address"] = df.iloc[address_row, 18]
                logger.debug(f"Owner address: {device_info['owner_address']}")

            break

    if not device_section_found:
        logger.debug("Device info section not found in the document")

    return device_info


def _extract_inspection_result(df: pd.DataFrame) -> Optional[str]:
    """Extract the inspection result from the DataFrame.

    Args:
        df: DataFrame containing the inspection data

    Returns:
        str: Inspection result or None if not found
    """
    logger.debug("Extracting inspection result from DataFrame")
    inspection_result = None
    section_found = False

    for row_idx in range(df.shape[0] - 30, df.shape[0]):
        if row_idx >= 0 and row_idx < df.shape[0] and pd.notna(df.iloc[row_idx, 0]):
            cell_value = df.iloc[row_idx, 0]
            if isinstance(cell_value, str) and "PUUTTEET JA HUOMAUTUKSET" in cell_value:
                logger.debug(f"Found inspection result section at row {row_idx}")
                section_found = True
                result_row = row_idx + 1
                if result_row < df.shape[0]:
                    if 0 < df.shape[1] and pd.notna(df.iloc[result_row, 0]) and df.iloc[result_row, 0] == 1:
                        inspection_result = "Käyttökunnossa"
                        logger.debug("Result: Käyttökunnossa (column 0)")
                    elif 1 < df.shape[1] and pd.notna(df.iloc[result_row, 1]) and df.iloc[result_row, 1] == 1:
                        inspection_result = "Korjattava"
                        logger.debug("Result: Korjattava (column 1)")
                    elif 2 < df.shape[1] and pd.notna(df.iloc[result_row, 2]) and df.iloc[result_row, 2] == 1:
                        inspection_result = "Ei käyttökunnossa"
                        logger.debug("Result: Ei käyttökunnossa (column 2)")
                    elif 15 < df.shape[1] and pd.notna(df.iloc[result_row, 15]):
                        result_text = str(df.iloc[result_row, 15])
                        logger.debug(f"Found result text: {result_text}")
                        if "käyttökunnossa" in result_text:
                            inspection_result = "Käyttökunnossa"
                            logger.debug("Result from text: Käyttökunnossa")
                        elif "korjattava" in result_text:
                            inspection_result = "Korjattava"
                            logger.debug("Result from text: Korjattava")
                        elif "ei ole käyttökunnossa" in result_text:
                            inspection_result = "Ei käyttökunnossa"
                            logger.debug("Result from text: Ei käyttökunnossa")
                break

    if not section_found:
        logger.debug("Inspection result section not found")
    elif not inspection_result:
        logger.debug("Inspection result section found, but result could not be determined")

    return inspection_result