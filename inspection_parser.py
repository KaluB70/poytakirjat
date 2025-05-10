"""
Functions to extract data from inspection files.
"""
import os
from datetime import datetime
from typing import Dict, Any, Optional, List

import pandas as pd

# Import logging
from logging_config import get_logger
from registry_updater import _find_existing_device

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

        # Debug dump of Excel content
        debug_excel_content(df)

        # Try coordinate-based extraction first
        coordinate_data = extract_data_by_coordinates(df)

        # Check if we got meaningful data
        has_valid_data = (
                pd.notna(coordinate_data.get("model")) or
                pd.notna(coordinate_data.get("serial_number")) or
                coordinate_data.get("inspection_date") is not None
        )

        if not has_valid_data:
            logger.warning("Coordinate-based extraction failed to find data. Falling back to text-based extraction.")
            # Fall back to the original extraction methods
            inspection_type = _extract_inspection_type(df)
            inspection_date = _extract_inspection_date(df)
            next_inspection_date = _extract_next_inspection_date(df)
            device_info = _extract_device_info(df)
            inspection_result = _extract_inspection_result(df)

            # Use text-based extraction results
            result = {
                "inspection_type": inspection_type,
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
        else:
            # Use coordinate-based extraction results
            result = {
                "inspection_type": _extract_inspection_type(df),
                "inspection_date": coordinate_data.get("inspection_date"),
                "next_inspection_date": coordinate_data.get("next_inspection_date"),
                "kympitys_date": coordinate_data.get("kympitys_date"),
                "model": coordinate_data.get("model"),
                "serial_number": coordinate_data.get("serial_number"),
                "owner": coordinate_data.get("owner"),
                "owner_address": None,
                "inspection_result": None,
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


def _extract_inspection_type(df: pd.DataFrame) -> str:
    """Extract the inspection type from the DataFrame.

    Args:
        df: DataFrame containing the inspection data

    Returns:
        str: Inspection type
    """
    logger.debug("Extracting inspection type")
    tarkastus_type_row = None
    for i in range(3):
        if pd.notna(df.iloc[i, 18]):
            tarkastus_type_row = i
            break

    tarkastus_type = df.iloc[tarkastus_type_row, 18] if tarkastus_type_row is not None else "Määräaikaistarkastus"
    logger.debug(f"Inspection type found: {tarkastus_type}")
    return tarkastus_type


def debug_excel_content(df: pd.DataFrame):
    """Print a sample of the Excel file content for debugging."""
    logger.debug("Excel file content sample:")

    # Get dimensions
    rows, cols = df.shape
    logger.debug(f"DataFrame dimensions: {rows} rows x {cols} columns")

    # Print column headers (Excel letters)
    col_letters = [chr(65 + i) if i < 26 else chr(64 + (i // 26)) + chr(65 + (i % 26)) for i in range(min(cols, 30))]
    logger.debug("Column indices: " + ", ".join(f"{col_letters[i]}({i})" for i in range(min(cols, 30))))

    # Print a sample of rows and columns
    sample_rows = min(rows, 20)
    sample_cols = min(cols, 20)

    for row in range(sample_rows):
        row_data = []
        for col in range(sample_cols):
            cell_value = df.iloc[row, col]
            if pd.notna(cell_value):
                row_data.append(f"{col_letters[col]}{row}:{str(cell_value)[:20]}")
        if row_data:  # Only print rows with data
            logger.debug(f"Row {row}: " + " | ".join(row_data))

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


def extract_data_by_coordinates(df: pd.DataFrame) -> Dict[str, Any]:
    """Extract data from specific cell coordinates according to the mapping.

    Args:
        df: DataFrame containing the inspection data

    Returns:
        Dict containing extracted data
    """
    result = {}

    # Convert Excel column letters to numeric indices
    def col_to_index(col_letter):
        logger.debug(f"Converting column letters to index: {col_letter}")
        index = sum((ord(c.upper()) - ord('A') + 1) * 26 ** i
                    for i, c in enumerate(reversed(col_letter))) - 1
        logger.debug(f"Column {col_letter} converted to index {index}")
        return index

    # Parse cell reference like "F-0:11" to row and column indices
    # There are multiple possible interpretations of this format
    def parse_cell_ref(ref):
        try:
            logger.debug(f"Parsing cell reference: {ref}")

            # Split the reference into parts
            if ':' in ref:
                col_part, row_part = ref.split(':')
                logger.debug(f"Split into col_part: {col_part}, row_part: {row_part}")
            else:
                logger.warning(f"Reference {ref} does not contain ':', assuming single cell")
                col_part, row_part = ref, "0"

            # Interpret column part based on different possibilities
            if '-' in col_part:
                # It could be a range (e.g., "A-C") or a special format
                start_col, end_part = col_part.split('-')
                logger.debug(f"Column range from {start_col} to {end_part}")

                # Try to interpret as Excel column
                try:
                    # First, try the most literal interpretation - direct cell reference
                    # e.g., "F-0" -> column F, no special meaning for "-0"
                    col_index = col_to_index(start_col)
                    logger.debug(f"Using column {start_col} (index {col_index})")
                except Exception as e:
                    logger.error(f"Failed to convert column {start_col}: {e}")
                    return None, None
            else:
                # Simple column reference
                try:
                    col_index = col_to_index(col_part)
                    logger.debug(f"Simple column {col_part} (index {col_index})")
                except Exception as e:
                    logger.error(f"Failed to convert column {col_part}: {e}")
                    return None, None

            # Convert row part to index
            try:
                # Excel rows are 1-based, but pandas rows are 0-based
                # If the input is using Excel row numbers, subtract 1
                row_index = int(row_part) - 1  # Adjusted for Excel to pandas conversion
                logger.debug(f"Row {row_part} converted to index {row_index}")
            except ValueError as e:
                logger.error(f"Failed to convert row {row_part}: {e}")
                return None, None

            logger.debug(f"Final indices: row={row_index}, col={col_index}")
            return row_index, col_index

        except Exception as e:
            logger.error(f"Error parsing cell reference {ref}: {e}")
            return None, None

    # Alternative interpretation if the first one fails
    def alt_parse_cell_ref(ref):
        try:
            logger.debug(f"Alternative parsing of cell reference: {ref}")

            # For references like "F-0:11", interpret as simply "F11"
            if ':' in ref:
                col_range, row = ref.split(':')

                # Take just the first column if it's a range
                if '-' in col_range:
                    col = col_range.split('-')[0]
                else:
                    col = col_range

                cell_ref = f"{col}{row}"
                logger.debug(f"Converted {ref} to Excel-style reference {cell_ref}")

                # Now parse using standard Excel reference
                col_index = col_to_index(col)
                row_index = int(row) - 1  # Excel is 1-indexed, pandas is 0-indexed

                logger.debug(f"Alternative indices: row={row_index}, col={col_index}")
                return row_index, col_index
            else:
                logger.warning(f"Alternative parsing can't handle {ref}")
                return None, None

        except Exception as e:
            logger.error(f"Error in alternative cell parsing {ref}: {e}")
            return None, None

    # Function to safely get a value from a cell
    def get_cell_value(row, col, label):
        try:
            if row is None or col is None:
                logger.warning(f"Invalid cell indices for {label}: row={row}, col={col}")
                return None

            if row < 0 or row >= df.shape[0] or col < 0 or col >= df.shape[1]:
                logger.warning(f"Cell indices out of bounds for {label}: row={row}, col={col}, shape={df.shape}")
                return None

            value = df.iloc[row, col]
            logger.debug(f"Value at {row},{col} for {label}: {value}")

            if pd.isna(value):
                logger.warning(f"Found NaN value for {label} at row={row}, col={col}")
                return None

            return value
        except Exception as e:
            logger.error(f"Error getting cell value for {label} at row={row}, col={col}: {e}")
            return None

    # Try both Excel row 1-based and 0-based interpretations
    # First, try with original references
    logger.debug("Trying original cell references...")

    # Extract device name from F-0:11
    logger.debug("Extracting model from F-0:11")
    row, col = parse_cell_ref("F-0:11")
    model = get_cell_value(row, col, "model")

    if model is None:
        # Try alternative parsing
        logger.debug("Trying alternative cell reference interpretation...")
        row, col = alt_parse_cell_ref("F-0:11")
        model = get_cell_value(row, col, "model")

        # If that fails, try direct Excel-style reference
        if model is None:
            logger.debug("Trying direct Excel references...")
            # Try "F11" (Excel style)
            try:
                col = col_to_index("F")
                row = 10  # Excel row 11 = pandas row 10
                model = get_cell_value(row, col, "model (direct)")
            except Exception as e:
                logger.error(f"Error with direct reference: {e}")

    if model is not None:
        result["model"] = model
        logger.debug(f"Found model: {model}")

    # Extract serial number from G-0:12
    logger.debug("Extracting serial number from G-0:12")
    row, col = parse_cell_ref("G-0:12")
    serial = get_cell_value(row, col, "serial_number")

    if serial is None:
        # Try alternative parsing
        row, col = alt_parse_cell_ref("G-0:12")
        serial = get_cell_value(row, col, "serial_number")

        # If that fails, try direct Excel-style reference
        if serial is None:
            try:
                col = col_to_index("G")
                row = 11  # Excel row 12 = pandas row 11
                serial = get_cell_value(row, col, "serial_number (direct)")
            except Exception as e:
                logger.error(f"Error with direct reference: {e}")

    if serial is not None:
        result["serial_number"] = serial
        logger.debug(f"Found serial number: {serial}")

    # Extract inspection date from AD-AG:6
    logger.debug("Extracting inspection date from AD-AG:6")
    row, col = parse_cell_ref("AD-AG:6")
    date_cell = get_cell_value(row, col, "inspection_date")

    if date_cell is None:
        # Try alternative parsing
        row, col = alt_parse_cell_ref("AD-AG:6")
        date_cell = get_cell_value(row, col, "inspection_date")

        # If that fails, try direct Excel-style reference
        if date_cell is None:
            try:
                col = col_to_index("AD")
                row = 5  # Excel row 6 = pandas row 5
                date_cell = get_cell_value(row, col, "inspection_date (direct)")
            except Exception as e:
                logger.error(f"Error with direct reference: {e}")

    inspection_date = None
    if date_cell is not None:
        try:
            logger.debug(f"Processing inspection date cell: {date_cell} (type: {type(date_cell)})")

            # Process the date based on type
            if isinstance(date_cell, datetime):
                inspection_date = date_cell
                logger.debug(f"Date is already a datetime: {inspection_date}")
            elif isinstance(date_cell, str):
                logger.debug(f"Date is a string, parsing: {date_cell}")

                # Try to parse Finnish date format (dd.mm.yyyy)
                if '.' in date_cell:
                    parts = date_cell.split('.')
                    if len(parts) >= 3:
                        day, month, year = map(int, parts)
                        inspection_date = datetime(year, month, day)
                        logger.debug(f"Parsed Finnish date format: {inspection_date}")
                else:
                    # Try other formats
                    try:
                        inspection_date = pd.to_datetime(date_cell).to_pydatetime()
                        logger.debug(f"Parsed with pd.to_datetime: {inspection_date}")
                    except Exception as e:
                        logger.error(f"Failed to parse date with pd.to_datetime: {e}")
            else:
                logger.warning(f"Unexpected date cell type: {type(date_cell)}")
        except Exception as e:
            logger.error(f"Error processing inspection date: {e}")

    if inspection_date:
        result["inspection_date"] = inspection_date
        logger.debug(f"Found inspection date: {inspection_date}")

        # Calculate next inspection date (add 1 year, set to first day of month)
        next_date = inspection_date.replace(year=inspection_date.year + 1, day=1)
        result["next_inspection_date"] = next_date
        logger.debug(f"Calculated next inspection date: {next_date}")

    # Extract worksite (owner) from Y-AB:6
    logger.debug("Extracting worksite (owner) from Y-AB:6")
    row, col = parse_cell_ref("Y-AB:6")
    owner = get_cell_value(row, col, "owner")

    if owner is None:
        # Try alternative parsing
        row, col = alt_parse_cell_ref("Y-AB:6")
        owner = get_cell_value(row, col, "owner")

        # If that fails, try direct Excel-style reference
        if owner is None:
            try:
                col = col_to_index("Y")
                row = 5  # Excel row 6 = pandas row 5
                owner = get_cell_value(row, col, "owner (direct)")
            except Exception as e:
                logger.error(f"Error with direct reference: {e}")

    if owner is not None:
        result["owner"] = owner
        logger.debug(f"Found owner: {owner}")

    # Extract kympitys data
    logger.debug("Extracting kympitys data from AB-AD:58 and Y-Z:58")
    year_row, year_col = parse_cell_ref("AB-AD:58")
    month_row, month_col = parse_cell_ref("Y-Z:58")

    year_cell = get_cell_value(year_row, year_col, "kympitys_year")
    month_cell = get_cell_value(month_row, month_col, "kympitys_month")

    if year_cell is None or month_cell is None:
        # Try alternative parsing
        year_row, year_col = alt_parse_cell_ref("AB-AD:58")
        month_row, month_col = alt_parse_cell_ref("Y-Z:58")

        year_cell = get_cell_value(year_row, year_col, "kympitys_year")
        month_cell = get_cell_value(month_row, month_col, "kympitys_month")

        # If that fails, try direct Excel-style reference
        if year_cell is None or month_cell is None:
            try:
                year_col = col_to_index("AB")
                year_row = 57  # Excel row 58 = pandas row 57
                month_col = col_to_index("Y")
                month_row = 57  # Excel row 58 = pandas row 57

                year_cell = get_cell_value(year_row, year_col, "kympitys_year (direct)")
                month_cell = get_cell_value(month_row, month_col, "kympitys_month (direct)")
            except Exception as e:
                logger.error(f"Error with direct reference: {e}")

    if year_cell is not None and month_cell is not None:
        try:
            logger.debug(f"Kympitys year cell: {year_cell}, month cell: {month_cell}")

            # Handle different data types
            if isinstance(year_cell, (int, float)) and not pd.isna(year_cell):
                year = int(year_cell)
            elif isinstance(year_cell, str) and year_cell.strip() and year_cell.strip().isdigit():
                year = int(year_cell.strip())
            else:
                logger.warning(f"Could not convert year value to int: {year_cell}")
                year = None

            if isinstance(month_cell, (int, float)) and not pd.isna(month_cell):
                month = int(month_cell)
            elif isinstance(month_cell, str) and month_cell.strip() and month_cell.strip().isdigit():
                month = int(month_cell.strip())
            else:
                logger.warning(f"Could not convert month value to int: {month_cell}")
                month = None

            if year and month and 1 <= month <= 12:
                kympitys_date = datetime(year, month, 1)
                result["kympitys_date"] = kympitys_date
                logger.debug(f"Created kympitys date: {kympitys_date}")
            else:
                logger.warning(f"Invalid kympitys values: year={year}, month={month}")
        except Exception as e:
            logger.error(f"Error processing kympitys date: {e}")

            # Log the extraction results
    logger.debug(f"Coordinate-based extraction results: {result}")

    # Check if we got any data at all
    if not result:
        logger.warning("Coordinate-based extraction found no data")

    return result

    # Update to _process_inspection_data function in registry_updater.py

def _process_inspection_data(df_existing: pd.DataFrame, inspection_data: List[Dict[str, Any]]) -> List[
        Dict[str, Any]]:
        """Process inspection data and update the existing dataframe.

        Args:
            df_existing: DataFrame containing existing registry data
            inspection_data: List of dictionaries containing inspection data

        Returns:
            List of dictionaries containing processing results
        """
        logger.debug("Processing inspection data")
        processed_records = []

        # Process each inspection data record
        for data in inspection_data:
            filename = data.get("filename", "unknown")
            logger.debug(f"Processing inspection data for file: {filename}")

            if "error" in data:
                logger.warning(f"Error in inspection data: {data['error']}")
                processed_records.append({
                    "status": "Error",
                    "message": data["error"],
                    "filename": filename
                })
                continue

            # Verify we have meaningful data
            if not data.get("model") and not data.get("serial_number") and not data.get("inspection_date"):
                logger.warning(f"No meaningful data found for file: {filename}")
                processed_records.append({
                    "status": "Error",
                    "message": "Extraction failed: No meaningful data found in the file",
                    "filename": filename
                })
                continue

            # Check if the device already exists in the registry
            logger.debug("Checking if device already exists in registry")
            found_idx = _find_existing_device(df_existing, data)

            if found_idx is not None:
                logger.debug(f"Device found in registry at index {found_idx}")
            else:
                logger.debug("Device not found in registry, will add new record")

            # Format dates for display
            inspection_date_formatted = None
            if data.get("inspection_date"):
                inspection_date_formatted = data["inspection_date"].strftime("%d.%m.%Y")

            next_inspection_date_formatted = None
            if data.get("next_inspection_date"):
                next_inspection_date_formatted = data["next_inspection_date"].strftime("1.%m.%Y")

            kympitys_date_formatted = None
            if data.get("kympitys_date"):
                kympitys_date_formatted = data["kympitys_date"].strftime("1.%m.%Y")

            # Prepare new data for update or insertion
            logger.debug("Preparing data for update/insertion")
            new_data = {
                "Aktiivinen": True,
                "Tilaaja": data.get("owner"),
                "Tilaajan laite": None,  # Leave empty as specified
                "Laitteen nimi": data.get("model"),
                "Laitteen sarjanumero": data.get("serial_number"),
                "Tarkastettu": inspection_date_formatted,
                "Seuraava tarkastus": next_inspection_date_formatted,
                "Seuraava kympitys": kympitys_date_formatted,
                "Huollettu/korjattu": None,  # Leave empty as specified
                "Työmaa": data.get("owner"),  # Same as Tilaaja
                "Lisätieto": None,  # Leave empty as specified
                "Verkkolaskutus1": None,  # Leave empty as specified
                "Verkkolaskutus2": None,  # Leave empty as specified
                "Maksuehto": None,  # Leave empty as specified
                "Tarkastuspöytäkirja": filename  # Use filename of processed file
            }

            # Log the data we're using
            for key, value in new_data.items():
                logger.debug(f"  {key}: {value}")

            if found_idx is not None:
                # Update existing record
                logger.debug("Updating existing record")
                for key, value in new_data.items():
                    if key in df_existing.columns:
                        df_existing.at[found_idx, key] = value
                        logger.debug(f"Updated field '{key}' with value '{value}'")
                status = "Updated"
            else:
                # Create a new row with all columns from the existing dataframe
                logger.debug("Creating new record")
                new_row = {}
                for col in df_existing.columns:
                    new_row[col] = None

                # Update with the data we have
                for key, value in new_data.items():
                    if key in new_row:
                        new_row[key] = value

                # Append the new row
                logger.debug("Appending new row to DataFrame")
                df_existing = pd.concat([df_existing, pd.DataFrame([new_row])], ignore_index=True)
                logger.debug(f"DataFrame shape after append: {df_existing.shape}")
                status = "Added"

            processed_records.append({
                "status": status,
                "serial_number": data.get("serial_number"),
                "model": data.get("model"),
                "owner": data.get("owner"),
                "filename": filename
            })
            logger.info(
                f"Device processed - Status: {status}, Serial: {data.get('serial_number')}, Model: {data.get('model')}")

        return processed_records