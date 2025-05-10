"""
Functions to update the customer registry with data from inspection files.
"""
import os
import pandas as pd
import openpyxl
from typing import List, Dict, Any, Tuple, Optional, Union

# Import logging
from logging_config import get_logger
logger = get_logger()

# Import from config instead of main
from config import OUTPUT_FILE

def update_customer_registry(registry_path: str, inspection_data: List[Dict[str, Any]]) -> Tuple[
    str, List[Dict[str, Any]]]:
    """Update the customer registry with data from inspection files.

    Args:
        registry_path: Path to the customer registry Excel file
        inspection_data: List of dictionaries containing inspection data

    Returns:
        Tuple containing the output file path and a list of processing results
    """
    logger.info(f"Updating customer registry: {os.path.basename(registry_path)}")
    logger.debug(f"Number of inspection files to process: {len(inspection_data)}")

    try:
        # Check that registry path exists
        if not registry_path or not os.path.exists(registry_path):
            logger.error(f"Customer registry file not found: {registry_path}")
            return "", [{
                "status": "Error",
                "message": f"Customer registry file not found: {registry_path}",
                "filename": "N/A"
            }]

        # Load the customer registry Excel file
        logger.debug(f"Loading workbook: {registry_path}")
        registry_wb = openpyxl.load_workbook(registry_path)
        logger.debug(f"Workbook loaded. Available sheets: {registry_wb.sheetnames}")

        # Check if the required sheet exists
        if "Kaikki" not in registry_wb.sheetnames:
            logger.error("Required sheet 'Kaikki' not found in registry file")
            return "", [{
                "status": "Error",
                "message": f"Required sheet 'Kaikki' not found in registry file",
                "filename": "N/A"
            }]

        registry_sheet = registry_wb["Kaikki"]
        logger.debug(f"Using sheet 'Kaikki' with dimensions: {registry_sheet.dimensions}")

        # Get existing data as dataframe
        logger.debug("Converting sheet data to DataFrame")
        data = []
        header = None

        # Read the sheet data
        logger.debug("Reading sheet data")
        row_count = 0
        for i, row in enumerate(registry_sheet.iter_rows(values_only=True)):
            if i == 0:
                header = list(row)
                logger.debug(f"Header row: {header}")
            else:
                data.append(dict(zip(header, row)))
                row_count += 1

        logger.debug(f"Read {row_count} data rows from registry")

        df_existing = pd.DataFrame(data)
        logger.debug(f"Created DataFrame with shape: {df_existing.shape}")

        # Process each inspection data record
        logger.info("Processing inspection data records")
        processed_records = _process_inspection_data(df_existing, inspection_data)
        logger.debug(f"Processed {len(processed_records)} records")

        # Added check for empty DataFrame
        if df_existing.empty:
            logger.error("DataFrame is empty after processing. No data to save.")
            return "", [{
                "status": "Error",
                "message": "No data to save in registry. DataFrame is empty after processing.",
                "filename": "N/A"
            }]

        # Save the updated registry
        logger.info("Saving updated registry")
        _save_updated_registry(registry_sheet, df_existing)
        logger.info(f"Registry saved to: {OUTPUT_FILE}")

        return OUTPUT_FILE, processed_records

    except Exception as e:
        logger.error(f"Error updating customer registry: {e}", exc_info=True)
        return "", [{
            "status": "Error",
            "message": f"Failed to update registry: {str(e)}",
            "filename": "N/A"
        }]

def _process_inspection_data(df_existing: pd.DataFrame, inspection_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
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

        # Check if the device already exists in the registry
        logger.debug("Checking if device already exists in registry")
        found_idx = _find_existing_device(df_existing, data)

        if found_idx is not None:
            logger.debug(f"Device found in registry at index {found_idx}")
        else:
            logger.debug("Device not found in registry, will add new record")

        # Prepare new data for update or insertion
        logger.debug("Preparing data for update/insertion")
        new_data = {
            "Aktiivinen": True,
            "Tilaaja": data.get("owner"),
            "Tilaajan osoite": data.get("owner_address"),
            "Laitteen nimi": data.get("model"),
            "Laitteen sarjanumero": data.get("serial_number"),
            "Tarkastettu": data.get("inspection_date"),
            "Seuraava tarkastus": data.get("next_inspection_date"),
            "Lisätieto": data.get("inspection_result"),
            "Tarkastuspöytäkirja": data.get("filename")
        }

        # Log the data we're using
        for key, value in new_data.items():
            logger.debug(f"  {key}: {value}")

        if found_idx is not None:
            # Update existing record
            logger.debug("Updating existing record")
            for key, value in new_data.items():
                if value is not None and key in df_existing.columns:
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
        logger.info(f"Device processed - Status: {status}, Serial: {data.get('serial_number')}, Model: {data.get('model')}")

    return processed_records

def _find_existing_device(df: pd.DataFrame, data: Dict[str, Any]) -> Optional[int]:
    """Find if a device already exists in the registry.

    Args:
        df: DataFrame containing existing registry data
        data: Dictionary containing inspection data

    Returns:
        int: Index of the existing device or None if not found
    """
    serial_to_find = str(data.get("serial_number", "")) if data.get("serial_number") is not None else ""
    model_to_find = str(data.get("model", "")) if data.get("model") is not None else ""

    logger.debug(f"Searching for device - Serial: {serial_to_find}, Model: {model_to_find}")

    if not serial_to_find:
        logger.debug("No serial number provided, device will be considered new")
        return None

    for idx, row in df.iterrows():
        # Convert everything to string for safer comparison
        row_serial = str(row.get("Laitteen sarjanumero", "")) if pd.notna(
            row.get("Laitteen sarjanumero", "")) else ""
        row_model = str(row.get("Laitteen nimi", "")) if pd.notna(row.get("Laitteen nimi", "")) else ""

        logger.debug(f"Comparing with registry row {idx} - Serial: '{row_serial}', Model: '{row_model}'")

        if row_serial and serial_to_find and row_serial == serial_to_find:
            logger.debug("Serial number match found")

            if not row_model or not model_to_find or row_model == model_to_find:
                logger.debug(f"Device found at index {idx}")
                return idx
            else:
                logger.debug("Serial number matched but model didn't match")

    logger.debug("Device not found in registry")
    return None

def _save_updated_registry(registry_sheet, df_existing: pd.DataFrame):
    """Save the updated registry to the Excel file.

    Args:
        registry_sheet: Excel worksheet object to update
        df_existing: DataFrame containing updated registry data
    """
    logger.debug("Saving updated registry to Excel file")

    # Verify DataFrame isn't empty before proceeding
    if df_existing.empty:
        logger.error("Cannot save an empty DataFrame to Excel")
        raise ValueError("DataFrame is empty, no data to save")

    # Make a copy of the workbook before clearing the sheet
    backup_wb = registry_sheet.parent

    try:
        # First, clear the sheet
        logger.debug("Clearing existing sheet")
        for row in registry_sheet.iter_rows():
            for cell in row:
                cell.value = None

        # Then write the updated dataframe
        header_row = list(df_existing.columns)
        logger.debug(f"Writing header row: {header_row}")
        registry_sheet.append(header_row)

        # Write data rows
        logger.debug(f"Writing {len(df_existing)} data rows")
        rows_written = 0
        for _, row in df_existing.iterrows():
            try:
                registry_sheet.append([row.get(col) for col in header_row])
                rows_written += 1
            except Exception as e:
                logger.error(f"Error writing row to Excel: {e}", exc_info=True)
                # Continue with other rows rather than failing completely

        logger.debug(f"Wrote {rows_written} rows to the worksheet")

        # Verify rows were actually written
        if rows_written == 0:
            logger.error("No rows were written to the Excel file")
            raise ValueError("Failed to write any data rows to Excel")

        # Save the workbook
        logger.debug(f"Saving workbook to: {OUTPUT_FILE}")
        try:
            registry_sheet.parent.save(OUTPUT_FILE)
            logger.info(f"Workbook saved successfully to: {OUTPUT_FILE}")
        except Exception as e:
            logger.error(f"Error saving workbook: {e}", exc_info=True)
            raise

    except Exception as e:
        logger.error(f"Error in _save_updated_registry: {e}", exc_info=True)
        # Don't raise here - we'll return an empty string and error message from the calling function
        raise