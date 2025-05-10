"""
Helper functions for file operations.
"""
import os
from pathlib import Path
from typing import Any

# Import logging
from logging_config import get_logger

logger = get_logger()


def save_uploaded_file(upload_dir: Path, file_name: str, file_content: Any) -> str:
    """Save an uploaded file to the specified directory.

    Args:
        upload_dir: Directory to save the file
        file_name: Name of the file
        file_content: Content of the file (file-like object)

    Returns:
        str: Path to the saved file
    """
    logger.debug(f"Saving uploaded file: {file_name} to directory: {upload_dir}")
    try:
        file_path = upload_dir / file_name
        with open(file_path, 'wb') as f:
            content = file_content.read()
            f.write(content)
        logger.debug(f"File saved successfully. Size: {len(content)} bytes")
        return str(file_path)
    except Exception as e:
        logger.error(f"Error saving file {file_name}: {e}", exc_info=True)
        raise


def is_valid_excel_file(file_path: str) -> bool:
    """Check if a file exists and has a valid Excel extension.

    Args:
        file_path: Path to the file

    Returns:
        bool: True if the file is a valid Excel file, False otherwise
    """
    valid_extensions = ['.xlsx', '.xls']

    # Check if the file exists
    exists = file_path and os.path.exists(file_path)

    # Check if it has a valid extension
    has_valid_ext = False
    if exists:
        has_valid_ext = any(file_path.lower().endswith(ext) for ext in valid_extensions)

    result = exists and has_valid_ext

    logger.debug(f"File validation: {file_path} - Exists: {exists}, Valid extension: {has_valid_ext}, Result: {result}")

    return result


def get_basename(file_path: str) -> str:
    """Extract the filename from a path.

    Args:
        file_path: Path to the file

    Returns:
        str: Filename without the path
    """
    basename = os.path.basename(file_path) if file_path else "unknown"
    logger.debug(f"Getting basename from {file_path}: {basename}")
    return basename