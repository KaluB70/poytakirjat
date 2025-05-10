# Updated config.py

"""
Configuration settings for the application.
"""
import os
import tempfile
from pathlib import Path

# Import logging if needed
from logging_config import get_logger

logger = get_logger()

# Application title
APP_TITLE = 'Tarkastuspöytäkirja to Asiakasrekisteri Data Transfer'
logger.debug(f"Application title: {APP_TITLE}")

# Setup temporary directories
TEMP_DIR = tempfile.mkdtemp()
UPLOAD_DIR = Path(tempfile.mkdtemp())
logger.debug(f"Temporary directory: {TEMP_DIR}")
logger.debug(f"Upload directory: {UPLOAD_DIR}")

os.makedirs(UPLOAD_DIR, exist_ok=True)

# Default path for the Asiakasrekisteri file
# You can modify this to point to your preferred location
DEFAULT_REGISTRY_PATH = "C:/Users/User/workspace/poytakirjat/resources/Asiakasrekisteri ja laitetiedot - Uusi.xlsx"
logger.debug(f"Default registry path: {DEFAULT_REGISTRY_PATH}")

# Output file path
OUTPUT_FILE = os.path.join(TEMP_DIR, 'Asiakasrekisteri_updated.xlsx')
logger.debug(f"Output file: {OUTPUT_FILE}")