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
OUTPUT_FILE = os.path.join(TEMP_DIR, 'Asiakasrekisteri_updated.xlsx')
logger.debug(f"Output file: {OUTPUT_FILE}")