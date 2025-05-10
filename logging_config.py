"""
Logging setup for the application.
"""
import logging
import os
import sys
from datetime import datetime

# Create logs directory if it doesn't exist
os.makedirs('logs', exist_ok=True)


# Configure logging
def setup_logging(level=logging.DEBUG):
    """Set up logging for the application.

    Args:
        level: Logging level (default: DEBUG)
    """
    # Create a logger
    logger = get_logger()
    logger.setLevel(level)

    # Clear any existing handlers
    if logger.handlers:
        logger.handlers.clear()

    # Create a file handler that logs even debug messages
    log_filename = f"logs/tarkastus_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(level)

    # Create a console handler with a higher log level
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)

    # Create a formatter and set it on the handlers
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # Add the handlers to the logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    # Log the start of the application
    logger.info(f"Logging initialized. Log file: {log_filename}")

    return logger


# Get logger
def get_logger():
    """Get the application logger."""
    return logging.getLogger('tarkastus_app')