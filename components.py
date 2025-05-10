"""
UI components for the Tarkastus application.
"""
from typing import List, Dict, Any, Callable, Optional
from nicegui import ui
import os

# Import logging
from logging_config import get_logger

logger = get_logger()


def create_app_header(title: str):
    """Create the application header.

    Args:
        title: Application title
    """
    logger.debug(f"Creating app header with title: {title}")
    with ui.header().classes('bg-blue-600 text-white'):
        ui.label(title).classes('text-h5')


def create_section_title(title: str):
    """Create a section title.

    Args:
        title: Section title
    """
    logger.debug(f"Creating section title: {title}")
    return ui.label(title).classes('text-h6')


class FileUploadComponent:
    """Component for handling file uploads."""

    def __init__(self, label: str, multiple: bool = False, accepted_types: str = '.xlsx,.xls'):
        """Initialize the FileUploadComponent.

        Args:
            label: Label for the upload button
            multiple: Whether to allow multiple file uploads
            accepted_types: Comma-separated list of accepted file types
        """
        logger.debug(f"Creating FileUploadComponent with label: {label}, multiple: {multiple}")
        self.upload = ui.upload(
            label=label,
            multiple=multiple,
            auto_upload=True
        ).props(f'accept={accepted_types}')
        self.upload.classes('w-full')

    def on_upload(self, callback: Callable):
        """Set callback for upload events.

        Args:
            callback: Function to call when a file is uploaded
        """
        logger.debug("Setting upload callback")
        self.upload.on_upload(callback)


class FileListComponent:
    """Component for displaying and managing a list of uploaded files."""

    def __init__(self, container_classes: str = 'w-full'):
        """Initialize the FileListComponent.

        Args:
            container_classes: CSS classes for the container
        """
        logger.debug("Creating FileListComponent")
        self.container = ui.column().classes(container_classes)
        self.files: List[str] = []
        self.update_callback: Optional[Callable] = None

    def set_files(self, files: List[str]):
        """Set the list of files.

        Args:
            files: List of file paths
        """
        logger.debug(f"Setting file list with {len(files)} files")
        self.files = files.copy()  # Make a copy to avoid reference issues
        self._update_view()

    def add_file(self, file_path: str):
        """Add a file to the list.

        Args:
            file_path: Path to the file
        """
        filename = os.path.basename(file_path)
        if file_path not in self.files:
            logger.debug(f"Adding file to list: {filename}")
            self.files.append(file_path)
            self._update_view()
        else:
            logger.debug(f"File already in list, not adding: {filename}")

    def remove_file(self, file_path: str):
        """Remove a file from the list.

        Args:
            file_path: Path to the file
        """
        filename = os.path.basename(file_path)
        if file_path in self.files:
            logger.debug(f"Removing file from list: {filename}")
            self.files.remove(file_path)
            self._update_view()
            if self.update_callback:
                logger.debug("Calling update callback after file removal")
                self.update_callback()
        else:
            logger.debug(f"File not in list, cannot remove: {filename}")

    def _update_view(self):
        """Update the UI view with the current list of files."""
        logger.debug(f"Updating file list view with {len(self.files)} files")
        self.container.clear()
        with self.container:
            ui.label(f'Ladatut tiedostot ({len(self.files)}):').classes('text-body1')
            for file_path in self.files:
                with ui.row().classes('w-full items-center'):
                    ui.label(os.path.basename(file_path)).classes('flex-grow')

                    # Create a button to remove the file
                    remove_btn = ui.button(icon='delete', color='red')

                    # Create a closure to capture the current file_path
                    def create_remove_handler(path):
                        return lambda: self.remove_file(path)

                    # Attach the handler
                    remove_btn.on_click(create_remove_handler(file_path))

    def set_update_callback(self, callback: Callable):
        """Set a callback to be called when files are updated.

        Args:
            callback: Function to call when files are updated
        """
        logger.debug("Setting update callback for FileListComponent")
        self.update_callback = callback

    def clear(self):
        """Clear the list of files."""
        logger.debug("Clearing file list")
        self.files = []
        self._update_view()


class ResultsTableComponent:
    """Component for displaying processing results in a table."""

    def __init__(self, container_classes: str = 'w-full'):
        """Initialize the ResultsTableComponent.

        Args:
            container_classes: CSS classes for the container
        """
        logger.debug("Creating ResultsTableComponent")
        self.container = ui.column().classes(container_classes)

    def show_results(self, results: List[Dict[str, Any]]):
        """Show processing results in a table.

        Args:
            results: List of dictionaries containing processing results
        """
        logger.debug(f"Showing results table with {len(results)} records")
        self.container.clear()
        with self.container:
            ui.label('Käsittelyn tulokset:').classes('text-h6 mt-4')

            # Create a compatible table format
            columns = [
                {'name': 'status', 'label': 'Tila', 'field': 'status'},
                {'name': 'model', 'label': 'Malli', 'field': 'model'},
                {'name': 'serial', 'label': 'Sarjanumero', 'field': 'serial_number'},
                {'name': 'owner', 'label': 'Omistaja', 'field': 'owner'},
                {'name': 'file', 'label': 'Tiedosto', 'field': 'filename'},
                {'name': 'message', 'label': 'Viesti', 'field': 'message'}
            ]

            # Make sure all records have all fields
            normalized_results = []
            for record in results:
                normalized_record = {}
                for col in columns:
                    field = col['field']
                    normalized_record[field] = record.get(field, "")
                normalized_results.append(normalized_record)

            # Count records by status
            status_counts = {}
            for record in normalized_results:
                status = record.get('status', '')
                if status in status_counts:
                    status_counts[status] += 1
                else:
                    status_counts[status] = 1

            logger.debug(f"Results by status: {status_counts}")
            ui.table(columns=columns, rows=normalized_results, row_key='filename')

    def clear(self):
        """Clear the results table."""
        logger.debug("Clearing results table")
        self.container.clear()