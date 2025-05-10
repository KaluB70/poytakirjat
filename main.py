"""
Entry point for the Tarkastuspöytäkirja to Asiakasrekisteri Data Transfer application.
"""
import os

from nicegui import ui

# Configure logging
from logging_config import setup_logging

# Initialize logging
logger = setup_logging()

logger.info("Application starting")

# Import configuration
from config import APP_TITLE, UPLOAD_DIR, OUTPUT_FILE
logger.info("Configuration loaded")

# Import the application modules
logger.info("Importing application modules")
try:
    from components import create_app_header, create_section_title, FileUploadComponent, FileListComponent, ResultsTableComponent
    from file_helpers import save_uploaded_file, is_valid_excel_file, get_basename
    from inspection_parser import extract_data_from_inspection_file
    from registry_updater import update_customer_registry

    logger.info("All modules imported successfully")
except Exception as e:
    logger.error(f"Error importing modules: {e}", exc_info=True)
    raise

class TarkastusApp:
    """Main application class for the Tarkastus app."""

    def __init__(self):
        """Initialize the application."""
        logger.info("Initializing TarkastusApp")
        self.registry_file = None
        self.uploaded_files = []
        self.processing_results = []

        # Create the application UI structure
        try:
            self._create_ui()
            logger.info("UI created successfully")
        except Exception as e:
            logger.error(f"Error creating UI: {e}", exc_info=True)
            raise

    def _create_ui(self):
        """Create the UI components and structure."""
        logger.debug("Creating UI components")
        # Create main UI components
        create_app_header(APP_TITLE)

        # Main container
        with ui.column().classes('w-full max-w-4xl mx-auto p-4 gap-4'):
            # Registry file upload section
            logger.debug("Creating registry file upload section")
            create_section_title('Lataa Asiakasrekisteri tiedosto')
            self.registry_upload = FileUploadComponent('Asiakasrekisteri ja laitetiedot Excel-tiedosto')
            self.registry_upload.on_upload(self._handle_registry_upload)

            ui.separator()

            # Inspection files upload section
            logger.debug("Creating inspection files upload section")
            create_section_title('Lataa Tarkastuspöytäkirja tiedostot')
            self.inspection_upload = FileUploadComponent('Tarkastuspöytäkirja Excel-tiedostot', multiple=True)
            self.inspection_upload.on_upload(self._handle_inspection_upload)

            # File list component
            logger.debug("Creating file list component")
            self.file_list = FileListComponent()
            self.file_list.set_update_callback(self._on_files_updated)

            ui.separator()

            # Action buttons
            logger.debug("Creating action buttons")
            with ui.row().classes('justify-end'):
                self.process_button = ui.button('Käsittele tiedostot', icon='play_arrow').props('color=primary')
                self.process_button.on_click(self._process_files)

                self.download_button = ui.button('Lataa päivitetty Asiakasrekisteri', icon='download')
                self.download_button.on_click(self._download_results)
                self.download_button.disable()

            # Results section
            logger.debug("Creating results container")
            self.results_container = ResultsTableComponent()

    def _handle_registry_upload(self, e):
        """Handle registry file upload events."""
        logger.info(f"Registry file upload event received: {e.name}")
        try:
            file_path = save_uploaded_file(UPLOAD_DIR, e.name, e.content)
            logger.debug(f"Registry file saved to: {file_path}")
            self.registry_file = file_path
            ui.notify(f'Asiakasrekisteri tiedosto ladattu: {e.name}', type='positive')
            logger.info(f"Registry file upload successful: {e.name}")
        except Exception as ex:
            error_msg = f"Error uploading registry file: {ex}"
            logger.error(error_msg, exc_info=True)
            ui.notify(error_msg, type='negative')

    def _handle_inspection_upload(self, e):
        """Handle inspection file upload events."""
        logger.info(f"Inspection file upload event received: {e.name}")
        try:
            file_path = save_uploaded_file(UPLOAD_DIR, e.name, e.content)
            logger.debug(f"Inspection file saved to: {file_path}")
            self.uploaded_files.append(file_path)
            self.file_list.add_file(file_path)
            ui.notify(f'Tarkastuspöytäkirja tiedosto ladattu: {e.name}', type='positive')
            logger.info(f"Inspection file upload successful: {e.name}")
        except Exception as ex:
            error_msg = f"Error uploading inspection file: {ex}"
            logger.error(error_msg, exc_info=True)
            ui.notify(error_msg, type='negative')

    def _on_files_updated(self):
        """Callback for when files are updated."""
        logger.debug("Files list updated")
        # Update the uploaded_files list to match what's in the UI component
        self.uploaded_files = self.file_list.files.copy()
        logger.debug(f"Updated files list: {[os.path.basename(f) for f in self.uploaded_files]}")

    def _process_files(self):
        """Process the uploaded files."""
        logger.info("Process files started")

        # Check if we have the required files
        if not self.registry_file:
            logger.warning("No registry file provided")
            ui.notify('Asiakasrekisteri tiedosto on pakollinen!', type='negative')
            return

        if not self.uploaded_files:
            logger.warning("No inspection files provided")
            ui.notify('Lataa vähintään yksi tarkastuspöytäkirja!', type='negative')
            return

        logger.info(f"Processing {len(self.uploaded_files)} inspection files")

        # Process each inspection file
        inspection_data = []
        for file_path in self.uploaded_files:
            logger.debug(f"Processing inspection file: {os.path.basename(file_path)}")
            # Make sure the file exists
            if is_valid_excel_file(file_path):
                try:
                    data = extract_data_from_inspection_file(file_path)
                    inspection_data.append(data)
                    logger.debug(f"Extracted data: {data}")
                except Exception as ex:
                    error_msg = f"Error extracting data from file {os.path.basename(file_path)}: {ex}"
                    logger.error(error_msg, exc_info=True)
                    ui.notify(error_msg, type='negative')
            else:
                logger.warning(f"Invalid Excel file: {file_path}")
                ui.notify(f'Tiedostoa {get_basename(file_path)} ei löydy tai se ei ole Excel-tiedosto!',
                        type='negative')

        if not inspection_data:
            logger.warning("No valid inspection data found")
            ui.notify('Ei tarkastuspöytäkirjoja käsiteltäväksi!', type='negative')
            return

        logger.info("Updating customer registry")
        # Update the registry
        try:
            output_path, results = update_customer_registry(self.registry_file, inspection_data)
            logger.info(f"Registry update completed. Output path: {output_path}")
            self.processing_results = results

            if output_path and os.path.exists(output_path):
                # Enable download button and configure it
                self.download_button.enable()
                logger.debug("Download button enabled")

                # Show results
                self.results_container.show_results(results)
                logger.debug("Results displayed")

                # Log processing summary
                summary = {}
                for record in results:
                    status = record.get('status', 'Unknown')
                    if status in summary:
                        summary[status] += 1
                    else:
                        summary[status] = 1
                logger.info(f"Processing summary: {summary}")

                ui.notify('Tiedostot käsitelty onnistuneesti!', type='positive')
            else:
                logger.warning(f"Output file not created: {output_path}")
                ui.notify('Virhe tiedostojen käsittelyssä!', type='negative')
        except Exception as ex:
            error_msg = f"Error updating registry: {ex}"
            logger.error(error_msg, exc_info=True)
            ui.notify(error_msg, type='negative')

    def _download_results(self):
        """Handle the download button click event."""
        logger.info("Download results requested")
        if os.path.exists(OUTPUT_FILE):
            # Get filename without path
            filename = os.path.basename(OUTPUT_FILE)
            logger.info(f"Initiating download of {filename}")
            ui.download(OUTPUT_FILE, filename)
        else:
            logger.warning(f"Output file not found: {OUTPUT_FILE}")
            ui.notify('Päivitettyä tiedostoa ei löydy!', type='negative')

def main():
    """Application entry point."""
    logger.info("Starting main function")
    try:
        # Create the application instance
        app = TarkastusApp()
        logger.info("TarkastusApp initialized")

        # Run the application
        logger.info(f"Starting NiceGUI with title: {APP_TITLE}")
        ui.run(title=APP_TITLE)
    except Exception as e:
        logger.critical(f"Fatal error in main function: {e}", exc_info=True)
        raise

if __name__ in {"__main__", "__mp_main__"}:
    logger.info("Application script executed directly")
    try:
        main()
    except Exception as e:
        logger.critical(f"Unhandled exception: {e}", exc_info=True)
        # Re-raise to show the error in the console
        raise