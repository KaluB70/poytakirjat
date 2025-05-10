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
from config import APP_TITLE, UPLOAD_DIR, OUTPUT_FILE, DEFAULT_REGISTRY_PATH

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
        self.registry_upload_container = None

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

            # Container for registry selection UI
            with ui.column().classes('w-full') as registry_section:
                # Add a note about the default registry file
                if os.path.exists(DEFAULT_REGISTRY_PATH):
                    # Status display for the selected registry
                    self.registry_status = ui.label('')

                    # Row for buttons
                    with ui.row().classes('w-full items-center justify-between'):
                        # Button to use the default file
                        use_default_btn = ui.button('Käytä oletustiedostoa', icon='file_open').props('color=primary')
                        use_default_btn.on_click(self._use_default_registry)

                        # Button to use custom file instead
                        self.custom_file_btn = ui.button('Käytä mukautettua tiedostoa', icon='upload_file').props(
                            'outline')
                        self.custom_file_btn.on_click(self._show_custom_registry_upload)

                # Container for the registry upload component (can be hidden/shown)
                self.registry_upload_container = ui.column().classes('w-full mt-2')
                with self.registry_upload_container:
                    self.registry_upload = FileUploadComponent('Asiakasrekisteri ja laitetiedot Excel-tiedosto')
                    self.registry_upload.on_upload(self._handle_registry_upload)

                # Initially hide the upload if default exists
                if os.path.exists(DEFAULT_REGISTRY_PATH):
                    self.registry_upload_container.set_visibility(False)

                    # Automatically use the default registry
                    self._use_default_registry()

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

    def _use_default_registry(self):
        """Use the default registry file."""
        logger.info(f"Using default registry file: {DEFAULT_REGISTRY_PATH}")
        if os.path.exists(DEFAULT_REGISTRY_PATH):
            self.registry_file = DEFAULT_REGISTRY_PATH
            self.registry_status.set_text(f'Käytetään oletustiedostoa: {os.path.basename(DEFAULT_REGISTRY_PATH)}')
            self.registry_status.classes('text-green-600')

            # Hide the upload component
            if self.registry_upload_container:
                self.registry_upload_container.set_visibility(False)

            ui.notify(f'Oletustiedosto valittu: {os.path.basename(DEFAULT_REGISTRY_PATH)}', type='positive')
        else:
            self.registry_status.set_text(f'Oletustiedostoa ei löydy: {os.path.basename(DEFAULT_REGISTRY_PATH)}')
            self.registry_status.classes('text-red-600')

            # Show the upload component as fallback
            if self.registry_upload_container:
                self.registry_upload_container.set_visibility(True)

            ui.notify(f'Oletustiedostoa ei löydy: {DEFAULT_REGISTRY_PATH}', type='negative')
            logger.warning(f"Default registry file not found: {DEFAULT_REGISTRY_PATH}")

    def _show_custom_registry_upload(self):
        """Show the custom registry upload component."""
        logger.debug("Showing custom registry upload")
        if self.registry_upload_container:
            self.registry_upload_container.set_visibility(True)

        # Clear the registry file selection
        self.registry_file = None
        self.registry_status.set_text('Valitse mukautettu tiedosto')
        self.registry_status.classes('text-gray-600')

    def _handle_registry_upload(self, elem):
        """Handle registry file upload events."""
        logger.info(f"Registry file upload event received: {elem.name}")
        try:
            file_path = save_uploaded_file(UPLOAD_DIR, elem.name, elem.content)
            logger.debug(f"Registry file saved to: {file_path}")
            self.registry_file = file_path

            # Update status
            self.registry_status.set_text(f'Valittu tiedosto: {elem.name}')
            self.registry_status.classes('text-green-600')

            ui.notify(f'Asiakasrekisteri tiedosto ladattu: {elem.name}', type='positive')
            logger.info(f"Registry file upload successful: {elem.name}")
        except Exception as ex:
            error_msg = f"Error uploading registry file: {ex}"
            logger.error(error_msg, exc_info=True)
            ui.notify(error_msg, type='negative')

    def _handle_inspection_upload(self, elem):
        """Handle inspection file upload events."""
        logger.info(f"Inspection file upload event received: {elem.name}")
        try:
            file_path = save_uploaded_file(UPLOAD_DIR, elem.name, elem.content)
            logger.debug(f"Inspection file saved to: {file_path}")
            self.uploaded_files.append(file_path)
            self.file_list.add_file(file_path)
            ui.notify(f'Tarkastuspöytäkirja tiedosto ladattu: {elem.name}', type='positive')
            logger.info(f"Inspection file upload successful: {elem.name}")
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