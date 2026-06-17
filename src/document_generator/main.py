import os
import io
import tempfile
import subprocess
import zipfile
from enum import StrEnum
from pathlib import Path
from typing import Any, Dict, List, Optional, TypedDict
import pandas as pd
import panel as pn
from docxtpl import DocxTemplate
from panelini import Panelini
from dotenv import load_dotenv
import logging

logger = logging.getLogger("DocumentGeneratorApp")


# Load environment variables

load_dotenv()

pn.extension("tabulator", notifications=True)


class FileFormat(StrEnum):
    WORD = "word"
    PDF = "pdf"
    ALL = "all"


class ZipFileInfo(TypedDict):
    file_format: FileFormat
    zip_buffer: io.BytesIO


class DocumentGeneratorApp:
    """
    Document Generator App - Renders Word templates with Excel data
    """

    def __init__(self):
        # Load LibreOffice path from environment
        self.libreoffice_path = os.getenv("LIBREOFFICE_PATH", "libreoffice")
        self.conversion_timeout = int(os.getenv("CONVERSION_TIMEOUT", "60"))

        # Data storage
        self.excel_data: pd.DataFrame = None
        self.templates: List[Dict[str, Any]] = []

        # Widgets
        self._create_widgets()

    def _create_widgets(self):
        """Create all UI widgets"""
        # File droppers
        self.excel_dropper = pn.widgets.FileInput(
            accept=".xlsx,.xls", multiple=False, name="📊 Upload Excel Data"
        )
        self.excel_dropper.param.watch(self._on_excel_upload, "value")

        self.template_dropper = pn.widgets.FileInput(
            accept=".docx", multiple=True, name="📄 Upload Word Templates"
        )
        self.template_dropper.param.watch(self._on_template_upload, "value")

        # Data table
        self.data_table = pn.widgets.Tabulator(
            pagination="local",
            page_size=20,
            selectable="checkbox",
            name="Excel Data",
            sizing_mode="stretch_width",
            height=400,
        )

        # Template table
        self.template_table_data = pd.DataFrame(
            columns=["File Name", "Naming Template"]
        )
        self.template_table = pn.widgets.Tabulator(
            self.template_table_data,
            disabled=False,
            editors={"Naming Template": {"type": "input"}},
            name="Templates",
            sizing_mode="stretch_width",
            height=200,
        )

        # Buttons
        self.clear_templates_btn = pn.widgets.Button(
            name="Clear Templates", button_type="warning", sizing_mode="stretch_width"
        )
        self.clear_templates_btn.on_click(self._clear_templates)

        self.preview_btn = pn.widgets.Button(
            name="Preview PDFs",
            button_type="primary",
            sizing_mode="stretch_width",
            disabled=True,
        )
        self.preview_btn.on_click(self._preview_documents)

        # File download widget (hidden)
        self.word_file_download = pn.widgets.FileDownload(
            name="Download Word Files (ZIP)",
            callback=self._download_word_documents,
            filename="documents_word.zip",
            button_type="success",
            sizing_mode="stretch_width",
            visible=True,
            disabled=True,
            embed=False,
            auto=True,
        )
        self.pdf_file_download = pn.widgets.FileDownload(
            name="Download PDF Files (ZIP)",
            callback=self._download_pdf_documents,
            filename="documents_pdf.zip",
            button_type="success",
            sizing_mode="stretch_width",
            visible=True,
            disabled=True,
            embed=False,
            auto=True,
        )
        self.all_formats_download = pn.widgets.FileDownload(
            name="Download Word & PDF Files (ZIP)",
            callback=self._download_all_formats,
            filename="documents.zip",
            button_type="success",
            sizing_mode="stretch_width",
            visible=True,
            disabled=True,
            embed=False,
            auto=True,
        )

        # The following three buttons are currently not in use:
        self.word_download_btn = pn.widgets.Button(
            name="Download Word Files (ZIP)",
            button_type="success",
            sizing_mode="stretch_width",
            disabled=True,
            visible=False,
        )
        self.word_download_btn.on_click(
            lambda event: self._trigger_download(event, FileFormat.WORD)
        )
        self.pdf_download_btn = pn.widgets.Button(
            name="Download PDF Files (ZIP)",
            button_type="success",
            sizing_mode="stretch_width",
            disabled=True,
            visible=False,
        )
        self.pdf_download_btn.on_click(
            lambda event: self._trigger_download(event, FileFormat.PDF)
        )
        self.all_formats_download_btn = pn.widgets.Button(
            name="Download Word & PDF Files (ZIP)",
            button_type="success",
            sizing_mode="stretch_width",
            disabled=True,
            visible=False,
        )
        self.all_formats_download_btn.on_click(
            lambda event: self._trigger_download(event, FileFormat.ALL)
        )

        # PDF preview area
        self.pdf_preview_area = pn.Column(
            name="PDF Previews", sizing_mode="stretch_width"
        )

        # Status indicator
        self.status = pn.pane.Markdown("Ready", sizing_mode="stretch_width")

    def _on_excel_upload(self, event):
        """Handle Excel file upload"""
        if event.new is None:
            return

        try:
            self.status.object = "⏳ Loading Excel file ..."

            # Parse Excel
            excel_bytes = io.BytesIO(event.new)
            self.excel_data = pd.read_excel(excel_bytes)

            # trim whitespace from column names
            self.excel_data.columns = self.excel_data.columns.str.strip()

            original_columns = self.excel_data.columns.tolist()
            # make all column names valid python identifiers
            # replace spaces and hyphens with underscores
            # add underscore prefix if column name starts with a digit
            # replace duplicate names with suffixes
            valid_columns = []
            seen = {}
            for col in original_columns:
                valid_col = col.strip().replace(" ", "_").replace("-", "_")
                # umlauts are allowed
                # valid_col = valid_col.replace('ä', 'ae').replace('ö', 'oe').replace('ü', 'ue').replace('ß', 'ss')
                if valid_col[0].isdigit():
                    valid_col = "_" + valid_col
                # replace all remaining invalid characters
                valid_col = "".join(
                    c if c.isalnum() or c == "_" else "_" for c in valid_col
                )
                if valid_col in seen:
                    seen[valid_col] += 1
                    valid_col = f"{valid_col}_{seen[valid_col]}"
                else:
                    seen[valid_col] = 0
                valid_columns.append(valid_col)

            # Update columns with valid names
            self.excel_data.columns = valid_columns
            self.excel_data = self.excel_data.loc[
                :, ~self.excel_data.columns.duplicated()
            ]

            # Update table
            self.data_table.value = self.excel_data

            status_msg = f"✅ Loaded {len(self.excel_data)} rows"
            for orig, valid in zip(original_columns, valid_columns):
                if orig != valid:
                    status_msg += f'\n- Warning: Column renamed: "{orig}" -> "{valid}"; use "{{{{ {valid} }}}}" in templates.'
            self.status.object = status_msg

            self._update_button_states()

            pn.state.notifications.success(f"Excel loaded: {len(self.excel_data)} rows")

        except Exception as e:
            self.status.object = f"❌ Error: {str(e)}"
            pn.state.notifications.error(f"Error loading Excel: {str(e)}")

    def _on_template_upload(self, event):
        """Handle template file upload"""
        if event.new is None:
            return

        try:
            # Get current templates
            current_templates = (
                self.template_table.value.to_dict("records")
                if len(self.template_table.value) > 0
                else []
            )

            # Add new templates
            filenames = self.template_dropper.filename
            if isinstance(filenames, str):
                filenames = [filenames]

            files = event.new
            if isinstance(files, bytes):
                files = [files]

            for filename, file_bytes in zip(filenames, files):
                # Create default naming template
                first_col = (
                    self.excel_data.columns[0] if self.excel_data is not None else "id"
                )
                base_name = Path(filename).stem
                naming_template = f"{base_name}_" + "{{ " + first_col + " }}"

                # Store template
                template_info = {
                    "File Name": filename,
                    "Naming Template": naming_template,
                    "bytes": file_bytes,
                }
                self.templates.append(template_info)

                # Add to table display
                current_templates.append(
                    {"File Name": filename, "Naming Template": naming_template}
                )

            # Update table
            self.template_table.value = pd.DataFrame(current_templates)

            self.status.object = f"✅ {len(self.templates)} template(s) loaded"
            self._update_button_states()

            pn.state.notifications.success(f"Added {len(filenames)} template(s)")

        except Exception as e:
            self.status.object = f"❌ Error: {str(e)}"
            pn.state.notifications.error(f"Error loading templates: {str(e)}")

    def _clear_templates(self, event):
        """Clear all templates"""
        self.templates = []
        self.template_table.value = pd.DataFrame(
            columns=["File Name", "Naming Template"]
        )
        self.status.object = "Templates cleared"
        self._update_button_states()
        pn.state.notifications.info("Templates cleared")

    def _update_button_states(self):
        """Update button enabled/disabled states"""
        has_data = self.excel_data is not None and len(self.excel_data) > 0
        has_templates = len(self.templates) > 0
        has_selection = (
            len(self.data_table.selection) > 0
            if self.data_table.value is not None
            else False
        )

        self.preview_btn.disabled = not (has_data and has_templates and has_selection)
        self.word_download_btn.disabled = not (has_data and has_templates and has_selection)
        self.word_file_download.disabled = not (has_data and has_templates and has_selection)
        self.pdf_download_btn.disabled = not (has_data and has_templates and has_selection)
        self.pdf_file_download.disabled = not (has_data and has_templates and has_selection)
        self.all_formats_download_btn.disabled = not (has_data and has_templates and has_selection)
        self.all_formats_download.disabled = not (has_data and has_templates and has_selection)

    @staticmethod
    def _create_empty_zip() -> io.BytesIO:
        """Create an empty but valid ZIP file"""
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            pass  # Create empty ZIP structure
        buffer.seek(0)
        return buffer

    def _create_zip_file_info(self, file_format: FileFormat, zip_buffer: io.BytesIO) -> ZipFileInfo:
        """Helper to create ZipFileInfo TypedDict"""
        return ZipFileInfo(
            file_format=file_format,
            zip_buffer=zip_buffer
        )

    @staticmethod
    def _render_template(
        template_bytes: bytes, row_data: Dict[str, Any], output_path: str
    ) -> str:
        """
        Render a Word template with row data

        Args:
            template_bytes: Template file bytes
            row_data: Dictionary representation of a single row
            output_path: Output path for rendered document

        Returns:
            Path to rendered document
        """
        # Create temp template file
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_template:
            tmp_template.write(template_bytes)
            tmp_template_path = tmp_template.name

        try:
            # Load template
            doc = DocxTemplate(tmp_template_path)

            # Render with row data
            doc.render(row_data)

            # Save result
            doc.save(output_path)

            return output_path

        finally:
            # Clean up temp template
            if os.path.exists(tmp_template_path):
                os.remove(tmp_template_path)

    def _convert_to_pdf(self, docx_path: str, pdf_dir=None) -> str:
        """
        Convert Word document to PDF using LibreOffice

        Args:
            docx_path: Path to Word document
            # pdf_dir: Directory to save PDF file to

        Returns:
            Path to PDF file
        """
        output_dir = os.path.dirname(docx_path)
        # if pdf_dir:
        #     output_dir = pdf_dir

        # Construct conversion command
        convert_cmd = [
            self.libreoffice_path,
            "--headless",
            "--convert-to",
            "pdf",
            docx_path,
            "--outdir",
            output_dir,
        ]

        # Run conversion
        result = subprocess.run(
            convert_cmd, capture_output=True, text=True, timeout=self.conversion_timeout
        )

        if result.returncode != 0:
            logging.error(f"LibreOffice conversion failed: {result.stderr}")
            raise Exception(f"LibreOffice conversion failed: {result.stderr}")

        # Determine PDF path
        pdf_path = docx_path.replace(".docx", ".pdf")

        # Wait for file to exist (with timeout)
        import time

        wait_time = 0
        while not os.path.exists(pdf_path) and wait_time < self.conversion_timeout:
            time.sleep(0.5)
            wait_time += 0.5

        if not os.path.exists(pdf_path):
            raise Exception(f"PDF file not created: {pdf_path}")

        return pdf_path

    @staticmethod
    def _get_output_filename(
        naming_template: str, row_data: Dict[str, Any], extension: str
    ) -> str:
        """
        Generate output filename from naming template

        Args:
            naming_template: Template string with placeholders
            row_data: Row data dictionary
            extension: File extension (.docx or .pdf)

        Returns:
            Generated filename
        """
        try:
            # Simple template rendering using string format
            filename = naming_template
            for key, value in row_data.items():
                filename = filename.replace("{{ " + key + " }}", str(value))

            # Ensure extension
            if not filename.endswith(extension):
                filename = filename + extension

            # Clean filename
            filename = "".join(
                c for c in filename if c.isalnum() or c in (" ", "_", "-", ".")
            )

            return filename
        except Exception as e:
            # Fallback to simple namin
            pn.state.notifications.warning(
                f"Error generating filename from template: {str(e)}. Using default naming."
            )
            return f"document_{row_data.get(list(row_data.keys())[0], 'output')}{extension}"

    def _preview_documents(self, event):
        """Generate and preview PDFs for selected rows"""
        try:
            self.status.object = "⏳ Generating previews ..."
            self.pdf_preview_area.clear()

            # Get selected rows
            selected_indices = self.data_table.selection
            if not selected_indices:
                pn.state.notifications.warning("No rows selected")
                return

            selected_rows = self.excel_data.iloc[selected_indices]

            # Get current naming templates from table
            template_configs = self.template_table.value.to_dict("records")

            preview_count = 0

            # Generate for each template and row combination
            for template_idx, template_info in enumerate(self.templates):
                template_bytes = template_info["bytes"]
                naming_template = template_configs[template_idx]["Naming Template"]

                for _, row in selected_rows.iterrows():
                    # Convert row to dictionary
                    row_data = row.to_dict()

                    # Convert all values to strings for template rendering
                    row_data = {
                        k: str(v) if pd.notna(v) else "" for k, v in row_data.items()
                    }

                    # Create temp directory for this document
                    with tempfile.TemporaryDirectory() as tmpdir:
                        # Generate output filename
                        docx_filename = self._get_output_filename(
                            naming_template, row_data, ".docx"
                        )
                        docx_path = os.path.join(tmpdir, docx_filename)

                        # Render template
                        self._render_template(template_bytes, row_data, docx_path)

                        # Convert to PDF
                        pdf_tmp_path = self._convert_to_pdf(docx_path)

                        # Read PDF bytes for preview
                        with open(pdf_tmp_path, "rb") as f:
                            pdf_bytes = f.read()

                        # Add to preview area
                        pdf_name = self._get_output_filename(
                            naming_template, row_data, ".pdf"
                        )
                        pdf_pane = pn.pane.PDF(
                            pdf_bytes,
                            name=pdf_name,
                            sizing_mode="stretch_width",
                            height=600,
                        )
                        self.pdf_preview_area.append(
                            pn.Card(
                                pdf_pane,
                                title=f"📄 {pdf_name}",
                                collapsible=True,
                                collapsed=preview_count > 0,
                            )
                        )

                        preview_count += 1

            self.status.object = f"✅ Generated {preview_count} preview(s)"
            pn.state.notifications.success(f"Generated {preview_count} preview(s)")

        except Exception as e:
            self.status.object = f"❌ Error: {str(e)}"
            pn.state.notifications.error(f"Error generating previews: {str(e)}")

    def _trigger_download(self, event, file_format: FileFormat = FileFormat.WORD):
        """Trigger the file download widget"""
        self.status.object = "⏳ Preparing download ..."
        # Trigger the FileDownload callback by incrementing clicks
        # self.file_download.clicks += 1
        # simulate click on self.file_download
        match file_format:
            case "word":
                self.word_file_download._click()
            case "pdf":
                self.pdf_file_download._click()
            case "all":
                self.all_formats_download._click()
        # self._download_documents()

    def _download_documents(self, as_pdf: bool = False) -> list[ZipFileInfo]:
        """
        Generate and return Word (or PDF) documents (depending on parameter
        'as_pdf') for selected rows as ZIP

        Returns:
            BytesIO object containing ZIP file
        """
        try:
            logger.info(f"Starting download: templates={len(self.templates)}, as_pdf={as_pdf}")

            # Early validation
            if not self.templates:
                pn.state.notifications.error("No templates uploaded")
                results: list[ZipFileInfo] = [
                    self._create_zip_file_info(FileFormat.WORD, self._create_empty_zip())
                ]
                if as_pdf:
                    results.append(
                        self._create_zip_file_info(FileFormat.PDF, self._create_empty_zip())
                    )
                return results

            # Get selected rows
            selected_indices = self.data_table.selection
            if not selected_indices:
                pn.state.notifications.warning("No rows selected")
                # Always return buffers for the formats the caller expects
                results: list[ZipFileInfo] = [
                    self._create_zip_file_info(FileFormat.WORD, self._create_empty_zip())
                ]
                if as_pdf:
                    results.append(
                        self._create_zip_file_info(FileFormat.PDF, self._create_empty_zip())
                    )
                return results

            selected_rows = self.excel_data.iloc[selected_indices]
            logger.info(f"Processing {len(selected_rows)} rows × {len(self.templates)} templates")

            # Get current naming templates from table
            template_configs = self.template_table.value.to_dict("records")

            # Create ZIP file in memory
            word_zip_buffer = io.BytesIO()
            pdf_zip_buffer = io.BytesIO()
            conversion_failures = []  # Track PDF conversion failures

            with zipfile.ZipFile(word_zip_buffer, "w", zipfile.ZIP_DEFLATED) as word_zipf, \
                zipfile.ZipFile(pdf_zip_buffer, "w", zipfile.ZIP_DEFLATED) as pdf_zipf:
                # Generate for each template and row combination
                for template_idx, template_info in enumerate(self.templates):
                    template_bytes = template_info["bytes"]
                    naming_template = template_configs[template_idx]["Naming Template"]

                    for _, row in selected_rows.iterrows():
                        # Convert row to dictionary
                        row_data = row.to_dict()

                        # Convert all values to strings for template rendering
                        row_data = {
                            k: str(v) if pd.notna(v) else ""
                            for k, v in row_data.items()
                        }

                        # Create temp file for this document
                        with tempfile.NamedTemporaryFile(
                            suffix=".docx", delete=False
                        ) as word_tmp_file:
                            word_tmp_path = word_tmp_file.name

                        # Ensure these are always defined for the finally block
                        pdf_tmp_path: Optional[str] = None

                        try:
                            # Render template
                            self._render_template(template_bytes, row_data, word_tmp_path)

                            # Generate output filename
                            word_output_filename = self._get_output_filename(
                                naming_template, row_data, ".docx"
                            )

                            # Add to ZIP
                            word_zipf.write(word_tmp_path, word_output_filename)

                            if as_pdf:
                                try:
                                    # Convert to PDF
                                    pdf_tmp_path = self._convert_to_pdf(word_tmp_path)

                                    # Generate PDF output filename
                                    pdf_output_filename = self._get_output_filename(
                                        naming_template, row_data, ".pdf"
                                    )

                                    # Add PDF to PDF ZIP
                                    pdf_zipf.write(pdf_tmp_path, pdf_output_filename)
                                except Exception as pdf_error:
                                    # Track failure but continue processing other documents
                                    conversion_failures.append({
                                        "document": word_output_filename,
                                        "error": str(pdf_error)
                                    })
                                    logger.warning(f"PDF conversion failed for {word_output_filename}: {pdf_error}")

                        finally:
                            # Clean up temp file
                            if os.path.exists(word_tmp_path):
                                os.remove(word_tmp_path)
                            # If PDF generation failed before pdf_tmp_path was set,
                            # avoid referencing an unassigned variable.
                            if as_pdf and pdf_tmp_path and os.path.exists(pdf_tmp_path):
                                os.remove(pdf_tmp_path)

            num_docs = len(selected_rows) * len(self.templates)

            # Check for conversion failures and notify user
            if conversion_failures:
                failure_count = len(conversion_failures)
                pn.state.notifications.warning(
                    f"⚠️ {failure_count} PDF conversion(s) failed. Word documents were generated successfully."
                )
                logger.error(f"PDF conversion failures: {conversion_failures}")
                self.status.object = f"⚠️ Generated {num_docs} document(s) but {failure_count} PDF conversion(s) failed"
            else:
                self.status.object = f"✅ Generated {num_docs} document(s)"
                pn.state.notifications.success(f"Downloading {num_docs} document(s)")

            # Reset buffer position
            word_zip_buffer.seek(0)
            pdf_zip_buffer.seek(0)

            logger.info(f"Created Word ZIP: {word_zip_buffer.getbuffer().nbytes} bytes")
            if as_pdf:
                logger.info(f"Created PDF ZIP: {pdf_zip_buffer.getbuffer().nbytes} bytes")

            return_object = [
                ZipFileInfo(
                    file_format=FileFormat.WORD,
                    zip_buffer=word_zip_buffer
                )
            ]
            if as_pdf:
                return_object.append(
                    ZipFileInfo(
                        file_format=FileFormat.PDF,
                        zip_buffer=pdf_zip_buffer
                    )
                )

            return return_object

        except Exception as e:
            self.status.object = f"❌ Error: {str(e)}"
            pn.state.notifications.error(f"Error generating documents: {str(e)}")
            results: list[ZipFileInfo] = [
                self._create_zip_file_info(FileFormat.WORD, self._create_empty_zip())
            ]
            if as_pdf:
                results.append(
                    self._create_zip_file_info(FileFormat.PDF, self._create_empty_zip())
                )
            return results

    @staticmethod
    def _get_zip_buffer(
        file_format: str, zip_infos: list[ZipFileInfo]
    ) -> io.BytesIO:
        """Return the zip buffer for the requested format.

        Panel's FileDownload cannot transfer None, so we always return a BytesIO,
        even if generation failed.
        """
        for info in zip_infos:
            if info["file_format"] == file_format:
                info["zip_buffer"].seek(0)
                return info["zip_buffer"]
        return DocumentGeneratorApp._create_empty_zip()

    def _download_word_documents(self) -> io.BytesIO:
        """
        Generate and return Word documents for selected rows as ZIP
        This is a callback for FileDownload widget

        Returns:
            BytesIO object containing ZIP file
        """
        results = self._download_documents(as_pdf=False)
        return self._get_zip_buffer(file_format=FileFormat.WORD, zip_infos=results)

    def _download_pdf_documents(self) -> io.BytesIO:
        """
        Generate and return PDF documents for selected rows as ZIP
        This is a callback for FileDownload widget

        Returns:
            BytesIO object containing ZIP file
        """
        results = self._download_documents(as_pdf=True)
        return self._get_zip_buffer(file_format=FileFormat.PDF, zip_infos=results)

    def _download_all_formats(self) -> io.BytesIO:
        # Create a temporary ZIP containing both Word and PDF ZIPs
        logger.info("Creating combined Word + PDF download")
        results = self._download_documents(as_pdf=True)
        combined_zip_buffer = io.BytesIO()
        with zipfile.ZipFile(combined_zip_buffer, "w", zipfile.ZIP_DEFLATED) as combined_zip:
            for result in results:
                result["zip_buffer"].seek(0)  # Defensive: ensure buffer at start
                combined_zip.writestr(
                    f"documents_{result['file_format']}.zip",
                    result["zip_buffer"].read()
                )
        combined_zip_buffer.seek(0)
        logger.info(f"Created combined ZIP: {combined_zip_buffer.tell()} bytes")
        return combined_zip_buffer

    def get_sidebar(self):
        """Get sidebar components"""
        return [
            pn.Card(
                collapsible=False,
                title="📤 Upload Files",
                objects=[
                    self.excel_dropper,
                    self.template_dropper,
                    self.clear_templates_btn,
                ],
            ),
            # pn.layout.Divider(),
            pn.Card(
                collapsible=False,
                title="⚙️ Actions",
                objects=[
                    self.preview_btn,
                    self.word_download_btn,  # currently not in use
                    self.pdf_download_btn,  # currently not in use
                    self.all_formats_download_btn,  # currently not in use
                    self.word_file_download,
                    self.pdf_file_download,
                    self.all_formats_download,
                ],
            ),
            # pn.layout.Divider(),
            pn.Card(
                collapsible=False,
                title="Status",
                objects=[
                    self.status,
                ],
            ),
        ]

    def get_main(self):
        """Get main area components"""
        # Watch for table selection changes
        self.data_table.param.watch(lambda e: self._update_button_states(), "selection")

        return [
            pn.Card(self.data_table, title="📊 Data Table", collapsible=False),
            pn.Card(
                self.template_table,
                title="📝 Templates Configuration",
                collapsible=False,
            ),
            pn.Card(
                self.pdf_preview_area,
                title="👁️ PDF Previews",
                collapsible=True,
                collapsed=True,
            ),
        ]


# Create the app

app = Panelini(title="Document Generator", sidebar_right_enabled=True)

# Initialize the document generator

doc_gen = DocumentGeneratorApp()

# Set sidebar and main content

app.sidebar_set(objects=doc_gen.get_sidebar())
app.main_set(objects=doc_gen.get_main())

# Make servable: Use for `panel serve` command in development

app.servable()

if __name__ == "__main__":
    pn.serve(app, port=int(os.getenv("INTERNAL_PORT")))
