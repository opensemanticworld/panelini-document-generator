import os
import io
import json
import tempfile
import subprocess
import zipfile
from pathlib import Path
from typing import List, Dict, Any
import pandas as pd
import panel as pn
from docxtpl import DocxTemplate
from panelini import Panelini
from dotenv import load_dotenv

# Load environment variables

load_dotenv()

pn.extension('tabulator', notifications=True)


class DocumentGeneratorApp:
    """
    Document Generator App - Renders Word templates with Excel data
    """
    
    def __init__(self):
        # Load LibreOffice path from environment
        self.libreoffice_path = os.getenv('LIBREOFFICE_PATH', 'libreoffice')
        self.conversion_timeout = int(os.getenv('CONVERSION_TIMEOUT', '60'))
        
        # Data storage
        self.excel_data: pd.DataFrame = None
        self.templates: List[Dict[str, Any]] = []
        
        # Widgets
        self._create_widgets()
        
    def _create_widgets(self):
        """Create all UI widgets"""
        # File droppers
        self.excel_dropper = pn.widgets.FileInput(
            accept='.xlsx,.xls',
            multiple=False,
            name='📊 Upload Excel Data'
        )
        self.excel_dropper.param.watch(self._on_excel_upload, 'value')
        
        self.template_dropper = pn.widgets.FileInput(
            accept='.docx',
            multiple=True,
            name='📄 Upload Word Templates'
        )
        self.template_dropper.param.watch(self._on_template_upload, 'value')
        
        # Data table
        self.data_table = pn.widgets.Tabulator(
            pagination='local',
            page_size=20,
            selectable='checkbox',
            name='Excel Data',
            sizing_mode='stretch_width',
            height=400
        )
        
        # Template table
        self.template_table_data = pd.DataFrame(
            columns=['File Name', 'Naming Template']
        )
        self.template_table = pn.widgets.Tabulator(
            self.template_table_data,
            disabled=False,
            editors={'Naming Template': {'type': 'input'}},
            name='Templates',
            sizing_mode='stretch_width',
            height=200
        )
        
        # Buttons
        self.clear_templates_btn = pn.widgets.Button(
            name='Clear Templates',
            button_type='warning',
            sizing_mode='stretch_width'
        )
        self.clear_templates_btn.on_click(self._clear_templates)
        
        self.preview_btn = pn.widgets.Button(
            name='Preview PDFs',
            button_type='primary',
            sizing_mode='stretch_width',
            disabled=True
        )
        self.preview_btn.on_click(self._preview_documents)
        
        # File download widget (hidden)
        self.file_download = pn.widgets.FileDownload(
            name='Download Word Files (ZIP)',
            callback=self._download_documents,
            filename='documents.zip',
            button_type='success',
            visible=True,
            disabled=True,
            embed=False,
            auto=True,
        )
        
        # currently no used
        self.download_btn = pn.widgets.Button(
            name='Download Word Files (ZIP)',
            button_type='success',
            sizing_mode='stretch_width',
            disabled=True,
            visible=False
        )
        self.download_btn.on_click(self._trigger_download)
        
        # PDF preview area
        self.pdf_preview_area = pn.Column(
            name='PDF Previews',
            sizing_mode='stretch_width'
        )
        
        # Status indicator
        self.status = pn.pane.Markdown('Ready', sizing_mode='stretch_width')
        
    def _on_excel_upload(self, event):
        """Handle Excel file upload"""
        if event.new is None:
            return
            
        try:
            self.status.object = '⏳ Loading Excel file...'
            
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
                valid_col = col.strip().replace(' ', '_').replace('-', '_')
                # umlauts are allowed
                # valid_col = valid_col.replace('ä', 'ae').replace('ö', 'oe').replace('ü', 'ue').replace('ß', 'ss')
                if valid_col[0].isdigit():
                    valid_col = '_' + valid_col
                # replace all remaining invalid characters
                valid_col = ''.join(c if c.isalnum() or c == '_' else '_' for c in valid_col)
                if valid_col in seen:
                    seen[valid_col] += 1
                    valid_col = f"{valid_col}_{seen[valid_col]}"
                else:
                    seen[valid_col] = 0
                valid_columns.append(valid_col)
            
            # Update columns with valid names
            self.excel_data.columns = valid_columns
            self.excel_data = self.excel_data.loc[:, ~self.excel_data.columns.duplicated()]
            
            # Update table
            self.data_table.value = self.excel_data
            
            status_msg = f'✅ Loaded {len(self.excel_data)} rows'
            for orig, valid in zip(original_columns, valid_columns):
                if orig != valid:
                    status_msg += f'\n- Warning: Column renamed: "{orig}" -> "{valid}"; use "{{{{ {valid} }}}}" in templates.'
            self.status.object = status_msg
            
            self._update_button_states()
            
            pn.state.notifications.success(
                f'Excel loaded: {len(self.excel_data)} rows'
            )
            
        except Exception as e:
            self.status.object = f'❌ Error: {str(e)}'
            pn.state.notifications.error(f'Error loading Excel: {str(e)}')
            
    def _on_template_upload(self, event):
        """Handle template file upload"""
        if event.new is None:
            return
            
        try:
            # Get current templates
            current_templates = self.template_table.value.to_dict('records') if len(self.template_table.value) > 0 else []
            
            # Add new templates
            filenames = self.template_dropper.filename
            if isinstance(filenames, str):
                filenames = [filenames]
                
            files = event.new
            if isinstance(files, bytes):
                files = [files]
            
            for filename, file_bytes in zip(filenames, files):
                # Create default naming template
                first_col = self.excel_data.columns[0] if self.excel_data is not None else 'id'
                base_name = Path(filename).stem
                naming_template = f"{base_name}_" + "{{ " + first_col + " }}"
                
                # Store template
                template_info = {
                    'File Name': filename,
                    'Naming Template': naming_template,
                    'bytes': file_bytes
                }
                self.templates.append(template_info)
                
                # Add to table display
                current_templates.append({
                    'File Name': filename,
                    'Naming Template': naming_template
                })
            
            # Update table
            self.template_table.value = pd.DataFrame(current_templates)
            
            self.status.object = f'✅ {len(self.templates)} template(s) loaded'
            self._update_button_states()
            
            pn.state.notifications.success(f'Added {len(filenames)} template(s)')
            
        except Exception as e:
            self.status.object = f'❌ Error: {str(e)}'
            pn.state.notifications.error(f'Error loading templates: {str(e)}')
            
    def _clear_templates(self, event):
        """Clear all templates"""
        self.templates = []
        self.template_table.value = pd.DataFrame(columns=['File Name', 'Naming Template'])
        self.status.object = 'Templates cleared'
        self._update_button_states()
        pn.state.notifications.info('Templates cleared')
        
    def _update_button_states(self):
        """Update button enabled/disabled states"""
        has_data = self.excel_data is not None and len(self.excel_data) > 0
        has_templates = len(self.templates) > 0
        has_selection = len(self.data_table.selection) > 0 if self.data_table.value is not None else False
        
        self.preview_btn.disabled = not (has_data and has_templates and has_selection)
        self.download_btn.disabled = not (has_data and has_templates and has_selection)
        self.file_download.disabled = not (has_data and has_templates and has_selection)
        
    def _render_template(self, template_bytes: bytes, row_data: Dict[str, Any], 
                        output_path: str) -> str:
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
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_template:
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
                
    def _convert_to_pdf(self, docx_path: str) -> str:
        """
        Convert Word document to PDF using LibreOffice
        
        Args:
            docx_path: Path to Word document
            
        Returns:
            Path to PDF file
        """
        output_dir = os.path.dirname(docx_path)
        
        # Construct conversion command
        convert_cmd = [
            self.libreoffice_path,
            '--headless',
            '--convert-to',
            'pdf',
            docx_path,
            '--outdir',
            output_dir
        ]
        
        # Run conversion
        result = subprocess.run(
            convert_cmd,
            capture_output=True,
            text=True,
            timeout=self.conversion_timeout
        )
        
        if result.returncode != 0:
            raise Exception(f"LibreOffice conversion failed: {result.stderr}")
        
        # Determine PDF path
        pdf_path = docx_path.replace('.docx', '.pdf')
        
        # Wait for file to exist (with timeout)
        import time
        wait_time = 0
        while not os.path.exists(pdf_path) and wait_time < self.conversion_timeout:
            time.sleep(0.5)
            wait_time += 0.5
            
        if not os.path.exists(pdf_path):
            raise Exception(f"PDF file not created: {pdf_path}")
            
        return pdf_path
        
    def _get_output_filename(self, naming_template: str, row_data: Dict[str, Any], 
                            extension: str) -> str:
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
            filename = "".join(c for c in filename if c.isalnum() or c in (' ', '_', '-', '.'))
            
            return filename
        except Exception as e:
            # Fallback to simple naming
            return f"document_{row_data.get(list(row_data.keys())[0], 'output')}{extension}"
            
    def _preview_documents(self, event):
        """Generate and preview PDFs for selected rows"""
        try:
            self.status.object = '⏳ Generating previews...'
            self.pdf_preview_area.clear()
            
            # Get selected rows
            selected_indices = self.data_table.selection
            if not selected_indices:
                pn.state.notifications.warning('No rows selected')
                return
                
            selected_rows = self.excel_data.iloc[selected_indices]
            
            # Get current naming templates from table
            template_configs = self.template_table.value.to_dict('records')
            
            preview_count = 0
            
            # Generate for each template and row combination
            for template_idx, template_info in enumerate(self.templates):
                template_bytes = template_info['bytes']
                naming_template = template_configs[template_idx]['Naming Template']
                
                for _, row in selected_rows.iterrows():
                    # Convert row to dictionary
                    row_data = row.to_dict()
                    
                    # Convert all values to strings for template rendering
                    row_data = {k: str(v) if pd.notna(v) else '' for k, v in row_data.items()}
                    
                    # Create temp directory for this document
                    with tempfile.TemporaryDirectory() as tmpdir:
                        # Generate output filename
                        docx_filename = self._get_output_filename(
                            naming_template, row_data, '.docx'
                        )
                        docx_path = os.path.join(tmpdir, docx_filename)
                        
                        # Render template
                        self._render_template(template_bytes, row_data, docx_path)
                        
                        # Convert to PDF
                        pdf_path = self._convert_to_pdf(docx_path)
                        
                        # Read PDF bytes for preview
                        with open(pdf_path, 'rb') as f:
                            pdf_bytes = f.read()
                        
                        # Add to preview area
                        pdf_name = self._get_output_filename(
                            naming_template, row_data, '.pdf'
                        )
                        pdf_pane = pn.pane.PDF(
                            pdf_bytes,
                            name=pdf_name,
                            sizing_mode='stretch_width',
                            height=600
                        )
                        self.pdf_preview_area.append(
                            pn.Card(
                                pdf_pane,
                                title=f"📄 {pdf_name}",
                                collapsible=True,
                                collapsed=preview_count > 0
                            )
                        )
                        
                        preview_count += 1
            
            self.status.object = f'✅ Generated {preview_count} preview(s)'
            pn.state.notifications.success(f'Generated {preview_count} preview(s)')
            
        except Exception as e:
            self.status.object = f'❌ Error: {str(e)}'
            pn.state.notifications.error(f'Error generating previews: {str(e)}')
    
    def _trigger_download(self, event):
        """Trigger the file download widget"""
        self.status.object = '⏳ Preparing download...'
        # Trigger the FileDownload callback by incrementing clicks
        #self.file_download.clicks += 1
        # simulate click on self.file_download
        self.file_download._click()
        #self._download_documents()
            
    def _download_documents(self) -> io.BytesIO:
        """
        Generate and return Word documents for selected rows as ZIP
        This is a callback for FileDownload widget
        
        Returns:
            BytesIO object containing ZIP file
        """
        try:
            # Get selected rows
            selected_indices = self.data_table.selection
            if not selected_indices:
                pn.state.notifications.warning('No rows selected')
                return io.BytesIO()
                
            selected_rows = self.excel_data.iloc[selected_indices]
            
            # Get current naming templates from table
            template_configs = self.template_table.value.to_dict('records')
            
            # Create ZIP file in memory
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Generate for each template and row combination
                for template_idx, template_info in enumerate(self.templates):
                    template_bytes = template_info['bytes']
                    naming_template = template_configs[template_idx]['Naming Template']
                    
                    for _, row in selected_rows.iterrows():
                        # Convert row to dictionary
                        row_data = row.to_dict()
                        
                        # Convert all values to strings for template rendering
                        row_data = {k: str(v) if pd.notna(v) else '' for k, v in row_data.items()}
                        
                        # Create temp file for this document
                        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp_file:
                            tmp_path = tmp_file.name
                        
                        try:
                            # Render template
                            self._render_template(template_bytes, row_data, tmp_path)
                            
                            # Generate output filename
                            output_filename = self._get_output_filename(
                                naming_template, row_data, '.docx'
                            )
                            
                            # Add to ZIP
                            zipf.write(tmp_path, output_filename)
                            
                        finally:
                            # Clean up temp file
                            if os.path.exists(tmp_path):
                                os.remove(tmp_path)
            
            num_docs = len(selected_rows) * len(self.templates)
            self.status.object = f'✅ Generated {num_docs} document(s)'
            pn.state.notifications.success(f'Downloading {num_docs} document(s)')
            
            # Reset buffer position
            zip_buffer.seek(0)
            return zip_buffer
            
        except Exception as e:
            self.status.object = f'❌ Error: {str(e)}'
            pn.state.notifications.error(f'Error generating documents: {str(e)}')
            return io.BytesIO()
            
    def get_sidebar(self):
        """Get sidebar components"""
        return [
            pn.pane.Markdown('## 📤 Upload Files'),
            self.excel_dropper,
            self.template_dropper,
            self.clear_templates_btn,
            pn.layout.Divider(),
            pn.pane.Markdown('## ⚙️ Actions'),
            self.preview_btn,
            self.download_btn,
            self.file_download,  # Hidden download widget
            pn.layout.Divider(),
            self.status
        ]
        
    def get_main(self):
        """Get main area components"""
        # Watch for table selection changes
        self.data_table.param.watch(lambda e: self._update_button_states(), 'selection')
        
        return [
            pn.Card(
                self.data_table,
                title='📊 Data Table',
                collapsible=False
            ),
            pn.Card(
                self.template_table,
                title='📝 Templates Configuration',
                collapsible=False
            ),
            pn.Card(
                self.pdf_preview_area,
                title='👁️ PDF Previews',
                collapsible=True,
                collapsed=True
            )
        ]


# Create the app

app = Panelini(
    title="Document Generator",
)

# Initialize the document generator

doc_gen = DocumentGeneratorApp()

# Set sidebar and main content

app.sidebar_set(objects=doc_gen.get_sidebar())
app.main_set(objects=doc_gen.get_main())

# Make servable

app.servable()

if __name__ == "__main__":
    app.serve(port=5010)