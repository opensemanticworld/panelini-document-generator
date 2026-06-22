import os
import io
import tempfile
import subprocess
import zipfile
from pathlib import Path
from typing import List, Dict, Any, TypedDict, Optional
import pandas as pd
import panel as pn
from docxtpl import DocxTemplate
from panelini import Panelini
from dotenv import load_dotenv

# Load environment variables

load_dotenv()

pn.extension("tabulator", notifications=True)


class DocumentGeneratorApp:
    # Copy from here on
    class PDFDocumentInfo(TypedDict):
        name: str
        path: str


    def _generate_pdf_documents(self) -> list[PDFDocumentInfo]:
        """Generate PDF documents for all selected rows and templates"""

        pdf_infos = []

        # This function is currently not used but can be implemented for direct PDF downloads
        self.status.object = "⏳ Generating PDF documents ..."
        # Get selected rows
        selected_indices = self.data_table.selection
        if not selected_indices:
            pn.state.notifications.warning("No rows selected")
            return []

        selected_rows = self.excel_data.iloc[selected_indices]

        # Get current naming templates from table
        template_configs = self.template_table.value.to_dict("records")

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
                    pdf_path = self._convert_to_pdf(docx_path)
                    pdf_name = self._get_output_filename(
                        naming_template, row_data, ".pdf"
                    )
                    pdf_info: DocumentGeneratorApp.PDFDocumentInfo = {
                        "name": pdf_name,
                        "path": pdf_path,
                    }
                    pdf_infos.append(pdf_info)

        return pdf_infos


    def _preview_documents(self, event):
        """Generate and preview PDFs for selected rows"""
        try:
            self.pdf_preview_area.clear()
            pdf_infos = self._generate_pdf_documents()

            preview_count = 0

            for pdf_info in pdf_infos:
                pdf_name = pdf_info["name"]
                # Read PDF bytes for preview
                with open(pdf_info["path"], "rb") as f:
                    pdf_bytes = f.read()

                # Add to preview area
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
