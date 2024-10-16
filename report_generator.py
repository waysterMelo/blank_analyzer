# report.py
import os
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

class ReportGenerator:
    def __init__(self):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "PDF Analysis Report"
        self.ws.append(["Arquivo PDF", "Página", "Status", "Porcentagem de Pixels Brancos", "OCR Realizado", "Texto Extraído"])

    def add_record(self, pdf_name, page_num, status, white_pixel_percentage, ocr_performed, extracted_text):
        self.ws.append([
            pdf_name,
            page_num,
            status,
            f"{white_pixel_percentage:.2%}",
            "Sim" if ocr_performed else "Não",
            extracted_text
        ])

    def finalize(self, output_path):
        # Create table with style
        tab = Table(displayName="PDFAnalysisTable", ref=f"A1:F{self.ws.max_row}")
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=True
        )
        tab.tableStyleInfo = style
        self.ws.add_table(tab)

        # Adjust column widths
        for col in self.ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = (max_length + 2)
            self.ws.column_dimensions[col_letter].width = adjusted_width

        # Save workbook
        self.wb.save(output_path)
