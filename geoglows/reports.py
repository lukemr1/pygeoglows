import plotly.io as pio
import datetime
import os
import io
pio.kaleido.scope.mathjax = None
from PyPDF2 import PdfMerger
from docx import Document
from docx.shared import Inches
import tempfile
import pandas as pd
from plotly.subplots import make_subplots

from .data import forecast
from ._plots import plots

today = datetime.date.today()

def _save_plots_to_docx(figures, filename, report_type):
    """
    Helper function that accepts plotly figures and outputs a DOCX report.
    """
    doc = Document()
    # Set 8.5 x 11 inches (Letter size)
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11)

    doc.add_heading('Hydrology Report', 0)
    doc.add_paragraph(f'Date Generated: {today}')
    doc.add_paragraph(f'Report Type: {report_type}')
    doc.add_paragraph('_' * 50)
    doc.add_paragraph('\n')

    for i, fig in enumerate(figures):
        # Save fig to bytes as png
        img_bytes = fig.to_image(format="png", width=1000, height=600, scale=2)
        img_stream = io.BytesIO(img_bytes)

        # Add to doc
        doc.add_picture(img_stream, width=Inches(7.0))
        doc.add_paragraph(f"Figure {i + 1}")
        doc.add_paragraph("\n")

    # Add Comments Section
    doc.add_page_break()
    doc.add_heading('User Comments', level=1)
    doc.add_paragraph("Please add any notes regarding this report below:")
    doc.add_paragraph("\n" * 10)

    doc.save(filename)

    abs_path = os.path.abspath(filename)
    print(f"Report saved to: {abs_path}")
    return abs_path


def forecast_report(riverids=None, user_data=None, date=None):

    figures = []

    if riverids is not None:
        if date is None:
            raise ValueError("date is required when providing riverids")

        formatted_date = pd.to_datetime(date).strftime('%Y%m%d')

        for r in riverids:
            data = forecast(river_id=r, date=formatted_date)
            fig = plots.forecast(data, plot_titles=["", f"Forecast for River: {r}"])
            figures.append(fig)

        report_date = date

    elif user_data is not None:

        fig = plots.forecast(user_data, plot_titles=["", f"Forecast for River"])
        figures.append(fig)

        report_date = date if date else today

    else:
        raise ValueError("Must provide either 'riverids' or 'data'")

    return _save_plots_to_docx(figures, f"forecast_report_{report_date}", "Forecast Report")