import datetime
import os
import io
from docx import Document
from docx.shared import Inches
import pandas as pd
from plotly.subplots import make_subplots
from pathlib import Path
from docx2pdf import convert

from .data import *
from ._plots import plots

today = datetime.date.today()

def _save_plots_to_docx(figures, filename, report_type, output_format='docx'):
    """
    Helper function that accepts plotly figures and outputs a DOCX or PDF report.
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
        paragraph = doc.add_paragraph()
        paragraph.alignment = 1
        paragraph.add_run().add_picture(img_stream, width=Inches(6.0))

        caption = doc.add_paragraph(f"Figure {i+1}")
        caption.alignment = 1
        doc.add_paragraph("\n")

    # Add Comments Section
    doc.add_page_break()
    doc.add_heading('User Comments', level=1)
    doc.add_paragraph("Please add any notes regarding this report below:")
    doc.add_paragraph("\n" * 10)

    downloads_path = str(Path.home() / "Downloads")

    if not filename.endswith('.docx'):
        filename = filename + '.docx'

    docx_path = os.path.join(downloads_path, filename)
    doc.save(docx_path)

    if output_format.lower() == 'pdf':
        pdf_path = docx_path.replace('.docx', '.pdf')
        try:
            convert(docx_path, pdf_path)
            os.remove(docx_path)
            abs_path = os.path.abspath(pdf_path)
        except Exception as e:
            print(f"PDF conversion failed: {e}. Saving as DOCX instead.")
            abs_path = os.path.abspath(docx_path)
        return abs_path
    else:
        abs_path = os.path.abspath(docx_path)
        print(f"Report saved to: {abs_path}")
        return abs_path

def _create_side_by_side_figure(fig_list, titles, overall_title):
    """
    Helper function to put two plotly figures side by side for easier visualization.
    """
    if len(fig_list) < 2:
        return fig_list[0] if fig_list else None

    num_plots = len(fig_list)
    combined_fig = make_subplots(rows=1, cols=num_plots, subplot_titles=titles)

    for i, fig in enumerate(fig_list):
        for trace in fig.data:
            combined_fig.add_trace(trace, row=1, col=i+1)

        if fig.layout.xaxis.title.text:
            combined_fig.update_xaxes(title_text=fig.layout.xaxis.title.text, row=1, col=i+1)
        if fig.layout.yaxis.title.text:
            combined_fig.update_yaxes(title_text=fig.layout.yaxis.title.text, row=1, col=i+1)

    combined_fig.update_layout(
        height=600,
        width=600*num_plots,
        title_text=overall_title,
        showlegend=True
    )

    return combined_fig


def forecast_report(riverids, dates, side_by_side=False, output_format='docx'):
    """"
    Generate a forecast report using one or more rivers/dates and saves as a docx/pdf to the Downloads folder.

    Args:
        riverids: int or list of int -- One or more River ID(s) to generate forecasts for.
        dates: str or list of str -- Date(s) for forecast in YYYY-MM-DD format.
        side_by_side: bool, optional. If True with multiple dates, display forecasts side by side for comparison. Default is False.
        output_format: {'docx', 'pdf'}, optional. Output format for report. Default is 'docx'.

    Returns:
        str - absolute path to the saved report file in the Downloads folder.

    Notes:
        PDF conversion on macOS may briefly show Microsoft Word opening and closing. This is normal as Word is used
        for the conversion process.
    """
    if not isinstance(riverids, list):
        riverids = [riverids]

    if not isinstance(dates, list):
        dates = [dates]

    figures = []

    if side_by_side and len(dates) >= 2:

        for r in riverids:
            fig_list = []
            subplot_titles = []

            for d in dates:
                formatted_date = pd.to_datetime(str(d)).strftime('%Y%m%d')
                data = forecast(river_id=r, date=formatted_date)
                fig = plots.forecast(data)
                fig_list.append(fig)
                subplot_titles.append(f"Forecast: {d}")

            combined_fig = _create_side_by_side_figure(
                fig_list,
                subplot_titles,
                f"Side-by-Side Forecast Comparison for River {r}"
            )
            figures.append(combined_fig)

    else:

        for d in dates:
            formatted_date = pd.to_datetime(str(d)).strftime('%Y%m%d')
            for r in riverids:
                data = forecast(river_id=r, date=formatted_date)
                fig = plots.forecast(data, plot_titles=["", f"Forecast for River: {r}"])
                figures.append(fig)

    report_date = dates[0] if len(dates) == 1 else f"{dates[0]}_to_{dates[-1]}"
    return _save_plots_to_docx(figures, f"forecast_report_{report_date}.docx", "Forecast Report", output_format)

def retrospective_report(riverids, output_format='docx'):

    figures = []

    if not isinstance(riverids, list):
        riverids = [riverids]

    for r in riverids:
        data = retrospective(river_id=r)
        fig = plots.retrospective(data, plot_titles=["", f"Retrospective for River: {r}"])
        figures.append(fig)

    return _save_plots_to_docx(figures, "retrospective_report.docx", "Retrospective Report", output_format)

## the plots from this report are really hard to read
def in_depth_retro(riverid, output_format='docx'):

    daily = retro_daily(riverid)
    monthly = retro_monthly(riverid)
    yearly = retro_yearly(riverid)

    fig1 = plots.daily_averages(daily, plot_titles=["", f"Daily Averages for River {riverid}"])
    fig2 = plots.monthly_averages(monthly, plot_titles=["", f"Monthly Averages for River {riverid}"])
    fig3 = plots.annual_averages(yearly, plot_titles=["", f"Annual Averages for River {riverid}"])

    return _save_plots_to_docx([fig1, fig2, fig3], f"in_depth_report_{riverid}.docx", "In Depth Retro", output_format)

## this one doesn't work super well, ask Riley
def return_period_comparison(riverids, date, output_format='docx'):

    formatted_date = pd.to_datetime(str(date)).strftime('%Y%m%d')

    figures = []

    if not isinstance(riverids, list):
        riverids = [riverids]

    for r in riverids:
        data = forecast(river_id=r, date=formatted_date)
        return_period = return_periods(river_id=r)
        fig = plots.forecast(df=data, rp_df=return_period)
        figures.append(fig)
    return _save_plots_to_docx(figures, f"return_period_comparison_{date}.docx", "Return Period Comparison", output_format)

def fdc_curves(riverids, output_format='docx'):

    figures = []

    for r in riverids:
        data = fdc(river_id=r)
        fig = plots.flow_duration_curve(df=data, plot_titles=[f" for River {r}"])
        figures.append(fig)

    return _save_plots_to_docx(figures, "fdc_curves_report.docx", "Flow Duration Curves", output_format)