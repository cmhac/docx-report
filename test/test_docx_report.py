"""Tests the docx_report module."""

from datetime import datetime
import os
from docx.document import Document as DocumentBaseClass
from docx import Document
import pandas as pd
from docx_report import DocxReport


# Sample data for testing
data = {
    "dates": [datetime(2021, 1, 1), datetime(2021, 1, 2), datetime(2021, 1, 3)],
    "values": [1.23, 4.56, 7.89],
}
df = pd.DataFrame(data)


def test_initialization():
    """Tests the initialization of a DocxReport object."""
    report = DocxReport("Test Report")
    # raise ValueError(type(report.doc))
    assert isinstance(report.doc, DocumentBaseClass)
    assert report.title == "Test Report"
    assert report.doc.paragraphs[1].text == "Test Report"


def test_cleanup_dataframe():
    """Tests the _cleanup_dataframe method."""
    report = DocxReport("Test Report")
    cleaned_df = report._cleanup_dataframe(df)  # pylint: disable=protected-access
    assert cleaned_df["dates"].dtype == "object"  # Dates converted to strings
    assert cleaned_df["values"][0] == 1.2  # Values rounded to 1 decimal place


def test_add_plot(tmp_path):
    """Tests the add_plot method."""
    report = DocxReport("Test Report")
    report.add_plot(df, "Test Plot", "X Axis", "Y Axis")
    report.save(tmp_path / "test.docx")
    doc = Document(tmp_path / "test.docx")
    # check that an image was added
    assert len(doc.inline_shapes) == 1
    assert os.path.exists(tmp_path / "temp.png") is False  # Temporary file deleted


def test_add_table():
    """Tests the add_table method."""
    report = DocxReport("Test Report")
    report.add_table(df)
    assert report.doc.tables[0].rows[0].cells[0].text == "index"  # header added
    assert report.doc.tables[0].rows[1].cells[1].text == "2021-01-01"  # Data added


def test_add_list_bullet():
    """Tests the add_list_bullet method."""
    report = DocxReport("Test Report")
    report.add_list_bullet("Test Bullet")
    assert report.doc.paragraphs[2].text == "Test Bullet"  # Bullet added


def test_save(tmp_path):
    """Tests the save method."""
    report = DocxReport("Test Report")
    report.save(tmp_path / "test.docx")
    assert os.path.exists(tmp_path / "test.docx")  # File saved
