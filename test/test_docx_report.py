"""Tests the docx_report module."""

from datetime import datetime
import os
from docx.document import Document
import matplotlib.pyplot as plt
import pandas as pd
import pytest
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
    assert isinstance(report.doc, Document)
    assert report.title == "Test Report"
    assert report.doc.paragraphs[1].text == "Test Report"
