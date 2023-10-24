# docx-report

This is a simple wrapper for the [`python-docx`](https://python-docx.readthedocs.io/en/latest/) package that makes creating Word documents in python easier. It contains convenience methods for creating tables and inserting data visualizations directly from Pandas dataframes, as well as simpler syntax for generally writing text to a document.

## Installation

```bash
pip install docx-report
```

## Usage

```python
from docx_report import DocxReport

# Create a new document
doc = DocxReport(title="My Report")

# Add a heading
doc.add_heading("My Heading")

# Add a paragraph
doc.add_paragraph("This is a paragraph.")

# Add a table
doc.add_table(df) # assuming you have a pandas dataframe called df

# Add a plot
doc.add_plot(df) # assuming you have a pandas dataframe called df

# Save the document
doc.save("my_report.docx")
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.
