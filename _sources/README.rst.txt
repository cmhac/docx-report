.. image:: https://codecov.io/gh/christopher-hacker/docx-report/branch/main/graph/badge.svg?token=019MXVQYN5
    :target: https://codecov.io/gh/christopher-hacker/docx-report
    :align: left

.. image:: https://github.com/christopher-hacker/docx-report/actions/workflows/test.yaml/badge.svg
    :target: https://github.com/christopher-hacker/docx-report/actions/workflows/test.yaml
    :alt: tests

.. image:: https://badge.fury.io/py/docx-report.svg
    :target: https://badge.fury.io/py/docx-report
    :alt: PyPI version

docx-report
===========

This is a simple wrapper for the `python-docx`_ package that makes creating Word documents in Python easier. It contains convenience methods for creating tables and inserting data visualizations directly from Pandas dataframes, as well as simpler syntax for generally writing text to a document.

Installation
------------

.. code-block:: bash

    pip install docx-report

Usage
-----

.. code-block:: python

    from docx_report import DocxReport

    # Create a new document
    doc = DocxReport(title="My Report")

    # Add a heading
    doc.add_heading("My Heading")

    # Add a paragraph
    doc.add_paragraph("This is a paragraph.")

    # Add a table
    doc.add_table(df)  # assuming you have a pandas dataframe called df

    # Add a plot
    doc.add_plot(df)  # assuming you have a pandas dataframe called df

    # Save the document
    doc.save("my_report.docx")

Contributing
------------

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

In order to contribute, you need python 3.10+ and poetry installed. If you have access to Github Codespaces, I strongly recommend creating a codespace for this project. The codespace will automatically install the correct version of python and poetry, create your virtual environment, and will also install the pre-commit hooks.

If you don't have access to Github Codespaces, you can create a virtual environment and install the pre-commit hooks manually using poetry:

.. code-block:: bash

    poetry install
    poetry run pre-commit install

The pre-commit hooks enforce code formatting and run the tests before allowing you to commit your changes. Those same checks will also run on Github Actions when you open a pull request, and are required to pass before your pull request can be merged.
