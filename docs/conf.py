# -*- coding: utf-8 -*-
"""Configuration file for the Sphinx documentation builder."""

# pylint: disable=invalid-name,redefined-builtin

#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

import os
import sys
import docx_report  # pylint: disable=unused-import

sys.path.insert(0, os.path.abspath(".."))

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

project = "docx-report"
copyright = "2023, Christopher Hacker"
author = "Christopher Hacker"

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = [
    "sphinx.ext.autodoc",
    "sphinx_autodoc_typehints",
    "sphinx.ext.napoleon",
    "recommonmark",
]

templates_path = ["_templates"]
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store"]


# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = "pydata_sphinx_theme"
html_static_path = ["_static"]

master_doc = "index"
