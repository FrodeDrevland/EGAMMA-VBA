# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

project = 'EGAMMA-VBA'
copyright = '2023-2026, Frode Drevland'
author = 'Frode Drevland'
release = '1.1.0'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = []

templates_path = ['_templates']
exclude_patterns = []



# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = 'sphinx_rtd_theme'
html_static_path = ['_static']

# -- Options for LaTeX / PDF output ------------------------------------------
# The PDF is the form of the documentation that ships in the release archive,
# because the audience for an Excel add-in cannot be expected to run Sphinx.
# Naming the output file explicitly keeps the build reproducible and lets the
# release workflow pick it up without globbing.

latex_documents = [
    ('index',
     'EGAMMA-VBA-manual.tex',
     'eGamma-VBA: expanded Gamma distribution for Microsoft Excel',
     'Frode Drevland',
     'manual'),
]

latex_elements = {
    'papersize': 'a4paper',
    'pointsize': '11pt',
}
