# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

import re
from pathlib import Path


def _library_version():
    """Read the version out of EGAMMA.bas.

    EGAMMA_LIB_VERSION is what EGAMMA_VERSION() reports to a worksheet, so it
    is the one place the library states its version. Deriving the manual's
    version from it means the PDF cannot claim a release the code does not,
    which is otherwise easy to miss: the title page is the last thing anyone
    looks at. Fails loudly rather than falling back to a guess.
    """
    source = Path(__file__).resolve().parents[2] / 'EGAMMA.bas'
    text = source.read_text(encoding='utf-8', errors='replace')
    match = re.search(r'EGAMMA_LIB_VERSION\s+As\s+String\s*=\s*"([^"]+)"', text)
    if match is None:
        raise RuntimeError(f'No EGAMMA_LIB_VERSION found in {source}')
    return match.group(1)


project = 'EGAMMA-VBA'
copyright = '2023-2026, Frode Drevland'
author = 'Frode Drevland'
release = _library_version()
version = release

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
