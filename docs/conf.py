"""Sphinx configuration."""
project = "Kompas 3D Wrapper"
author = "Nick"
copyright = "2023, Nick"
extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.napoleon",
    "sphinx_click",
    "myst_parser",
]
autodoc_typehints = "description"
html_theme = "furo"
