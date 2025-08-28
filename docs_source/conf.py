from datetime import date

project = "WizardDocx"
author = "Mattia Rubino"
release = "1.0.0"
copyright = f"{date.today().year} {author}"

pypi_slug = "WizardDocx"
pypi_url  = f"https://pypi.org/project/{pypi_slug}/"

github_user = "textwizard-dev"
github_repo = "WizardDocx"
has_github  = True
github_url  = f"https://github.com/{github_user}/{github_repo}"

github_url_subst = github_url or "#"   


autodoc_mock_imports = [
    "pytesseract"
]

extensions = [
    "myst_parser",
    "sphinx.ext.autodoc",
    "sphinx.ext.napoleon",
    "sphinx_autodoc_typehints",
    "sphinx.ext.viewcode",
    "sphinx.ext.todo",
    "sphinx_copybutton",
    "sphinx_design",
    "sphinx.ext.ifconfig",
]
source_suffix = {".rst": "restructuredtext", ".md": "markdown"}
myst_enable_extensions = ["deflist","colon_fence","linkify","attrs_block","smartquotes"]
autodoc_typehints = "description"
todo_include_todos = True

html_theme = "furo"
html_static_path = ["_static"]
templates_path = ["_templates"]
html_logo = "_static/img/logo_textwizard.png"
html_favicon = "_static/img/logo_textwizard.png"
html_css_files = ["css/custom.css"]

exclude_patterns = ["_build","Thumbs.db",".DS_Store"]
master_doc = "index"
rst_prolog = f"""
.. |PYPI_URL| replace:: {pypi_url}
.. |GITHUB_URL| replace:: {github_url_subst}
"""

def setup(app):
    app.add_config_value("has_github", has_github, "env")
