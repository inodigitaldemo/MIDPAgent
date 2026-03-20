# MIDPAgent

An AI Agent that listens to a master information delivery plan in a SharePoint list and works on tasks in it.

This repository also includes a Markdown → HTML → DOCX conversion toolchain with branded Word headers/footers.

## Configuration

This project reads settings from `config.json` (not environment variables).

1. Copy `config.example.json` to `config.json`.
2. Fill in:
   - `sharepoint.site_url`
   - `sharepoint.list_name`
   - `sharepoint.client_id`
   - `sharepoint.client_secret`
   - `azure.ai_project_endpoint`

SharePoint authentication uses **app credentials only** (`client_id` + `client_secret`).

## Markdown to DOCX Tooling

### What it does

The conversion flow is:

1. Markdown (`.md`) → standalone HTML (`.html`) using Pandoc
2. HTML → Word (`.docx`) using Pandoc
3. Post-process DOCX to add **real Word header/footer content** (not just body HTML), including page numbers

### Key files

- `md_to_docx/console_app.py`: command-line entry point
- `md_to_docx/converter.py`: conversion and DOCX branding logic
- `DocTemplates/HTML/template.html`: HTML template used during Markdown → HTML
- `DocTemplates/CSS/style.css`: shared CSS and branding color variables

### Install dependencies

Use your virtual environment and install requirements:

```powershell
pip install -r requirements.txt
```

`pypandoc_binary` is used so Pandoc is bundled with the Python package.

### Usage

Run the converter as a module:

```powershell
python -m md_to_docx.console_app <input.md> [--output-dir DIR] [--css PATH]
```

Examples:

```powershell
# Convert a markdown file in place
python -m md_to_docx.console_app MDFiles/ExampleDocument.md

# Write outputs to a custom folder
python -m md_to_docx.console_app docs/spec.md --output-dir dist/

# Use a custom stylesheet
python -m md_to_docx.console_app report.md --css DocTemplates/CSS/style.css
```

### Output

For an input file named `MyDoc.md`, the tool generates:

- `MyDoc.html` (intermediate)
- `MyDoc.docx` (final output)

By default, output is written to the same folder as the input Markdown unless `--output-dir` is provided.

### Branding behavior

The DOCX branding step applies:

- A header band with title + subtitle
- A footer band with left/right text
- Dynamic page numbering (`Page X`) via native Word field codes

Brand colors are read from CSS custom properties in `DocTemplates/CSS/style.css`:

- `--document-header-bg`
- `--document-footer-bg`
- `--document-header-text`
- `--document-footer-text`

Current defaults are black backgrounds (`#000000`) with text in `#ffb81c`.
