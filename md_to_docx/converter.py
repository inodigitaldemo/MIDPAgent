"""
md_to_docx.converter
~~~~~~~~~~~~~~~~~~~~
Core conversion helpers that implement the Markdown → HTML → DOCX pipeline.

All functions accept :class:`pathlib.Path` objects (or strings) and return
the :class:`pathlib.Path` of the newly created output file so callers can
chain the steps easily (DRY principle – each step is defined once).
"""

from pathlib import Path

import pypandoc


# ---------------------------------------------------------------------------
# Default CSS path (DocTemplates/CSS/style.css, relative to the repo root)
# ---------------------------------------------------------------------------

_REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CSS_PATH: Path = _REPO_ROOT / "DocTemplates" / "CSS" / "style.css"


# ---------------------------------------------------------------------------
# Individual conversion steps
# ---------------------------------------------------------------------------

def md_to_html(
    md_path: Path | str,
    html_path: Path | str | None = None,
    css_path: Path | str | None = DEFAULT_CSS_PATH,
) -> Path:
    """Convert a Markdown file to a self-contained HTML file.

    Parameters
    ----------
    md_path:
        Path to the source ``.md`` file.
    html_path:
        Destination path for the generated ``.html`` file.  Defaults to the
        same directory and stem as *md_path* with a ``.html`` extension.
    css_path:
        Path to a CSS stylesheet to embed in the HTML.  Defaults to
        ``DocTemplates/CSS/style.css``.  Pass ``None`` to omit styling.

    Returns
    -------
    :class:`pathlib.Path`
        The path of the generated HTML file.
    """
    md_path = Path(md_path)

    if html_path is None:
        html_path = md_path.with_suffix(".html")
    html_path = Path(html_path)

    extra_args = ["--standalone"]
    if css_path is not None:
        css_path = Path(css_path)
        if css_path.exists():
            extra_args.extend(["--css", str(css_path)])

    pypandoc.convert_file(
        str(md_path),
        "html",
        outputfile=str(html_path),
        extra_args=extra_args,
    )
    return html_path


def html_to_docx(
    html_path: Path | str,
    docx_path: Path | str | None = None,
) -> Path:
    """Convert an HTML file to a DOCX file.

    Parameters
    ----------
    html_path:
        Path to the source ``.html`` file.
    docx_path:
        Destination path for the generated ``.docx`` file.  Defaults to the
        same directory and stem as *html_path* with a ``.docx`` extension.

    Returns
    -------
    :class:`pathlib.Path`
        The path of the generated DOCX file.
    """
    html_path = Path(html_path)

    if docx_path is None:
        docx_path = html_path.with_suffix(".docx")
    docx_path = Path(docx_path)

    pypandoc.convert_file(str(html_path), "docx", outputfile=str(docx_path))
    return docx_path


# ---------------------------------------------------------------------------
# Convenience wrapper
# ---------------------------------------------------------------------------

def convert_md_to_docx(
    md_path: Path | str,
    output_dir: Path | str | None = None,
    css_path: Path | str | None = DEFAULT_CSS_PATH,
) -> tuple[Path, Path]:
    """Convert a Markdown file to DOCX via an intermediate HTML file.

    The two-step pipeline is:
    1. ``.md``  →  ``.html``  (via :func:`md_to_html`)
    2. ``.html`` → ``.docx``  (via :func:`html_to_docx`)

    Parameters
    ----------
    md_path:
        Path to the source ``.md`` file.
    output_dir:
        Directory in which to write the ``.html`` and ``.docx`` files.
        Defaults to the same directory as *md_path*.
    css_path:
        CSS stylesheet for the HTML step.  See :func:`md_to_html`.

    Returns
    -------
    tuple[Path, Path]
        ``(html_path, docx_path)`` – the paths of the generated files.
    """
    md_path = Path(md_path)

    if output_dir is None:
        output_dir = md_path.parent
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    html_path = output_dir / md_path.with_suffix(".html").name
    docx_path = output_dir / md_path.with_suffix(".docx").name

    md_to_html(md_path, html_path, css_path)
    html_to_docx(html_path, docx_path)

    return html_path, docx_path
