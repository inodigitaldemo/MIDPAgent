"""
md_to_docx.console_app
~~~~~~~~~~~~~~~~~~~~~~
Example console application: converts a ``.md`` file to DOCX via HTML.

Usage
-----
    python -m md_to_docx.console_app <input.md> [--output-dir DIR] [--css PATH]

Examples
--------
    # Convert README.md in the current directory:
    python -m md_to_docx.console_app README.md

    # Specify a custom output directory:
    python -m md_to_docx.console_app docs/spec.md --output-dir dist/

    # Use a custom CSS stylesheet:
    python -m md_to_docx.console_app report.md --css custom/styles.css
"""

import argparse
import sys
from pathlib import Path

from md_to_docx.converter import DEFAULT_CSS_PATH, convert_md_to_docx


def build_parser() -> argparse.ArgumentParser:
    """Return the argument parser for the console application."""
    parser = argparse.ArgumentParser(
        prog="md_to_docx",
        description=(
            "Convert a Markdown (.md) file to DOCX via an intermediate HTML file."
        ),
    )
    parser.add_argument(
        "input",
        metavar="INPUT.md",
        type=Path,
        help="Path to the source Markdown file.",
    )
    parser.add_argument(
        "--output-dir",
        metavar="DIR",
        type=Path,
        default=None,
        help=(
            "Directory in which to write the generated .html and .docx files. "
            "Defaults to the same directory as INPUT.md."
        ),
    )
    parser.add_argument(
        "--css",
        metavar="PATH",
        type=Path,
        default=DEFAULT_CSS_PATH,
        help=(
            f"Path to a CSS stylesheet to embed in the HTML output. "
            f"Defaults to '{DEFAULT_CSS_PATH}'."
        ),
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """Entry point for the console application.

    Parameters
    ----------
    argv:
        Argument list (defaults to ``sys.argv[1:]`` when ``None``).

    Returns
    -------
    int
        Exit code (0 on success, non-zero on failure).
    """
    parser = build_parser()
    args = parser.parse_args(argv)

    md_path: Path = args.input
    if not md_path.exists():
        print(f"ERROR: Input file not found: {md_path}", file=sys.stderr)
        return 1
    if md_path.suffix.lower() != ".md":
        print(
            f"ERROR: Input file must have a .md extension, got: '{md_path.suffix}' ({md_path})",
            file=sys.stderr,
        )
        return 1

    css_path: Path | None = args.css
    if css_path is not None and not css_path.exists():
        print(f"WARNING: CSS file not found, continuing without styling: {css_path}", file=sys.stderr)
        css_path = None

    print(f"Input:      {md_path}")
    print(f"CSS:        {css_path or '(none)'}")
    print(f"Output dir: {args.output_dir or md_path.parent}")
    print()

    try:
        html_path, docx_path = convert_md_to_docx(
            md_path=md_path,
            output_dir=args.output_dir,
            css_path=css_path,
        )
    except Exception as exc:
        print(f"ERROR: Conversion failed – {exc}", file=sys.stderr)
        return 1

    print(f"HTML generated: {html_path}")
    print(f"DOCX generated: {docx_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
