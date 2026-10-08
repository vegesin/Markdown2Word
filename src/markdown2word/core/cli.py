"""Headless entry point shared by automation and the compatibility wrapper."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .pipeline import ConversionConfig, PipelineRunner


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert Markdown without starting the GUI")
    parser.add_argument("markdown", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--template", type=Path)
    parser.add_argument("--pandoc", type=Path)
    parser.add_argument("--no-formulas", action="store_true", help="Use editable native Word equations instead of MathType (no Office required)")
    parser.add_argument("--no-numbers", action="store_true")
    parser.add_argument("--number-format", choices=["Simple", "ChapterDash", "ChapterDot"], default="Simple")
    parser.add_argument("--visio", action="store_true")
    parser.add_argument("--no-auto-word", action="store_true")
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--preserve-table-style", action="store_true")
    parser.add_argument("--no-auto-captions", action="store_true")
    parser.add_argument('--reference-superscript', action='store_true')
    parser.add_argument("--timeout", type=int, default=1800)
    args = parser.parse_args()
    config = ConversionConfig(markdown=args.markdown, output=args.output,
        template=args.template, python=Path(sys.executable), pandoc=args.pandoc,
        convert_formulas=not args.no_formulas, add_numbers=not (args.no_numbers or args.no_formulas),
        number_format=args.number_format, create_visio=args.visio,
        auto_start_word=not args.no_auto_word, overwrite=args.force,
        preserve_table_style=args.preserve_table_style, auto_captions=not args.no_auto_captions,
        stage_timeout=args.timeout, reference_superscript=args.reference_superscript)
    runner = PipelineRunner()
    try:
        return runner.run(config, lambda text: print(text, end="", flush=True))
    except (ValueError, OSError, RuntimeError) as exc:
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
