#!/usr/bin/env python3
"""
Markdown to DOCX Converter (English / Journal-focused)

Converts markdown files to Word documents using pandoc, then optionally
applies Water Research / Elsevier journal formatting.

Usage:
    python convert_md_to_docx.py input.md output.docx
    python convert_md_to_docx.py input.md output.docx --journal
    python convert_md_to_docx.py --batch input_dir/ output_dir/
"""

import sys
import os
import argparse
import subprocess
import re
import tempfile
from pathlib import Path


def check_pandoc_installed():
    try:
        result = subprocess.run(['pandoc', '--version'],
                              capture_output=True, text=True)
        return result.returncode == 0
    except FileNotFoundError:
        return False


def preprocess_markdown(content):
    """Convert single newlines to double newlines so each line becomes a paragraph."""
    lines = content.split('\n')
    processed = []
    for i, line in enumerate(lines):
        current = line.rstrip()
        processed.append(current)
        if (i + 1 < len(lines)
            and current and lines[i + 1].rstrip()
            and not current.startswith('#')
            and not current.startswith('-')
            and not current.startswith('*')
            and not current.startswith('+')
            and not re.match(r'^\d+\.', current)
            and not current.startswith('```')
            and not current.startswith('|')):
            processed.append('')  # blank line = paragraph break
    return '\n'.join(processed)


def convert_md_to_docx(input_file, output_file, reference_doc=None,
                       toc=False, standalone=True, metadata=None,
                       journal=False):
    """Convert a markdown file to DOCX using pandoc."""
    if not check_pandoc_installed():
        print("Error: pandoc is not installed.")
        print("  macOS:   brew install pandoc")
        print("  Linux:   sudo apt-get install pandoc")
        print("  Windows: choco install pandoc")
        return False

    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found.")
        return False

    with open(input_file, 'r', encoding='utf-8') as f:
        original_content = f.read()

    processed_content = preprocess_markdown(original_content)

    temp_fd, temp_file = tempfile.mkstemp(suffix='.md', text=True)
    try:
        with os.fdopen(temp_fd, 'w', encoding='utf-8') as f:
            f.write(processed_content)

        cmd = ['pandoc', temp_file, '-o', output_file]
        if standalone:
            cmd.append('--standalone')
        if toc:
            cmd.append('--toc')
        if reference_doc and os.path.exists(reference_doc):
            cmd.extend(['--reference-doc', reference_doc])

        if metadata:
            for key in ('title', 'author', 'date'):
                if key in metadata:
                    cmd.extend(['--metadata', f'{key}={metadata[key]}'])

        print(f"Converting '{input_file}' to '{output_file}'...")
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"Error during conversion:\n{result.stderr}")
            return False

        print(f"Converted: {output_file}")

        # Apply journal formatting if requested
        if journal:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            format_script = os.path.join(script_dir, 'format_journal.py')
            if os.path.exists(format_script):
                print("Applying journal formatting...")
                fmt_result = subprocess.run(
                    [sys.executable, format_script, output_file],
                    capture_output=True, text=True
                )
                if fmt_result.returncode == 0:
                    print(fmt_result.stdout.strip())
                else:
                    print(f"Warning: Journal formatting failed:\n{fmt_result.stderr}")

        return True

    except Exception as e:
        print(f"Error: {str(e)}")
        return False
    finally:
        if os.path.exists(temp_file):
            os.unlink(temp_file)


def batch_convert(input_dir, output_dir, pattern='*.md', **kwargs):
    """Batch convert all markdown files in a directory."""
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    md_files = list(input_path.glob(pattern))
    if not md_files:
        print(f"No markdown files found in '{input_dir}'")
        return

    print(f"Found {len(md_files)} markdown file(s)")
    success = 0
    for md_file in md_files:
        docx_file = output_path / f"{md_file.stem}.docx"
        if convert_md_to_docx(str(md_file), str(docx_file), **kwargs):
            success += 1
    print(f"\n{success}/{len(md_files)} files converted successfully")


def main():
    parser = argparse.ArgumentParser(
        description='Convert Markdown to DOCX (with optional journal formatting)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python convert_md_to_docx.py paper.md paper.docx
  python convert_md_to_docx.py paper.md paper.docx --journal
  python convert_md_to_docx.py paper.md paper.docx --toc --title "My Paper"
  python convert_md_to_docx.py --batch input_dir/ output_dir/
        """
    )

    parser.add_argument('input', help='Input markdown file or directory (batch mode)')
    parser.add_argument('output', nargs='?', help='Output DOCX file or directory (batch mode)')
    parser.add_argument('--reference-doc', help='Reference DOCX file for styling')
    parser.add_argument('--toc', action='store_true', help='Generate table of contents')
    parser.add_argument('--no-standalone', action='store_true', help='Do not create standalone document')
    parser.add_argument('--journal', action='store_true',
                        help='Apply Water Research / Elsevier journal formatting after conversion')
    parser.add_argument('--title', help='Document title')
    parser.add_argument('--author', help='Document author')
    parser.add_argument('--date', help='Document date')
    parser.add_argument('--batch', action='store_true', help='Batch convert all .md files')
    parser.add_argument('--pattern', default='*.md', help='File pattern for batch mode (default: *.md)')

    args = parser.parse_args()

    metadata = {}
    for key in ('title', 'author', 'date'):
        val = getattr(args, key, None)
        if val:
            metadata[key] = val

    if args.batch:
        if not args.output:
            print("Error: Output directory required for batch mode")
            sys.exit(1)
        batch_convert(
            args.input, args.output,
            pattern=args.pattern,
            reference_doc=args.reference_doc,
            toc=args.toc,
            standalone=not args.no_standalone,
            metadata=metadata or None,
            journal=args.journal,
        )
    else:
        if not args.output:
            print("Error: Output file required")
            parser.print_help()
            sys.exit(1)
        success = convert_md_to_docx(
            args.input, args.output,
            reference_doc=args.reference_doc,
            toc=args.toc,
            standalone=not args.no_standalone,
            metadata=metadata or None,
            journal=args.journal,
        )
        sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
