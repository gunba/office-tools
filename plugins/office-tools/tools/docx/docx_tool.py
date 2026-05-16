#!/usr/bin/env python3
"""
docx_tool.py - Read and edit Word documents.

Reading uses python-docx (pure Python, safe).
Editing uses Word COM automation (preserves all formatting, styles, headers/footers, etc.).

Usage:
    python tools/docx_tool.py read <file>
    python tools/docx_tool.py replace <file> --find TEXT --replace TEXT [--output FILE]
    python tools/docx_tool.py replace <file> --replacements '[{"find":"x","replace":"y"}, ...]' [--output FILE]
"""

# === office-tools bootstrap ============================================
import os
import sys

_PLUGIN_ROOT = os.environ.get("CLAUDE_PLUGIN_ROOT") or os.environ.get("PLUGIN_ROOT")
if _PLUGIN_ROOT and _PLUGIN_ROOT not in sys.path:
    sys.path.insert(0, _PLUGIN_ROOT)
else:
    sys.path.insert(
        0,
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
    )

from bootstrap.ensure_python_deps import run as _ensure_python_deps  # noqa: E402

_ensure_python_deps()
# === end bootstrap =====================================================

import argparse
import json

import docx
from docx.text.paragraph import Paragraph
from docx.table import Table


def cmd_read(args):
    doc = docx.Document(args.file)
    output = []

    # Iterate body elements in document order
    for child in doc.element.body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'p':
            para = Paragraph(child, doc)
            style_name = para.style.name if para.style else "Normal"
            text = para.text

            if not text.strip():
                output.append("")
                continue

            if style_name.startswith("Heading"):
                try:
                    level = int(style_name.split()[-1])
                except ValueError:
                    level = 1
                output.append(f"{'#' * level} {text}")
            elif "List" in style_name or "Bullet" in style_name:
                output.append(f"- {text}")
            elif "Number" in style_name:
                output.append(f"1. {text}")
            else:
                output.append(text)

        elif tag == 'tbl':
            table = Table(child, doc)
            rows = []
            for row in table.rows:
                cells = []
                for cell in row.cells:
                    cells.append(cell.text.strip().replace('\n', ' '))
                rows.append(cells)

            if rows:
                max_cols = max(len(r) for r in rows)
                for r in rows:
                    while len(r) < max_cols:
                        r.append("")

                output.append("")
                output.append("| " + " | ".join(rows[0]) + " |")
                output.append("| " + " | ".join(["---"] * max_cols) + " |")
                for row_data in rows[1:]:
                    output.append("| " + " | ".join(row_data) + " |")
                output.append("")

    # Clean up excessive blank lines
    result = "\n".join(output)
    while "\n\n\n" in result:
        result = result.replace("\n\n\n", "\n\n")

    print(result.strip())


def cmd_replace(args):
    """Find and replace text using Word COM automation to preserve formatting."""
    import win32com.client

    filepath = os.path.abspath(args.file)
    output_path = os.path.abspath(args.output) if args.output else None

    if not os.path.exists(filepath):
        print(f"Error: File not found: {filepath}")
        sys.exit(1)

    # Build list of replacements
    replacements = []
    if args.replacements:
        replacements = json.loads(args.replacements)
    elif args.find and args.replace is not None:
        replacements = [{"find": args.find, "replace": args.replace}]
    else:
        print("Error: Provide --find/--replace or --replacements JSON")
        sys.exit(1)

    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        doc = word.Documents.Open(filepath)

        for rep in replacements:
            count = 0
            while True:
                rng = doc.Content
                find_obj = rng.Find
                find_obj.ClearFormatting()
                find_obj.Text = rep["find"]
                find_obj.Forward = True
                find_obj.Wrap = 0        # wdFindStop
                find_obj.MatchCase = False
                find_obj.MatchWholeWord = False
                find_obj.MatchWildcards = False
                if find_obj.Execute():
                    rng.Text = rep["replace"]
                    count += 1
                else:
                    break
            print(f"  '{rep['find']}' -> '{rep['replace']}' ({count} replacement{'s' if count != 1 else ''})")

        if output_path:
            doc.SaveAs2(output_path, FileFormat=12)  # 12 = docx
            print(f"\nSaved to: {output_path}")
        else:
            doc.Save()
            print(f"\nSaved: {filepath}")

        doc.Close(False)

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass


def main():
    parser = argparse.ArgumentParser(description="Read and edit Word documents")
    sub = parser.add_subparsers(dest="command")

    # read
    p_read = sub.add_parser("read", help="Read document as markdown")
    p_read.add_argument("file")

    # replace
    p_rep = sub.add_parser("replace", help="Find and replace text (Word COM)")
    p_rep.add_argument("file")
    p_rep.add_argument("--find", "-f", help="Text to find")
    p_rep.add_argument("--replace", "-r", help="Replacement text")
    p_rep.add_argument("--replacements", help='Batch JSON: [{"find":"x","replace":"y"}, ...]')
    p_rep.add_argument("--output", "-o", help="Save to new file instead of editing in-place")

    args = parser.parse_args()

    commands = {
        "read": cmd_read,
        "replace": cmd_replace,
    }

    if args.command in commands:
        commands[args.command](args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
