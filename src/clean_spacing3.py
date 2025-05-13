#!/usr/bin/env python3
"""
zero_spc_regex.py

Usage:
    python zero_spc_regex.py input.pptx [output.pptx]

What it does:
1. Unzips the input PPTX (really a ZIP archive) into a temp directory.
2. Locates every XML part under 'ppt/' (slides, layouts, masters, notes, etc.),
   except any relationship folders (`_rels`) and the root [Content_Types].xml.
3. Reads each XML file as raw text and applies a single regex:
      replace any spc="..." with spc="0"
4. Writes the text back, preserving all other namespaces, comments,
   processing instructions and formatting exactly.
5. Re-zips everything into the specified output file (or `<input>.fixed.pptx`).

This avoids ElementTree’s reserialization pitfalls and keeps PowerPoint from “repairing” your file.
"""

import sys
import os
import zipfile
import tempfile
import shutil
import re

def zero_spc_in_xml_file(path):
    """Load XML as text, replace spc="…" with spc="0", write back."""
    with open(path, 'r', encoding='utf-8') as f:
        text = f.read()
    # Only replace attributes, not other text:
    new_text = re.sub(r'\bspc="[^"]*"', 'spc="0"', text)
    if new_text != text:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(new_text)

def process_pptx(input_pptx, output_pptx=None):
    if not output_pptx:
        base, ext = os.path.splitext(input_pptx)
        output_pptx = f"{base}.fixed{ext}"

    # 1) Unzip to temp
    tmpdir = tempfile.mkdtemp(prefix="pptx_edit_")
    try:
        with zipfile.ZipFile(input_pptx, 'r') as zin:
            zin.extractall(tmpdir)

        # 2) Walk the ppt/ folder
        ppt_root = os.path.join(tmpdir, 'ppt')
        for dirpath, dirnames, filenames in os.walk(ppt_root):
            # skip any _rels directories entirely
            if os.path.basename(dirpath).lower() == '_rels':
                continue

            for fn in filenames:
                if not fn.lower().endswith('.xml'):
                    continue
                # skip content types at root
                rel_path = os.path.relpath(os.path.join(dirpath, fn), tmpdir)
                if rel_path == '[Content_Types].xml':
                    continue
                # at this point, it's an XML part under ppt/ (slides, masters, etc.)
                full_path = os.path.join(dirpath, fn)
                zero_spc_in_xml_file(full_path)

        # 3) Re-zip everything (preserve folder structure)
        with zipfile.ZipFile(output_pptx, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(tmpdir):
                for file in files:
                    absfn = os.path.join(root, file)
                    arcname = os.path.relpath(absfn, tmpdir)
                    zout.write(absfn, arcname)

        print(f"✅ Done. Fixed PPTX written to: {output_pptx}")

    finally:
        shutil.rmtree(tmpdir)

if __name__ == "__main__":
    if not 2 <= len(sys.argv) <= 3:
        print(__doc__)
        sys.exit(1)
    inp = sys.argv[1]
    outp = sys.argv[2] if len(sys.argv) == 3 else None
    process_pptx(inp, outp)
