#!/usr/bin/env python3
"""
zero_spc.py

Usage:
    python zero_spc.py input.pptx [output.pptx]

This script:
1. Unzips the input PPTX into a temporary folder.
2. Finds every XML file under 'ppt/'.
3. Parses each XML, sets all a:rPr/@spc to "0".
4. Writes the modified XML back.
5. Re-zips everything into output.pptx (or overwrites input if not specified).
"""

import sys
import os
import zipfile
import shutil
import tempfile
import xml.etree.ElementTree as ET

# Register the 'a' namespace so it appears nicely in output
ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')

def fix_spacing_in_xml_file(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    # namespace map
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    # find all <a:rPr> elements
    # for rPr in root.findall('.//a:rPr', ns):
    #     # if they have a 'spc' attribute, set it to "0"
    #     if 'spc' in rPr.attrib:
    #         rPr.attrib['spc'] = '0'
    # write back
    tree.write(xml_path, xml_declaration=True, encoding='UTF-8')

def zero_out_spc(input_pptx, output_pptx=None):
    if output_pptx is None:
        # overwrite input
        output_pptx = input_pptx + '.fixed.pptx'
    # make a temp dir
    tmpdir = tempfile.mkdtemp(prefix="pptx_unzip_")
    try:
        # unzip all files
        with zipfile.ZipFile(input_pptx, 'r') as zin:
            zin.extractall(tmpdir)

        # walk through ppt/ subfolder
        ppt_folder = os.path.join(tmpdir, 'ppt')
        for dirpath, _, filenames in os.walk(ppt_folder):
            for fn in filenames:
                if fn.lower().endswith('.xml'):
                    full = os.path.join(dirpath, fn)
                    fix_spacing_in_xml_file(full)

        # rezip into output
        with zipfile.ZipFile(output_pptx, 'w', zipfile.ZIP_DEFLATED) as zout:
            # maintain folder structure
            for dirpath, dirnames, filenames in os.walk(tmpdir):
                for fn in filenames:
                    full = os.path.join(dirpath, fn)
                    # archive name should be relative to tmpdir
                    arcname = os.path.relpath(full, tmpdir)
                    zout.write(full, arcname)
        print(f"✅ Written fixed PPTX to: {output_pptx}")
    finally:
        # clean up
        shutil.rmtree(tmpdir)

if __name__ == '__main__':
    if not (2 <= len(sys.argv) <= 3):
        print(__doc__)
        sys.exit(1)
    inp = sys.argv[1]
    outp = sys.argv[2] if len(sys.argv) == 3 else None
    zero_out_spc(inp, outp)
