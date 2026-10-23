#!/usr/bin/env python3
"""Verify PPTX gradient stop ordering using only Python's standard library.

Usage: python regression-gradient-stop-order.py --officecli /path/to/officecli
The input is generated from scratch; no private presentation or fonts are needed.
"""
import argparse
import os
import pathlib
import re
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'


def run(cli, *args):
    env = dict(os.environ, OFFICECLI_SKIP_UPDATE='1', OFFICECLI_NO_AUTO_RESIDENT='1')
    result = subprocess.run([cli, *map(str, args)], env=env, capture_output=True, text=True)
    if result.returncode:
        raise RuntimeError(result.stdout + result.stderr)
    return result.stdout


def gradient(stops, radial=False):
    colors = ''.join(f'<a:gs pos="{pos}"><a:srgbClr val="{color}"/></a:gs>'
                     for pos, color in stops)
    shading = '<a:path path="circle"/>' if radial else '<a:lin ang="10800000" scaled="1"/>'
    return f'<a:gradFill><a:gsLst>{colors}</a:gsLst>{shading}</a:gradFill>'


def shape(identifier, fill='', outline='', preset='rect'):
    return ET.fromstring(f'''<p:sp xmlns:p="{P}" xmlns:a="{A}">
      <p:nvSpPr><p:cNvPr id="{identifier}" name="case-{identifier}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="{identifier * 600000}" y="1000000"/><a:ext cx="500000" cy="1000000"/></a:xfrm>
        <a:prstGeom prst="{preset}"><a:avLst/></a:prstGeom>{fill}{outline}</p:spPr>
    </p:sp>''')


def verify(cli, directory):
    base = directory / 'base.pptx'
    run(cli, 'create', base)
    run(cli, 'add', base, '/', '--type', 'slide')
    unordered = [(100000, '112233'), (0, '445566'), (7000, '778899'), (90000, 'AABBCC')]
    equal = [(100000, 'FFFFFF'), (50000, 'FF0000'), (0, '000000'), (50000, '0000FF')]
    already_sorted = [(0, '123456'), (50000, '654321'), (100000, 'ABCDEF')]
    with zipfile.ZipFile(base) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    root = ET.fromstring(parts['ppt/slides/slide1.xml'])
    tree = root.find(f'{{{P}}}cSld/{{{P}}}spTree')
    for node in list(tree):
        if node.tag not in (f'{{{P}}}nvGrpSpPr', f'{{{P}}}grpSpPr'):
            tree.remove(node)
    tree.extend([
        shape(2, gradient(unordered)),
        shape(3, gradient(unordered, radial=True)),
        shape(4, '<a:noFill/>', '<a:ln w="25400">' + gradient(unordered) + '</a:ln>', 'roundRect'),
        shape(5, gradient(equal)),
        shape(6, gradient(already_sorted)),
    ])
    parts['ppt/slides/slide1.xml'] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    fixture = directory / 'gradient-stop-order.pptx'
    with zipfile.ZipFile(fixture, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, contents in parts.items():
            archive.writestr(name, contents)
    html = run(cli, 'view', fixture, 'html')
    (directory / 'gradient-stop-order.html').write_text(html)
    checks = {
        'linear fill': 'linear-gradient(270deg, #445566 0%, #778899 7%, #AABBCC 90%, #112233 100%)' in html,
        'radial fill': re.search(r'radial-gradient\([^\n]*#445566 0%, #778899 7%, #AABBCC 90%, #112233 100%', html) is not None,
        'stable equal-position stops': '#000000 0%, #FF0000 50%, #0000FF 50%, #FFFFFF 100%' in html,
        'already sorted stops unchanged': '#123456 0%, #654321 50%, #ABCDEF 100%' in html,
    }
    svg_stops = re.findall(r'<stop offset="([\d.]+)%" stop-color="(#[\dA-Fa-f]{6})"', html)
    checks['SVG gradient outline'] = svg_stops == [('0', '#445566'), ('7', '#778899'), ('90', '#AABBCC'), ('100', '#112233')]
    for name, passed in checks.items():
        print(('PASS' if passed else 'FAIL') + ': ' + name)
    assert all(checks.values()), 'Gradient stop ordering regression'


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--officecli', default='officecli')
    parser.add_argument('--output-dir', type=pathlib.Path)
    args = parser.parse_args()
    if args.output_dir:
        args.output_dir.mkdir(parents=True, exist_ok=True)
        verify(args.officecli, args.output_dir)
    else:
        with tempfile.TemporaryDirectory(prefix='officecli-gradient-') as directory:
            verify(args.officecli, pathlib.Path(directory))
