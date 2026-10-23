#!/usr/bin/env python3
"""Generate minimal brace shapes and verify their HTML open paths and transforms."""
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
    result = subprocess.run([cli, *map(str, args)], capture_output=True, text=True,
                            env=dict(os.environ, OFFICECLI_SKIP_UPDATE='1', OFFICECLI_NO_AUTO_RESIDENT='1'))
    if result.returncode:
        raise RuntimeError(result.stdout + result.stderr)
    return result.stdout


def shape(identifier, preset, transform='', adjustments='', outline=None):
    if outline is None:
        outline = '<a:ln w="19050"><a:solidFill><a:srgbClr val="264B8C"/></a:solidFill></a:ln>'
    return ET.fromstring(f'''<p:sp xmlns:p="{P}" xmlns:a="{A}">
      <p:nvSpPr><p:cNvPr id="{identifier}" name="case-{identifier}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm {transform}><a:off x="{identifier*800000}" y="1000000"/><a:ext cx="500000" cy="2000000"/></a:xfrm>
        <a:prstGeom prst="{preset}"><a:avLst>{adjustments}</a:avLst></a:prstGeom><a:noFill/>{outline}</p:spPr>
    </p:sp>''')


def verify(cli, directory):
    base = directory / 'base.pptx'
    run(cli, 'create', base)
    run(cli, 'add', base, '/', '--type', 'slide')
    with zipfile.ZipFile(base) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    root = ET.fromstring(parts['ppt/slides/slide1.xml'])
    tree = root.find(f'{{{P}}}cSld/{{{P}}}spTree')
    for node in list(tree):
        if node.tag not in (f'{{{P}}}nvGrpSpPr', f'{{{P}}}grpSpPr'):
            tree.remove(node)
    tree.extend([
        shape(2, 'leftBrace'),
        shape(3, 'rightBrace'),
        shape(4, 'leftBrace', 'rot="16200000"'),
        shape(5, 'rightBrace', 'flipH="1"'),
        shape(6, 'leftBrace', adjustments='<a:gd name="adj2" fmla="val 30000"/><a:gd name="adj1" fmla="val 16000"/>'),
        shape(7, 'leftBrace', outline='<a:ln><a:noFill/></a:ln>'),
        shape(8, 'rect'),
        shape(9, 'leftBrace', outline='<a:ln w="19050"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></a:ln>'),
    ])
    group = ET.fromstring(f'''<p:grpSp xmlns:p="{P}" xmlns:a="{A}">
      <p:nvGrpSpPr><p:cNvPr id="10" name="group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="1000000" y="3500000"/><a:ext cx="3000000" cy="2000000"/>
      <a:chOff x="0" y="0"/><a:chExt cx="3000000" cy="2000000"/></a:xfrm></p:grpSpPr>
    </p:grpSp>''')
    group.append(shape(11, 'rightBrace'))
    tree.append(group)
    parts['ppt/slides/slide1.xml'] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    fixture = directory / 'brace-geometry.pptx'
    with zipfile.ZipFile(fixture, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, contents in parts.items():
            archive.writestr(name, contents)
    html = run(cli, 'view', fixture, 'html')
    (directory / 'brace-geometry.html').write_text(html)
    def block(identifier):
        marker = f'data-path="/slide[1]/shape[@id={identifier}]"'
        start = html.index(marker)
        end = html.find('data-path=', start + len(marker))
        return html[start:end if end >= 0 else len(html)]
    def stroke_path(identifier):
        paths = re.findall(r'<path d="([^"]+)"[^>]*stroke=', block(identifier))
        return paths[0] if paths else ''
    checks = {
        'left brace is an open arc path': stroke_path(2).count(' A ') == 4 and not stroke_path(2).endswith('Z'),
        'right brace is an open arc path': stroke_path(3).count(' A ') == 4 and not stroke_path(3).endswith('Z'),
        'left/right paths differ': bool(stroke_path(2)) and stroke_path(2) != stroke_path(3),
        'rotation retained': 'rotate(270deg)' in block(4) and bool(stroke_path(4)),
        'flip retained': 'scaleX(-1)' in block(5) and bool(stroke_path(5)),
        'named adjustments change geometry': bool(stroke_path(6)) and stroke_path(6) != stroke_path(2),
        'explicit noFill has no stroked path': 'stroke=' not in block(7),
        'ordinary rectangle unchanged': 'border:1.5pt solid #264B8C' in block(8),
        'gradient outline uses the brace path': 'stroke="url(#' in block(9) and '<linearGradient' in block(9),
        'grouped brace uses SVG': bool(re.search(r'data-path="/slide\[1\]/group\[@id=10\]".*?<path d="[^"]+"[^>]*stroke=', html, re.S)),
    }
    for name, passed in checks.items():
        print(('PASS' if passed else 'FAIL') + ': ' + name)
    assert all(checks.values()), 'Brace geometry regression'


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--officecli', default='officecli')
    parser.add_argument('--output-dir', type=pathlib.Path)
    args = parser.parse_args()
    if args.output_dir:
        args.output_dir.mkdir(parents=True, exist_ok=True)
        verify(args.officecli, args.output_dir)
    else:
        with tempfile.TemporaryDirectory(prefix='officecli-brace-') as directory:
            verify(args.officecli, pathlib.Path(directory))
