#!/usr/bin/env python3
"""Reproduce PPTX outlines losing their theme color after a local width override."""
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


def shape(identifier, outline='', preset='rect'):
    height = 0 if preset == 'line' else 700000
    return ET.fromstring(f'''<p:sp xmlns:p="{P}" xmlns:a="{A}">
      <p:nvSpPr><p:cNvPr id="{identifier}" name="case-{identifier}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="1000000" y="{identifier*800000}"/><a:ext cx="2000000" cy="{height}"/></a:xfrm>
      <a:prstGeom prst="{preset}"><a:avLst/></a:prstGeom><a:noFill/>{outline}</p:spPr>
      <p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>
      <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
      <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
      <a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>
    </p:sp>''')


def verify(cli, directory):
    base = directory / 'base.pptx'
    run(cli, 'create', base)
    run(cli, 'add', base, '/', '--type', 'slide')
    with zipfile.ZipFile(base) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    theme = ET.fromstring(parts['ppt/theme/theme1.xml'])
    accent = theme.find(f'.//{{{A}}}clrScheme/{{{A}}}accent1')[0].get('val')
    root = ET.fromstring(parts['ppt/slides/slide1.xml'])
    tree = root.find(f'{{{P}}}cSld/{{{P}}}spTree')
    for node in list(tree):
        if node.tag not in (f'{{{P}}}nvGrpSpPr', f'{{{P}}}grpSpPr'):
            tree.remove(node)
    tree.extend([
        shape(2, '<a:ln w="19050"/>'),
        shape(3, '<a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>'),
        shape(4, '<a:ln w="19050"><a:noFill/></a:ln>'),
        shape(5),
        shape(6, '<a:ln w="19050"/>', 'line'),
        shape(7, '<a:ln><a:prstDash val="dash"/></a:ln>'),
    ])
    parts['ppt/slides/slide1.xml'] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    fixture = directory / 'outline-inheritance.pptx'
    with zipfile.ZipFile(fixture, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, contents in parts.items():
            archive.writestr(name, contents)
    html = run(cli, 'view', fixture, 'html')
    (directory / 'outline-inheritance.html').write_text(html)
    def block(identifier):
        marker = f'data-path="/slide[1]/shape[@id={identifier}]"'
        start = html.index(marker)
        end = html.find('data-path=', start + len(marker))
        return html[start:end if end >= 0 else len(html)]
    checks = {
        'local width inherits theme color': f'border:1.5pt solid #{accent}' in block(2),
        'explicit local color wins': 'border:1.5pt solid #FF0000' in block(3),
        'explicit noFill stays invisible': 'border:' not in block(4),
        'theme-only outline unchanged': f'solid #{accent}' in block(5),
        'line shape inherits theme color': f'stroke="#{accent}"' in block(6),
        'dash override inherits a visible theme stroke': f'stroke="#{accent}"' in block(7) and 'stroke-dasharray=' in block(7),
    }
    for name, passed in checks.items():
        print(('PASS' if passed else 'FAIL') + ': ' + name)
    assert all(checks.values()), 'Theme outline inheritance regression'


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--officecli', default='officecli')
    parser.add_argument('--output-dir', type=pathlib.Path)
    args = parser.parse_args()
    if args.output_dir:
        args.output_dir.mkdir(parents=True, exist_ok=True)
        verify(args.officecli, args.output_dir)
    else:
        with tempfile.TemporaryDirectory(prefix='officecli-outline-') as directory:
            verify(args.officecli, pathlib.Path(directory))
