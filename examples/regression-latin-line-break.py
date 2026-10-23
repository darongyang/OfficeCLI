#!/usr/bin/env python3
"""Verify latinLnBrk without prescribing a particular word, font, or font size."""
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
    r = subprocess.run([cli, *map(str, args)], capture_output=True, text=True,
                       env=dict(os.environ, OFFICECLI_SKIP_UPDATE='1', OFFICECLI_NO_AUTO_RESIDENT='1'))
    if r.returncode:
        raise RuntimeError(r.stdout + r.stderr)
    return r.stdout


def shape(identifier, paragraph='', levels='', placeholder=False, text='UnbreakableToken'):
    ph = '<p:ph type="body" idx="1"/>' if placeholder else ''
    return ET.fromstring(f'''<p:sp xmlns:p="{P}" xmlns:a="{A}">
      <p:nvSpPr><p:cNvPr id="{identifier}" name="case-{identifier}"/><p:cNvSpPr txBox="1"/><p:nvPr>{ph}</p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="{identifier*700000}" y="1000000"/><a:ext cx="600000" cy="3000000"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
      <p:txBody><a:bodyPr wrap="square"/><a:lstStyle>{levels}</a:lstStyle>
      <a:p><a:pPr {paragraph}/><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>{text}</a:t></a:r></a:p></p:txBody>
    </p:sp>''')


def verify(cli, directory):
    base = directory / 'base.pptx'
    run(cli, 'create', base)
    run(cli, 'add', base, '/', '--type', 'slide')
    with zipfile.ZipFile(base) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    root = ET.fromstring(parts['ppt/slides/slide1.xml'])
    tree = root.find(f'{{{P}}}cSld/{{{P}}}spTree')
    for node in list(tree):
        if node.tag not in (f'{{{P}}}nvGrpSpPr', f'{{{P}}}grpSpPr'):
            tree.remove(node)
    tree.extend([
        shape(2, 'latinLnBrk="0"'),
        shape(3, 'latinLnBrk="1"'),
        shape(4, levels='<a:lvl1pPr latinLnBrk="0"/>'),
        shape(5, 'latinLnBrk="1"', '<a:lvl1pPr latinLnBrk="0"/>'),
        shape(6, 'latinLnBrk="0"', text='first second third'),
        shape(7),
        shape(8, placeholder=True),
        shape(9, 'latinLnBrk="0"', placeholder=True),
        shape(10, levels='<a:defPPr latinLnBrk="1"/>'),
        shape(11, levels='<a:defPPr latinLnBrk="1"/><a:lvl1pPr latinLnBrk="0"/>'),
    ])
    master = ET.fromstring(parts['ppt/slideMasters/slideMaster1.xml'])
    styles = master.find(f'{{{P}}}txStyles')
    if styles is None:
        styles = ET.SubElement(master, f'{{{P}}}txStyles')
    body = styles.find(f'{{{P}}}bodyStyle')
    if body is None:
        body = ET.SubElement(styles, f'{{{P}}}bodyStyle')
    level = body.find(f'{{{A}}}lvl1pPr')
    if level is None:
        level = ET.SubElement(body, f'{{{A}}}lvl1pPr')
    level.set('latinLnBrk', '1')
    parts['ppt/slideMasters/slideMaster1.xml'] = ET.tostring(master, encoding='utf-8', xml_declaration=True)
    parts['ppt/slides/slide1.xml'] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    fixture = directory / 'latin-line-break.pptx'
    with zipfile.ZipFile(fixture, 'w', zipfile.ZIP_DEFLATED) as archive:
        for n, b in parts.items():
            archive.writestr(n, b)
    html = run(cli, 'view', fixture, 'html')
    (directory / 'latin-line-break.html').write_text(html)
    def paragraph(identifier):
        start = html.index(f'data-path="/slide[1]/shape[@id={identifier}]"')
        return re.search(r'<div class="para" style="([^"]+)"', html[start:])[1]
    checks = {
        'explicit false prevents intra-word breaks': 'overflow-wrap:normal' in paragraph(2),
        'explicit true allows intra-word breaks': 'overflow-wrap:break-word' in paragraph(3),
        'shape list style inherited': 'overflow-wrap:normal' in paragraph(4),
        'paragraph overrides its list style': 'overflow-wrap:break-word' in paragraph(5),
        'ordinary space wrapping stays enabled': 'overflow-wrap:normal' in paragraph(6) and 'nowrap' not in paragraph(6),
        'Office default is false': 'overflow-wrap:normal' in paragraph(7),
        'master body style inherited': 'overflow-wrap:break-word' in paragraph(8),
        'paragraph overrides master style': 'overflow-wrap:normal' in paragraph(9),
        'shape default paragraph inherited': 'overflow-wrap:break-word' in paragraph(10),
        'level overrides default paragraph': 'overflow-wrap:normal' in paragraph(11),
        'font size unchanged': all('font-size:24pt' in paragraph(i) for i in range(2, 10)),
    }
    # A non-placeholder text box still inherits the presentation's text defaults.
    presentation = ET.fromstring(parts['ppt/presentation.xml'])
    defaults = presentation.find(f'{{{P}}}defaultTextStyle')
    if defaults is None:
        defaults = ET.SubElement(presentation, f'{{{P}}}defaultTextStyle')
    level = defaults.find(f'{{{A}}}lvl1pPr')
    if level is None:
        level = ET.SubElement(defaults, f'{{{A}}}lvl1pPr')
    level.set('latinLnBrk', '1')
    other = styles.find(f'{{{P}}}otherStyle')
    if other is not None:
        styles.remove(other)
    parts['ppt/presentation.xml'] = ET.tostring(presentation, encoding='utf-8', xml_declaration=True)
    parts['ppt/slideMasters/slideMaster1.xml'] = ET.tostring(master, encoding='utf-8', xml_declaration=True)
    global_fixture = directory / 'global-line-break.pptx'
    with zipfile.ZipFile(global_fixture, 'w', zipfile.ZIP_DEFLATED) as archive:
        for n, b in parts.items():
            archive.writestr(n, b)
    html = run(cli, 'view', global_fixture, 'html')
    (directory / 'global-line-break.html').write_text(html)
    checks['presentation default inherited by plain text box'] = 'overflow-wrap:break-word' in paragraph(7)
    checks['local false overrides presentation default'] = 'overflow-wrap:normal' in paragraph(2)
    for name, ok in checks.items():
        print(('PASS' if ok else 'FAIL') + ': ' + name)
    assert all(checks.values()), 'Latin line-breaking regression'


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--officecli', default='officecli')
    parser.add_argument('--output-dir', type=pathlib.Path)
    args = parser.parse_args()
    if args.output_dir:
        args.output_dir.mkdir(parents=True, exist_ok=True)
        verify(args.officecli, args.output_dir)
    else:
        with tempfile.TemporaryDirectory(prefix='officecli-latin-break-') as directory:
            verify(args.officecli, pathlib.Path(directory))
