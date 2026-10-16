"""Independent, authored OOXML fixtures and a headless PPTX round-trip gate.

Fixtures use no Office templates, network or installed applications.
An optional local user file is copied for testing and never modified.
Generated files and images belong in a private plan/ directory, never in git.
Render comparisons test save stability, NOT parity with Microsoft PowerPoint.
"""

from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
import posixpath
import shutil
from pathlib import Path
import struct
import subprocess
import time
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
import zipfile
import zlib

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
HEAD = ' '.join(f'xmlns:{key}="{value}"' for key, value in NS.items())
REL = NS["r"]
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def rels(entries):
    return DECL + f'<Relationships xmlns="{REL_NS}">' + ''.join(
        f'<Relationship Id="{rid}" Type="{REL}/{kind}" Target="{target}"/>'
        for rid, kind, target in entries
    ) + '</Relationships>'


def xfrm(x, y, w, h, rotation=0):
    return (f'<a:xfrm rot="{rotation}"><a:off x="{x}" y="{y}"/>'
            f'<a:ext cx="{w}" cy="{h}"/></a:xfrm>')


def run(text, size=2400, extra=''):
    return (f'<a:r><a:rPr lang="en-US" sz="{size}" {extra}>'
            '<a:solidFill><a:srgbClr val="203040"/></a:solidFill>'
            '<a:latin typeface="Inter"/></a:rPr>'
            f'<a:t>{escape(text)}</a:t></a:r>')


def body(runs):
    return ('<p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" '
            'rIns="91440" bIns="45720" anchor="t"/><a:lstStyle/><a:p>'
            f'{runs}<a:endParaRPr lang="en-US" sz="2400"/></a:p></p:txBody>')


def shape(i, name, geometry, transform, fill, text=''):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="{name}"/>'
            '<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' + transform +
            f'<a:prstGeom prst="{geometry}"><a:avLst/></a:prstGeom>{fill}'
            '<a:ln w="12700"><a:solidFill><a:srgbClr val="305070"/>'
            '</a:solidFill></a:ln></p:spPr>' + text + '</p:sp>')


def tree(shapes):
    return ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/>'
            '<p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
            '<a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/>'
            '<a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' + shapes + '</p:spTree>')


def png_pattern():
    # Tiny diagnostic pixel data, not illustrative/decorative artwork.
    def chunk(tag, data):
        return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data))
    rows = b''.join(b'\0' + b''.join(bytes((x * 64, y * 64, 80, 255)) for x in range(4)) for y in range(4))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 4, 4, 8, 6, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(rows)) + chunk(b'IEND', b'')


def fixtures(folder):
    folder.mkdir(parents=True, exist_ok=True)
    solid = '<a:solidFill><a:srgbClr val="D7E7F2"/></a:solidFill>'
    cases = []
    cases.append(('text', shape(2, 'Latin and Hangul', 'rect', xfrm(600000, 600000, 7800000, 2000000),
                               '<a:noFill/>', body(run('PPTX compatibility ') + run('굵은 한글 123', 2800, 'b="1"') + run(' italic underline', 1800, 'i="1" u="sng"')))))
    cases.append(('geometry', ''.join(shape(i + 2, name, name, xfrm(600000 + i * 2400000, 900000, 1800000, 2000000, i * 900000), solid)
                                      for i, name in enumerate(('rect', 'ellipse', 'triangle')))))
    child = shape(3, 'Grouped rectangle', 'rect', xfrm(100000, 150000, 1200000, 600000), solid, body(run('Grouped', 1600)))
    group = ('<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="Rotated group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
             '<p:grpSpPr><a:xfrm rot="900000"><a:off x="1500000" y="1000000"/><a:ext cx="5000000" cy="2800000"/>'
             '<a:chOff x="0" y="0"/><a:chExt cx="2500000" cy="1400000"/></a:xfrm></p:grpSpPr>' + child + '</p:grpSp>')
    cases.append(('group', group))
    grad = ('<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="4060A0"/></a:gs>'
            '<a:gs pos="100000"><a:srgbClr val="F0C060"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill>')
    cases.append(('gradient', shape(2, 'Gradient', 'roundRect', xfrm(700000, 700000, 6500000, 3000000), grad)))
    rows = []
    for row in (('Region', 'Amount'), ('북쪽', '120'), ('남쪽', '240')):
        cells = ''.join('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>' + run(t, 1800) + '</a:p></a:txBody><a:tcPr/></a:tc>' for t in row)
        rows.append('<a:tr h="600000">' + cells + '</a:tr>')
    table = ('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Amounts"/><p:cNvGraphicFramePr/><p:nvPr/>'
             '</p:nvGraphicFramePr><p:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="6000000" cy="1800000"/></p:xfrm>'
             '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/>'
             '<a:tblGrid><a:gridCol w="3000000"/><a:gridCol w="3000000"/></a:tblGrid>' + ''.join(rows) + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>')
    cases.append(('table', table))
    pic = ('<p:pic><p:nvPicPr><p:cNvPr id="2" name="Cropped test pixels"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
           '<p:blipFill><a:blip r:embed="rId2"/><a:srcRect l="10000" t="10000" r="10000" b="10000"/>'
           '<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>' + xfrm(1600000, 900000, 4600000, 3000000) +
           '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
    cases.append(('picture', pic))
    colors = [('dk1','000000'),('lt1','FFFFFF'),('dk2','203040'),('lt2','F0F2F4'),('accent1','305070'),('accent2','60A080'),
              ('accent3','C08060'),('accent4','8060A0'),('accent5','60A0C0'),('accent6','C0A060'),('hlink','2040C0'),('folHlink','803080')]
    fillstyles = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    linestyles = '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>' * 3
    theme = (f'<a:theme xmlns:a="{NS["a"]}" name="Authored compatibility theme"><a:themeElements><a:clrScheme name="Authored">' +
             ''.join(f'<a:{n}><a:srgbClr val="{v}"/></a:{n}>' for n,v in colors) + '</a:clrScheme><a:fontScheme name="Authored">' +
             ''.join(f'<a:{kind}><a:latin typeface="Inter"/><a:ea typeface=""/><a:cs typeface=""/></a:{kind}>' for kind in ('majorFont','minorFont')) +
             f'</a:fontScheme><a:fmtScheme name="Authored"><a:fillStyleLst>{fillstyles}</a:fillStyleLst><a:lnStyleLst>{linestyles}</a:lnStyleLst>' +
             '<a:effectStyleLst>' + '<a:effectStyle><a:effectLst/></a:effectStyle>' * 3 + f'</a:effectStyleLst><a:bgFillStyleLst>{fillstyles}</a:bgFillStyleLst></a:fmtScheme>' +
             '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>')
    cm = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    parts = {
        '_rels/.rels': rels([('rId1','officeDocument','ppt/presentation.xml')]),
        'ppt/theme/theme1.xml': DECL + theme,
        'ppt/presentation.xml': DECL + f'<p:presentation {HEAD}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>' +
            ''.join(f'<p:sldId id="{256+i}" r:id="rId{i+2}"/>' for i in range(len(cases))) + '</p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
        'ppt/_rels/presentation.xml.rels': rels([('rId1','slideMaster','slideMasters/slideMaster1.xml')] + [(f'rId{i+2}','slide',f'slides/slide{i+1}.xml') for i in range(len(cases))]),
        'ppt/slideMasters/slideMaster1.xml': DECL + f'<p:sldMaster {HEAD}><p:cSld name="Authored blank">' + tree('') + '</p:cSld>' + cm + '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>',
        'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels([('rId1','slideLayout','../slideLayouts/slideLayout1.xml'),('rId2','theme','../theme/theme1.xml')]),
        'ppt/slideLayouts/slideLayout1.xml': DECL + f'<p:sldLayout {HEAD} type="blank" preserve="1"><p:cSld name="Authored blank">' + tree('') + '</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>',
        'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels([('rId1','slideMaster','../slideMasters/slideMaster1.xml')]),
        'ppt/media/pattern.png': png_pattern(),
        'ppt/commentAuthors.xml': DECL + f'<p:cmAuthorLst {HEAD}><p:cmAuthor id="0" name="Fixture author" initials="FX" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>',
        'ppt/comments/comment1.xml': DECL + f'<p:cmLst {HEAD}><p:cm authorId="0" dt="2026-10-16T00:00:00.000" idx="1"><p:pos x="0" y="0"/><p:text>Authored compatibility comment</p:text></p:cm></p:cmLst>',
    }
    presentation_rels = ET.fromstring(parts['ppt/_rels/presentation.xml.rels'])
    ET.SubElement(presentation_rels, f'{{{REL_NS}}}Relationship', {
        'Id':'rId20', 'Type':REL+'/commentAuthors', 'Target':'commentAuthors.xml'})
    parts['ppt/_rels/presentation.xml.rels'] = ET.tostring(presentation_rels, encoding='utf-8', xml_declaration=True)
    for i, (name, shapes) in enumerate(cases, 1):
        parts[f'ppt/slides/slide{i}.xml'] = DECL + f'<p:sld {HEAD}><p:cSld name="{name}"><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>' + tree(shapes) + '</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
        entries = [('rId1','slideLayout','../slideLayouts/slideLayout1.xml')]
        if name == 'picture':
            entries.append(('rId2','image','../media/pattern.png'))
        if name == 'text':
            entries.append(('rId3','comments','../comments/comment1.xml'))
        parts[f'ppt/slides/_rels/slide{i}.xml.rels'] = rels(entries)
    overrides = [('ppt/presentation.xml','presentation.main'),('ppt/slideMasters/slideMaster1.xml','slideMaster'),('ppt/slideLayouts/slideLayout1.xml','slideLayout')]
    overrides += [(f'ppt/slides/slide{i}.xml','slide') for i in range(1,len(cases)+1)]
    overrides += [('ppt/commentAuthors.xml','commentAuthors'),('ppt/comments/comment1.xml','comments')]
    parts['[Content_Types].xml'] = DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' + ''.join(f'<Override PartName="/{n}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.{t}+xml"/>' for n,t in overrides) + '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'
    out = folder / 'authored-core.pptx'
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in parts.items():
            z.writestr(n, data)
    return out


def signature(path):
    with zipfile.ZipFile(path) as z:
        invalid_xml = []
        for name in z.namelist():
            if name.endswith(('.xml', '.rels')):
                try:
                    ET.fromstring(z.read(name))
                except ET.ParseError as error:
                    invalid_xml.append({'part':name,'error':str(error)})
        slides = sorted((n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')), key=lambda n:int(n.rsplit('slide',1)[1][:-4]))
        result = []
        for n in slides:
            root = ET.fromstring(z.read(n))
            shapes = {}
            for tag in ('sp','pic','graphicFrame','grpSp','cxnSp'):
                for s in root.findall('.//p:'+tag, NS):
                    nv = next(iter(s), None)
                    name = nv.find('p:cNvPr', NS) if nv is not None else None
                    if name is None:
                        continue
                    trans = s.find('p:spPr/a:xfrm', NS)
                    if trans is None:
                        trans = s.find('p:xfrm', NS)
                    if trans is None:
                        trans = s.find('p:grpSpPr/a:xfrm', NS)
                    geom = dict(trans.attrib) if trans is not None else {}
                    if 'rot' not in geom:
                        geom['rot'] = '0'
                    if trans is not None:
                        for c in trans:
                            geom[c.tag.split('}')[-1]] = dict(c.attrib)
                    key = name.get('name','')
                    while key in shapes:
                        key += '#duplicate'
                    shapes[key] = {'kind': tag, 'geometry':geom, 'text':[t.text or '' for t in s.findall('.//a:t',NS)]}
            result.append(shapes)
        media = sorted(hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist() if n.startswith('ppt/media/') and not n.endswith('/'))
        broken = []
        for name in z.namelist():
            if not name.endswith('.rels'):
                continue
            base = posixpath.dirname(posixpath.dirname(name))
            for rel in ET.fromstring(z.read(name)):
                if rel.get('TargetMode') == 'External':
                    continue
                target = rel.get('Target','')
                resolved = posixpath.normpath(posixpath.join(base, target)).lstrip('/')
                if resolved not in z.namelist():
                    broken.append({'part':name,'target':target})
        page = ET.fromstring(z.read('ppt/presentation.xml')).find('p:sldSz',NS)
        return {'slides':result,'media':media,'page_size':dict(page.attrib) if page is not None else {},'broken_relationships':broken,'invalid_xml':invalid_xml}


def invoke(exe, args, log):
    start = time.perf_counter()
    try:
        r = subprocess.run([str(exe), *map(str,args)], capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=180)
    except subprocess.TimeoutExpired as error:
        def decoded(value):
            return value.decode('utf-8',errors='replace') if isinstance(value,bytes) else value or ''
        log.write_text(decoded(error.stdout)+decoded(error.stderr)+'\nTimed out after 180 seconds.\n',encoding='utf-8')
        raise
    log.write_text(r.stdout + r.stderr, encoding='utf-8')
    if r.returncode:
        raise RuntimeError(f'{args[0]} failed ({r.returncode}); see {log}')
    return round((time.perf_counter()-start)*1000, 2)


def compare_pixels(a, b):
    from PIL import Image, ImageChops
    ia, ib = Image.open(a).convert('RGBA'), Image.open(b).convert('RGBA')
    if ia.size != ib.size:
        return {'same_size':False, 'before':ia.size,'after':ib.size,'pass':False}
    diff = ImageChops.difference(ia, ib)
    rgb = diff.convert('RGB').tobytes()
    count = sum(max(rgb[i:i+3]) > 8 for i in range(0, len(rgb), 3))
    return {'same_size':True,'changed_pixels_over_8':count,'pixel_count':ia.width*ia.height,'changed_fraction':count/(ia.width*ia.height), 'pass':count == 0}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--out', type=Path, required=True)
    parser.add_argument('--cli', type=Path)
    parser.add_argument('--input', type=Path, help='optional local PPTX; copied, never modified')
    args = parser.parse_args()
    out = args.out.resolve()
    private_root = Path(__file__).resolve().parent.parent/'plan'
    if not out.is_relative_to(private_root) or out == private_root:
        parser.error('output must be inside this checkout\'s ignored plan/ directory')
    if out.exists() and any(out.iterdir()):
        parser.error('output must be a new or empty directory; prior evidence is never overwritten')
    if args.input:
        if args.input.suffix.lower() != '.pptx' or not args.input.is_file():
            parser.error('--input must be an existing local PPTX')
        out.mkdir(parents=True, exist_ok=True)
        src = out/'input-copy.pptx'
        shutil.copyfile(args.input,src)
    else:
        src = fixtures(out)
    sig = signature(src)
    if not args.input:
        assert len(sig['slides']) == 6 and len(sig['media']) == 1 and not sig['broken_relationships'] and not sig['invalid_xml']
    report = {'scope':'save stability only, not PowerPoint reference parity', 'source':'local copy' if args.input else 'authored OOXML', 'input_sha256':hashlib.sha256(src.read_bytes()).hexdigest(), 'checks':[]}
    if args.cli:
        saved, again = out/'saved.pptx', out/'saved-again.pptx'
        report['timings_ms'] = {
            'open_save': invoke(args.cli, ['convert',src,saved], out/'convert.log'),
            'reopen_save': invoke(args.cli, ['convert',saved,again], out/'reopen.log'),
            'render_before': invoke(args.cli, ['render',src,'--all','--scale','1',out/'before'], out/'render-before.log'),
            'render_after': invoke(args.cli, ['render',saved,'--all','--scale','1',out/'after'], out/'render-after.log'),
        }
        saved_sig, again_sig = signature(saved), signature(again)
        for key in ('slides','media','page_size'):
            report['checks'].append({'name':key+'_retained','pass':sig[key] == saved_sig[key], 'before':sig[key], 'after':saved_sig[key]})
            report['checks'].append({'name':key+'_second_save_stable','pass':saved_sig[key] == again_sig[key]})
        report['checks'].append({'name':'xml_well_formed','pass':not sig['invalid_xml'] and not saved_sig['invalid_xml'] and not again_sig['invalid_xml'], 'before':sig['invalid_xml'], 'after':saved_sig['invalid_xml'], 'second_save':again_sig['invalid_xml']})
        report['checks'].append({'name':'source_media_bytes_retained','pass':not (Counter(sig['media'])-Counter(saved_sig['media']))})
        report['checks'].append({'name':'slide_dimensions_retained','pass':all(sig['page_size'].get(k)==saved_sig['page_size'].get(k) for k in ('cx','cy'))})
        report['checks'].append({'name':'relationships_resolve','pass':not saved_sig['broken_relationships'] and not again_sig['broken_relationships'], 'after':saved_sig['broken_relationships']})
        before = sorted((out/'before').glob('*.png'))
        after = sorted((out/'after').glob('*.png'))
        report['checks'].append({'name':'render_count','pass':len(before)==len(after)==len(sig['slides']),'before':len(before),'after':len(after)})
        for a,b in zip(before,after):
            report['checks'].append({'name':'render_'+a.name, **compare_pixels(a,b)})
        report['pass'] = all(c['pass'] for c in report['checks'])
    (out/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'directory':str(out),'pass':report.get('pass'),'checks':[{k:v for k,v in c.items() if k not in ('before','after')} for c in report['checks']]},ensure_ascii=False,indent=2))
    return 0 if report.get('pass',True) else 1


if __name__ == '__main__':
    raise SystemExit(main())
