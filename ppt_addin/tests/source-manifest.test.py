"""Compile the real C#5 reader and exercise adversarial OOXML ZIP fixtures.

python ppt_addin/tests/source-manifest.test.py [--real own-local-snapshot.pptx]
No Office host is launched and no user presentation is changed.
"""
import argparse
import base64
import hashlib
import json
import os
from pathlib import Path
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
PNG = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==')


def rels(*items):
    root = ET.Element('Relationships', xmlns=REL)
    for rid, target, kind, *external in items:
        attrs = {'Id': rid, 'Target': target, 'Type': R + '/' + kind}
        if external:
            attrs['TargetMode'] = 'External'
        ET.SubElement(root, 'Relationship', attrs)
    return ET.tostring(root)


def picture(shape_id, relation, crop='', linked=False):
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{shape_id}" name="same name"/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:{"link" if linked else "embed"}="{relation}"/>{crop}</p:blipFill></p:pic>')


def part(pictures, kind='sld', extra=''):
    return (f'<p:{kind} xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">'
            f'<p:cSld><p:spTree>{pictures}</p:spTree></p:cSld>{extra}</p:{kind}>')


def base_parts():
    return {
        'ppt/presentation.xml': f'<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldMasterIdLst><p:sldMasterId id="1" r:id="m2"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="901" r:id="second"/><p:sldId id="777" r:id="first"/></p:sldIdLst></p:presentation>',
        'ppt/_rels/presentation.xml.rels': rels(('first', 'slides/slide1.xml', 'slide'), ('second', 'slides/slide2.xml', 'slide'), ('m2', 'slideMasters/slideMaster2.xml', 'slideMaster')),
        'ppt/slides/slide1.xml': part(picture(10, 'blue')),
        'ppt/slides/_rels/slide1.xml.rels': rels(('blue', '../media/image1.png', 'image')),
        'ppt/slides/slide2.xml': part(picture(10, 'red', '<a:srcRect l="10000" t="20000" r="30000" b="40000"/>') + picture(11, 'redAgain')),
        'ppt/slides/_rels/slide2.xml.rels': rels(('red', '../media/image9.png', 'image'), ('redAgain', '../media/image8.png', 'image')),
        'ppt/slideMasters/slideMaster2.xml': part(picture(90, 'red'), 'sldMaster', '<p:sldLayoutIdLst><p:sldLayoutId id="99" r:id="layout7"/></p:sldLayoutIdLst>'),
        'ppt/slideMasters/_rels/slideMaster2.xml.rels': rels(('red', '../media/image9.png', 'image'), ('layout7', '../slideLayouts/slideLayout7.xml', 'slideLayout')),
        'ppt/slideLayouts/slideLayout7.xml': part(picture(91, 'blue'), 'sldLayout'),
        'ppt/slideLayouts/_rels/slideLayout7.xml.rels': rels(('blue', '../media/image1.png', 'image')),
        'ppt/media/image9.png': PNG,
        'ppt/media/image8.png': PNG,
        'ppt/media/image1.png': PNG + b'different embedded image bytes',
    }


def write_package(path, parts):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--real', type=Path)
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix='ppt-source-manifest-tests-') as work:
        work = Path(work)
        runner = work / 'runner.cs'
        runner.write_text('''using System; using System.Text; using System.Web.Script.Serialization;
class Runner { static void Main(string[] args) { Console.OutputEncoding = new UTF8Encoding(false);
var map=Dongsidaye.PowerPointTools.PictureSourceManifest.Read(args[0],args[1]);
var json=new JavaScriptSerializer();json.MaxJsonLength=16777216;Console.Write(json.Serialize(map)); } }''', encoding='utf-8')
        csc = Path(os.environ['WINDIR']) / 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
        exe = work / 'runner.exe'
        compile_result = subprocess.run([str(csc), '/nologo', '/target:exe', '/langversion:5', '/out:' + str(exe),
            '/reference:System.dll', '/reference:System.Core.dll', '/reference:System.Xml.dll', '/reference:System.Xml.Linq.dll',
            '/reference:System.IO.Compression.dll', '/reference:System.Web.Extensions.dll', str(ROOT / 'PictureSourceManifest.cs'), str(runner)], capture_output=True, text=True)
        assert compile_result.returncode == 0, compile_result.stdout + compile_result.stderr

        def read(parts, label):
            package = work / (label + '.pptx')
            write_package(package, parts)
            output = work / label
            proc = subprocess.run([str(exe), str(package), str(output)], capture_output=True, encoding='utf-8-sig')
            assert proc.returncode == 0, proc.stderr
            result = json.loads(proc.stdout)
            for entry in result.values():
                assert Path(entry['path']).resolve().is_relative_to(output.resolve()), 'media escaped owned output directory'
                assert entry['fingerprint'] == 'sha256:' + hashlib.sha256(Path(entry['path']).read_bytes()).hexdigest()
            return result

        result = read(base_parts(), 'mapping')
        assert set(result) == {'slide:901:10', 'slide:901:11', 'slide:777:10', 'master:1:90', 'layout:1:1:91'}
        assert result['slide:901:10']['crop'] == {'l': .1, 't': .2, 'r': .7, 'b': .6}
        assert result['slide:901:10']['fingerprint'] == result['slide:901:11']['fingerprint']
        assert result['slide:901:10']['fingerprint'] != result['slide:777:10']['fingerprint']
        print('PASS exact embedded media SHA-256, SlideID mapping, master/layout order, crop, duplicate media')

        parts = base_parts()
        parts['ppt/slides/slide2.xml'] = part(picture(10, 'external') + picture(11, 'escape') + picture(12, 'encoded') + picture(13, 'red', linked=True) + picture(14, 'dup'))
        parts['ppt/slides/_rels/slide2.xml.rels'] = rels(('external', 'https://example.invalid/image.png', 'image', True), ('escape', '../../../outside.png', 'image'), ('encoded', '%2e%2e/%2e%2e/%2e%2e/outside.png', 'image'), ('red', '../media/image9.png', 'image'), ('dup', 'https://example.invalid/image.png', 'image', True), ('dup', '../media/image9.png', 'image'))
        parts['../outside.png'] = b'must not extract'
        result = read(parts, 'unsafe-relations')
        assert not any(key.startswith('slide:901:') for key in result)
        print('PASS external/linked images, duplicate relationship IDs, traversal and encoded traversal rejected')

        parts = base_parts()
        parts['ppt/slides/slide2.xml'] = part(picture(10, 'red') + picture(10, 'redAgain') + picture('010', 'red') + picture(12, 'red', '<a:srcRect l="60000" r="60000"/>'))
        result = read(parts, 'ambiguous-objects')
        assert not any(key.startswith('slide:901:') for key in result)
        parts['ppt/slides/slide2.xml'] = '<!DOCTYPE x [<!ENTITY ext SYSTEM "file:///C:/Windows/win.ini">]>' + part(picture(13, 'red'))
        assert not any(key.startswith('slide:901:') for key in read(parts, 'dtd'))
        print('PASS duplicate shape IDs, noncanonical IDs, invalid crops and DTD rejected')

        parts = base_parts()
        parts['ppt/slides/slide2.xml'] = part(picture(10, 'red').replace('<a:blip r:embed="red"/>', '<a:blip r:embed="red"><a:extLst><a:ext uri="svg"><s:svgBlip xmlns:s="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="svg"/></a:ext></a:extLst></a:blip>'))
        result = read(parts, 'alternate-svg')
        assert 'slide:901:10' not in result
        print('PASS alternate SVG payload is not identified by its PNG fallback')

        parts = base_parts()
        parts['ppt/media/image9.png'] = b'X' * (32 * 1024 * 1024 + 1)
        result = read(parts, 'oversize-media')
        assert 'slide:901:10' not in result and 'master:1:90' not in result and 'slide:901:11' in result
        print('PASS oversized media skipped while independent valid pictures remain')

        parts = base_parts()
        large = b'X' * (32 * 1024 * 1024)
        parts['ppt/slides/slide2.xml'] = part(''.join(picture(100+i, 'large'+str(i)) for i in range(9)))
        parts['ppt/slides/_rels/slide2.xml.rels'] = rels(*[('large'+str(i), '../media/large'+str(i)+'.png', 'image') for i in range(9)])
        for i in range(9):
            parts['ppt/media/large'+str(i)+'.png'] = large
        result = read(parts, 'total-budget')
        assert 0 < sum(key.startswith('slide:901:') for key in result) <= 7
        print('PASS cumulative 256 MiB expansion budget limits many individually valid media entries')

        if args.real:
            proc = subprocess.run([str(exe), str(args.real.resolve()), str(work / 'real')], capture_output=True, encoding='utf-8-sig')
            assert proc.returncode == 0, proc.stderr
            result = json.loads(proc.stdout)
            with zipfile.ZipFile(args.real) as z:
                pres = ET.fromstring(z.read('ppt/presentation.xml'))
                slide_id = list(pres.find('{' + P + '}sldIdLst'))[5].attrib['id']
            check = result[f'slide:{slide_id}:60']['fingerprint']
            floorplan = result[f'slide:{slide_id}:3']['fingerprint']
            assert check != floorplan, 'slide-6 checkmark and floor plan must not share a source identity'
            print(f'PASS local snapshot: {len(result)} mapped pictures; slide 6 shape 60 and shape 3 have distinct original-media SHA-256')


if __name__ == '__main__':
    main()
