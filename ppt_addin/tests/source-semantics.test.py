"""Synthetic OOXML regression probes for picture preview semantics.

This suite compiles the production reader without starting PowerPoint. Tests
marked FAIL identify cases that must use native rendering until their complete
drawing semantics are implemented. All media and packages are generated here.
"""
import importlib.util
import json
import os
from pathlib import Path
import struct
import subprocess
import tempfile
import zlib

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('manifest_fixture', Path(__file__).with_name('source-manifest.test.py'))
fixture = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fixture)


def synthetic_png():
    # Four different colored quadrants make rotations/crops semantically real.
    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data) & 0xffffffff)
    colors = ((255, 0, 0), (0, 255, 0), (0, 0, 255), (255, 255, 0))
    pixels = b''.join(b'\0' + b''.join(bytes(colors[(y >= 4) * 2 + (x >= 4)]) for x in range(8)) for y in range(8))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 8, 8, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(pixels)) + chunk(b'IEND', b'')


def picture(shape_id=10, xfrm='', blip_children='', geometry='', fill='', crop='', placeholder=False):
    xfrm = xfrm or '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm>'
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{shape_id}" name="Same Picture Name"/>'
            f'<p:nvPr>{"<p:ph type=\"pic\" idx=\"1\"/>" if placeholder else ""}</p:nvPr></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="red">{blip_children}</a:blip>{crop}{fill}</p:blipFill>'
            f'<p:spPr>{xfrm}{geometry}</p:spPr></p:pic>')


def main():
    failures = []
    with tempfile.TemporaryDirectory(prefix='ppt-source-semantics-') as scratch:
        scratch = Path(scratch)
        runner = scratch / 'runner.cs'
        runner.write_text('''using System; using System.Text; using System.Web.Script.Serialization;
class Runner { static void Main(string[] args) { Console.OutputEncoding=new UTF8Encoding(false);
var data=Dongsidaye.PowerPointTools.PictureSourceManifest.Read(args[0],args[1]);
var json=new JavaScriptSerializer();json.MaxJsonLength=16777216;Console.Write(json.Serialize(data)); } }''', encoding='utf-8')
        exe = scratch / 'runner.exe'
        command = [str(Path(os.environ['WINDIR']) / 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'), '/nologo', '/target:exe', '/langversion:5', '/out:' + str(exe)]
        command += ['/reference:' + name for name in ('System.dll', 'System.Core.dll', 'System.Xml.dll', 'System.Xml.Linq.dll', 'System.IO.Compression.dll', 'System.Web.Extensions.dll')]
        result = subprocess.run(command + [str(ROOT / 'PictureSourceManifest.cs'), str(runner)], capture_output=True, text=True)
        assert result.returncode == 0, result.stdout + result.stderr

        def read(xml, label, update=None):
            parts = fixture.base_parts()
            parts['ppt/media/image9.png'] = synthetic_png()
            parts['ppt/media/image8.png'] = synthetic_png()
            parts['ppt/slides/slide2.xml'] = fixture.part(xml)
            if update:
                parts.update(update)
            package = scratch / (label + '.pptx')
            fixture.write_package(package, parts)
            proc = subprocess.run([str(exe), str(package), str(scratch / label)], capture_output=True, encoding='utf-8-sig')
            assert proc.returncode == 0, proc.stderr
            return json.loads(proc.stdout)

        def test(label, work):
            try:
                work()
                print('PASS ' + label)
            except AssertionError as error:
                failures.append(label)
                print('FAIL ' + label + ': ' + str(error))

        def native_required(xml, label):
            entry = read(xml, label).get('slide:901:10', {})
            assert entry.get('display') is None, 'unsupported visual semantics are still authorized for raw-source canvas preview'

        def plain_transform():
            xfrm = '<a:xfrm rot="5400000" flipH="1"><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm>'
            data = read(picture(xfrm=xfrm), 'plain-transform')['slide:901:10']
            assert data['display'] == dict(width=100, height=50, rotation=90, flipH=True, flipV=False)
            assert data['crop'] == dict(l=0, t=0, r=1, b=1)
        test('plain rotation and mirror are faithfully described', plain_transform)

        test('grayscale/brightness require native preview', lambda: native_required(picture(blip_children='<a:grayscl/><a:lum bright="30000" contrast="10000"/>'), 'color-effect'))
        test('picture transparency requires native preview', lambda: native_required(picture(blip_children='<a:alphaModFix amt="25000"/>'), 'alpha-effect'))
        test('ellipse picture crop requires native preview', lambda: native_required(picture(geometry='<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'), 'ellipse-clip'))
        test('picture shadow and reflection require native preview', lambda: native_required(picture(geometry='<a:effectLst><a:outerShdw blurRad="63500" dist="63500" dir="0"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>'), 'shadow-effect'))
        test('tiled picture fill requires native preview', lambda: native_required(picture(fill='<a:tile tx="0" ty="0" sx="50000" sy="50000"/>'), 'tile-fill'))
        test('non-default destination fill rectangle requires native preview', lambda: native_required(picture(fill='<a:stretch><a:fillRect l="10000" t="20000" r="10000" b="20000"/></a:stretch>'), 'fill-rectangle'))
        test('picture fill with independent rotation requires native preview', lambda: native_required(picture().replace('<p:blipFill>', '<p:blipFill rotWithShape="0">'), 'independent-fill-rotation'))
        test('theme style references require native preview', lambda: native_required(picture().replace('</p:pic>', '<p:style><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef></p:style></p:pic>'), 'theme-effect'))

        def local_appearance():
            gray = read(picture(blip_children='<a:grayscl/>'), 'gray-stamp')['slide:901:10']
            alpha = read(picture(blip_children='<a:alphaModFix amt="25000"/>'), 'alpha-stamp')['slide:901:10']
            assert gray['fingerprint'] == alpha['fingerprint']
            assert gray['display'] is None and alpha['display'] is None
            assert gray['appearance'] != alpha['appearance'], 'native preview cache must invalidate between complex styles'
        test('appearance tracks style changes with unchanged source media', local_appearance)

        def inherited_appearance():
            xml = picture(placeholder=True)
            theme_a = f'<a:theme xmlns:a="{fixture.A}" name="A"><a:themeElements><a:clrScheme name="A"><a:accent1><a:srgbClr val="FF0000"/></a:accent1></a:clrScheme></a:themeElements></a:theme>'
            theme_b = theme_a.replace('FF0000', '0000FF')
            a = read(xml, 'theme-red', {'ppt/theme/theme1.xml': theme_a})['slide:901:10']
            b = read(xml, 'theme-blue', {'ppt/theme/theme1.xml': theme_b})['slide:901:10']
            assert a['fingerprint'] == b['fingerprint'] and a['display'] is None and b['display'] is None
            assert a['appearance'] != b['appearance'], 'theme change must invalidate an unchanged placeholder picture'
            altered = fixture.part(picture(91, geometry='<a:effectLst><a:softEdge rad="12700"/></a:effectLst>'), 'sldLayout')
            c = read(xml, 'layout-effect', {'ppt/slideLayouts/slideLayout7.xml': altered})['slide:901:10']
            base = read(xml, 'layout-base')['slide:901:10']
            assert c['appearance'] != base['appearance'], 'layout appearance change must invalidate an unchanged placeholder'
        test('theme and layout dependency changes invalidate native appearance', inherited_appearance)

        def group_transform():
            group = '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="30" name="group"/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="0" y="0"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></p:grpSpPr>' + picture() + '</p:grpSp>'
            display = read(group, 'parent-transform')['slide:901:10']['display']
            assert display is None or (display['rotation'] == 90 and display['width'] == 200 and display['height'] == 100), 'parent group rotation and scaling were ignored'
        test('nested picture composes group transform or falls back', group_transform)

        def placeholder():
            xml = picture(placeholder=True).replace('<a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm>', '')
            assert read(xml, 'inherited-placeholder')['slide:901:10']['display'] is None
        test('placeholder with inherited size is not guessed', placeholder)

        def crop_pair():
            xml = picture(10, crop='<a:srcRect b="50000"/>') + picture(11, crop='<a:srcRect t="50000"/>')
            data = read(xml, 'two-crops')
            assert data['slide:901:10']['fingerprint'] == data['slide:901:11']['fingerprint']
            assert data['slide:901:10']['crop'] != data['slide:901:11']['crop']
        test('same media retains distinct source windows', crop_pair)

        def multiple_masters():
            parts = fixture.base_parts()
            presentation = parts['ppt/presentation.xml'].replace('<p:sldMasterId id="1" r:id="m2"/>', '<p:sldMasterId id="1" r:id="m2"/><p:sldMasterId id="2" r:id="m1"/>')
            update = {
                'ppt/presentation.xml': presentation,
                'ppt/_rels/presentation.xml.rels': fixture.rels(('first', 'slides/slide1.xml', 'slide'), ('second', 'slides/slide2.xml', 'slide'), ('m2', 'slideMasters/slideMaster2.xml', 'slideMaster'), ('m1', 'slideMasters/slideMaster1.xml', 'slideMaster')),
                'ppt/slideMasters/slideMaster1.xml': fixture.part(picture(90), 'sldMaster'),
                'ppt/slideMasters/_rels/slideMaster1.xml.rels': fixture.rels(('red', '../media/image1.png', 'image')),
            }
            data = read(picture(), 'two-masters', update)
            assert data['master:1:90']['fingerprint'] != data['master:2:90']['fingerprint']
        test('same shape ID and name across different masters stays isolated', multiple_masters)

    print('RESULT: ' + str(len(failures)) + ' unprotected semantic cases')
    raise SystemExit(1 if failures else 0)


if __name__ == '__main__':
    main()
