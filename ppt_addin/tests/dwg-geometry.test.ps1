$ErrorActionPreference = 'Stop'
$sourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$scratch = Join-Path $env:TEMP ('ppt-dwg-geometry-tests-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch | Out-Null
$harness = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Dongsidaye.PowerPointTools
{
    internal static class GeometryRegression
    {
        private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
        private static int failures;

        private static void Check(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private static void Near(double actual, double expected, string message)
        {
            Check(Math.Abs(actual - expected) < 0.000001, message + ": " + actual + " / " + expected);
        }

        private static void Test(string name, Action action)
        {
            try { action(); Console.WriteLine("PASS " + name); }
            catch (Exception error) { failures++; Console.Error.WriteLine("FAIL " + name + " " + error.Message); }
        }

        private static void Reject(Action action, string message)
        {
            try { action(); }
            catch (InvalidOperationException) { return; }
            throw new InvalidOperationException(message);
        }

        private static DwgOcrPoint Point(double x, double y)
        {
            return new DwgOcrPoint { x = x, y = y };
        }

        private static DwgOcrPoint[] Rectangle(double left, double top, double right, double bottom)
        {
            return new[] { Point(left, top), Point(right, top), Point(right, bottom), Point(left, bottom) };
        }

        private static string Drawing(string body)
        {
            return "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 200 100'>" + body + "</svg>";
        }

        private static DwgOcrPage Page()
        {
            return new DwgOcrPage { width = 1000, height = 500, viewBox = new[] { -40d, 80d, 200d, 100d } };
        }

        private static DwgOcrLine Line(DwgOcrPoint[] points)
        {
            return new DwgOcrLine { text = "source", color = "#123456", confidence = 0.9, points = points };
        }

        // Independently evaluate the SVG path subset emitted by PrepareGeometry.
        // The test checks the visible result of even-odd clip intersections,
        // including overlap, rather than comparing generated path strings.
        private static List<List<DwgOcrPoint>> PathPolygons(string data)
        {
            string[] tokens = Regex.Matches(data, @"[MLhvZ]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?")
                .Cast<Match>().Select(match => match.Value).ToArray();
            var polygons = new List<List<DwgOcrPoint>>();
            List<DwgOcrPoint> current = null;
            double x = 0, y = 0;
            int index = 0;
            while (index < tokens.Length)
            {
                string command = tokens[index++];
                if (command == "Z")
                {
                    if (current == null || current.Count < 3) throw new InvalidOperationException("Invalid closed clip contour");
                    polygons.Add(current); current = null; continue;
                }
                if (command == "M" || command == "L")
                {
                    x = Double.Parse(tokens[index++], CultureInfo.InvariantCulture);
                    y = Double.Parse(tokens[index++], CultureInfo.InvariantCulture);
                    if (command == "M") current = new List<DwgOcrPoint>();
                }
                else if (command == "h") x += Double.Parse(tokens[index++], CultureInfo.InvariantCulture);
                else if (command == "v") y += Double.Parse(tokens[index++], CultureInfo.InvariantCulture);
                else throw new InvalidOperationException("Unknown test clip command " + command);
                if (current == null) throw new InvalidOperationException("Clip contour has no start");
                current.Add(Point(x, y));
            }
            return polygons;
        }

        private static bool Inside(IList<DwgOcrPoint> polygon, double x, double y)
        {
            bool inside = false;
            for (int i = 0, previous = polygon.Count - 1; i < polygon.Count; previous = i++)
            {
                DwgOcrPoint first = polygon[i], second = polygon[previous];
                if ((first.y > y) != (second.y > y) &&
                    x < (second.x - first.x) * (y - first.y) / (second.y - first.y) + first.x) inside = !inside;
            }
            return inside;
        }

        private static bool ClipVisible(XDocument document, XElement element, double x, double y)
        {
            foreach (XElement parent in element.Ancestors())
            {
                string reference = (string)parent.Attribute("clip-path");
                if (reference == null) continue;
                string id = reference.Substring("url(#".Length).TrimEnd(')');
                XElement clip = document.Descendants(Svg + "clipPath").Single(item => (string)item.Attribute("id") == id);
                XElement path = clip.Element(Svg + "path");
                Check((string)path.Attribute("clip-rule") == "evenodd", "Clip fill rule changed");
                bool visible = false;
                foreach (var polygon in PathPolygons((string)path.Attribute("d")))
                    if (Inside(polygon, x, y)) visible = !visible;
                if (!visible) return false;
            }
            return true;
        }

        public static int Main()
        {
            Test("SourceId restores the complete paragraph once and removes successful slots", delegate {
                string preview = Drawing("<g transform='translate(10 20)'><text data-ppt-text-id='paragraph'><tspan x='2'>first</tspan><tspan x='2' dy='18'>second</tspan></text><text data-ppt-text-id='native'>converted</text></g>");
                string geometry = Drawing("<g transform='translate(10 20)'><g data-ppt-text-slot='paragraph'/><g data-ppt-text-slot='native'/></g>");
                XDocument result = XDocument.Parse(DwgOcrLayout.PrepareGeometry(preview, geometry,
                    new HashSet<string>(new[] { "paragraph" }), new List<DwgOcrPoint[]>()));
                XElement text = result.Descendants(Svg + "text").Single();
                Check((string)text.Attribute("data-ppt-text-id") == "paragraph", "Wrong original paragraph restored");
                Check(text.Elements(Svg + "tspan").Select(item => item.Value).SequenceEqual(new[] { "first", "second" }), "Partial paragraph lost or duplicated");
                Check(!result.Descendants().Any(item => item.Attribute("data-ppt-text-slot") != null), "A native slot remains");
            });

            Test("fallback text keeps ancestor transforms and stays outside OCR cuts", delegate {
                string parents = "<g id='outer' transform='translate(20 10)' fill='#112233' style='font-size:12px'><g id='inner' transform='rotate(25)' font-family='Arial'>";
                string preview = Drawing(parents + "<text data-ppt-text-id='tiny' x='3' y='4'>tiny</text></g></g>");
                string geometry = Drawing(parents + "<path id='outline' d='M0 0 L30 30'/><g data-ppt-text-slot='tiny'/></g></g>");
                XDocument result = XDocument.Parse(DwgOcrLayout.PrepareGeometry(preview, geometry,
                    new HashSet<string>(new[] { "tiny" }), new List<DwgOcrPoint[]> { Rectangle(0, 0, 100, 80) }));
                XElement text = result.Descendants(Svg + "text").Single();
                Check(!text.Ancestors().Any(item => ((string)item.Attribute("clip-path") ?? "").Contains("ppt-ocr-")), "OCR erased a retained paragraph");
                Check(text.Ancestors().Any(item => (string)item.Attribute("transform") == "translate(20 10)" && (string)item.Attribute("fill") == "#112233" && (string)item.Attribute("style") == "font-size:12px"), "Outer style or placement lost");
                Check(text.Ancestors().Any(item => (string)item.Attribute("transform") == "rotate(25)" && (string)item.Attribute("font-family") == "Arial"), "Inner transform or font lost");
                Check(result.Descendants().Count(item => (string)item.Attribute("id") == "outer") == 1, "Cloned ancestors duplicated SVG IDs");
                XElement outline = result.Descendants().Single(item => (string)item.Attribute("id") == "outline");
                Check(!ClipVisible(result, outline, 20, 20), "Selected geometry was not clipped");
            });

            Test("unconverted complex text also survives an overlapping OCR selection", delegate {
                string drawing = Drawing("<g transform='scale(2)' opacity='0.5'><text transform='skewX(12)' x='10' y='20'>kept</text><rect id='art' width='200' height='100'/></g>");
                XDocument result = XDocument.Parse(DwgOcrLayout.PrepareGeometry(drawing, drawing,
                    new HashSet<string>(), new List<DwgOcrPoint[]> { Rectangle(0, 0, 100, 80) }));
                XElement text = result.Descendants(Svg + "text").Single();
                Check((string)text.Attribute("transform") == "skewX(12)", "Unsupported text transform changed");
                Check(text.Ancestors().Any(item => (string)item.Attribute("opacity") == "0.5"), "Retained opacity changed");
                Check(!text.Ancestors().Any(item => ((string)item.Attribute("clip-path") ?? "").Contains("ppt-ocr-")), "Retained complex text was clipped");
            });

            Test("OCR pixels map into a nonzero SVG viewBox without mutating source coordinates", delegate {
                DwgOcrLine line = Line(Rectangle(100, 100, 300, 150));
                DwgOcrPoint[] points = DwgOcrLayout.ToSvgPoints(Page(), line);
                Near(points[0].x, -20, "Nonzero viewBox X ignored"); Near(points[0].y, 100, "Nonzero viewBox Y ignored");
                Near(points[2].x, 20, "OCR width scale changed"); Near(points[2].y, 110, "OCR height scale changed");
                Near(line.points[0].x, 100, "Input OCR coordinates mutated");
                DwgTextItem text = DwgOcrLayout.ToText(Page(), line, "edited", "ocr-7");
                Check(text.Text == "edited" && text.SourceId == "ocr-7" && text.Color == "#123456", "Reviewed OCR text identity changed");
                Near(text.CenterX, 0, "OCR center X changed"); Near(text.CenterY, 105, "OCR center Y changed");
                Near(text.InkWidth, 40, "OCR ink width changed"); Near(text.InkHeight, 10, "OCR ink height changed");
            });

            Test("rotated OCR quadrilateral preserves its reading axis", delegate {
                DwgOcrLine line = Line(new[] { Point(100, 100), Point(200, 200), Point(175, 225), Point(75, 125) });
                DwgTextItem text = DwgOcrLayout.ToText(Page(), line, "rotated", "ocr-8");
                Near(text.Rotation, 45, "OCR rotation lost");
                Near(text.Width, Math.Sqrt(800), "Rotated OCR width changed");
                Near(text.Height, Math.Sqrt(50), "Rotated OCR height changed");
                Near(text.CenterX, -12.5, "Rotated center X changed"); Near(text.CenterY, 112.5, "Rotated center Y changed");
            });

            Test("invalid, nonfinite and out-of-page OCR coordinates are rejected", delegate {
                var invalidPage = Page(); invalidPage.viewBox[0] = Double.NaN;
                Reject(delegate { DwgOcrLayout.ToSvgPoints(invalidPage, Line(Rectangle(0, 0, 50, 50))); }, "NaN viewBox accepted");
                invalidPage = Page(); invalidPage.viewBox[3] = Double.PositiveInfinity;
                Reject(delegate { DwgOcrLayout.ToSvgPoints(invalidPage, Line(Rectangle(0, 0, 50, 50))); }, "Infinite viewBox accepted");
                Reject(delegate { DwgOcrLayout.ToSvgPoints(Page(), Line(new[] { Point(Double.NaN, 0), Point(10, 0), Point(10, 10), Point(0, 10) })); }, "NaN point accepted");
                Reject(delegate { DwgOcrLayout.ToSvgPoints(Page(), Line(Rectangle(-1, 0, 50, 50))); }, "Outside-page point accepted");
                Reject(delegate { DwgOcrLayout.ToSvgPoints(Page(), Line(new[] { Point(0, 0), Point(1, 1), Point(2, 2) })); }, "Incomplete OCR box accepted");
                Reject(delegate { DwgOcrLayout.ToText(Page(), Line(Rectangle(20, 20, 20, 40)), "flat", "ocr-9"); }, "Zero-width OCR text accepted");
                Reject(delegate { DwgOcrLayout.ToText(Page(), Line(Rectangle(0, 0, 50, 50)), "line1\nline2", "ocr-9"); }, "Multiline OCR override accepted");
            });

            Test("only selected OCR areas are hidden and overlap cannot reveal cut geometry", delegate {
                string drawing = Drawing("<rect id='background-art' width='200' height='100'/>");
                XDocument one = XDocument.Parse(DwgOcrLayout.PrepareGeometry(drawing, drawing, new HashSet<string>(),
                    new List<DwgOcrPoint[]> { Rectangle(10, 10, 30, 30) }));
                XElement oneArt = one.Descendants(Svg + "rect").Single();
                Check(!ClipVisible(one, oneArt, 20, 20), "Selected original outline remains visible");
                Check(ClipVisible(one, oneArt, 80, 20), "Unselected art was cleared");
                Check(ClipVisible(one, oneArt, 35, 20), "Neighboring unselected area was cleared");
                XDocument overlap = XDocument.Parse(DwgOcrLayout.PrepareGeometry(drawing, drawing, new HashSet<string>(),
                    new List<DwgOcrPoint[]> { Rectangle(10, 10, 30, 30), Rectangle(20, 10, 40, 30) }));
                XElement overlapArt = overlap.Descendants(Svg + "rect").Single();
                Check(!ClipVisible(overlap, overlapArt, 15, 20), "First selection not hidden");
                Check(!ClipVisible(overlap, overlapArt, 25, 20), "Overlapping OCR cuts used XOR and revealed the source");
                Check(!ClipVisible(overlap, overlapArt, 35, 20), "Second selection not hidden");
                Check(ClipVisible(overlap, overlapArt, 80, 20), "Outside geometry was cleared");
                Check(overlap.Descendants(Svg + "rect").Count() == 1, "Source vector duplicated or deleted");
            });

            return failures == 0 ? 0 : 1;
        }
    }
}
'@
try {
    $harnessPath = Join-Path $scratch 'GeometryRegression.cs'
    [IO.File]::WriteAllText($harnessPath, $harness, [Text.UTF8Encoding]::new($false))
    $office = Get-ChildItem -Path (Join-Path $env:windir 'assembly\GAC_MSIL\office\*\office.dll') | Select-Object -First 1 -ExpandProperty FullName
    $powerPoint = Get-ChildItem -Path (Join-Path $env:windir 'assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\*\Microsoft.Office.Interop.PowerPoint.dll') | Select-Object -First 1 -ExpandProperty FullName
    if (!$office -or !$powerPoint) { throw 'Office interop assemblies were not found for compile-only tests.' }
    $compiler = Join-Path $env:windir 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
    $executable = Join-Path $scratch 'GeometryRegression.exe'
    $arguments = @('/nologo', '/target:exe', '/langversion:5', '/codepage:65001', ('/out:' + $executable),
        '/reference:System.dll', '/reference:System.Core.dll', '/reference:System.Xml.dll', '/reference:System.Xml.Linq.dll',
        ('/reference:' + $office), ('/reference:' + $powerPoint),
        (Join-Path $sourceRoot 'DwgOcrLayout.cs'), (Join-Path $sourceRoot 'DwgTextInsertion.cs'), $harnessPath)
    & $compiler @arguments
    if ($LASTEXITCODE -ne 0) { throw 'DWG geometry harness compilation failed.' }
    & $executable
    if ($LASTEXITCODE -ne 0) { throw 'DWG geometry regressions failed.' }
}
finally {
    $target = [IO.Path]::GetFullPath($scratch)
    $temporaryRoot = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\'
    if ($target.StartsWith($temporaryRoot, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($target).StartsWith('ppt-dwg-geometry-tests-')) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
}
