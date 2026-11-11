using System.Xml.Linq;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.Image;
using ACadSharp.IO;
using ACadSharp.Tables;
using CSMath;
using Dongsidaye.CadConverter;
using ImageColor = SixLabors.ImageSharp.Color;

int failures = 0;
void Check(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}
void Near(double value, double expected, string message, double tolerance = 1e-5)
{
    Check(Math.Abs(value - expected) <= tolerance, message + $" ({value} vs {expected})");
}
void Test(string name, Action work)
{
    try { work(); Console.WriteLine("PASS " + name); }
    catch (Exception error) { failures++; Console.Error.WriteLine("FAIL " + name + " " + error.Message); }
}
string Svg(string content) => "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 500 300\"><g font-family=\"Arial\" fill=\"#123456\">" + content + "</g></svg>";
XNamespace svgNamespace = "http://www.w3.org/2000/svg";

Test("same-row anchors, baseline, color and native fallback slots", () =>
{
    var result = EditableTextLayout.Extract(Svg("<line x1=\"0\" y1=\"0\" x2=\"1\" y2=\"1\"/><text x=\"30\" y=\"40\" font-size=\"20\">Label 中文</text>"));
    Check(result.Texts.Count == 1, "Expected one editable row");
    var item = result.Texts[0];
    Near(item.X, 30, "Anchor X changed"); Near(item.Y, 40, "Anchor Y changed"); Near(item.FontSize, 20, "Font size changed");
    Check(item.CenterY < item.Y && item.InkWidth > 0 && item.InkHeight > 0, "Alphabetic baseline was not measured");
    Check(item.Color == "#123456", "Inherited fill lost");
    Check(result.HasGeometry, "Line geometry lost");
    Check(XDocument.Parse(result.GeometrySvg).Descendants(svgNamespace + "text").Count() == 0, "Editable label is duplicated in geometry");
    Check(result.GeometrySvg.Contains("data-ppt-text-slot=\"" + item.SourceId + "\""), "Fallback slot missing");
    Check(result.PreviewSvg.Contains("data-ppt-text-id=\"" + item.SourceId + "\""), "Preview identity missing");
});

Test("nested transforms compose in SVG order with nonuniform width", () =>
{
    string label = "<text x=\"3\" y=\"4\" font-size=\"10\">ABCDE</text>";
    var unscaled = EditableTextLayout.Extract(Svg(label)).Texts.Single();
    var transformed = EditableTextLayout.Extract(Svg("<g transform=\"translate(100 50) rotate(90) scale(2 3)\">" + label + "</g>")).Texts.Single();
    Near(transformed.X, 88, "Transform order X is wrong"); Near(transformed.Y, 56, "Transform order Y is wrong");
    Near(transformed.Rotation, 90, "Rotation lost");
    Near(transformed.Width, unscaled.Width * 2, "Width axis not preserved");
    Near(transformed.FontSize, 30, "Vertical font size not preserved");
    Near(transformed.InkHeight, unscaled.InkHeight * 3, "Ink height not preserved");
});

Test("MTEXT tspans retain explicit line baselines and one source identity", () =>
{
    var result = EditableTextLayout.Extract(Svg("<text x=\"15\" y=\"25\" font-size=\"12\" dominant-baseline=\"hanging\"><tspan x=\"15\">第一行</tspan><tspan x=\"15\" dy=\"18\">第二行</tspan><tspan x=\"15\" dy=\"18\">Third</tspan></text>"));
    Check(result.Texts.Count == 3, "Line count changed");
    Near(result.Texts[1].Y - result.Texts[0].Y, 18, "First line spacing changed");
    Near(result.Texts[2].Y - result.Texts[1].Y, 18, "Second line spacing changed");
    Check(result.Texts.Select(item => item.SourceId).Distinct().Count() == 1, "MTEXT fallback identity split across lines");
    Check(result.Texts[0].CenterY > result.Texts[0].Y, "Hanging baseline not preserved");
});

Test("center and end anchors preserve source textLength", () =>
{
    var start = EditableTextLayout.Extract(Svg("<text x=\"200\" y=\"80\" font-size=\"16\" textLength=\"100\">ABCDE</text>")).Texts.Single();
    var middle = EditableTextLayout.Extract(Svg("<text x=\"200\" y=\"80\" font-size=\"16\" textLength=\"100\" text-anchor=\"middle\">ABCDE</text>")).Texts.Single();
    var end = EditableTextLayout.Extract(Svg("<text x=\"200\" y=\"80\" font-size=\"16\" textLength=\"100\" text-anchor=\"end\">ABCDE</text>")).Texts.Single();
    Near(start.Width, 100, "textLength ignored");
    Near(start.CenterX - middle.CenterX, 50, "Middle anchor not centered");
    Near(start.CenterX - end.CenterX, 100, "End anchor not preserved");
});

Test("unsupported skew, reflection, clip and malformed coordinates preserve SVG text", () =>
{
    foreach (string attributes in new[] { "transform=\"skewX(12)\"", "transform=\"scale(-1 1)\"", "clip-path=\"url(#clip)\"", "x=\"1 2 3\"", "transform=\"matrix(1e308 0 0 1 0 0)\"" })
    {
        var result = EditableTextLayout.Extract(Svg("<text y=\"40\" font-size=\"12\" " + attributes + ">保留</text>"));
        Check(result.Texts.Count == 0, "Unsupported label was partially extracted");
        Check(XDocument.Parse(result.GeometrySvg).Descendants(svgNamespace + "text").Single().Value == "保留", "Unsupported text was removed");
        Check(result.Warnings.Count > 0, "Unsupported layout had no warning");
    }
});

Test("mixed unsupported label does not lose successfully extracted sibling", () =>
{
    var result = EditableTextLayout.Extract(Svg("<text x=\"10\" y=\"20\" font-size=\"12\">native</text><text x=\"20\" y=\"40\" font-size=\"12\"><textPath href=\"#curve\">curved</textPath></text>"));
    Check(result.Texts.Single().Text == "native", "Valid sibling was dropped");
    Check(XDocument.Parse(result.GeometrySvg).Descendants(svgNamespace + "text").Single().Value == "curved", "Curved text was lost");
});

Test("vertical writing is retained instead of becoming horizontal editable text", () =>
{
    var result = EditableTextLayout.Extract(Svg("<text x=\"20\" y=\"40\" font-size=\"12\" writing-mode=\"vertical-rl\">竖排文字</text>"));
    Check(result.Texts.Count == 0 && result.GeometrySvg.Contains("竖排文字"), "Vertical glyph layout was lost");
});

Test("span-specific transparency retains the whole source paragraph", () =>
{
    var result = EditableTextLayout.Extract(Svg("<text x=\"20\" y=\"40\" font-size=\"12\"><tspan x=\"20\">opaque</tspan><tspan x=\"20\" dy=\"18\" fill-opacity=\"0.5\">translucent</tspan></text>"));
    Check(result.Texts.Count == 0 && XDocument.Parse(result.GeometrySvg).Descendants(svgNamespace + "text").Count() == 1, "Partial paragraph opacity was discarded");
});

Test("rotated ink expands equal preview and geometry viewBoxes", () =>
{
    var result = EditableTextLayout.Extract(Svg("<text x=\"0\" y=\"0\" font-size=\"20\" transform=\"rotate(-30)\">扩展画布</text>"));
    string viewBox = XDocument.Parse(result.PreviewSvg).Root!.Attribute("viewBox")!.Value;
    Check(viewBox == XDocument.Parse(result.GeometrySvg).Root!.Attribute("viewBox")!.Value, "Preview and geometry no longer share coordinates");
    Check(double.Parse(viewBox.Split(' ')[1], System.Globalization.CultureInfo.InvariantCulture) < 0, "Rotated ink was clipped by the old bounds");
});

Test("DWG TEXT width factor is corrected exactly once", () =>
{
    var drawing = new CadDocument(ACadVersion.AC1032);
    var source = new TextEntity { Value = "Width factor", InsertPoint = new XYZ(0, 0, 0), Height = 10, WidthFactor = 0.6 };
    drawing.ModelSpace.Entities.Add(source);
    string label = $"<text x=\"10\" y=\"20\" font-size=\"14\" data-handle=\"{source.Handle:X}\">Width factor</text>";
    var normal = EditableTextLayout.Extract(Svg(label)).Texts.Single();
    var corrected = EditableTextLayout.Extract(Svg(label), drawing);
    Near(corrected.Texts.Single().Width, normal.Width * 0.6, "DWG width factor was lost");
    var again = EditableTextLayout.Extract(corrected.PreviewSvg, drawing);
    Near(again.Texts.Single().Width, corrected.Texts.Single().Width, "Width factor was applied twice");
});

Test("ambiguous block labels retain original SVG instead of guessing width", () =>
{
    var drawing = new CadDocument(ACadVersion.AC1032);
    var block = new BlockRecord("ambiguous");
    block.Entities.Add(new TextEntity { Value = "same", Height = 10, WidthFactor = 0.6 });
    block.Entities.Add(new TextEntity { Value = "same", Height = 10, WidthFactor = 1.3 });
    drawing.BlockRecords.Add(block);
    var result = EditableTextLayout.Extract(Svg("<text x=\"10\" y=\"20\" font-size=\"14\" data-block=\"ambiguous\">same</text>"), drawing);
    Check(result.Texts.Count == 0 && result.GeometrySvg.Contains(">same</text>"), "Ambiguous source was guessed or removed");
});

Test("actual DWG TEXT, MTEXT and block ATTRIB render through the same extractor", () =>
{
    var drawing = new CadDocument(ACadVersion.AC1032);
    drawing.ModelSpace.Entities.Add(new TextEntity { Value = "旋转宽文字", InsertPoint = new XYZ(20, 160, 0), Height = 12, Rotation = Math.PI / 6, WidthFactor = 0.8 });
    drawing.ModelSpace.Entities.Add(new MText { Value = "第一行\\P第二行\\P第三行", InsertPoint = new XYZ(20, 120, 0), Height = 10, RectangleWidth = 100, AttachmentPoint = AttachmentPointType.TopLeft, LineSpacing = 1.2 });
    var block = new BlockRecord("text-block");
    block.Entities.Add(new TextEntity { Value = "块内文字", Height = 8 });
    drawing.BlockRecords.Add(block);
    var insert = new Insert(block) { InsertPoint = new XYZ(180, 50, 0), Rotation = Math.PI / 4 };
    insert.Attributes.Add(new AttributeEntity { Tag = "label", Value = "属性文字", InsertPoint = new XYZ(180, 65, 0), Height = 8 });
    drawing.ModelSpace.Entities.Add(insert);
    string folder = Path.Combine(Path.GetTempPath(), "cad-text-layout-" + Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(folder);
    try
    {
        string path = Path.Combine(folder, "fixture.dwg");
        DwgWriter.Write(path, drawing);
        CadDocument read = DwgReader.Read(path);
        var exporter = new ImageExporter();
        exporter.Configuration.BackgroundColor = ImageColor.Transparent;
        exporter.Configuration.ForegroundColor = ImageColor.Black;
        exporter.Configuration.Svg.EmitSize = false;
        exporter.AddModelSpace(read);
        IReadOnlyList<RenderedPage> pages = exporter.Render(ImageExportFormat.Svg);
        try
        {
            var result = EditableTextLayout.Extract(((RenderedSvgPage)pages.Single()).Content, read);
            Check(result.Texts.Count == 6, "DWG text/MTEXT/attribute rows were omitted");
            Check(result.Texts.Any(item => item.Text == "属性文字"), "Attribute value missing");
            Check(result.Texts.Any(item => Math.Abs(item.Rotation + 45) < 0.1), "Block rotation lost");
            Check(result.Texts.Count(item => item.Text.StartsWith("第", StringComparison.Ordinal)) == 3, "MTEXT lines changed");
        }
        finally { foreach (RenderedPage page in pages) page.Dispose(); }
    }
    finally { Directory.Delete(folder, true); }
});

if (failures > 0) Environment.ExitCode = 1;
