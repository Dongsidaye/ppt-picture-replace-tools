using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace Dongsidaye.PowerPointTools
{
    internal sealed class DwgOcrPoint
    {
        public double x { get; set; }
        public double y { get; set; }
    }

    internal sealed class DwgOcrLine
    {
        public string text { get; set; }
        public double confidence { get; set; }
        public DwgOcrPoint[] points { get; set; }
        public DwgOcrPoint[] inkPoints { get; set; }
        public string color { get; set; }
    }

    internal sealed class DwgOcrPage
    {
        public int width { get; set; }
        public int height { get; set; }
        public double[] viewBox { get; set; }
        public DwgOcrLine[] lines { get; set; }
    }

    internal sealed class DwgOcrSelection
    {
        public int index { get; set; }
        public string text { get; set; }
    }

    internal sealed class DwgInsertionOptions
    {
        public string font { get; set; }
        public DwgOcrSelection[] ocrSelections { get; set; }
    }

    internal static class DwgOcrLayout
    {
        internal static DwgOcrPoint[] ToSvgPoints(DwgOcrPage page, DwgOcrLine line)
        {
            if (page == null || page.width <= 0 || page.height <= 0 || page.viewBox == null || page.viewBox.Length != 4 ||
                page.viewBox.Any(value => !Finite(value)) || page.viewBox[2] <= 0 || page.viewBox[3] <= 0 ||
                line == null || line.points == null || line.points.Length != 4)
                throw new InvalidOperationException("OCR 文字区域无效，请重新识别。");
            return line.points.Select(point => {
                if (!Finite(point.x) || !Finite(point.y) || point.x < 0 || point.y < 0 || point.x > page.width || point.y > page.height)
                    throw new InvalidOperationException("OCR 坐标超出图纸范围。");
                return new DwgOcrPoint { x = page.viewBox[0] + point.x * page.viewBox[2] / page.width,
                    y = page.viewBox[1] + point.y * page.viewBox[3] / page.height };
            }).ToArray();
        }

        internal static DwgTextItem ToText(DwgOcrPage page, DwgOcrLine line, string text, string sourceId)
        {
            if (String.IsNullOrWhiteSpace(text) || text.Length > 2000 || text.IndexOfAny(new[] {'\r','\n'}) >= 0)
                throw new InvalidOperationException("请核对 OCR 的单行文字内容。");
            DwgOcrPoint[] points = ToSvgPoints(page, line.inkPoints != null && line.inkPoints.Length == 4
                ? new DwgOcrLine { points = line.inkPoints } : line);
            double width = Distance(points[0], points[1]), height = Distance(points[0], points[3]);
            if (width <= 0 || height <= 0) throw new InvalidOperationException("OCR 文字尺寸无效。");
            return new DwgTextItem { SourceId = sourceId, Text = text,
                CenterX = points.Average(point => point.x), CenterY = points.Average(point => point.y),
                Width = width, Height = height, InkWidth = width, InkHeight = height, FontSize = height / 0.8,
                Rotation = Math.Atan2(points[1].y - points[0].y, points[1].x - points[0].x) * 180 / Math.PI,
                FontFamily = "Microsoft YaHei", Color = line.color ?? "#000000" };
        }

        internal static string PrepareGeometry(string previewSvg, string geometrySvg, HashSet<string> fallback,
            IList<DwgOcrPoint[]> replacements)
        {
            XDocument preview = XDocument.Parse(previewSvg), geometry = XDocument.Parse(geometrySvg);
            XNamespace svg = "http://www.w3.org/2000/svg";
            foreach (XElement slot in geometry.Descendants().Where(element => element.Attribute("data-ppt-text-slot") != null).ToArray())
            {
                string id = (string)slot.Attribute("data-ppt-text-slot");
                if (fallback.Contains(id))
                {
                    XElement original = preview.Descendants(svg + "text").Single(element => (string)element.Attribute("data-ppt-text-id") == id);
                    slot.ReplaceWith(new XElement(original));
                }
                else slot.Remove();
            }
            XElement root = geometry.Root;
            if (replacements.Count > 0)
            {
                // Keep retained CAD text outside OCR cuts. Clone its ancestor
                // transforms/styles so a nearby outline selection cannot erase it.
                var retainedText = new List<XElement>();
                foreach (XElement text in root.Descendants(svg + "text").Where(element => !element.Ancestors(svg + "defs").Any()).ToArray())
                {
                    XElement copy = new XElement(text);
                    foreach (XElement parent in text.Ancestors().TakeWhile(parent => parent != root))
                        copy = new XElement(parent.Name, parent.Attributes().Where(attribute => attribute.Name.LocalName != "id"), copy);
                    retainedText.Add(copy);
                    text.Remove();
                }
                string[] values = ((string)root.Attribute("viewBox")).Split(new[] {' ', ',', '\t', '\r', '\n'}, StringSplitOptions.RemoveEmptyEntries);
                double[] box = values.Select(value => Double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
                XElement definitions = new XElement(svg + "defs");
                string clipId = "ppt-ocr-" + Guid.NewGuid().ToString("N");
                string outer = "M" + N(box[0]) + " " + N(box[1]) + "h" + N(box[2]) + "v" + N(box[3]) + "h" + N(-box[2]) + "Z";
                // Sequential clips subtract the union of the selected areas,
                // including overlapping OCR boxes, without covering other art.
                XElement content = new XElement(svg + "g");
                foreach (XElement element in root.Elements().Where(element => element.Name != svg + "defs").ToArray())
                { element.Remove(); content.Add(element); }
                int index = 0;
                foreach (DwgOcrPoint[] polygon in replacements)
                {
                    string id = clipId + (++index).ToString(CultureInfo.InvariantCulture);
                    string cut = "M" + String.Join(" L", polygon.Select(point => N(point.x) + " " + N(point.y))) + "Z";
                    definitions.Add(new XElement(svg + "clipPath", new XAttribute("id", id), new XAttribute("clipPathUnits", "userSpaceOnUse"),
                        new XElement(svg + "path", new XAttribute("clip-rule", "evenodd"), new XAttribute("d", outer + cut))));
                    content = new XElement(svg + "g", new XAttribute("clip-path", "url(#" + id + ")"), content);
                }
                root.AddFirst(definitions);
                root.Add(content);
                root.Add(retainedText);
            }
            return geometry.ToString(SaveOptions.DisableFormatting);
        }

        private static double Distance(DwgOcrPoint first, DwgOcrPoint second)
        { return Math.Sqrt(Math.Pow(first.x - second.x, 2) + Math.Pow(first.y - second.y, 2)); }
        private static bool Finite(double value) { return !Double.IsNaN(value) && !Double.IsInfinity(value); }
        private static string N(double value) { return value.ToString("0.########", CultureInfo.InvariantCulture); }
    }
}
