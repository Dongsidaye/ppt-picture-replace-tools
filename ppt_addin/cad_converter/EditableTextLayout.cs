using System.Globalization;
using System.Numerics;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using ACadSharp;
using ACadSharp.Entities;
using SixLabors.Fonts;

namespace Dongsidaye.CadConverter;

internal sealed class EditableTextItem
{
    public string SourceId { get; init; } = "";
    public string Text { get; init; } = "";
    public double X { get; init; }
    public double Y { get; init; }
    public double CenterX { get; init; }
    public double CenterY { get; init; }
    public double Width { get; init; }
    public double Height { get; init; }
    public double InkWidth { get; init; }
    public double InkHeight { get; init; }
    public double FontSize { get; init; }
    public string FontFamily { get; init; } = "Arial";
    public double Rotation { get; init; }
    public string Color { get; init; } = "#000000";
    public bool Bold { get; init; }
    public bool Italic { get; init; }
}

internal sealed class EditableTextLayoutResult
{
    public string PreviewSvg { get; init; } = "";
    public string GeometrySvg { get; init; } = "";
    public bool HasGeometry { get; init; }
    public List<EditableTextItem> Texts { get; init; } = new();
    public List<string> Warnings { get; init; } = new();
}

/// <summary>Preserves the converter's line placement while separating editable labels.</summary>
internal static class EditableTextLayout
{
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
    private static readonly Regex TransformPart = new(@"([a-zA-Z]+)\s*\(([^)]*)\)", RegexOptions.Compiled);
    private static readonly Regex Numbers = new(@"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?", RegexOptions.Compiled);
    private static readonly HashSet<string> GeometryNames = new(StringComparer.Ordinal)
    {
        "path", "line", "polyline", "polygon", "circle", "ellipse", "rect", "image", "text", "use"
    };
    private const float ReferenceSize = 100f;

    public static EditableTextLayoutResult Extract(string fullSvg, CadDocument? source = null)
    {
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 16 * 1024 * 1024 };
        using var textReader = new StringReader(fullSvg);
        using var reader = XmlReader.Create(textReader, settings);
        XDocument preview = XDocument.Load(reader);
        XElement root = preview.Root ?? throw new InvalidDataException("SVG 根元素缺失。");
        if (root.Name != Svg + "svg") throw new InvalidDataException("文字布局输入不是 SVG。");

        var items = new List<EditableTextItem>();
        var warnings = new List<string>();
        var extracted = new List<XElement>();
        int sourceIndex = 0;
        foreach (XElement element in root.Descendants(Svg + "text").ToArray())
        {
            element.SetAttributeValue("data-ppt-text-id", "ppt-text-" + (++sourceIndex).ToString(CultureInfo.InvariantCulture));
            try
            {
                CorrectSourceWidth(element, source, warnings);
                List<EditableTextItem> lines = ExtractElement(element);
                if (lines.Count == 0) continue;
                items.AddRange(lines);
                extracted.Add(element);
            }
            catch (LayoutUnsupported error)
            {
                string label = element.Value.Trim();
                if (label.Length > 24) label = label[..24] + "…";
                warnings.Add($"文字“{label}”保留为 SVG。{error.Message}");
            }
            catch (Exception error) when (error is ArgumentException or InvalidOperationException or FontException)
            {
                warnings.Add("部分文字无法用当前系统字体测量，已保留在 SVG 中。");
            }
        }
        ExpandViewBox(root, items);
        string previewSvg = preview.ToString(SaveOptions.DisableFormatting);
        foreach (XElement element in extracted)
            element.ReplaceWith(new XElement(Svg + "g", new XAttribute("data-ppt-text-slot", (string)element.Attribute("data-ppt-text-id")!)));
        bool hasGeometry = root.Descendants().Any(element => element.Name.Namespace == Svg && GeometryNames.Contains(element.Name.LocalName));
        return new EditableTextLayoutResult
        {
            PreviewSvg = previewSvg,
            GeometrySvg = preview.ToString(SaveOptions.DisableFormatting),
            HasGeometry = hasGeometry,
            Texts = items,
            Warnings = warnings.Distinct(StringComparer.Ordinal).Take(20).ToList()
        };
    }

    private static void CorrectSourceWidth(XElement element, CadDocument? source, List<string> warnings)
    {
        if (source == null || element.Attribute("textLength") != null || element.Attribute("data-width-corrected") != null) return;
        string transform = (string?)element.Attribute("transform") ?? "";
        if (Regex.IsMatch(transform, @"\bscale\s*\(", RegexOptions.IgnoreCase)) return;

        TextEntity? entity = null;
        string handle = (string?)element.Attribute("data-handle") ?? "";
        if (ulong.TryParse(handle, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out ulong identity))
        {
            entity = source.GetCadObject(identity) as TextEntity;
        }
        else
        {
            string blockName = (string?)element.Attribute("data-block") ?? "";
            if (blockName.Length == 0) return;
            var candidates = source.BlockRecords.Where(block => block.Name == blockName)
                .SelectMany(block => block.Entities.OfType<TextEntity>())
                .Where(candidate => candidate.Value == element.Value).ToList();
            if (candidates.Count == 1) entity = candidates[0];
            else if (candidates.Any(candidate => Math.Abs(candidate.WidthFactor - 1d) > 1e-8))
            {
                throw new LayoutUnsupported("图块中文字存在多个可能来源，无法确认原始字宽。");
            }
        }
        if (entity == null) return;
        if (Math.Abs(entity.ObliqueAngle) > 1e-8) throw new LayoutUnsupported("倾斜文字需要保留原图形排版。");
        double factor = entity.WidthFactor;
        if (!double.IsFinite(factor) || factor <= 0) throw new LayoutUnsupported("原始字宽系数无效。");
        if (Math.Abs(factor - 1d) <= 1e-8) return;
        double x = Coordinate(element, "x", 0), y = Coordinate(element, "y", 0);
        string stretch = $"translate({Number(x)} {Number(y)}) scale({Number(factor)} 1) translate({Number(-x)} {Number(-y)})";
        element.SetAttributeValue("transform", (transform + " " + stretch).Trim());
        element.SetAttributeValue("data-width-corrected", "true");
        warnings.Add("已根据 DWG 原始字宽系数校正可编辑文字宽度。");
    }

    private static List<EditableTextItem> ExtractElement(XElement element)
    {
        XElement[] context = element.AncestorsAndSelf().Reverse().ToArray();
        Matrix matrix = Matrix.Identity;
        var style = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (XElement ancestor in context)
        {
            if (ancestor.Name.Namespace != Svg || ancestor.Name.LocalName is not ("svg" or "g" or "text"))
                throw new LayoutUnsupported("裁剪、定义或特殊绘制上下文不能直接转换成文本框。");
            ApplyStyle(style, ancestor);
            if (style.TryGetValue("display", out string? display) && display == "none") return new List<EditableTextItem>();
            if (style.TryGetValue("visibility", out string? visibility) && visibility is "hidden" or "collapse") return new List<EditableTextItem>();
            foreach (string property in new[] { "clip-path", "mask", "filter" })
                if (style.TryGetValue(property, out string? value) && value != "none") throw new LayoutUnsupported("含裁剪、蒙版或滤镜。");
            foreach (string property in new[] { "opacity", "fill-opacity" })
                if (style.TryGetValue(property, out string? value) && Math.Abs(Scalar(value) - 1d) > 1e-9) throw new LayoutUnsupported("含透明度效果。");
            matrix *= ParseTransform((string?)ancestor.Attribute("transform"));
        }
        if (element.Attribute("rotate") != null) throw new LayoutUnsupported("含逐字符旋转。");
        EnsureOrthogonal(matrix);

        var spans = element.Elements().ToArray();
        if (spans.Length == 0)
        {
            string value = TextValue(element);
            if (string.IsNullOrWhiteSpace(value)) return new List<EditableTextItem>();
            return new List<EditableTextItem> { MakeItem(element, value, Coordinate(element, "x", 0) + Coordinate(element, "dx", 0), Coordinate(element, "y", 0) + Coordinate(element, "dy", 0), style, matrix) };
        }
        if (spans.Any(span => span.Name != Svg + "tspan" || span.HasElements) ||
            element.Nodes().OfType<XText>().Any(node => !string.IsNullOrWhiteSpace(node.Value)))
            throw new LayoutUnsupported("含混排或嵌套字符布局。");
        if (element.Attribute("textLength") != null) throw new LayoutUnsupported("多行段落包含整段宽度约束。");

        double cursorX = Coordinate(element, "x", 0) + Coordinate(element, "dx", 0);
        double cursorY = Coordinate(element, "y", 0) + Coordinate(element, "dy", 0);
        var result = new List<EditableTextItem>();
        foreach (XElement span in spans)
        {
            if (span.Attribute("rotate") != null) throw new LayoutUnsupported("含逐字符旋转。");
            cursorX = Coordinate(span, "x", cursorX) + Coordinate(span, "dx", 0);
            cursorY = Coordinate(span, "y", cursorY) + Coordinate(span, "dy", 0);
            var lineStyle = new Dictionary<string, string>(style, StringComparer.OrdinalIgnoreCase);
            ApplyStyle(lineStyle, span);
            Matrix lineMatrix = matrix * ParseTransform((string?)span.Attribute("transform"));
            EnsureOrthogonal(lineMatrix);
            string value = TextValue(span);
            if (string.IsNullOrWhiteSpace(value)) continue;
            EditableTextItem item = MakeItem(span, value, cursorX, cursorY, lineStyle, lineMatrix);
            result.Add(item);
            cursorX += item.Width / lineMatrix.ScaleX;
        }
        return result;
    }

    private static EditableTextItem MakeItem(XElement element, string value, double x, double y, Dictionary<string, string> style, Matrix matrix)
    {
        foreach (string property in new[] { "clip-path", "mask", "filter" })
            if (style.TryGetValue(property, out string? effect) && effect != "none") throw new LayoutUnsupported("含行级裁剪、蒙版或滤镜。");
        foreach (string property in new[] { "opacity", "fill-opacity" })
            if (style.TryGetValue(property, out string? opacity) && Math.Abs(Scalar(opacity) - 1d) > 1e-9) throw new LayoutUnsupported("含行级透明度效果。");
        if (style.TryGetValue("writing-mode", out string? writingMode) && writingMode is not ("horizontal-tb" or "lr-tb" or "lr")) throw new LayoutUnsupported("竖向文字布局暂保留为图形。");
        if (style.TryGetValue("text-decoration", out string? decoration) && decoration != "none") throw new LayoutUnsupported("含文字装饰效果。");
        if (style.TryGetValue("font-stretch", out string? stretch) && stretch is not ("normal" or "100%")) throw new LayoutUnsupported("含字体伸缩效果。");
        if (value.IndexOfAny(new[] { '\r', '\n' }) >= 0) throw new LayoutUnsupported("文本行未提供独立的基线坐标。");
        double fontSize = style.TryGetValue("font-size", out string? fontSizeValue) ? Scalar(fontSizeValue) : 16d;
        if (fontSize <= 0 || fontSize > 1e8) throw new LayoutUnsupported("字号无效。");
        string requested = style.TryGetValue("font-family", out string? familyValue) ? familyValue : "Arial";
        FontFamily family = ResolveFamily(requested);
        string weight = style.TryGetValue("font-weight", out string? weightValue) ? weightValue : "normal";
        bool bold = weight is "bold" or "bolder" || double.TryParse(weight, NumberStyles.Float, CultureInfo.InvariantCulture, out double numericWeight) && numericWeight >= 600;
        bool italic = style.TryGetValue("font-style", out string? fontStyle) && fontStyle is "italic" or "oblique";
        FontStyle selectedStyle = (bold ? FontStyle.Bold : FontStyle.Regular) | (italic ? FontStyle.Italic : FontStyle.Regular);
        Font font = family.CreateFont(ReferenceSize, selectedStyle);
        string baseline = style.TryGetValue("dominant-baseline", out string? baselineValue) ? baselineValue : "alphabetic";
        TextBaseline textBaseline = baseline switch
        {
            "alphabetic" or "auto" => TextBaseline.Alphabetic,
            "hanging" => TextBaseline.Hanging,
            "central" => TextBaseline.Central,
            "middle" => TextBaseline.Middle,
            "text-before-edge" or "text-top" => TextBaseline.TextTop,
            _ => throw new LayoutUnsupported("文字基线类型暂不支持。")
        };
        var options = new TextOptions(font) { Dpi = 72, TextBaseline = textBaseline, Origin = Vector2.Zero, WrappingLength = -1 };
        var fallbacks = new List<FontFamily>();
        foreach (string fallbackName in new[] { "Microsoft YaHei", "SimSun", "Arial Unicode MS", "Arial" })
            if (SystemFonts.TryGet(fallbackName, out FontFamily fallback) && fallback.Name != family.Name) fallbacks.Add(fallback);
        options.FallbackFontFamilies = fallbacks;
        FontRectangle advance = TextMeasurer.MeasureAdvance(value, options);
        FontRectangle ink = TextMeasurer.MeasureBounds(value, options);
        double unit = fontSize / ReferenceSize;
        double naturalWidth = advance.Width * unit;
        if (!double.IsFinite(naturalWidth) || naturalWidth <= 0 || ink.Width <= 0 || ink.Height <= 0) throw new LayoutUnsupported("无法测量文字尺寸。");
        double width = naturalWidth;
        if (style.TryGetValue("letter-spacing", out string? spacing) && spacing != "normal") width += Scalar(spacing) * Math.Max(0, value.EnumerateRunes().Count() - 1);
        XAttribute? textLength = element.Attribute("textLength") ?? element.Parent?.Attribute("textLength");
        if (textLength != null) width = Scalar(textLength.Value);
        if (width <= 0 || !double.IsFinite(width)) throw new LayoutUnsupported("文字宽度无效。");
        double widthFactor = width / naturalWidth;
        string anchor = style.TryGetValue("text-anchor", out string? anchorValue) ? anchorValue : "start";
        double anchorOffset = anchor switch { "start" => 0d, "middle" => width / 2d, "end" => width, _ => throw new LayoutUnsupported("文字水平对齐类型暂不支持。") };
        double left = ink.X * unit * widthFactor - anchorOffset;
        double top = ink.Y * unit;
        double inkWidth = ink.Width * unit * widthFactor;
        double inkHeight = ink.Height * unit;
        (double originX, double originY) = matrix.Apply(x, y);
        (double centerX, double centerY) = matrix.Apply(x + left + inkWidth / 2d, y + top + inkHeight / 2d);
        string color = style.TryGetValue("fill", out string? fill) ? NormalizeColor(fill, style) : "#000000";
        return new EditableTextItem
        {
            SourceId = (string?)element.AncestorsAndSelf().First(item => item.Name == Svg + "text").Attribute("data-ppt-text-id") ?? "",
            Text = value, X = originX, Y = originY,
            CenterX = centerX, CenterY = centerY,
            Width = width * matrix.ScaleX,
            Height = Math.Max(fontSize, advance.Height * unit) * matrix.ScaleY,
            InkWidth = inkWidth * matrix.ScaleX, InkHeight = inkHeight * matrix.ScaleY,
            FontSize = fontSize * matrix.ScaleY, FontFamily = family.Name,
            Rotation = Math.Atan2(matrix.B, matrix.A) * 180d / Math.PI,
            Color = color, Bold = bold, Italic = italic
        };
    }

    private static FontFamily ResolveFamily(string requested)
    {
        foreach (string value in requested.Split(','))
        {
            string name = value.Trim().Trim('\'', '"');
            if (SystemFonts.TryGet(name, out FontFamily family)) return family;
        }
        foreach (string fallback in new[] { "Arial", "Microsoft YaHei", "SimSun" })
            if (SystemFonts.TryGet(fallback, out FontFamily family)) return family;
        throw new LayoutUnsupported("系统没有可用的替代字体。");
    }

    private static string NormalizeColor(string value, Dictionary<string, string> style)
    {
        if (value == "currentColor") value = style.TryGetValue("color", out string? color) ? color : "#000000";
        if (Regex.IsMatch(value, "^#[0-9a-fA-F]{6}$")) return value.ToUpperInvariant();
        if (Regex.IsMatch(value, "^#[0-9a-fA-F]{3}$")) return "#" + string.Concat(value.Skip(1).Select(character => new string(character, 2))).ToUpperInvariant();
        if (value == "black") return "#000000";
        if (value == "white") return "#FFFFFF";
        throw new LayoutUnsupported("文字填充颜色或效果暂不支持。");
    }

    private static void ApplyStyle(Dictionary<string, string> style, XElement element)
    {
        foreach (XAttribute attribute in element.Attributes())
            if (attribute.Name.NamespaceName.Length == 0) style[attribute.Name.LocalName] = attribute.Value;
        string? css = (string?)element.Attribute("style");
        if (css == null) return;
        foreach (string declaration in css.Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            int separator = declaration.IndexOf(':');
            if (separator < 1) throw new LayoutUnsupported("文字样式无效。");
            style[declaration[..separator].Trim()] = declaration[(separator + 1)..].Trim();
        }
    }

    private static string TextValue(XElement element)
    {
        string value = element.Value;
        bool preserve = element.AncestorsAndSelf().Select(item => (string?)item.Attribute(XNamespace.Xml + "space")).FirstOrDefault(item => item != null) == "preserve";
        return preserve ? value : Regex.Replace(value, @"\s+", " ");
    }

    private static double Coordinate(XElement element, string name, double fallback)
    {
        XAttribute? attribute = element.Attribute(name);
        return attribute == null ? fallback : Scalar(attribute.Value);
    }

    private static double Scalar(string value)
    {
        value = value.Trim();
        if (value.EndsWith("px", StringComparison.Ordinal)) value = value[..^2];
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) || !double.IsFinite(result))
            throw new LayoutUnsupported("存在无法还原的坐标或样式数值。");
        return result;
    }

    private static Matrix ParseTransform(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return Matrix.Identity;
        Matrix matrix = Matrix.Identity;
        int end = 0;
        foreach (Match part in TransformPart.Matches(value))
        {
            if (!SeparatorsOnly(value[end..part.Index])) throw new LayoutUnsupported("SVG 变换无效。");
            string arguments = part.Groups[2].Value;
            MatchCollection matches = Numbers.Matches(arguments);
            int numberEnd = 0;
            var values = new List<double>();
            foreach (Match number in matches)
            {
                if (!SeparatorsOnly(arguments[numberEnd..number.Index])) throw new LayoutUnsupported("SVG 变换参数无效。");
                values.Add(Scalar(number.Value));
                numberEnd = number.Index + number.Length;
            }
            if (!SeparatorsOnly(arguments[numberEnd..])) throw new LayoutUnsupported("SVG 变换参数无效。");
            Matrix local;
            switch (part.Groups[1].Value)
            {
                case "matrix" when values.Count == 6:
                    local = new Matrix(values[0], values[1], values[2], values[3], values[4], values[5]); break;
                case "translate" when values.Count is 1 or 2:
                    local = new Matrix(1, 0, 0, 1, values[0], values.Count == 2 ? values[1] : 0); break;
                case "scale" when values.Count is 1 or 2:
                    local = new Matrix(values[0], 0, 0, values.Count == 2 ? values[1] : values[0], 0, 0); break;
                case "rotate" when values.Count is 1 or 3:
                    double radians = values[0] * Math.PI / 180d;
                    local = new Matrix(Math.Cos(radians), Math.Sin(radians), -Math.Sin(radians), Math.Cos(radians), 0, 0);
                    if (values.Count == 3) local = new Matrix(1, 0, 0, 1, values[1], values[2]) * local * new Matrix(1, 0, 0, 1, -values[1], -values[2]);
                    break;
                default: throw new LayoutUnsupported("包含倾斜或未知 SVG 变换。");
            }
            matrix *= local;
            end = part.Index + part.Length;
        }
        if (!SeparatorsOnly(value[end..])) throw new LayoutUnsupported("SVG 变换无效。");
        return matrix;
    }

    private static bool SeparatorsOnly(string value) => value.All(character => char.IsWhiteSpace(character) || character == ',');

    private static void EnsureOrthogonal(Matrix matrix)
    {
        double x = matrix.ScaleX, y = matrix.ScaleY;
        double determinant = matrix.A * matrix.D - matrix.B * matrix.C;
        double dot = matrix.A * matrix.C + matrix.B * matrix.D;
        if (!matrix.IsFinite || !double.IsFinite(x) || !double.IsFinite(y) || !double.IsFinite(determinant) || !double.IsFinite(dot) ||
            x < 1e-10 || y < 1e-10 || determinant <= 0 || Math.Abs(dot) > x * y * 1e-7)
            throw new LayoutUnsupported("包含剪切、反射或退化变换。");
    }

    private static void ExpandViewBox(XElement root, List<EditableTextItem> texts)
    {
        string[] values = ((string?)root.Attribute("viewBox") ?? "").Split(new[] { ' ', ',', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (values.Length != 4 || texts.Count == 0) return;
        double[] bounds = values.Select(Scalar).ToArray();
        double left = bounds[0], top = bounds[1], right = left + bounds[2], bottom = top + bounds[3];
        foreach (EditableTextItem text in texts)
        {
            double angle = text.Rotation * Math.PI / 180d;
            double halfWidth = (Math.Abs(Math.Cos(angle)) * text.InkWidth + Math.Abs(Math.Sin(angle)) * text.InkHeight) / 2d;
            double halfHeight = (Math.Abs(Math.Sin(angle)) * text.InkWidth + Math.Abs(Math.Cos(angle)) * text.InkHeight) / 2d;
            left = Math.Min(left, text.CenterX - halfWidth); top = Math.Min(top, text.CenterY - halfHeight);
            right = Math.Max(right, text.CenterX + halfWidth); bottom = Math.Max(bottom, text.CenterY + halfHeight);
        }
        root.SetAttributeValue("viewBox", $"{Number(left)} {Number(top)} {Number(right - left)} {Number(bottom - top)}");
        double scale = 1600d / Math.Max(right - left, bottom - top);
        root.SetAttributeValue("width", Number((right - left) * scale));
        root.SetAttributeValue("height", Number((bottom - top) * scale));
    }

    private static string Number(double value) => value.ToString("0.#########", CultureInfo.InvariantCulture);

    private sealed class LayoutUnsupported(string message) : Exception(message);

    private readonly record struct Matrix(double A, double B, double C, double D, double E, double F)
    {
        internal static Matrix Identity => new(1, 0, 0, 1, 0, 0);
        internal double ScaleX => Math.Sqrt(A * A + B * B);
        internal double ScaleY => Math.Sqrt(C * C + D * D);
        internal bool IsFinite => new[] { A, B, C, D, E, F }.All(double.IsFinite);
        internal (double X, double Y) Apply(double x, double y) => (A * x + C * y + E, B * x + D * y + F);
        public static Matrix operator *(Matrix left, Matrix right) => new(
            left.A * right.A + left.C * right.B, left.B * right.A + left.D * right.B,
            left.A * right.C + left.C * right.D, left.B * right.C + left.D * right.D,
            left.A * right.E + left.C * right.F + left.E, left.B * right.E + left.D * right.F + left.F);
    }
}
