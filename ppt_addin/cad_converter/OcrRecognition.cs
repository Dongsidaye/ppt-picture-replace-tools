using System.Globalization;
using System.Text.Json.Serialization;
using System.Xml;
using System.Xml.Linq;
using Microsoft.ML.OnnxRuntime;
using RapidOcrNet;
using SkiaSharp;
using Svg.Skia;

namespace Dongsidaye.CadConverter;

internal sealed record OcrPoint(
    [property: JsonPropertyName("x")] double X,
    [property: JsonPropertyName("y")] double Y);

internal sealed record OcrLine(
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("confidence")] double Confidence,
    [property: JsonPropertyName("points")] IReadOnlyList<OcrPoint> Points,
    [property: JsonPropertyName("color")] string Color,
    [property: JsonPropertyName("inkPoints")] IReadOnlyList<OcrPoint>? InkPoints = null);

internal sealed record OcrPage(
    [property: JsonPropertyName("width")] int Width,
    [property: JsonPropertyName("height")] int Height,
    [property: JsonPropertyName("viewBox")] double[] ViewBox,
    [property: JsonPropertyName("lines")] IReadOnlyList<OcrLine> Lines);

/// <summary>Optional offline OCR of geometric letter outlines in a converted drawing.</summary>
internal static class OcrRecognition
{
    private const long MaximumSvgBytes = 16L * 1024 * 1024;
    private const long MaximumPixels = 32L * 1024 * 1024;
    private const int MaximumSide = 4096;
    private const int TileSide = 1536;
    private const int TileOverlap = 192;
    private const int MaximumLines = 5000;
    private static readonly XNamespace SvgNamespace = "http://www.w3.org/2000/svg";

    internal static OcrPage ReadSvg(string svgPath, string modelsFolder, CancellationToken cancellationToken = default)
    {
        XDocument document = ReadGeometry(svgPath);
        double[] viewBox = ReadViewBox(document.Root!);
        using SKBitmap image = Rasterize(document, cancellationToken);
        IReadOnlyList<OcrLine> lines = ReadImage(image, modelsFolder, cancellationToken);
        return new OcrPage(image.Width, image.Height, viewBox, lines);
    }

    private static XDocument ReadGeometry(string path)
    {
        FileInfo input = new(path);
        if (!input.Exists || input.Length < 32 || input.Length > MaximumSvgBytes)
        {
            throw new InvalidDataException("OCR SVG is missing or exceeds the size limit.");
        }
        XmlReaderSettings settings = new()
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaximumSvgBytes
        };
        using XmlReader reader = XmlReader.Create(path, settings);
        XDocument document = XDocument.Load(reader);
        XElement root = document.Root ?? throw new InvalidDataException("OCR SVG has no root element.");
        if (root.Name != SvgNamespace + "svg")
        {
            throw new InvalidDataException("OCR input must be an SVG drawing.");
        }

        // Existing text remains editable through the direct text path. OCR is
        // restricted to geometry, including geometry left by exploded letters.
        root.Descendants(SvgNamespace + "text").Remove();
        HashSet<string> allowed = new(StringComparer.Ordinal)
        {
            "svg", "g", "defs", "path", "line", "circle", "ellipse", "rect",
            "polyline", "polygon", "use", "clipPath", "pattern", "title", "desc",
            "mask", "linearGradient", "radialGradient", "stop"
        };
        int elements = 0;
        foreach (XElement element in root.DescendantsAndSelf())
        {
            if (++elements > 250000 || element.Name.Namespace != SvgNamespace || !allowed.Contains(element.Name.LocalName))
            {
                throw new InvalidDataException("OCR SVG contains unsupported content.");
            }
            foreach (XAttribute attribute in element.Attributes())
            {
                string name = attribute.Name.LocalName;
                string value = attribute.Value.Trim();
                if (attribute.Name == XNamespace.Xml + "base" || name.StartsWith("on", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("OCR SVG contains an active reference.");
                }
                if (name == "href" && value.Length > 0 && !value.StartsWith('#'))
                {
                    throw new InvalidDataException("OCR SVG may only reference local geometry.");
                }
                if (value.Contains("url(", StringComparison.OrdinalIgnoreCase) &&
                    !System.Text.RegularExpressions.Regex.IsMatch(value, "^url\\(\\s*['\" ]?#[a-zA-Z0-9_\\-:.]+['\" ]?\\s*\\)$"))
                {
                    throw new InvalidDataException("OCR SVG contains an external style reference.");
                }
            }
        }
        return document;
    }

    private static double[] ReadViewBox(XElement root)
    {
        string[] values = ((string?)root.Attribute("viewBox") ?? "")
            .Split([' ', ',', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        if (values.Length != 4)
        {
            throw new InvalidDataException("OCR SVG has no valid drawing extents.");
        }
        double[] bounds = values.Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        if (bounds.Any(value => !double.IsFinite(value)) || bounds[2] <= 0 || bounds[3] <= 0)
        {
            throw new InvalidDataException("OCR SVG has invalid drawing extents.");
        }
        return bounds;
    }

    private static SKBitmap Rasterize(XDocument document, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using SKSvg svg = new();
        using MemoryStream content = new();
        document.Save(content, SaveOptions.DisableFormatting);
        content.Position = 0;
        SKPicture picture = svg.Load(content) ?? throw new InvalidDataException("OCR SVG could not be rendered.");
        SKRect bounds = picture.CullRect;
        if (!float.IsFinite(bounds.Width) || !float.IsFinite(bounds.Height) || bounds.Width <= 0 || bounds.Height <= 0)
        {
            throw new InvalidDataException("OCR SVG has invalid rendering dimensions.");
        }
        double scale = MaximumSide / Math.Max(bounds.Width, bounds.Height);
        int width = Math.Clamp((int)Math.Ceiling(bounds.Width * scale), 1, MaximumSide);
        int height = Math.Clamp((int)Math.Ceiling(bounds.Height * scale), 1, MaximumSide);
        if ((long)width * height > MaximumPixels)
        {
            throw new InvalidDataException("OCR drawing exceeds the pixel limit.");
        }
        SKBitmap bitmap = new(width, height, SKColorType.Bgra8888, SKAlphaType.Premul);
        try
        {
            using SKCanvas canvas = new(bitmap);
            canvas.Clear(SKColors.White);
            canvas.Scale((float)scale);
            canvas.Translate(-bounds.Left, -bounds.Top);
            canvas.DrawPicture(picture);
            cancellationToken.ThrowIfCancellationRequested();
            return bitmap;
        }
        catch
        {
            bitmap.Dispose();
            throw;
        }
    }

    private static IReadOnlyList<OcrLine> ReadImage(SKBitmap image, string modelsFolder, CancellationToken cancellationToken)
    {
        string detector = Model(modelsFolder, "ch_PP-OCRv5_mobile_det.onnx");
        string classifier = Model(modelsFolder, "ch_PP-LCNet_x0_25_textline_ori_cls_mobile.onnx");
        string recognizer = Model(modelsFolder, "ch_PP-OCRv5_rec_mobile.onnx");
        string dictionary = Model(modelsFolder, "ppocrv5_dict.txt");
        // Each short-lived import uses several tile and recognition shapes.
        // Avoid retaining a separate pooled allocation pattern for those shapes.
        using SessionOptions session = new()
        {
            IntraOpNumThreads = Math.Max(1, Math.Min(4, Environment.ProcessorCount)),
            InterOpNumThreads = 1,
            EnableCpuMemArena = false,
            EnableMemoryPattern = false,
            ExecutionMode = ExecutionMode.ORT_SEQUENTIAL
        };
        using RapidOcr engine = new();
        engine.InitModels(detector, classifier, recognizer, dictionary, session);
        RapidOcrOptions options = RapidOcrOptions.Default with
        {
            ImgResize = TileSide,
            TextScore = 0.5f,
            ReturnWordBox = false
        };
        List<OcrLine> result = [];
        // Start with complete lines. Detail tiles can split a long line at a
        // border, so a containing whole-page result wins over those fragments.
        OcrResult fullPage = engine.DetectAsync(image, options, cancellationToken: cancellationToken).GetAwaiter().GetResult();
        AddDetections(fullPage, result, image, 0, 0);
        if (image.Width <= TileSide && image.Height <= TileSide)
        {
            return result;
        }
        foreach (int top in TileOrigins(image.Height))
        {
            foreach (int left in TileOrigins(image.Width))
            {
                cancellationToken.ThrowIfCancellationRequested();
                SKRectI bounds = new(left, top, Math.Min(image.Width, left + TileSide), Math.Min(image.Height, top + TileSide));
                using SKBitmap tile = new(bounds.Width, bounds.Height);
                if (!image.ExtractSubset(tile, bounds))
                {
                    throw new InvalidDataException("OCR drawing tile could not be read.");
                }
                OcrResult detected = engine.DetectAsync(tile, options, cancellationToken: cancellationToken).GetAwaiter().GetResult();
                AddDetections(detected, result, image, left, top);
            }
        }
        return result.OrderBy(line => line.Points.Min(point => point.Y)).ThenBy(line => line.Points.Min(point => point.X)).ToArray();
    }

    private static void AddDetections(OcrResult detected, List<OcrLine> result, SKBitmap image, int left, int top)
    {
        foreach (TextBlock block in detected.TextBlocks)
        {
            if (string.IsNullOrWhiteSpace(block.Text) || block.BoxPoints.Length != 4)
            {
                continue;
            }
            double confidence = block.CharScores?.Average() ?? 0;
            if (!double.IsFinite(confidence) || confidence < 0.5)
            {
                continue;
            }
            OcrPoint[] points = ReadingOrderPoints(block).Select(point => new OcrPoint(
                Math.Clamp(point.X + left, 0, image.Width), Math.Clamp(point.Y + top, 0, image.Height))).ToArray();
            OcrLine line = new(block.Text.Trim(), Math.Clamp(confidence, 0, 1), points,
                ReadInkColor(image, points), ReadInkBounds(image, points));
            AddDistinct(result, line);
            if (result.Count > MaximumLines)
            {
                throw new InvalidDataException("OCR found too many text regions. Import a smaller drawing area.");
            }
        }
    }

    private static IReadOnlyList<OcrPoint> ReadInkBounds(SKBitmap image, IReadOnlyList<OcrPoint> polygon)
    {
        double deltaX = polygon[1].X - polygon[0].X;
        double deltaY = polygon[1].Y - polygon[0].Y;
        double length = Math.Sqrt(deltaX * deltaX + deltaY * deltaY);
        if (length < 1)
        {
            return polygon;
        }
        double ux = deltaX / length, uy = deltaY / length;
        double vx = -uy, vy = ux;
        if ((polygon[3].X - polygon[0].X) * vx + (polygon[3].Y - polygon[0].Y) * vy < 0)
        {
            vx = -vx;
            vy = -vy;
        }
        int left = Math.Max(0, (int)Math.Floor(polygon.Min(point => point.X)));
        int top = Math.Max(0, (int)Math.Floor(polygon.Min(point => point.Y)));
        int right = Math.Min(image.Width, (int)Math.Ceiling(polygon.Max(point => point.X)));
        int bottom = Math.Min(image.Height, (int)Math.Ceiling(polygon.Max(point => point.Y)));
        double area = (double)(right - left) * (bottom - top);
        int step = Math.Max(1, (int)Math.Ceiling(Math.Sqrt(area / 200000)));
        double minimumU = double.PositiveInfinity, maximumU = double.NegativeInfinity;
        double minimumV = double.PositiveInfinity, maximumV = double.NegativeInfinity;
        for (int y = top; y < bottom; y += step)
        {
            for (int x = left; x < right; x += step)
            {
                double centerX = x + 0.5, centerY = y + 0.5;
                if (!InsidePolygon(centerX, centerY, polygon))
                {
                    continue;
                }
                SKColor color = image.GetPixel(x, y);
                if (color.Alpha < 40 || (color.Red > 244 && color.Green > 244 && color.Blue > 244))
                {
                    continue;
                }
                double relativeX = centerX - polygon[0].X, relativeY = centerY - polygon[0].Y;
                double u = relativeX * ux + relativeY * uy;
                double v = relativeX * vx + relativeY * vy;
                minimumU = Math.Min(minimumU, u);
                maximumU = Math.Max(maximumU, u);
                minimumV = Math.Min(minimumV, v);
                maximumV = Math.Max(maximumV, v);
            }
        }
        if (!double.IsFinite(minimumU) || maximumU - minimumU < 1 || maximumV - minimumV < 1)
        {
            return polygon;
        }
        // Sampling can land just inside the outermost antialiased stroke.
        // Include one sample step so the tight box does not trim that stroke.
        double padding = step;
        minimumU -= padding;
        maximumU += padding;
        minimumV -= padding;
        maximumV += padding;
        OcrPoint Map(double u, double v) => new(
            Math.Clamp(polygon[0].X + u * ux + v * vx, 0, image.Width),
            Math.Clamp(polygon[0].Y + u * uy + v * vy, 0, image.Height));
        return [Map(minimumU, minimumV), Map(maximumU, minimumV), Map(maximumU, maximumV), Map(minimumU, maximumV)];
    }

    private static IEnumerable<SKPointI> ReadingOrderPoints(TextBlock block)
    {
        SKPointI[] points = block.BoxPoints;
        double width = Math.Sqrt(Math.Pow(points[1].X - points[0].X, 2) + Math.Pow(points[1].Y - points[0].Y, 2));
        double height = Math.Sqrt(Math.Pow(points[3].X - points[0].X, 2) + Math.Pow(points[3].Y - points[0].Y, 2));
        // RapidOcrNet 4.2.0 OcrUtils rotates tall rectified crops clockwise
        // when their integer height is at least 1.5 times their integer width.
        // Undo that normalization in the output polygon's reading direction.
        int first = (int)height >= (int)width * 1.5 ? 3 : 0;
        // DetectOnce resets AngleIndex to zero whenever a proposed 180-degree
        // flip is below ClsThresh. Index 1 therefore records an applied flip,
        // without duplicating or guessing the classifier confidence threshold.
        if (block.AngleIndex == 1)
        {
            first = (first + 2) % points.Length;
        }
        for (int i = 0; i < points.Length; i++)
        {
            yield return points[(first + i) % points.Length];
        }
    }

    private static string ReadInkColor(SKBitmap image, IReadOnlyList<OcrPoint> points)
    {
        int left = Math.Max(0, (int)Math.Floor(points.Min(point => point.X)));
        int top = Math.Max(0, (int)Math.Floor(points.Min(point => point.Y)));
        int right = Math.Min(image.Width, (int)Math.Ceiling(points.Max(point => point.X)));
        int bottom = Math.Min(image.Height, (int)Math.Ceiling(points.Max(point => point.Y)));
        double area = (double)(right - left) * (bottom - top);
        int step = Math.Max(1, (int)Math.Ceiling(Math.Sqrt(area / 4096)));
        Dictionary<int, (int Count, long Red, long Green, long Blue)> colors = [];
        for (int y = top; y < bottom; y += step)
        {
            for (int x = left; x < right; x += step)
            {
                if (!InsidePolygon(x + 0.5, y + 0.5, points))
                {
                    continue;
                }
                SKColor color = image.GetPixel(x, y);
                if (color.Alpha < 40 || (color.Red > 244 && color.Green > 244 && color.Blue > 244))
                {
                    continue;
                }
                int key = ((color.Red >> 5) << 6) | ((color.Green >> 5) << 3) | (color.Blue >> 5);
                colors.TryGetValue(key, out var value);
                colors[key] = (value.Count + 1, value.Red + color.Red, value.Green + color.Green, value.Blue + color.Blue);
            }
        }
        if (colors.Count == 0)
        {
            return "#000000";
        }
        var dominant = colors.Values.MaxBy(value => value.Count);
        return $"#{dominant.Red / dominant.Count:X2}{dominant.Green / dominant.Count:X2}{dominant.Blue / dominant.Count:X2}";
    }

    private static bool InsidePolygon(double x, double y, IReadOnlyList<OcrPoint> polygon)
    {
        bool positive = false, negative = false;
        for (int i = 0; i < polygon.Count; i++)
        {
            OcrPoint first = polygon[i], second = polygon[(i + 1) % polygon.Count];
            double cross = (second.X - first.X) * (y - first.Y) - (second.Y - first.Y) * (x - first.X);
            positive |= cross > 0;
            negative |= cross < 0;
        }
        return !positive || !negative;
    }

    private static string Model(string folder, string file)
    {
        string path = Path.Combine(folder, file);
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("OCR model is missing. Reinstall the add-in.", file);
        }
        return path;
    }

    private static IEnumerable<int> TileOrigins(int length)
    {
        int last = Math.Max(0, length - TileSide);
        int previous = -1;
        for (int position = 0; ; position += TileSide - TileOverlap)
        {
            int start = Math.Min(position, last);
            if (start != previous)
            {
                yield return start;
                previous = start;
            }
            if (start == last)
            {
                yield break;
            }
        }
    }

    private static void AddDistinct(List<OcrLine> lines, OcrLine candidate)
    {
        for (int i = 0; i < lines.Count; i++)
        {
            OcrLine existing = lines[i];
            (double overlap, double existingArea, double candidateArea) = Overlap(existing, candidate);
            if (overlap < 0.7)
            {
                continue;
            }
            if (candidateArea > existingArea * 1.25 ||
                (candidateArea >= existingArea * 0.8 && candidate.Confidence > existing.Confidence))
            {
                lines[i] = candidate;
            }
            return;
        }
        lines.Add(candidate);
    }

    private static (double Ratio, double FirstArea, double SecondArea) Overlap(OcrLine first, OcrLine second)
    {
        double firstLeft = first.Points.Min(point => point.X), firstTop = first.Points.Min(point => point.Y);
        double firstRight = first.Points.Max(point => point.X), firstBottom = first.Points.Max(point => point.Y);
        double secondLeft = second.Points.Min(point => point.X), secondTop = second.Points.Min(point => point.Y);
        double secondRight = second.Points.Max(point => point.X), secondBottom = second.Points.Max(point => point.Y);
        double intersection = Math.Max(0, Math.Min(firstRight, secondRight) - Math.Max(firstLeft, secondLeft)) *
            Math.Max(0, Math.Min(firstBottom, secondBottom) - Math.Max(firstTop, secondTop));
        double firstArea = (firstRight - firstLeft) * (firstBottom - firstTop);
        double secondArea = (secondRight - secondLeft) * (secondBottom - secondTop);
        double minimum = Math.Min(firstArea, secondArea);
        return (minimum > 0 ? intersection / minimum : 0, firstArea, secondArea);
    }
}
