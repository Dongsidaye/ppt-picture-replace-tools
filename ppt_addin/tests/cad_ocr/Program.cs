using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Dongsidaye.CadConverter;
using SkiaSharp;

namespace CadOcrSmoke;

internal static class Program
{
    private sealed record Check(string Name, bool Passed, string Error = "");

    private static int Main(string[] args)
    {
        string? folder = null;
        bool temporary = args.Length == 1;
        try
        {
            if (args.Length is < 1 or > 2)
            {
                Console.Error.WriteLine("Usage: OcrSmoke <models-folder> [new-evidence-folder]");
                return 2;
            }
            string models = Path.GetFullPath(args[0]);
            folder = temporary
                ? Path.Combine(Path.GetTempPath(), "ppt-cad-ocr-test-" + Guid.NewGuid().ToString("N"))
                : Path.GetFullPath(args[1]);
            if (Directory.Exists(folder))
            {
                throw new IOException("Use a new evidence folder so existing files remain intact.");
            }
            Directory.CreateDirectory(folder);

            string drawing = Path.Combine(folder, "letter-outlines.svg");
            WriteOutlineFixture(drawing);
            byte[] sourceHash = SHA256.HashData(File.ReadAllBytes(drawing));
            OcrPage page = OcrRecognition.ReadSvg(drawing, models);
            List<Check> checks = [];
            string text = string.Join(" ", page.Lines.Select(line => line.Text));
            string normalized = string.Concat(text.Where(character => !char.IsWhiteSpace(character)));

            CheckResult(checks, "Chinese character outlines", () =>
            {
                Require(normalized.Contains("中文文字识别", StringComparison.Ordinal), "The Chinese outline text was not recognized.");
                Require(normalized.Contains("图纸标注", StringComparison.Ordinal), "The second Chinese outline label was not recognized.");
            });
            CheckResult(checks, "Latin character outlines", () =>
            {
                Require(normalized.Contains("DWG", StringComparison.Ordinal), "DWG was not recognized.");
                Require(normalized.Contains("ABC", StringComparison.Ordinal), "ABC was not recognized.");
            });
            CheckResult(checks, "Numeric character outlines", () =>
            {
                Require(normalized.Contains("123", StringComparison.Ordinal), "123 was not recognized.");
                Require(normalized.Contains("456", StringComparison.Ordinal), "456 was not recognized.");
            });
            CheckResult(checks, "Red outline text retains its ink color", () =>
            {
                OcrLine[] redLabels = page.Lines.Where(line => line.Text.Contains("图纸", StringComparison.Ordinal)).ToArray();
                Require(redLabels.Length > 0 && redLabels.All(line => IsRed(line.Color)), "Red label color was not recovered.");
            });
            CheckResult(checks, "Source drawing and coordinate bounds", () =>
            {
                Require(sourceHash.SequenceEqual(SHA256.HashData(File.ReadAllBytes(drawing))), "OCR changed its input SVG.");
                Require(page.Width <= 4096 && page.Height <= 4096, "Rasterization exceeded the side limit.");
                Require(page.ViewBox.SequenceEqual(new double[] { 0, 0, 1800, 380 }), "Drawing extents changed.");
                Require(page.Lines.All(line => line.Points.Count == 4 && line.Points.All(point =>
                    point.X >= 0 && point.X <= page.Width && point.Y >= 0 && point.Y <= page.Height)), "A detected polygon is outside the raster.");
                Require(page.Lines.All(line => line.Confidence >= 0.5 && line.Confidence <= 1), "A line has invalid confidence.");
            });
            CheckResult(checks, "Ink bounds exclude detector padding", () =>
            {
                foreach (OcrLine line in page.Lines)
                {
                    CheckInkBounds(line);
                }
            });
            CheckResult(checks, "Cancellation is observed", () =>
            {
                using CancellationTokenSource cancelled = new();
                cancelled.Cancel();
                try
                {
                    OcrRecognition.ReadSvg(drawing, models, cancelled.Token);
                }
                catch (OperationCanceledException)
                {
                    return;
                }
                throw new InvalidOperationException("A cancelled OCR operation completed.");
            });
            CheckResult(checks, "External SVG resources are rejected", () =>
            {
                string external = Path.Combine(folder, "external-resource.svg");
                File.WriteAllText(external,
                    "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 100 100\"><image href=\"https://example.com/image.png\"/></svg>");
                try
                {
                    OcrRecognition.ReadSvg(external, models);
                }
                catch (InvalidDataException)
                {
                    return;
                }
                throw new InvalidOperationException("An external SVG resource was accepted.");
            });
            List<object> orientations = [];
            foreach (int expected in new[] { 180, 90 })
            {
                CheckResult(checks, $"{expected}-degree character outlines retain their reading direction", () =>
                {
                    string rotated = Path.Combine(folder, $"letter-outlines-{expected}.svg");
                    WriteOutlineFixture(rotated, expected);
                    OcrPage rotatedPage = OcrRecognition.ReadSvg(rotated, models);
                    OcrLine line = rotatedPage.Lines.First(candidate =>
                        string.Concat(candidate.Text.Where(character => !char.IsWhiteSpace(character))) == "中文文字识别DWG123");
                    double actual = Math.Atan2(line.Points[1].Y - line.Points[0].Y, line.Points[1].X - line.Points[0].X) * 180 / Math.PI;
                    orientations.Add(new { expected, actual, line.Text, line.Points, line.InkPoints });
                    double difference = Math.Abs(((actual - expected + 540) % 360) - 180);
                    Require(difference < 3, $"Expected {expected} degrees but received {actual:0.###}.");
                    CheckInkBounds(line);
                });
            }

            var report = new
            {
                passed = checks.Count(check => check.Passed),
                failed = checks.Count(check => !check.Passed),
                checks,
                ocr = page,
                orientations,
                note = "Recognition checks ignore whitespace because OCR word segmentation can change spaces."
            };
            string json = JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(Path.Combine(folder, "report.json"), json);
            Console.WriteLine(json);
            return checks.All(check => check.Passed) ? 0 : 1;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error.GetBaseException().Message);
            return 1;
        }
        finally
        {
            if (temporary && folder != null)
            {
                string root = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                string target = Path.GetFullPath(folder);
                if (target.StartsWith(root, StringComparison.OrdinalIgnoreCase) && Path.GetFileName(target).StartsWith("ppt-cad-ocr-test-", StringComparison.Ordinal))
                {
                    try { Directory.Delete(target, recursive: true); } catch (IOException) { }
                }
            }
        }
    }

    private static void WriteOutlineFixture(string output, int rotation = 0)
    {
        using SKTypeface typeface = SKTypeface.FromFamilyName("Microsoft YaHei");
        if (!typeface.FamilyName.Contains("YaHei", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("The Windows OCR fixture requires the Microsoft YaHei font.");
        }
        using SKFont font = new(typeface, 90);
        using SKPath first = font.GetTextPath("中文文字识别 DWG 123", new SKPoint(45, 150));
        using SKPath second = font.GetTextPath("图纸标注 ABC 456", new SKPoint(45, 290));
        string viewBox = rotation == 90 ? "0 0 380 1800" : "0 0 1800 380";
        string transform = rotation switch
        {
            180 => "translate(1800 380) rotate(180)",
            90 => "translate(380 0) rotate(90)",
            _ => "translate(0 0)"
        };
        string svg = $"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"{viewBox}\"><g transform=\"{transform}\"><path d=\"{first.ToSvgPathData()}\"/><path fill=\"#CC2222\" d=\"{second.ToSvgPathData()}\"/></g></svg>";
        Require(!svg.Contains("<text", StringComparison.OrdinalIgnoreCase), "The fixture must contain outlines rather than SVG text.");
        File.WriteAllText(output, svg);
    }

    private static void CheckResult(List<Check> results, string name, Action work)
    {
        try
        {
            work();
            results.Add(new Check(name, true));
            Console.WriteLine("PASS " + name);
        }
        catch (Exception error)
        {
            results.Add(new Check(name, false, error.Message));
            Console.WriteLine("FAIL " + name + " - " + error.Message);
        }
    }

    private static bool IsRed(string color)
    {
        if (color.Length != 7 || color[0] != '#')
        {
            return false;
        }
        byte red = byte.Parse(color.AsSpan(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        byte green = byte.Parse(color.AsSpan(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        byte blue = byte.Parse(color.AsSpan(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        return red > 140 && green < 80 && blue < 80 && red > green * 2;
    }

    private static void CheckInkBounds(OcrLine line)
    {
        Require(line.InkPoints is { Count: 4 }, "No tight ink polygon was returned.");
        IReadOnlyList<OcrPoint> ink = line.InkPoints!;
        double inkWidth = Distance(ink[0], ink[1]), inkHeight = Distance(ink[0], ink[3]);
        double detectionWidth = Distance(line.Points[0], line.Points[1]), detectionHeight = Distance(line.Points[0], line.Points[3]);
        Require(inkWidth > 0 && inkHeight > 0, "The tight ink bounds are empty.");
        Require(inkHeight < detectionHeight * 0.95, "The tight ink box still includes the detector's vertical padding.");
        Require(inkWidth <= detectionWidth * 1.01, "The tight ink box enlarged the line width.");
        double detectionAngle = Math.Atan2(line.Points[1].Y - line.Points[0].Y, line.Points[1].X - line.Points[0].X);
        double inkAngle = Math.Atan2(ink[1].Y - ink[0].Y, ink[1].X - ink[0].X);
        Require(Math.Abs(Math.Atan2(Math.Sin(inkAngle - detectionAngle), Math.Cos(inkAngle - detectionAngle))) < 0.001,
            "The tight ink box changed the text reading direction.");
    }

    private static double Distance(OcrPoint first, OcrPoint second)
    {
        double x = second.X - first.X, y = second.Y - first.Y;
        return Math.Sqrt(x * x + y * y);
    }

    private static void Require(bool condition, string message)
    {
        if (!condition)
        {
            throw new InvalidOperationException(message);
        }
    }
}
