using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.Image;
using ACadSharp.IO;
using ImageColor = SixLabors.ImageSharp.Color;

namespace Dongsidaye.CadConverter;

internal static class Program
{
    private const int MaximumWarnings = 20;
    private const long MaximumInputBytes = 256L * 1024 * 1024;
    private static readonly List<string> Warnings = new();
    private static readonly HashSet<string> DistinctWarnings = new(StringComparer.Ordinal);
    private static int omittedWarnings;
    private static int editableTextCount;
    private static int ocrLineCount;

    public static int Main(string[] args)
    {
        bool createdOutput = false;
        string? output = null;
        string? resultPath = null;
        string? privateInput = null;
        int entityCount = 0;
        try
        {
            if (args.Length != 3 && (args.Length != 4 || args[3] != "--ocr"))
            {
                throw new ArgumentException("Usage: DwgToSvg.exe <input.dwg> <output.svg> <result.json>");
            }

            string input = Path.GetFullPath(args[0]);
            output = Path.GetFullPath(args[1]);
            resultPath = Path.GetFullPath(args[2]);
            ValidatePaths(input, output, resultPath);
            privateInput = CopyPrivateSnapshot(input, Path.GetDirectoryName(output)!);
            ValidateDrawingHeader(privateInput);

            CadDocument document = DwgReader.Read(privateInput, (_, notification) => RecordNotification(notification));
            entityCount = document.ModelSpace.Entities.Count;
            if (entityCount == 0)
            {
                throw new InvalidDataException("DWG model space contains no entities.");
            }
            RemoveExternalRasterContent(document);

            ImageExporter exporter = new();
            exporter.Configuration.Width = 1600;
            exporter.Configuration.Height = 1600;
            exporter.Configuration.SetPadding(12);
            exporter.Configuration.BackgroundColor = ImageColor.Transparent;
            exporter.Configuration.ForegroundColor = ImageColor.Black;
            exporter.Configuration.LayerVisibility = LayerVisibilityMode.Screen;
            exporter.Configuration.Svg.NonScalingStroke = true;
            exporter.Configuration.Svg.EmitSize = false;
            exporter.Configuration.OnNotification += (_, notification) => RecordNotification(notification);
            exporter.AddModelSpace(document);

            IReadOnlyList<RenderedPage> pages = exporter.Render(ImageExportFormat.Svg);
            try
            {
                if (pages.Count != 1 || pages[0] is not RenderedSvgPage svgPage)
                {
                    throw new InvalidDataException("The DWG did not produce one SVG model-space page.");
                }

                string svg = PrepareSvg(svgPage.Content);
                EditableTextLayoutResult layout = EditableTextLayout.Extract(svg, document);
                svg = layout.PreviewSvg;
                editableTextCount = layout.Texts.Count;
                foreach (string warning in layout.Warnings) RecordNotification(new NotificationEventArgs(warning, NotificationType.Warning));
                string folder = Path.GetDirectoryName(output)!;
                WriteOwnedText(Path.Combine(folder, "geometry.svg"), layout.GeometrySvg);
                WriteOwnedText(Path.Combine(folder, "text-layout.json"), JsonSerializer.Serialize(layout.Texts));
                if (args.Length == 4 && layout.HasGeometry)
                {
                    string models = Path.Combine(AppContext.BaseDirectory, "models", "ocr");
                    OcrPage ocr = OcrRecognition.ReadSvg(Path.Combine(folder, "geometry.svg"), models);
                    ocrLineCount = ocr.Lines.Count;
                    WriteOwnedText(Path.Combine(folder, "ocr-layout.json"), JsonSerializer.Serialize(ocr));
                }
                using FileStream stream = new(output, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                createdOutput = true;
                byte[] bytes = new UTF8Encoding(false).GetBytes(svg);
                stream.Write(bytes);
            }
            finally
            {
                foreach (RenderedPage page in pages)
                {
                    page.Dispose();
                }
            }

            WriteResult(resultPath, true, "", entityCount);
            return 0;
        }
        catch (Exception error)
        {
            if (createdOutput && output != null)
            {
                try { File.Delete(output); } catch (IOException) { }
            }

            string message = error.GetBaseException().Message;
            if (resultPath != null)
            {
                try { WriteResult(resultPath, false, message, entityCount); }
                catch (Exception resultError) when (resultError is IOException or UnauthorizedAccessException)
                {
                    Console.Error.WriteLine(resultError.Message);
                }
            }
            Console.Error.WriteLine(message);
            return 1;
        }
        finally
        {
            if (privateInput != null)
            {
                try { File.Delete(privateInput); }
                catch (Exception error) when (error is IOException or UnauthorizedAccessException)
                {
                    Console.Error.WriteLine("Could not remove the private input snapshot: " + error.Message);
                }
            }
        }
    }

    private static string CopyPrivateSnapshot(string input, string outputDirectory)
    {
        string snapshot = Path.Combine(outputDirectory, "private-input-" + Guid.NewGuid().ToString("N") + ".dwg");
        bool created = false;
        try
        {
            using FileStream source = new(input, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (source.Length > MaximumInputBytes)
            {
                throw new InvalidDataException("The DWG exceeds the 256 MiB input limit.");
            }
            using FileStream destination = new(snapshot, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            created = true;
            byte[] buffer = new byte[65536];
            long copied = 0;
            int count;
            while ((count = source.Read(buffer, 0, buffer.Length)) > 0)
            {
                if (count > MaximumInputBytes - copied)
                {
                    throw new InvalidDataException("The DWG exceeds the 256 MiB input limit.");
                }
                destination.Write(buffer, 0, count);
                copied += count;
            }
            return snapshot;
        }
        catch
        {
            if (created)
            {
                try { File.Delete(snapshot); }
                catch (Exception error) when (error is IOException or UnauthorizedAccessException)
                {
                    Console.Error.WriteLine("Could not remove the partial input snapshot: " + error.Message);
                }
            }
            throw;
        }
    }

    private static void ValidatePaths(string input, string output, string resultPath)
    {
        if (!string.Equals(Path.GetExtension(input), ".dwg", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Select a DWG drawing.");
        }
        if (!string.Equals(Path.GetExtension(output), ".svg", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(Path.GetExtension(resultPath), ".json", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Output paths must end in .svg and .json.");
        }
        if (string.Equals(input, output, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(input, resultPath, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(output, resultPath, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Input and output paths must be different.");
        }
        if (!string.Equals(Path.GetDirectoryName(output), Path.GetDirectoryName(resultPath), StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("SVG and result files must use the same conversion directory.");
        }
        if (!Directory.Exists(Path.GetDirectoryName(output)))
        {
            throw new DirectoryNotFoundException("The conversion output directory does not exist.");
        }
        if (File.Exists(output) || File.Exists(resultPath))
        {
            throw new IOException("Conversion output files already exist.");
        }
    }

    private static void ValidateDrawingHeader(string input)
    {
        using FileStream stream = new(input, FileMode.Open, FileAccess.Read, FileShare.Read);
        Span<byte> signature = stackalloc byte[6];
        if (stream.Read(signature) != signature.Length || !Encoding.ASCII.GetString(signature).StartsWith("AC10", StringComparison.Ordinal))
        {
            throw new InvalidDataException("The selected file does not have a supported DWG signature.");
        }
    }

    private static void RecordNotification(NotificationEventArgs notification)
    {
        if (notification.NotificationType == NotificationType.None || string.IsNullOrWhiteSpace(notification.Message))
        {
            return;
        }
        string text = notification.Message.Length > 600 ? notification.Message[..600] : notification.Message;
        if (Warnings.Count >= MaximumWarnings)
        {
            omittedWarnings++;
            return;
        }
        if (!DistinctWarnings.Add(text))
        {
            return;
        }
        Warnings.Add(text);
    }

    private static void RemoveExternalRasterContent(CadDocument document)
    {
        int removed = 0;
        foreach (var block in document.BlockRecords)
        {
            foreach (Entity entity in block.Entities.ToArray())
            {
                if (entity is RasterImage or Ole2Frame)
                {
                    block.Entities.Remove(entity);
                    removed++;
                }
            }
        }
        if (removed > 0)
        {
            RecordNotification(new NotificationEventArgs($"External raster images and OLE objects were omitted ({removed} entities).", NotificationType.Warning));
        }
    }

    private static string PrepareSvg(string content)
    {
        XmlReaderSettings settings = new() { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
        using StringReader text = new(content);
        using XmlReader reader = XmlReader.Create(text, settings);
        XDocument svg = XDocument.Load(reader);
        XNamespace ns = "http://www.w3.org/2000/svg";
        XElement root = svg.Root ?? throw new InvalidDataException("The converter returned an empty SVG.");
        if (root.Name != ns + "svg")
        {
            throw new InvalidDataException("The converter returned an invalid SVG root.");
        }
        double[] bounds = ((string?)root.Attribute("viewBox") ?? "")
            .Split(new[] { ' ', ',', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        if (bounds.Length != 4 || bounds.Any(value => !double.IsFinite(value)) || bounds[2] <= 0 || bounds[3] <= 0)
        {
            throw new InvalidDataException("The drawing has invalid SVG extents.");
        }
        string[] visibleElements = ["line", "path", "polyline", "polygon", "circle", "ellipse", "text"];
        if (!root.Descendants().Any(element => element.Name.Namespace == ns && visibleElements.Contains(element.Name.LocalName)))
        {
            throw new InvalidDataException("The drawing contains no supported visible model-space geometry.");
        }
        double scale = 1600 / Math.Max(bounds[2], bounds[3]);
        root.SetAttributeValue("width", Math.Max(1, bounds[2] * scale).ToString("0.###", CultureInfo.InvariantCulture));
        root.SetAttributeValue("height", Math.Max(1, bounds[3] * scale).ToString("0.###", CultureInfo.InvariantCulture));
        return svg.ToString(SaveOptions.DisableFormatting);
    }

    private static void WriteResult(string resultPath, bool ok, string error, int entityCount)
    {
        string warning = string.Join("\n", Warnings);
        if (omittedWarnings > 0)
        {
            warning += $"\n({omittedWarnings} additional notifications omitted.)";
        }
        byte[] bytes = JsonSerializer.SerializeToUtf8Bytes(new { ok, error, warning, entityCount, textCount = editableTextCount, ocrCount = ocrLineCount });
        using FileStream stream = new(resultPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        stream.Write(bytes);
    }

    private static void WriteOwnedText(string path, string content)
    {
        byte[] bytes = new UTF8Encoding(false).GetBytes(content);
        using var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        file.Write(bytes);
    }
}
