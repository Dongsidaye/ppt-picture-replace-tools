using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Tables;
using CSMath;

if (args.Length is < 1 or > 2)
{
    throw new ArgumentException("Usage: SmokeTest <new-output-folder> [DwgToSvg.exe]");
}
string folder = Path.GetFullPath(args[0]);
if (Directory.Exists(folder))
{
    throw new IOException("Use a new fixture output folder.");
}
Directory.CreateDirectory(folder);
string input = Path.Combine(folder, "vector-中文.dwg");
string output = Path.Combine(folder, "converted.svg");
string resultPath = Path.Combine(folder, "result.json");

CadDocument drawing = new(ACadVersion.AC1032);
drawing.ModelSpace.Entities.Add(new Line { StartPoint = new XYZ(0, 0, 0), EndPoint = new XYZ(240, 0, 0) });
drawing.ModelSpace.Entities.Add(new Circle { Center = new XYZ(70, 70, 0), Radius = 35 });
LwPolyline polyline = new() { IsClosed = true };
polyline.Vertices.Add(new LwPolyline.Vertex(120, 20));
polyline.Vertices.Add(new LwPolyline.Vertex(220, 20));
polyline.Vertices.Add(new LwPolyline.Vertex(220, 110));
polyline.Vertices.Add(new LwPolyline.Vertex(120, 110));
drawing.ModelSpace.Entities.Add(polyline);
drawing.ModelSpace.Entities.Add(new TextEntity { Value = "中文 DWG导入", InsertPoint = new XYZ(5, 125, 0), Height = 12 });
BlockRecord block = new("SmokeBlock");
block.Entities.Add(new Line { StartPoint = new XYZ(0, 0, 0), EndPoint = new XYZ(20, 20, 0) });
block.Entities.Add(new Line { StartPoint = new XYZ(0, 20, 0), EndPoint = new XYZ(20, 0, 0) });
drawing.BlockRecords.Add(block);
drawing.ModelSpace.Entities.Add(new Insert(block) { InsertPoint = new XYZ(190, 125, 0) });
drawing.ModelSpace.Entities.Add(new Ray { StartPoint = new XYZ(20, 20, 0), Direction = new XYZ(1, 0, 0) });
DwgWriter.Write(input, drawing);

if (args.Length == 1)
{
    Console.WriteLine(input);
    return;
}

byte[] before = SHA256.HashData(File.ReadAllBytes(input));
ProcessStartInfo start = new(Path.GetFullPath(args[1]))
{
    UseShellExecute = false,
    CreateNoWindow = true,
    RedirectStandardOutput = true,
    RedirectStandardError = true
};
start.ArgumentList.Add(input);
start.ArgumentList.Add(output);
start.ArgumentList.Add(resultPath);
using Process process = Process.Start(start) ?? throw new IOException("Converter did not start.");
Task<string> errors = process.StandardError.ReadToEndAsync();
Task<string> standard = process.StandardOutput.ReadToEndAsync();
if (!process.WaitForExit(90000))
{
    process.Kill(entireProcessTree: true);
    throw new TimeoutException("Converter fixture timed out.");
}
if (process.ExitCode != 0)
{
    throw new Exception($"Conversion failed with code {process.ExitCode}: {await errors}");
}
using JsonDocument result = JsonDocument.Parse(File.ReadAllText(resultPath));
if (!result.RootElement.GetProperty("ok").GetBoolean() || result.RootElement.GetProperty("entityCount").GetInt32() != 6)
{
    throw new Exception("Unexpected converter result.");
}
string warning = result.RootElement.GetProperty("warning").GetString() ?? "";
if (!warning.Contains("not implemented", StringComparison.OrdinalIgnoreCase))
{
    throw new Exception("The unsupported RAY entity was silently omitted.");
}
XDocument svg = XDocument.Load(output);
XNamespace ns = "http://www.w3.org/2000/svg";
if (!svg.Descendants(ns + "path").Any() && !svg.Descendants(ns + "polyline").Any())
{
    throw new Exception("The result has no vector paths.");
}
if (!svg.Descendants(ns + "text").Any(element => element.Value.Contains("中文")))
{
    throw new Exception("The Chinese CAD text was lost.");
}
if (svg.Descendants(ns + "image").Any())
{
    throw new Exception("The vector fixture was rasterized.");
}
if (!before.SequenceEqual(SHA256.HashData(File.ReadAllBytes(input))))
{
    throw new Exception("The input drawing was modified.");
}
Console.WriteLine(JsonSerializer.Serialize(new { ok = true, entityCount = 6, vector = true, chinese = true, block = true, unsupportedEntityWarned = true, sourceUnchanged = true }));
