using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Xml;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Dongsidaye.PowerPointTools
{
    public sealed partial class Connect
    {
        private sealed class DwgImportRequest
        {
            internal string Id, Owner, Folder, Input, Output, OutputHash, Result, Status = "准备转换…";
            internal int Generation;
            internal volatile bool Completed;
            internal bool Inserted;
            internal bool OcrEnabled, Inserting, InsertionCompleted;
            internal string Font = "Microsoft YaHei", PreviewSvg, GeometrySvg, InsertionResult;
            internal int InsertionDone, InsertionTotal;
            internal DwgTextItem[] Texts = new DwgTextItem[0];
            internal DwgOcrPage OcrPage;
            internal PowerPoint.Presentation Presentation;
            internal readonly CancellationTokenSource Cancellation = new CancellationTokenSource();
        }

        private DwgImportRequest dwgImport;
        private static JavaScriptSerializer DwgJson()
        {
            return new JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024 };
        }

        internal string BeginDwgImport(string optionsJson)
        {
            if (!connected || stopping || FileDialogPending) throw new InvalidOperationException("请先打开文稿并完成文件选择。");
            var options = DwgJson().Deserialize<Dictionary<string, object>>(optionsJson);
            if (options == null || !options.ContainsKey("path") || !options.ContainsKey("documentKey"))
                throw new InvalidOperationException("缺少 DWG 文件或文稿信息。");
            string path = Convert.ToString(options["path"]), expectedOwner = Convert.ToString(options["documentKey"]);
            string owner = DocumentKey;
            if (String.IsNullOrEmpty(owner) || owner != expectedOwner) throw new InvalidOperationException("文稿已切换，请重新选择 DWG。");
            if (dwgImport != null && (!dwgImport.Completed || dwgImport.Inserting)) throw new InvalidOperationException("上一次导入正在结束，请稍后重试。");
            if (!String.Equals(Path.GetExtension(path), ".dwg", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("请选择 DWG 文件。");
            var request = new DwgImportRequest {
                Id = Guid.NewGuid().ToString("N"), Owner = owner, Generation = browserGeneration,
                Input = Path.GetFullPath(path), Presentation = application.ActiveWindow.Presentation,
                OcrEnabled = options.ContainsKey("ocr") && Convert.ToBoolean(options["ocr"]),
                Font = options.ContainsKey("font") ? Convert.ToString(options["font"]) : "Microsoft YaHei"
            };
            request.Folder = Path.Combine(Path.GetTempPath(), "DongsidayeDwgImport", request.Id);
            request.Output = Path.Combine(request.Folder, "drawing.svg");
            DwgImportRequest previous = dwgImport;
            dwgImport = request;
            if (previous != null) DeleteDwgFolder(previous);
            Task.Run(delegate { ConvertDwg(request); });
            return request.Id;
        }

        private static void ConvertDwg(DwgImportRequest request)
        {
            try
            {
                Directory.CreateDirectory(request.Folder);
                if ((File.GetAttributes(request.Folder) & FileAttributes.ReparsePoint) != 0) throw new IOException("转换目录无效。");
                string helper = Path.Combine(Path.GetDirectoryName(typeof(Connect).Assembly.Location), "cad", "DwgToSvg.exe");
                if (!File.Exists(helper)) throw new FileNotFoundException("转换组件缺失，请重新安装插件。");
                string resultPath = Path.Combine(request.Folder, "result.json");
                request.Status = request.OcrEnabled ? "正在转换模型空间并识别轮廓文字…" : "正在转换模型空间…";
                var start = new ProcessStartInfo(helper, QuoteDwgArgument(request.Input) + " " + QuoteDwgArgument(request.Output) + " " + QuoteDwgArgument(resultPath) + (request.OcrEnabled ? " --ocr" : "")) {
                    UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden,
                    WorkingDirectory = request.Folder
                };
                using (Process process = Process.Start(start))
                {
                    var elapsed = Stopwatch.StartNew();
                    try
                    {
                        while (!process.WaitForExit(150))
                        {
                            request.Cancellation.Token.ThrowIfCancellationRequested();
                            if (elapsed.Elapsed.TotalSeconds > 180) throw new TimeoutException("转换超过 3 分钟，请简化图纸后重试。");
                            try
                            {
                                process.Refresh();
                                if (process.PrivateMemorySize64 > 1536L * 1024 * 1024) throw new InvalidOperationException("图纸转换所需内存过大，请分区导出后导入。");
                            }
                            catch (InvalidOperationException)
                            {
                                // The helper can exit between WaitForExit and
                                // the resource sample. Its finished result wins.
                                if (process.HasExited) break;
                                throw;
                            }
                        }
                    }
                    finally
                    {
                        try { if (!process.HasExited) { process.Kill(); process.WaitForExit(5000); } } catch { }
                    }
                    request.Cancellation.Token.ThrowIfCancellationRequested();
                    if (!File.Exists(resultPath) || new FileInfo(resultPath).Length > 65536) throw new IOException("转换器未返回有效结果。");
                    var result = DwgJson().Deserialize<Dictionary<string, object>>(File.ReadAllText(resultPath, Encoding.UTF8));
                    if (result == null || process.ExitCode != 0 || !result.ContainsKey("ok") || !Convert.ToBoolean(result["ok"]))
                        throw new InvalidOperationException(result != null && result.ContainsKey("error") ? Convert.ToString(result["error"]) : "DWG 转换失败。");
                    request.Status = "正在检查 SVG…";
                    DwgSvg.Validate(request.Output);
                    byte[] svgBytes = File.ReadAllBytes(request.Output);
                    using (SHA256 hash = SHA256.Create()) request.OutputHash = Convert.ToBase64String(hash.ComputeHash(svgBytes));
                    string preview = Convert.ToBase64String(svgBytes);
                    request.PreviewSvg = Encoding.UTF8.GetString(svgBytes);
                    string geometryPath = Path.Combine(request.Folder, "geometry.svg");
                    DwgSvg.Validate(geometryPath);
                    request.GeometrySvg = File.ReadAllText(geometryPath, Encoding.UTF8);
                    request.Texts = DwgJson().Deserialize<DwgTextItem[]>(ReadDwgMetadata(Path.Combine(request.Folder, "text-layout.json")));
                    if (request.Texts == null || request.Texts.Length > 20000) throw new InvalidDataException("图纸文字数量过多，请分区导入。");
                    string ocrPath = Path.Combine(request.Folder, "ocr-layout.json");
                    if (request.OcrEnabled && File.Exists(ocrPath)) request.OcrPage = DwgJson().Deserialize<DwgOcrPage>(ReadDwgMetadata(ocrPath));
                    request.Result = DwgJson().Serialize(new {
                        state = "done", ok = true, cancelled = false, status = "转换完成", previewBase64 = preview,
                        warning = result.ContainsKey("warning") ? Convert.ToString(result["warning"]) : "",
                        entityCount = result.ContainsKey("entityCount") ? result["entityCount"] : 0,
                        textCount = request.Texts.Length, ocr = request.OcrPage
                    });
                }
            }
            catch (OperationCanceledException) { request.Result = DwgJson().Serialize(new { state = "done", ok = false, cancelled = true, error = "已取消转换。" }); }
            catch (Exception error) { request.Result = DwgJson().Serialize(new { state = "done", ok = false, error = error.Message }); }
            finally
            {
                request.Completed = true;
                if (request.Cancellation.IsCancellationRequested) DeleteDwgFolder(request);
            }
        }

        private static string ReadDwgMetadata(string path)
        {
            if (new FileInfo(path).Length > 16L * 1024 * 1024) throw new IOException("文字识别数据过大，请分区导入。");
            return File.ReadAllText(path, Encoding.UTF8);
        }

        private static string QuoteDwgArgument(string value)
        {
            // These are generated file paths, with no trailing directory separator.
            if (value.IndexOf('"') >= 0) throw new ArgumentException("文件路径无效。");
            return "\"" + value + "\"";
        }

        internal string PollDwgImport(string id)
        {
            DwgImportRequest request = RequireDwgRequest(id);
            if (!request.Completed) return DwgJson().Serialize(new { state = "pending", status = request.Status });
            return request.Result;
        }

        private DwgImportRequest RequireDwgRequest(string id)
        {
            if (dwgImport == null || id != dwgImport.Id) throw new InvalidOperationException("转换请求已过期，请重新选择 DWG。");
            if (!IsCurrentBrowser(dwgImport.Generation) || DocumentKey != dwgImport.Owner || !SameComIdentity(dwgImport.Presentation, application.ActiveWindow.Presentation))
            {
                dwgImport.Cancellation.Cancel();
                throw new InvalidOperationException("文稿已切换，请回到原文稿重新导入。");
            }
            return dwgImport;
        }

        internal bool CancelDwgImport(string id)
        {
            if (dwgImport == null || id != dwgImport.Id) return false;
            dwgImport.Cancellation.Cancel();
            if (dwgImport.Completed && !dwgImport.Inserting) DeleteDwgFolder(dwgImport);
            return true;
        }

        private void CancelDwgImports()
        {
            if (dwgImport == null) return;
            dwgImport.Cancellation.Cancel();
            if (dwgImport.Completed && !dwgImport.Inserting) DeleteDwgFolder(dwgImport);
            dwgImport.Presentation = null;
        }

        private static void DeleteDwgFolder(DwgImportRequest request)
        {
            string root = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "DongsidayeDwgImport")) + Path.DirectorySeparatorChar;
            string folder = Path.GetFullPath(request.Folder);
            if (!folder.StartsWith(root, StringComparison.OrdinalIgnoreCase) || Path.GetFileName(folder) != request.Id) return;
            try { Directory.Delete(folder, true); } catch { }
        }

        internal string InsertDwgImport(string id)
        {
            try
            {
                DwgImportRequest request = RequireDwgRequest(id);
                CheckDwgReady(request);
                request.Inserting = true;
                return InsertDwgEditable(request, new DwgInsertionOptions(), false).GetAwaiter().GetResult();
            }
            catch (Exception error)
            {
                return DwgJson().Serialize(new { ok = false, error = error.Message });
            }
        }
    }

    internal static class DwgSvg
    {
        internal static double[] Validate(string path)
        {
            var file = new FileInfo(path);
            if (!file.Exists || file.Length < 32 || file.Length > 16L * 1024 * 1024) throw new IOException("SVG 无效或超过 16 MB，请简化图纸后导入。");
            XDocument document;
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 16L * 1024 * 1024 };
            using (var reader = XmlReader.Create(path, settings)) document = XDocument.Load(reader);
            XNamespace svg = "http://www.w3.org/2000/svg";
            if (document.Root == null || document.Root.Name != svg + "svg") throw new IOException("转换结果不是 SVG。");
            int count = 0;
            var allowed = new HashSet<string>(new[] { "svg", "g", "defs", "path", "line", "circle", "ellipse", "rect", "polyline", "polygon", "text", "tspan", "use", "clipPath", "pattern", "image", "title", "desc", "mask", "linearGradient", "radialGradient", "stop" }, StringComparer.Ordinal);
            foreach (XElement element in document.Root.DescendantsAndSelf())
            {
                if (++count > 250000) throw new IOException("图纸矢量对象过多，请分区导入。");
                if (element.Name.Namespace != svg || !allowed.Contains(element.Name.LocalName)) throw new IOException("SVG 包含不支持的元素。");
                foreach (XAttribute attribute in element.Attributes())
                {
                    string name = attribute.Name.LocalName, value = attribute.Value.Trim();
                    if (attribute.Name == XNamespace.Xml + "base") throw new IOException("SVG 包含外部基地址。");
                    if (name.StartsWith("on", StringComparison.OrdinalIgnoreCase)) throw new IOException("SVG 包含活动事件。");
                    if (name == "href" && value.Length > 0 && !value.StartsWith("#", StringComparison.Ordinal) &&
                        !value.StartsWith("data:image/png;base64,", StringComparison.OrdinalIgnoreCase) && !value.StartsWith("data:image/jpeg;base64,", StringComparison.OrdinalIgnoreCase)) throw new IOException("SVG 包含外部文件引用。");
                    if (value.IndexOf("url(", StringComparison.OrdinalIgnoreCase) >= 0 && !System.Text.RegularExpressions.Regex.IsMatch(value, @"^url\(\s*['"" ]?#[a-zA-Z0-9_\-:.]+['"" ]?\s*\)$")) throw new IOException("SVG 包含外部样式引用。");
                }
            }
            string[] box = ((string)document.Root.Attribute("viewBox") ?? "").Split(new[] { ' ', ',', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            double width, height, left, top;
            if (box.Length != 4 || !Double.TryParse(box[2], NumberStyles.Float, CultureInfo.InvariantCulture, out width) ||
                !Double.TryParse(box[3], NumberStyles.Float, CultureInfo.InvariantCulture, out height) ||
                !Double.TryParse(box[0], NumberStyles.Float, CultureInfo.InvariantCulture, out left) ||
                !Double.TryParse(box[1], NumberStyles.Float, CultureInfo.InvariantCulture, out top) ||
                Double.IsNaN(left) || Double.IsNaN(top) || Double.IsInfinity(left) || Double.IsInfinity(top) ||
                Double.IsNaN(width) || Double.IsNaN(height) || Double.IsInfinity(width) || Double.IsInfinity(height) || width <= 0 || height <= 0)
                throw new IOException("SVG 缺少有效的图纸范围。");
            return new[] { width, height, left, top };
        }
    }
}
