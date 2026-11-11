using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using Dongsidaye.PowerPointTools;

namespace Dongsidaye.WpsCad
{
    internal sealed class JobSpec
    {
        public string id { get; set; }
        public string phase { get; set; }
        public string input { get; set; }
        public bool ocr { get; set; }
        public string font { get; set; }
        public double pageWidth { get; set; }
        public double pageHeight { get; set; }
        public DwgOcrSelection[] ocrSelections { get; set; }
        public string[] fallbackIds { get; set; }
    }

    internal sealed class SavedState
    {
        public string id { get; set; }
        public string mode { get; set; }
        public bool ocr { get; set; }
        public Dictionary<string, string> hashes { get; set; }
        public PreparedState prepared { get; set; }
        public string planHash { get; set; }
    }

    internal sealed class PreparedState
    {
        public string font { get; set; }
        public double left { get; set; }
        public double top { get; set; }
        public double width { get; set; }
        public double height { get; set; }
        public double scale { get; set; }
        public double[] viewBox { get; set; }
        public List<DwgOcrPoint[]> replacements { get; set; }
    }

    internal sealed class TextPlan
    {
        public DwgTextItem item { get; set; }
        public bool ocr { get; set; }
        public double fontSize { get; set; }
        public InkMetrics metrics { get; set; }
        public double targetWidth { get; set; }
        public double targetHeight { get; set; }
        public double centerX { get; set; }
        public double centerY { get; set; }
        public string unsupported { get; set; }
    }

    internal sealed class JobContext
    {
        internal readonly string Id, Phase, Prefix, Folder, RequestPath, StatusPath, CancelPath;
        internal readonly Stopwatch Elapsed = Stopwatch.StartNew();
        internal JobSpec Request;

        internal JobContext(string requestPath)
        {
            string fullPath = Path.GetFullPath(requestPath);
            if (!Path.IsPathRooted(requestPath)) throw new InvalidOperationException("CAD 请求必须使用绝对路径。");
            SafeFiles.RejectReparse(fullPath);
            string temp = SafeFiles.ExistingLongPath(Path.GetTempPath());
            if (!String.Equals(SafeFiles.ExistingLongPath(Path.GetDirectoryName(fullPath)), temp, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("CAD 请求必须位于当前用户临时目录。");
            Match match = Regex.Match(Path.GetFileName(fullPath), @"^DongsidayeWpsCad_([a-f0-9]{32})\.(convert|prepare|finalize|cleanup)\.json$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success) throw new InvalidOperationException("CAD 请求文件名无效。");
            Id = match.Groups[1].Value.ToLowerInvariant();
            Phase = match.Groups[2].Value.ToLowerInvariant();
            // Keep the caller's TEMP spelling for the JS file mailbox. WPS can
            // expose ADMINI~1 while .NET expands it to Administrator.
            RequestPath = requestPath;
            int separator = Math.Max(requestPath.LastIndexOf('\\'), requestPath.LastIndexOf('/'));
            if (separator < 0) throw new InvalidOperationException("CAD 请求必须使用完整临时目录路径。");
            Prefix = requestPath.Substring(0, separator + 1) + "DongsidayeWpsCad_" + Id;
            Folder = Prefix;
            StatusPath = Prefix + "." + Phase + ".status.json";
            CancelPath = Prefix + ".cancel.flag";
            SafeFiles.RejectReparse(RequestPath);
        }

        internal void Check()
        {
            if (Phase != "cleanup" && File.Exists(CancelPath)) throw new OperationCanceledException("已取消 CAD 导入。");
            if (Elapsed.Elapsed.TotalSeconds > 180) throw new TimeoutException("CAD 处理超过 3 分钟，请简化图纸后重试。");
        }

        internal string Child(string name) { return Path.Combine(Folder, name); }

        internal void Pending(string message)
        {
            SafeFiles.WriteJson(StatusPath, new { state = "pending", ok = true, message = message });
        }

        internal void Complete(bool ok, string output, string error, bool cancelled)
        {
            SafeFiles.WriteJson(StatusPath, new { state = "done", ok = ok, output = output ?? "", error = error ?? "", cancelled = cancelled });
        }
    }

    internal static class SafeFiles
    {
        internal const long MaximumBytes = 16L * 1024 * 1024;
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetLongPathName(string shortPath, StringBuilder longPath, uint capacity);

        internal static string ExistingLongPath(string path)
        {
            string full = Path.GetFullPath(path);
            RejectReparse(full);
            var expanded = new StringBuilder(32768);
            uint length = GetLongPathName(full, expanded, (uint)expanded.Capacity);
            if (length == 0 || length >= expanded.Capacity) throw new IOException("无法确认 CAD 临时目录的真实路径。");
            return Path.GetFullPath(expanded.ToString()).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
        internal static JavaScriptSerializer Json()
        {
            return new JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024, RecursionLimit = 256 };
        }

        internal static void RejectReparse(string path)
        {
            string current = Path.GetFullPath(path);
            for (;;)
            {
                if ((File.Exists(current) || Directory.Exists(current)) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("CAD 路径包含重解析链接。");
                string parent = Path.GetDirectoryName(current);
                if (String.IsNullOrEmpty(parent) || parent == current) break;
                current = parent;
            }
        }

        internal static byte[] Read(string path, long maximum)
        {
            RejectReparse(path);
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (stream.Length < 1 || stream.Length > maximum) throw new IOException("CAD 数据文件为空或超过大小限制。");
                byte[] result = new byte[(int)stream.Length];
                int offset = 0;
                while (offset < result.Length)
                {
                    int count = stream.Read(result, offset, result.Length - offset);
                    if (count == 0) throw new IOException("CAD 数据读取不完整。");
                    offset += count;
                }
                return result;
            }
        }

        internal static string Text(string path, long maximum) { return Encoding.UTF8.GetString(Read(path, maximum)).TrimStart('\uFEFF'); }
        internal static T ReadJson<T>(string path, long maximum) { return Json().Deserialize<T>(Text(path, maximum)); }

        internal static string Hash(byte[] bytes)
        {
            using (SHA256 hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
        }

        internal static void WriteJson(string path, object value)
        {
            // WPS exposes both UTF-8 text and binary-string readers. ASCII JSON
            // keeps Chinese text and paths correct through either interface.
            string json = Json().Serialize(value);
            var ascii = new StringBuilder(json.Length);
            foreach (char character in json)
            {
                if (character > 127) ascii.Append("\\u").Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                else ascii.Append(character);
            }
            if (ascii.Length > 32 * 1024 * 1024) throw new InvalidOperationException("CAD 处理结果超过 32 MB，请分区导入。");
            WriteText(path, ascii.ToString());
        }

        internal static void WriteText(string path, string value)
        {
            RejectReparse(path);
            string temporary = path + ".tmp-" + Guid.NewGuid().ToString("N");
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    byte[] bytes = new UTF8Encoding(false).GetBytes(value);
                    stream.Write(bytes, 0, bytes.Length);
                    stream.Flush(true);
                }
                RejectReparse(path);
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        internal static void DeleteFolder(JobContext job)
        {
            if (!Directory.Exists(job.Folder)) return;
            string expected = Path.Combine(Path.GetFullPath(Path.GetTempPath()), "DongsidayeWpsCad_" + job.Id);
            if (!String.Equals(ExistingLongPath(job.Folder), ExistingLongPath(expected), StringComparison.OrdinalIgnoreCase)) throw new IOException("拒绝清理非本次 CAD 目录。");
            RejectReparse(job.Folder);
            var pending = new Stack<string>();
            pending.Push(job.Folder);
            int count = 0;
            while (pending.Count > 0)
            {
                foreach (string item in Directory.EnumerateFileSystemEntries(pending.Pop()))
                {
                    if (++count > 10000) throw new IOException("CAD 临时目录包含异常数量的文件。");
                    FileAttributes attributes = File.GetAttributes(item);
                    if ((attributes & FileAttributes.ReparsePoint) != 0) throw new IOException("CAD 临时目录包含链接，已停止清理。");
                    if ((attributes & FileAttributes.Directory) != 0) pending.Push(item);
                }
            }
            Directory.Delete(job.Folder, true);
        }
    }

    internal static class Program
    {
        private static readonly string[] ImmutableNames = { "drawing.svg", "geometry.svg", "text-layout.json", "result.json" };

        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Length == 1 && args[0] == "--serve") return CadServer.Run();
            JobContext job = null;
            FileStream jobLock = null;
            bool cleanup = false;
            bool requestValidated = false;
            try
            {
                if (args.Length != 2 || args[0] != "--job") return 2;
                job = new JobContext(args[1]);
                job.Request = SafeFiles.ReadJson<JobSpec>(job.RequestPath, 8L * 1024 * 1024);
                if (job.Request == null || !String.Equals(job.Request.id, job.Id, StringComparison.OrdinalIgnoreCase) || job.Request.phase != job.Phase)
                    throw new InvalidOperationException("CAD 请求身份或阶段不一致。");
                requestValidated = true;
                jobLock = AcquireLock(job);
                job.Check();
                string output;
                switch (job.Phase)
                {
                    case "convert": output = Convert(job); break;
                    case "prepare": output = Prepare(job); break;
                    case "finalize": output = Finalize(job); break;
                    case "cleanup": Cleanup(job); cleanup = true; output = ""; break;
                    default: throw new InvalidOperationException("未知 CAD 处理阶段。");
                }
                job.Complete(true, output, "", false);
                return 0;
            }
            catch (Exception error)
            {
                if (job != null)
                {
                    bool cancelled = error is OperationCanceledException;
                    FileStream completionLock = jobLock;
                    if (completionLock == null)
                    {
                        try
                        {
                            SafeFiles.RejectReparse(job.Prefix + ".lock");
                            completionLock = new FileStream(job.Prefix + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read);
                        }
                        catch { }
                    }
                    if (completionLock != null)
                    {
                        try
                        {
                            if (cancelled || error is TimeoutException) { try { SafeFiles.DeleteFolder(job); } catch { } }
                            try { job.Complete(false, "", error.Message, cancelled); } catch { }
                        }
                        finally { if (completionLock != jobLock) completionLock.Dispose(); }
                    }
                    else if (requestValidated)
                    {
                        // A different phase may still own the folder. Report
                        // only this queued phase and leave its live owner alone.
                        string activePhase = ReadActivePhase(job);
                        if (!String.IsNullOrEmpty(activePhase) && activePhase != job.Phase)
                        {
                            try { job.Complete(false, "", error.Message, cancelled); } catch { }
                        }
                    }
                }
                return 1;
            }
            finally
            {
                if (jobLock != null) jobLock.Dispose();
                if (cleanup && job != null)
                {
                    try { SafeFiles.RejectReparse(job.Prefix + ".lock"); File.Delete(job.Prefix + ".lock"); } catch { }
                }
            }
        }

        private static FileStream AcquireLock(JobContext job)
        {
            var elapsed = Stopwatch.StartNew();
            string path = job.Prefix + ".lock";
            for (;;)
            {
                job.Check();
                SafeFiles.RejectReparse(path);
                FileStream stream;
                try { stream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read); }
                catch (IOException)
                {
                    if (elapsed.Elapsed.TotalSeconds >= 5) throw new InvalidOperationException("上一个 CAD 阶段仍在结束，请稍后重新导入。");
                    Thread.Sleep(75);
                    continue;
                }
                try
                {
                    byte[] bytes = Encoding.ASCII.GetBytes(job.Phase);
                    stream.SetLength(0);
                    stream.Write(bytes, 0, bytes.Length);
                    stream.Flush(true);
                    return stream;
                }
                catch { stream.Dispose(); throw; }
            }
        }

        private static string ReadActivePhase(JobContext job)
        {
            try
            {
                string path = job.Prefix + ".lock";
                SafeFiles.RejectReparse(path);
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream, Encoding.ASCII))
                {
                    if (stream.Length > 32) return "";
                    string phase = reader.ReadToEnd();
                    return new[] { "convert", "prepare", "finalize", "cleanup" }.Contains(phase) ? phase : "";
                }
            }
            catch { return ""; }
        }

        private static string Convert(JobContext job)
        {
            if (Directory.Exists(job.Folder) || File.Exists(job.Folder)) throw new InvalidOperationException("此 CAD 转换任务已经存在，请重新选择文件。");
            string input = Path.GetFullPath(job.Request.input ?? "");
            if (!String.Equals(Path.GetExtension(input), ".dwg", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("请选择 DWG 文件。");
            Directory.CreateDirectory(job.Folder);
            SafeFiles.RejectReparse(job.Folder);
            string converter = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DwgToSvg.exe");
            SafeFiles.RejectReparse(converter);
            if (!File.Exists(converter)) throw new FileNotFoundException("CAD 转换组件缺失，请重新安装插件。");
            job.Pending(job.Request.ocr ? "正在转换图纸并识别轮廓文字…" : "正在转换模型空间…");
            var start = new ProcessStartInfo(converter, Quote(input) + " " + Quote(job.Child("drawing.svg")) + " " + Quote(job.Child("result.json")) + (job.Request.ocr ? " --ocr" : "")) {
                UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden, WorkingDirectory = job.Folder
            };
            using (Process process = Process.Start(start))
            {
                try
                {
                    while (!process.WaitForExit(150))
                    {
                        job.Check();
                        try
                        {
                            process.Refresh();
                            if (process.PrivateMemorySize64 > 1536L * 1024 * 1024) throw new InvalidOperationException("图纸转换所需内存过大，请分区导入。");
                        }
                        catch (InvalidOperationException) { if (process.HasExited) break; throw; }
                    }
                    job.Check();
                    var result = SafeFiles.ReadJson<Dictionary<string, object>>(job.Child("result.json"), 65536);
                    if (process.ExitCode != 0 || result == null || !result.ContainsKey("ok") || !System.Convert.ToBoolean(result["ok"]))
                        throw new InvalidOperationException(result != null && result.ContainsKey("error") ? System.Convert.ToString(result["error"]) : "DWG 转换失败。");
                    double[] size = DwgSvg.Validate(job.Child("drawing.svg"));
                    DwgSvg.Validate(job.Child("geometry.svg"));
                    DwgTextItem[] texts = ReadTexts(job);
                    DwgOcrPage ocr = ReadOcr(job, job.Request.ocr);
                    var hashes = new Dictionary<string, string>(StringComparer.Ordinal);
                    foreach (string name in ImmutableNames) hashes.Add(name, SafeFiles.Hash(SafeFiles.Read(job.Child(name), SafeFiles.MaximumBytes)));
                    if (job.Request.ocr && File.Exists(job.Child("ocr-layout.json"))) hashes.Add("ocr-layout.json", SafeFiles.Hash(SafeFiles.Read(job.Child("ocr-layout.json"), SafeFiles.MaximumBytes)));
                    var state = new SavedState { id = job.Id, mode = "converted", ocr = job.Request.ocr, hashes = hashes };
                    SafeFiles.WriteJson(job.Child("state.json"), state);
                    string output = job.Child("converted.json");
                    SafeFiles.WriteJson(output, new {
                        warning = result.ContainsKey("warning") ? System.Convert.ToString(result["warning"]) : "",
                        entityCount = result.ContainsKey("entityCount") ? result["entityCount"] : 0,
                        textCount = texts.Length, ocr = ocr,
                        previewBase64 = System.Convert.ToBase64String(SafeFiles.Read(job.Child("drawing.svg"), SafeFiles.MaximumBytes)),
                        outputHash = hashes["drawing.svg"], size = size
                    });
                    return output;
                }
                finally
                {
                    if (!process.HasExited) { try { process.Kill(); process.WaitForExit(5000); } catch { } }
                }
            }
        }

        private static string Prepare(JobContext job)
        {
            SavedState state = VerifyState(job);
            string font = job.Request.font;
            if (!new[] { "Microsoft YaHei", "SimSun", "Arial" }.Contains(font, StringComparer.Ordinal)) throw new InvalidOperationException("请选择支持的导入字体。");
            double pageWidth = job.Request.pageWidth, pageHeight = job.Request.pageHeight;
            if (!Finite(pageWidth) || !Finite(pageHeight) || pageWidth <= 0 || pageHeight <= 0 || pageWidth > 100000 || pageHeight > 100000)
                throw new InvalidOperationException("幻灯片页面尺寸无效。");
            double[] size = DwgSvg.Validate(job.Child("drawing.svg"));
            double scale = Math.Min(pageWidth * 0.8 / size[0], pageHeight * 0.8 / size[1]);
            double width = size[0] * scale, height = size[1] * scale;
            if (!Finite(scale) || !Finite(width) || !Finite(height) || scale <= 0 || width < 0.01 || height < 0.01) throw new InvalidOperationException("图纸缩放后尺寸无效。");
            var prepared = new PreparedState { font = font, left = (pageWidth - width) / 2, top = (pageHeight - height) / 2,
                width = width, height = height, scale = scale, viewBox = new[] { size[2], size[3], size[0], size[1] }, replacements = new List<DwgOcrPoint[]>() };
            var plans = new List<TextPlan>();
            DwgTextItem[] texts = ReadTexts(job);
            for (int index = 0; index < texts.Length; index++)
            {
                job.Check();
                plans.Add(FontFitting.Plan(texts[index], false, prepared));
                if (index % 100 == 0) job.Pending("正在匹配文字字体 " + index + " / " + texts.Length + "…");
            }
            DwgOcrSelection[] selections = job.Request.ocrSelections ?? new DwgOcrSelection[0];
            if (selections.Length > 500) throw new InvalidOperationException("一次最多恢复 500 行 OCR 文字。");
            DwgOcrPage page = ReadOcr(job, state.ocr);
            var selected = new HashSet<int>();
            foreach (DwgOcrSelection selection in selections)
            {
                job.Check();
                if (selection == null || !selected.Add(selection.index) || page == null || selection.index < 0 || selection.index >= page.lines.Length)
                    throw new InvalidOperationException("OCR 选择无效或重复，请重新核对。");
                DwgOcrLine line = page.lines[selection.index];
                DwgTextItem item = DwgOcrLayout.ToText(page, line, selection.text, "wps-ocr:" + selection.index);
                TextPlan plan = FontFitting.Plan(item, true, prepared);
                if (!String.IsNullOrEmpty(plan.unsupported)) throw new InvalidOperationException("OCR 文字无法按原图排版：" + plan.unsupported);
                plans.Add(plan);
                prepared.replacements.Add(DwgOcrLayout.ToSvgPoints(page, line));
            }
            string output = job.Child("plan.json");
            SafeFiles.WriteJson(output, new { texts = plans, left = prepared.left, top = prepared.top, width = width, height = height, scale = scale, viewBox = prepared.viewBox });
            state.mode = "prepared";
            state.prepared = prepared;
            state.planHash = SafeFiles.Hash(SafeFiles.Read(output, 48L * 1024 * 1024));
            SafeFiles.WriteJson(job.Child("state.json"), state);
            return output;
        }

        private static string Finalize(JobContext job)
        {
            SavedState state = VerifyState(job);
            if (state.mode != "prepared" || state.prepared == null || String.IsNullOrEmpty(state.planHash)) throw new InvalidOperationException("请先完成文字排版准备。");
            if (SafeFiles.Hash(SafeFiles.Read(job.Child("plan.json"), 48L * 1024 * 1024)) != state.planHash) throw new InvalidOperationException("文字排版数据已变化，请重新导入。");
            var allowed = new HashSet<string>(ReadTexts(job).Select(item => item.SourceId), StringComparer.Ordinal);
            string[] ids = job.Request.fallbackIds ?? new string[0];
            if (ids.Length > 20000 || ids.Any(id => String.IsNullOrEmpty(id) || !allowed.Contains(id))) throw new InvalidOperationException("SVG 文字回退编号无效。");
            var fallback = new HashSet<string>(ids, StringComparer.Ordinal);
            string svg = DwgOcrLayout.PrepareGeometry(SafeFiles.Text(job.Child("drawing.svg"), SafeFiles.MaximumBytes),
                SafeFiles.Text(job.Child("geometry.svg"), SafeFiles.MaximumBytes), fallback, state.prepared.replacements);
            string path = job.Child("insertion.svg");
            SafeFiles.WriteText(path, svg);
            DwgSvg.Validate(path);
            XNamespace svgNamespace = "http://www.w3.org/2000/svg";
            var visible = new HashSet<string>(new[] { "path", "line", "polyline", "polygon", "circle", "ellipse", "rect", "text", "image", "use" });
            bool hasGeometry = XDocument.Parse(svg).Descendants().Any(element => element.Name.Namespace == svgNamespace && visible.Contains(element.Name.LocalName) && !element.Ancestors(svgNamespace + "defs").Any());
            string output = job.Child("finalized.json");
            SafeFiles.WriteJson(output, new { path = path, hasGeometry = hasGeometry, left = state.prepared.left, top = state.prepared.top,
                width = state.prepared.width, height = state.prepared.height });
            state.mode = "finalized";
            SafeFiles.WriteJson(job.Child("state.json"), state);
            return output;
        }

        private static SavedState VerifyState(JobContext job)
        {
            job.Check();
            SafeFiles.RejectReparse(job.Folder);
            SavedState state = SafeFiles.ReadJson<SavedState>(job.Child("state.json"), 1024 * 1024);
            if (state == null || state.id != job.Id || !new[] { "converted", "prepared", "finalized" }.Contains(state.mode) || state.hashes == null)
                throw new InvalidOperationException("CAD 转换状态无效，请重新导入。");
            var expected = new HashSet<string>(ImmutableNames, StringComparer.Ordinal);
            if (state.ocr && File.Exists(job.Child("ocr-layout.json"))) expected.Add("ocr-layout.json");
            if (!expected.SetEquals(state.hashes.Keys)) throw new InvalidOperationException("CAD 转换文件清单已变化。");
            foreach (string name in expected)
            {
                job.Check();
                if (SafeFiles.Hash(SafeFiles.Read(job.Child(name), SafeFiles.MaximumBytes)) != state.hashes[name]) throw new InvalidOperationException("CAD 转换内容已变化，请重新导入。");
            }
            return state;
        }

        private static DwgTextItem[] ReadTexts(JobContext job)
        {
            DwgTextItem[] texts = SafeFiles.ReadJson<DwgTextItem[]>(job.Child("text-layout.json"), SafeFiles.MaximumBytes);
            if (texts == null || texts.Length > 20000 || texts.Any(item => item == null || String.IsNullOrEmpty(item.SourceId))) throw new InvalidOperationException("CAD 文字数据无效或超过 20000 行。");
            return texts;
        }

        private static DwgOcrPage ReadOcr(JobContext job, bool enabled)
        {
            if (!enabled || !File.Exists(job.Child("ocr-layout.json"))) return null;
            DwgOcrPage page = SafeFiles.ReadJson<DwgOcrPage>(job.Child("ocr-layout.json"), SafeFiles.MaximumBytes);
            if (page == null || page.lines == null || page.lines.Length > 20000) throw new InvalidOperationException("OCR 数据无效或过多。");
            return page;
        }

        private static void Cleanup(JobContext job)
        {
            SafeFiles.DeleteFolder(job);
            foreach (string phase in new[] { "convert", "prepare", "finalize", "cleanup" })
            {
                foreach (string suffix in new[] { "." + phase + ".json", "." + phase + ".status.json" })
                {
                    string file = job.Prefix + suffix;
                    SafeFiles.RejectReparse(file);
                    if (File.Exists(file)) File.Delete(file);
                }
            }
            SafeFiles.RejectReparse(job.CancelPath);
            if (File.Exists(job.CancelPath)) File.Delete(job.CancelPath);
        }

        private static string Quote(string value)
        {
            if (value.IndexOf('"') >= 0 || value.EndsWith("\\", StringComparison.Ordinal)) throw new InvalidOperationException("CAD 文件路径无效。");
            return "\"" + value + "\"";
        }

        internal static bool Finite(double value) { return !Double.IsNaN(value) && !Double.IsInfinity(value); }
    }
}
