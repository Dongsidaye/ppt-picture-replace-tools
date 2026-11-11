using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Dongsidaye.PowerPointTools
{
    public sealed partial class Connect
    {
        internal string BeginDwgInsertion(string id, string optionsJson)
        {
            DwgImportRequest request = RequireDwgRequest(id);
            CheckDwgReady(request);
            var options = DwgJson().Deserialize<DwgInsertionOptions>(optionsJson) ?? new DwgInsertionOptions();
            request.Inserting = true;
            request.InsertionCompleted = false;
            try { dispatcher.BeginInvoke((MethodInvoker)async delegate { await InsertDwgEditable(request, options, true); }); }
            catch { request.Inserting = false; request.InsertionCompleted = true; throw; }
            return DwgJson().Serialize(new { ok = true, pending = true });
        }

        internal string PollDwgInsertion(string id)
        {
            DwgImportRequest request = RequireDwgRequest(id);
            return request.InsertionCompleted ? request.InsertionResult : DwgJson().Serialize(new {
                state = "pending", done = request.InsertionDone, total = request.InsertionTotal });
        }

        private static void CheckDwgReady(DwgImportRequest request)
        {
            if (!request.Completed || request.Cancellation.IsCancellationRequested || request.Inserted || request.Inserting || String.IsNullOrEmpty(request.Result))
                throw new InvalidOperationException("转换尚未完成、已取消或已经插入。");
            var result = DwgJson().Deserialize<Dictionary<string, object>>(request.Result);
            if (!result.ContainsKey("ok") || !Convert.ToBoolean(result["ok"])) throw new InvalidOperationException("请先成功转换 DWG。");
        }

        private async Task<string> InsertDwgEditable(DwgImportRequest request, DwgInsertionOptions options, bool allowYield)
        {
            var created = new List<PowerPoint.Shape>();
            PowerPoint.Shape group = null;
            int nativeCount = 0, recognizedCount = 0;
            var fallback = new HashSet<string>(StringComparer.Ordinal);
            try
            {
                RequireDwgRequest(request.Id);
                request.Cancellation.Token.ThrowIfCancellationRequested();
                if (FileDialogPending) throw new InvalidOperationException("请先完成文件选择。");
                PowerPoint.DocumentWindow window = application.ActiveWindow;
                if (window.ViewType != PowerPoint.PpViewType.ppViewNormal && window.ViewType != PowerPoint.PpViewType.ppViewSlide)
                    throw new InvalidOperationException("请切换到普通幻灯片页面后插入。");
                PowerPoint.Slide slide = window.View.Slide as PowerPoint.Slide;
                if (slide == null) throw new InvalidOperationException("请先选择要插入的幻灯片。");
                double[] size;
                using (var source = new FileStream(request.Output, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (SHA256 hash = SHA256.Create())
                {
                    if (request.OutputHash != Convert.ToBase64String(hash.ComputeHash(source))) throw new IOException("SVG 已变化，请重新转换并检查预览。");
                    size = DwgSvg.Validate(request.Output);
                }
                float pageWidth = request.Presentation.PageSetup.SlideWidth, pageHeight = request.Presentation.PageSetup.SlideHeight;
                double scale = Math.Min(pageWidth * 0.8 / size[0], pageHeight * 0.8 / size[1]);
                float width = (float)(size[0] * scale), height = (float)(size[1] * scale);
                if (width < 0.01f || height < 0.01f) throw new IOException("图纸长宽比过大，请分区导入。");
                float left = (pageWidth - width) / 2, top = (pageHeight - height) / 2;
                string font = options.font ?? request.Font;
                if (!new[] { "Microsoft YaHei", "SimSun", "Arial" }.Contains(font)) font = "Microsoft YaHei";
                DwgOcrSelection[] selections = options.ocrSelections ?? new DwgOcrSelection[0];
                if (selections.Length > 500 || selections.Select(selection => selection.index).Distinct().Count() != selections.Length)
                    throw new InvalidOperationException("OCR 选择过多或包含重复项，请减少本次选择。");
                if (selections.Length > 0 && (!request.OcrEnabled || request.OcrPage == null || request.OcrPage.lines == null))
                    throw new InvalidOperationException("本次导入没有可用的 OCR 结果。");
                var ocrTexts = new List<DwgTextItem>();
                var polygons = new List<DwgOcrPoint[]>();
                foreach (DwgOcrSelection selection in selections)
                {
                    if (selection.index < 0 || selection.index >= request.OcrPage.lines.Length) throw new InvalidOperationException("OCR 选择已失效。");
                    DwgOcrLine line = request.OcrPage.lines[selection.index];
                    ocrTexts.Add(DwgOcrLayout.ToText(request.OcrPage, line, selection.text, "ocr-" + selection.index));
                    polygons.Add(DwgOcrLayout.ToSvgPoints(request.OcrPage, line));
                }
                request.InsertionTotal = request.Texts.Length + ocrTexts.Count;
                foreach (var sourceGroup in request.Texts.GroupBy(item => item.SourceId))
                {
                    var currentGroup = new List<PowerPoint.Shape>();
                    bool failed = false;
                    foreach (DwgTextItem item in sourceGroup)
                    {
                        RequireDwgRequest(request.Id);
                        request.Cancellation.Token.ThrowIfCancellationRequested();
                        if (!failed)
                        {
                            try
                            {
                                PowerPoint.Shape text = DwgTextInsertion.Add(slide, item, left, top, size[2], size[3], scale, font);
                                if (text != null) { currentGroup.Add(text); created.Add(text); text.Name = "DWG_Text_" + Guid.NewGuid().ToString("N"); }
                            }
                            catch (Exception error)
                            {
                                Log("DWG text fallback " + sourceGroup.Key + ": " + error.Message);
                                PowerPoint.Shape unfinished = error.Data["DwgCreatedShape"] as PowerPoint.Shape;
                                if (unfinished != null) { currentGroup.Add(unfinished); created.Add(unfinished); }
                                failed = true;
                                fallback.Add(sourceGroup.Key);
                                foreach (PowerPoint.Shape text in currentGroup) { text.Delete(); created.Remove(text); }
                                currentGroup.Clear();
                            }
                        }
                        request.InsertionDone++;
                        if (allowYield && request.InsertionDone % 8 == 0) await Task.Delay(1);
                    }
                    nativeCount += currentGroup.Count;
                }
                foreach (DwgTextItem item in ocrTexts)
                {
                    RequireDwgRequest(request.Id);
                    request.Cancellation.Token.ThrowIfCancellationRequested();
                    PowerPoint.Shape text;
                    try { text = DwgTextInsertion.Add(slide, item, left, top, size[2], size[3], scale, font); }
                    catch (Exception error) { if (error.Data["DwgCreatedShape"] is PowerPoint.Shape) created.Add((PowerPoint.Shape)error.Data["DwgCreatedShape"]); throw; }
                    if (text != null) { created.Add(text); text.Name = "DWG_OCR_" + Guid.NewGuid().ToString("N"); recognizedCount++; }
                    request.InsertionDone++;
                    if (allowYield && request.InsertionDone % 8 == 0) await Task.Delay(1);
                }
                RequireDwgRequest(request.Id);
                request.Cancellation.Token.ThrowIfCancellationRequested();
                string geometry = DwgOcrLayout.PrepareGeometry(request.PreviewSvg, request.GeometrySvg, fallback, polygons);
                string insertionPath = Path.Combine(request.Folder, "insertion.svg");
                File.WriteAllText(insertionPath, geometry, new UTF8Encoding(false));
                DwgSvg.Validate(insertionPath);
                XNamespace svg = "http://www.w3.org/2000/svg";
                bool hasGeometry = XDocument.Parse(geometry).Descendants().Any(element => element.Name.Namespace == svg &&
                    new[] { "path", "line", "polyline", "polygon", "circle", "ellipse", "rect", "text", "image", "use" }.Contains(element.Name.LocalName) &&
                    !element.Ancestors(svg + "defs").Any());
                if (hasGeometry)
                {
                    PowerPoint.Shape picture = slide.Shapes.AddPicture(insertionPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, left, top, width, height);
                    created.Add(picture);
                    picture.Name = "DWG_Geometry_" + Guid.NewGuid().ToString("N");
                    for (int index = 0; index < created.Count - 1; index++)
                    {
                        picture.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                        if (allowYield && index % 16 == 15) { await Task.Delay(1); RequireDwgRequest(request.Id); request.Cancellation.Token.ThrowIfCancellationRequested(); }
                    }
                }
                if (created.Count == 0) throw new InvalidOperationException("图纸没有可插入的可见内容。");
                group = created.Count == 1 ? created[0] : slide.Shapes.Range(created.Select(shape => (object)shape.Name).ToArray()).Group();
                request.Inserted = true;
                try { group.Name = "DWG_" + Path.GetFileNameWithoutExtension(request.Input); group.AlternativeText = "DWG 模型空间 · 可编辑文字 " + (nativeCount + recognizedCount); group.Select(Office.MsoTriState.msoTrue); }
                catch (Exception error) { Log("DWG selection: " + error.Message); }
                DeleteDwgFolder(request);
                request.InsertionResult = DwgJson().Serialize(new { state = "done", ok = true, textCount = nativeCount, ocrCount = recognizedCount, fallbackCount = fallback.Count });
            }
            catch (Exception error)
            {
                bool clean = true;
                if (group != null) { try { group.Delete(); } catch { clean = false; } }
                else foreach (PowerPoint.Shape shape in created.AsEnumerable().Reverse()) { try { shape.Delete(); } catch { clean = false; } }
                if (!clean) request.Inserted = true;
                request.InsertionResult = DwgJson().Serialize(new { state = "done", ok = false,
                    cancelled = error is OperationCanceledException,
                    error = clean ? error.Message : "导入中断且部分对象未能清理，请检查当前页面后再操作。" });
            }
            finally { request.Inserting = false; request.InsertionCompleted = true; if (request.Cancellation.IsCancellationRequested) DeleteDwgFolder(request); }
            return request.InsertionResult;
        }
    }
}
