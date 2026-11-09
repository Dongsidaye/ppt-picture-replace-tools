using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Dongsidaye.PowerPointTools
{
    // Office calls this standard dual interface through its native vtable.
    // IDispatch-only exposes seven slots and cannot safely receive the five
    // subsequent IDTExtensibility2 lifecycle calls (SAFEARRAY** for custom).
    [ComVisible(true), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), TypeLibType(4160), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, [In] int connectMode, [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection([In] int removeMode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComVisible(true), Guid("5BA1B2A7-D0BB-4471-9C03-84A520F68DA1"), ProgId("Dongsidaye.PowerPointTools"), ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class Connect : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        internal static Connect Current;
        internal readonly string Root = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        private PowerPoint.Application application;
        private Office.COMAddIn addin;
        private Office.ICTPFactory factory;
        private Office.CustomTaskPane pane;
        private ToolboxPane paneControl;
        private WebView2 parkedBrowser;
        private int paneWindowHwnd;
        private bool migratingPane;
        private Control dispatcher;
        private Timer bootstrapTimer;
        private Bridge bridge;
        private bool connected;
        private bool creatingPane;
        private bool paneCreationFailed;
        private bool browserReady;
        private bool browserRecoveryPending;
        private bool discardingBrowser;
        private int browserGeneration;
        private bool stopping;
        private bool requestedVisible;
        private string lastError = "";
        private string reportPath;
        private bool selfTestActive;
        private int bootstrapAttempts;
        private readonly Queue<string> pendingScripts = new Queue<string>();
        private readonly HashSet<string> pendingEvents = new HashSet<string>(StringComparer.Ordinal);
        private readonly Dictionary<string, object> images = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        private readonly List<Image> imageSources = new List<Image>();
        private readonly List<PowerPoint.Presentation> ownedScratch = new List<PowerPoint.Presentation>();
        private sealed class FileDialogRequest
        {
            internal string Id, Kind, Title, Filter, DocumentKey, Result;
            internal int Generation;
            internal PowerPoint.Presentation Presentation;
            internal bool Completed;
        }
        private FileDialogRequest activeFileDialog;
        private readonly Dictionary<string, FileDialogRequest> fileDialogs = new Dictionary<string, FileDialogRequest>();
        private readonly Queue<string> completedFileDialogs = new Queue<string>();
        private sealed class PictureSnapshotRequest
        {
            internal string Id, DocumentKey, Folder, Result;
            internal int Generation;
            internal bool Completed;
            internal PowerPoint.Presentation Presentation;
        }
        private PictureSnapshotRequest pictureSnapshot;
        private readonly Dictionary<string, PictureSnapshotRequest> pictureSnapshots = new Dictionary<string, PictureSnapshotRequest>();
        private readonly Queue<string> pictureSnapshotHistory = new Queue<string>();
        private Dictionary<string, string> iconMap;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();

        internal object ApplicationObject { get { return application; } }
        internal Bridge NativeBridge { get { return bridge; } }
        internal int BrowserGeneration { get { return browserGeneration; } }
        internal bool IsCurrentBrowser(int generation) { return connected && !stopping && generation == browserGeneration; }
        internal string DocumentKey
        {
            get
            {
                try
                {
                    int hwnd = application.ActiveWindow.HWND;
                    return hwnd == 0 ? "" : "window:" + hwnd.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
                catch { return ""; }
            }
        }
        internal string LogPath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Dongsidaye", "PowerPointTools", "host.log"); } }

        public void OnConnection(object app, int connectMode, object addInInst, ref Array custom)
        {
            Log("OnConnection entered, mode=" + connectMode);
            try
            {
                Current = this;
                application = (PowerPoint.Application)app;
                addin = addInInst as Office.COMAddIn;
                if (addin != null) addin.Object = this;
                connected = true;
                stopping = false;
                dispatcher = new Control();
                IntPtr forceHandle = dispatcher.Handle;
                bridge = new Bridge(this);
                application.WindowSelectionChange += OnSelection;
                application.PresentationOpen += OnOpen;
                ((PowerPoint.EApplication_Event)application).NewPresentation += OnNew;
                application.WindowActivate += OnActivate;
                application.PresentationClose += OnClose;
                bootstrapTimer = new Timer();
                bootstrapTimer.Interval = 500;
                bootstrapTimer.Tick += Bootstrap;
                bootstrapTimer.Start();
                Log("Connected to PowerPoint " + application.Version);
            }
            catch (Exception ex) { Fail("OnConnection", ex); }
        }

        public void CTPFactoryAvailable(Office.ICTPFactory value)
        {
            factory = value;
            Log("Custom task pane factory available");
            QueueOnUi(delegate { EnsurePane(); });
        }

        public void OnDisconnection(int removeMode, ref Array custom) { Disconnect(); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { QueueOnUi(delegate { EnsurePane(); }); }
        public void OnBeginShutdown(ref Array custom) { Disconnect(); }

        private void Bootstrap(object sender, EventArgs args)
        {
            bootstrapAttempts++;
            EnsurePane();
            if (pane != null || bootstrapAttempts >= 20)
            {
                bootstrapTimer.Stop();
                if (pane == null && factory == null) Fail("Bootstrap", new InvalidOperationException("PowerPoint did not supply a CustomTaskPane factory. No task pane was created."));
            }
        }

        private void Disconnect()
        {
            if (stopping) return;
            stopping = true;
            connected = false;
            browserReady = false;
            browserGeneration++;
            if (activeFileDialog != null) CompleteFileDialog(activeFileDialog, false, true, "", "插件已断开，文件选择已取消。");
            if (pictureSnapshot != null && !pictureSnapshot.Completed)
            {
                pictureSnapshot.Result = "{\"state\":\"done\",\"error\":\"插件已重载，请重新读取图片。\"}";
                pictureSnapshot.Completed = true;
            }
            try
            {
                if (bootstrapTimer != null) { bootstrapTimer.Stop(); bootstrapTimer.Dispose(); bootstrapTimer = null; }
                if (application != null)
                {
                    application.WindowSelectionChange -= OnSelection;
                    application.PresentationOpen -= OnOpen;
                    ((PowerPoint.EApplication_Event)application).NewPresentation -= OnNew;
                    application.WindowActivate -= OnActivate;
                    application.PresentationClose -= OnClose;
                }
            }
            catch (Exception ex) { Fail("Detach events", ex); }
            CloseOwnedScratch();
            try { if (pane != null) pane.Delete(); } catch (Exception ex) { Log("Delete pane: " + ex.Message); }
            try { if (paneControl != null) paneControl.Dispose(); } catch (Exception ex) { Log("Dispose pane: " + ex.Message); }
            try { if (parkedBrowser != null) parkedBrowser.Dispose(); } catch { }
            parkedBrowser = null;
            foreach (Image image in imageSources) { try { image.Dispose(); } catch { } }
            imageSources.Clear();
            images.Clear();
            pendingScripts.Clear();
            pendingEvents.Clear();
            paneControl = null;
            pane = null;
            factory = null;
            application = null;
            try { if (addin != null) addin.Object = null; } catch { }
            addin = null;
            reportPath = null;
            selfTestActive = false;
            try { if (dispatcher != null) dispatcher.Dispose(); } catch { }
            dispatcher = null;
            if (Current == this) Current = null;
            Log("Disconnected");
        }

        internal void TrackScratch(object value)
        {
            PowerPoint.Presentation presentation = value as PowerPoint.Presentation;
            if (presentation == null || !Marshal.IsComObject(value)) throw new ArgumentException("Scratch tracking requires a native PowerPoint presentation.", "value");
            if (stopping) throw new InvalidOperationException("The toolbox is disconnecting.");
            foreach (PowerPoint.Presentation existing in ownedScratch)
                if (SameComIdentity(existing, presentation)) return;
            ownedScratch.Add(presentation);
        }

        internal void ReleaseScratch(object value)
        {
            for (int index = ownedScratch.Count - 1; index >= 0; index--)
                if (SameComIdentity(ownedScratch[index], value)) ownedScratch.RemoveAt(index);
        }

        private static bool SameComIdentity(object first, object second)
        {
            if (Object.ReferenceEquals(first, second)) return true;
            if (first == null || second == null || !Marshal.IsComObject(first) || !Marshal.IsComObject(second)) return false;
            IntPtr left = IntPtr.Zero;
            IntPtr right = IntPtr.Zero;
            try
            {
                left = Marshal.GetIUnknownForObject(first);
                right = Marshal.GetIUnknownForObject(second);
                return left == right;
            }
            catch { return false; }
            finally
            {
                if (left != IntPtr.Zero) Marshal.Release(left);
                if (right != IntPtr.Zero) Marshal.Release(right);
            }
        }

        private void CloseOwnedScratch()
        {
            // The engine explicitly registers only presentations that it creates.
            // Never enumerate all open presentations, and never alter the Saved
            // flag of a document outside this ownership list.
            PowerPoint.Presentation[] scratch = ownedScratch.ToArray();
            ownedScratch.Clear();
            foreach (PowerPoint.Presentation presentation in scratch)
            {
                try
                {
                    presentation.Saved = Office.MsoTriState.msoTrue;
                    presentation.Close();
                }
                catch (Exception ex) { Log("Close owned scratch: " + ex.Message); }
            }
        }

        internal bool FileDialogPending { get { return activeFileDialog != null && !activeFileDialog.Completed; } }

        internal string BeginPictureSnapshot()
        {
            if (!connected || stopping || dispatcher == null || dispatcher.IsDisposed) throw new InvalidOperationException("插件尚未连接。");
            string key = DocumentKey;
            if (String.IsNullOrEmpty(key)) throw new InvalidOperationException("请先打开文稿。");
            if (pictureSnapshot != null && !pictureSnapshot.Completed)
            {
                if (pictureSnapshot.DocumentKey == key && pictureSnapshot.Generation == browserGeneration) return pictureSnapshot.Id;
                throw new InvalidOperationException("正在读取另一份文稿，请稍后重试。");
            }
            var request = new PictureSnapshotRequest { Id = Guid.NewGuid().ToString("N"), DocumentKey = key, Generation = browserGeneration, Presentation = application.ActiveWindow.Presentation };
            pictureSnapshot = request; pictureSnapshots.Add(request.Id, request);
            try { dispatcher.BeginInvoke((MethodInvoker)delegate { BuildPictureSnapshot(request); }); }
            catch (Exception ex) { request.Result = json.Serialize(new { state = "done", error = ex.Message }); request.Completed = true; }
            return request.Id;
        }

        internal string PollPictureSnapshot(string id)
        {
            PictureSnapshotRequest request;
            if (!pictureSnapshots.TryGetValue(id ?? "", out request)) throw new InvalidOperationException("图片读取请求已过期。");
            return request.Completed ? request.Result : "{\"state\":\"pending\"}";
        }

        private async void BuildPictureSnapshot(PictureSnapshotRequest request)
        {
            if (request.Completed) return;
            try
            {
                if (!IsCurrentBrowser(request.Generation) || DocumentKey != request.DocumentKey) throw new InvalidOperationException("文稿已切换，请重新读取图片。");
                request.Folder = Path.Combine(Path.GetTempPath(), "DongsidayePictureSources", request.Id);
                Directory.CreateDirectory(request.Folder);
                string snapshot = Path.Combine(request.Folder, "source.pptx");
                PowerPoint.Presentation source = request.Presentation;
                if (!SameComIdentity(source, application.ActiveWindow.Presentation)) throw new InvalidOperationException("文稿已切换。");
                Office.MsoTriState saved = source.Saved;
                source.SaveCopyAs(snapshot, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoFalse);
                if (source.Saved != saved) throw new InvalidOperationException("文稿保存状态发生变化，已停止本次图片读取。");
                var entries = await Task.Run(delegate { return PictureSourceManifest.Read(snapshot, request.Folder); });
                try { File.Delete(snapshot); } catch { }
                if (!IsCurrentBrowser(request.Generation) || DocumentKey != request.DocumentKey || !SameComIdentity(source, application.ActiveWindow.Presentation) || source.Saved != saved) throw new InvalidOperationException("文稿已切换或修改，已丢弃旧图片读取结果。");
                request.Result = json.Serialize(new { state = "done", documentKey = request.DocumentKey, entries = entries });
            }
            catch (Exception ex) { request.Result = json.Serialize(new { state = "done", error = ex.Message }); }
            finally
            {
                request.Completed = true;
                request.Presentation = null;
                pictureSnapshotHistory.Enqueue(request.Id);
                while (pictureSnapshotHistory.Count > 2)
                {
                    string expired = pictureSnapshotHistory.Dequeue();
                    PictureSnapshotRequest old;
                    if (pictureSnapshots.TryGetValue(expired, out old))
                    {
                        pictureSnapshots.Remove(expired);
                        try { if (!String.IsNullOrEmpty(old.Folder)) Directory.Delete(old.Folder, true); } catch { }
                    }
                }
            }
        }

        internal string BeginFileDialog(string requestJson)
        {
            if (!connected || stopping || dispatcher == null || dispatcher.IsDisposed) throw new InvalidOperationException("插件尚未连接。");
            if (FileDialogPending) throw new InvalidOperationException("请先完成或取消当前文件选择。");
            Dictionary<string, object> options = json.Deserialize<Dictionary<string, object>>(requestJson);
            if (options == null) throw new ArgumentException("文件选择参数无效。");
            object value;
            string kind = options.TryGetValue("kind", out value) ? Convert.ToString(value) : "file";
            if (kind != "file" && kind != "folder") throw new ArgumentException("不支持的文件选择类型。");
            string expectedKey = options.TryGetValue("documentKey", out value) ? Convert.ToString(value) : "";
            string key = DocumentKey;
            if (String.IsNullOrEmpty(key) || (!String.IsNullOrEmpty(expectedKey) && expectedKey != key)) throw new InvalidOperationException("文稿已切换，请重新选择文件。");
            var request = new FileDialogRequest
            {
                Id = Guid.NewGuid().ToString("N"), Kind = kind, DocumentKey = key,
                Generation = browserGeneration, Presentation = application.ActiveWindow.Presentation,
                Title = options.TryGetValue("title", out value) ? Convert.ToString(value) : "选择文件",
                Filter = options.TryGetValue("filter", out value) ? Convert.ToString(value) : "*.*"
            };
            if (request.Title.Length > 256 || request.Filter.Length > 2048) throw new ArgumentException("文件选择参数过长。");
            activeFileDialog = request;
            fileDialogs.Add(request.Id, request);
            try
            {
                // FileDialog.Show must run after the synchronous WebView2 host call
                // returns. A nested modal loop inside that callback can deadlock JS.
                dispatcher.BeginInvoke((MethodInvoker)delegate { ShowNativeFileDialog(request); });
            }
            catch (Exception ex) { CompleteFileDialog(request, false, false, "", ex.Message); }
            return request.Id;
        }

        internal string PollFileDialog(string requestId)
        {
            FileDialogRequest request;
            if (!fileDialogs.TryGetValue(requestId ?? "", out request)) throw new InvalidOperationException("文件选择请求已过期，请重新选择。");
            return request.Completed ? request.Result : "{\"state\":\"pending\"}";
        }

        private bool FileDialogDocumentMatches(FileDialogRequest request)
        {
            try { return IsCurrentBrowser(request.Generation) && request.DocumentKey == DocumentKey && SameComIdentity(request.Presentation, application.ActiveWindow.Presentation); }
            catch { return false; }
        }

        private void ShowNativeFileDialog(FileDialogRequest request)
        {
            if (request.Completed) return;
            try
            {
                if (!FileDialogDocumentMatches(request)) throw new InvalidOperationException("文稿已切换或插件已重载，请重新选择文件。");
                Office.FileDialog dialog = application.FileDialog[request.Kind == "folder" ? Office.MsoFileDialogType.msoFileDialogFolderPicker : Office.MsoFileDialogType.msoFileDialogFilePicker];
                dialog.Title = request.Title;
                dialog.AllowMultiSelect = false;
                if (request.Kind == "file")
                {
                    dialog.Filters.Clear();
                    dialog.Filters.Add("文件", String.IsNullOrWhiteSpace(request.Filter) ? "*.*" : request.Filter, Type.Missing);
                }
                int choice = dialog.Show();
                if (!FileDialogDocumentMatches(request)) throw new InvalidOperationException("文稿已切换，已丢弃本次文件选择；请在当前文稿中重试。");
                if (choice != -1 || dialog.SelectedItems.Count == 0) CompleteFileDialog(request, false, true, "", "");
                else CompleteFileDialog(request, true, false, dialog.SelectedItems.Item(1), "");
            }
            catch (Exception ex) { CompleteFileDialog(request, false, false, "", ex.Message); }
            finally
            {
                if (!request.Completed) CompleteFileDialog(request, false, false, "", "文件选择未完成。");
                if (Object.ReferenceEquals(activeFileDialog, request)) activeFileDialog = null;
                request.Presentation = null;
            }
        }

        private void CompleteFileDialog(FileDialogRequest request, bool ok, bool cancelled, string path, string error)
        {
            if (request.Completed) return;
            request.Result = json.Serialize(new { state = "done", ok = ok, cancelled = cancelled, path = path ?? "", error = error ?? "" });
            request.Completed = true;
            request.Presentation = null;
            if (Object.ReferenceEquals(activeFileDialog, request)) activeFileDialog = null;
            completedFileDialogs.Enqueue(request.Id);
            while (completedFileDialogs.Count > 16) fileDialogs.Remove(completedFileDialogs.Dequeue());
        }

        private void EnsurePane()
        {
            if (!connected || stopping || creatingPane || paneCreationFailed || browserRecoveryPending || discardingBrowser || factory == null) return;
            // Office cannot attach a native task pane until a document window exists.
            // PresentationOpen/WindowActivate will retry; do not log the same startup
            // failure every timer tick while PowerPoint is displaying its start page.
            PowerPoint.DocumentWindow activeWindow;
            int activeHwnd;
            try
            {
                if (application.Windows.Count == 0) return;
                activeWindow = application.ActiveWindow;
                activeHwnd = activeWindow.HWND;
                if (activeHwnd == 0) return;
            }
            catch { return; }
            if (pane != null && paneControl != null)
            {
                if (activeHwnd != paneWindowHwnd) MigratePane();
                return;
            }
            creatingPane = true;
            try
            {
                Log("Creating native CustomTaskPane for window " + activeHwnd);
                pane = factory.CreateCTP("Dongsidaye.PowerPointToolsPane", "东四大爷的工具箱", activeWindow);
                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                pane.Width = 450;
                paneControl = pane.ContentControl as ToolboxPane;
                if (paneControl == null) throw new InvalidOperationException("The native CustomTaskPane did not return the toolbox control.");
                paneWindowHwnd = activeHwnd;
                pane.Visible = requestedVisible;
                if (parkedBrowser != null)
                {
                    paneControl.AttachBrowser(parkedBrowser, this);
                    parkedBrowser = null;
                    DrainScripts();
                }
                else paneControl.Initialize(this);
                Log("Native CustomTaskPane created");
            }
            catch (Exception ex)
            {
                try { if (pane != null) pane.Delete(); } catch { }
                pane = null;
                paneControl = null;
                paneCreationFailed = true;
                Fail("Create CustomTaskPane", ex);
            }
            finally { creatingPane = false; }
        }

        private async void MigratePane()
        {
            if (migratingPane || !browserReady || paneControl == null || stopping) return;
            migratingPane = true;
            int generation = browserGeneration;
            Office.CustomTaskPane nextPane = null;
            ToolboxPane nextControl = null;
            WebView2 sharedBrowser = null;
            ToolboxPane oldControl = paneControl;
            Office.CustomTaskPane oldPane = pane;
            try
            {
                // This executes after the originating COM call has returned to JS,
                // ending the previous geometry session before its UI is reparented.
                await oldControl.RunScript("if(typeof window.__pptEmit==='function'){window.__pptEmit('WindowActivate');}");
                if (!IsCurrentBrowser(generation) || application.Windows.Count == 0 || oldControl != paneControl) return;
                PowerPoint.DocumentWindow target = application.ActiveWindow;
                int targetHwnd = target.HWND;
                if (targetHwnd == 0 || targetHwnd == paneWindowHwnd) return;
                nextPane = factory.CreateCTP("Dongsidaye.PowerPointToolsPane", "东四大爷的工具箱", target);
                nextPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                nextPane.Width = 450;
                nextPane.Visible = false;
                nextControl = nextPane.ContentControl as ToolboxPane;
                if (nextControl == null) throw new InvalidOperationException("The destination window did not return a toolbox control.");
                sharedBrowser = oldControl.DetachBrowser();
                if (sharedBrowser == null) throw new InvalidOperationException("The existing toolbox browser is unavailable for migration.");
                nextControl.AttachBrowser(sharedBrowser, this);
                nextPane.Visible = requestedVisible;
                pane = nextPane;
                paneControl = nextControl;
                paneWindowHwnd = targetHwnd;
                nextPane = null;
                nextControl = null;
                sharedBrowser = null;
                try { oldPane.Delete(); } catch (Exception ex) { Log("Remove old window pane: " + ex.Message); }
                try { oldControl.Dispose(); } catch (Exception ex) { Log("Dispose old pane control: " + ex.Message); }
                Log("Moved persistent toolbox to PowerPoint window " + targetHwnd);
            }
            catch (Exception ex)
            {
                if (!IsCurrentBrowser(generation))
                {
                    // A process failure/disconnect can arrive while Office pumps
                    // messages during CTP creation. Do not reattach that generation.
                    try
                    {
                        WebView2 obsolete = nextControl == null ? sharedBrowser : nextControl.DetachBrowser();
                        if (obsolete != null) obsolete.Dispose();
                    }
                    catch { }
                    try { if (nextPane != null) nextPane.Delete(); } catch { }
                    return;
                }
                // Restore the original live browser if creation/reparenting fails.
                if (nextControl != null)
                {
                    WebView2 moved = nextControl.DetachBrowser();
                    if (moved != null) sharedBrowser = moved;
                }
                if (sharedBrowser != null)
                {
                    try
                    {
                        if (oldControl != null && !oldControl.IsDisposed) oldControl.AttachBrowser(sharedBrowser, this);
                        else ParkBrowser(oldControl, sharedBrowser);
                    }
                    catch (Exception restore) { Fail("Restore browser after pane migration", restore); }
                }
                try { if (nextPane != null) nextPane.Delete(); } catch { }
                Fail("Move CustomTaskPane to active window", ex);
            }
            finally { if (generation == browserGeneration) migratingPane = false; }
        }

        internal bool ParkBrowser(ToolboxPane source, WebView2 browser)
        {
            if (stopping || discardingBrowser || browserRecoveryPending || dispatcher == null || dispatcher.IsDisposed) return false;
            try
            {
                if (paneControl == source) { paneControl = null; pane = null; paneWindowHwnd = 0; }
                if (browser.CoreWebView2 == null)
                {
                    browserReady = false;
                    Log("Document window closed before browser initialization; a new window will retry initialization");
                    return false;
                }
                // Closing an Office window disposes its ActiveX control. Keep the
                // browser on a hidden, process-owned parent until a new window exists.
                if (browser.Parent != null) browser.Parent.Controls.Remove(browser);
                dispatcher.Controls.Add(browser);
                parkedBrowser = browser;
                Log("Parked persistent toolbox after document window closure");
                return true;
            }
            catch (Exception ex) { Fail("Keep toolbox browser alive", ex); return false; }
        }

        internal void ShowPane(string title)
        {
            requestedVisible = true;
            int generation = browserGeneration;
            QueueOnUi(delegate
            {
                if (!IsCurrentBrowser(generation)) return;
                RecoverBrowserForUserRequest();
                EnsurePane();
                if (pane == null) { Fail("ShowPane", new InvalidOperationException("Native CustomTaskPane is unavailable.")); return; }
                pane.Visible = true;
            });
        }

        internal void HidePane()
        {
            requestedVisible = false;
            int generation = browserGeneration;
            QueueOnUi(delegate { if (IsCurrentBrowser(generation) && pane != null) pane.Visible = false; });
        }

        internal void BrowserReady(int generation)
        {
            if (!IsCurrentBrowser(generation) || browserRecoveryPending) return;
            browserReady = true;
            lastError = "";
            Log("WebView2 local toolbox ready");
            FlushEvents();
            DrainScripts();
            EnsurePane();
        }

        internal void BrowserFailed(int generation, string error)
        {
            if (!IsCurrentBrowser(generation)) return;
            browserReady = false;
            browserRecoveryPending = true;
            browserGeneration++;
            // Abandon both already-enqueued actions and BeginInvoke callbacks from
            // this generation. Recovery never replays partially executed edits.
            pendingScripts.Clear();
            pendingEvents.Clear();
            executingScripts = false;
            migratingPane = false;
            selfTestActive = false;
            reportPath = null;
            Fail("WebView2", new InvalidOperationException(error));
        }

        private void RecoverBrowserForUserRequest()
        {
            if (!browserRecoveryPending || stopping || !connected) return;
            // Only a fresh Ribbon action, self-test request, or explicit pane open
            // invokes recovery. Background Office events must not replay work.
            discardingBrowser = true;
            Office.CustomTaskPane failedPane = pane;
            ToolboxPane failedControl = paneControl;
            WebView2 failedParked = parkedBrowser;
            pane = null;
            paneControl = null;
            parkedBrowser = null;
            paneWindowHwnd = 0;
            try
            {
                try { if (failedPane != null) failedPane.Delete(); } catch (Exception ex) { Log("Remove failed pane: " + ex.Message); }
                try { if (failedControl != null) failedControl.Dispose(); } catch (Exception ex) { Log("Dispose failed browser control: " + ex.Message); }
                try { if (failedParked != null) failedParked.Dispose(); } catch (Exception ex) { Log("Dispose failed parked browser: " + ex.Message); }
                CloseOwnedScratch();
                paneCreationFailed = false;
                browserRecoveryPending = false;
                Log("Rebuilding toolbox after browser failure; previous commands were discarded");
            }
            finally { discardingBrowser = false; }
        }

        internal void QueueOnUi(Action work)
        {
            if (stopping || dispatcher == null || dispatcher.IsDisposed) return;
            try
            {
                dispatcher.BeginInvoke((MethodInvoker)delegate
                {
                    if (stopping) return;
                    try { work(); } catch (Exception ex) { Fail("UI callback", ex); }
                });
            }
            catch (Exception ex) { if (!stopping) Fail("Queue UI callback", ex); }
        }

        private void QueueScript(string script)
        {
            int generation = browserGeneration;
            QueueOnUi(delegate
            {
                if (!IsCurrentBrowser(generation)) return;
                RecoverBrowserForUserRequest();
                // Preserve previously accepted commands instead of silently dropping
                // an action or a self-test while the browser starts up.
                if (pendingScripts.Count >= 256)
                {
                    Fail("Queue command", new InvalidOperationException("Too many pending toolbox commands. Wait for initialization."));
                    return;
                }
                pendingScripts.Enqueue(script);
                EnsurePane();
                DrainScripts();
            });
        }

        private bool executingScripts;
        private async void DrainScripts()
        {
            if (executingScripts || !browserReady || paneControl == null || stopping) return;
            executingScripts = true;
            int generation = browserGeneration;
            try
            {
                while (pendingScripts.Count > 0 && browserReady && IsCurrentBrowser(generation))
                {
                    string script = pendingScripts.Dequeue();
                    await paneControl.RunScript(script);
                }
            }
            catch (Exception ex)
            {
                if (IsCurrentBrowser(generation)) BrowserFailed(generation, "Execute toolbox callback: " + ex.Message);
            }
            finally { if (generation == browserGeneration) executingScripts = false; }
        }

        private void Emit(string name)
        {
            if (stopping || !connected || browserRecoveryPending || discardingBrowser || !pendingEvents.Add(name)) return;
            QueueOnUi(delegate { FlushEvents(); });
        }
        private void FlushEvents()
        {
            if (!browserReady || stopping) return;
            foreach (string name in pendingEvents)
                pendingScripts.Enqueue("if(typeof window.__pptEmit==='function'){window.__pptEmit(" + json.Serialize(name) + ");}");
            pendingEvents.Clear();
            DrainScripts();
        }
        private void RetryForDocument()
        {
            QueueOnUi(delegate
            {
                paneCreationFailed = false;
                EnsurePane();
            });
        }
        private void OnSelection(PowerPoint.Selection selection) { Emit("WindowSelectionChange"); }
        private void OnOpen(PowerPoint.Presentation presentation) { RetryForDocument(); Emit("PresentationOpen"); }
        private void OnNew(PowerPoint.Presentation presentation) { RetryForDocument(); Emit("NewPresentation"); }
        private void OnActivate(PowerPoint.Presentation presentation, PowerPoint.DocumentWindow window) { RetryForDocument(); Emit("WindowActivate"); }
        private void OnClose(PowerPoint.Presentation presentation) { Emit("PresentationClose"); }

        public string GetCustomUI(string ribbonId)
        {
            try { return File.ReadAllText(Path.Combine(Root, "ui", "ribbon.xml"), Encoding.UTF8); }
            catch (Exception ex) { Fail("Read Ribbon", ex); return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"/>"; }
        }

        public void OnAction(Office.IRibbonControl control)
        {
            try
            {
                string id = control.Id;
                string documentKey = DocumentKey;
                QueueScript("if(typeof window.__pptDispatch==='function'){window.__pptDispatch(" + json.Serialize(id) + "," + json.Serialize(documentKey) + ");}");
            }
            catch (Exception ex) { Fail("Ribbon action", ex); }
        }

        public void OnLoad(Office.IRibbonUI ribbon) { }
        public void OnAddInLoad(Office.IRibbonUI ribbon) { }

        public object GetImage(Office.IRibbonControl control)
        {
            try
            {
                string id = control.Id;
                object result;
                if (images.TryGetValue(id, out result)) return result;
                if (iconMap == null) iconMap = json.Deserialize<Dictionary<string, string>>(File.ReadAllText(Path.Combine(Root, "ui", "icon-map.json"), Encoding.UTF8));
                string relative;
                if (iconMap == null || !iconMap.TryGetValue(id, out relative)) return null;
                string uiRoot = Path.GetFullPath(Path.Combine(Root, "ui")) + Path.DirectorySeparatorChar;
                string path = Path.GetFullPath(Path.Combine(uiRoot, relative));
                if (!path.StartsWith(uiRoot, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Icon is outside the UI directory.");
                Bitmap bitmap;
                using (Image source = Image.FromFile(path)) bitmap = new Bitmap(source);
                imageSources.Add(bitmap);
                result = PictureConverter.Convert(bitmap);
                images[id] = result;
                return result;
            }
            catch (Exception ex) { Fail("Ribbon icon", ex); return null; }
        }

        public string Status()
        {
            return json.Serialize(new { connected = connected, factoryAvailable = factory != null, paneCreated = pane != null, browserReady = browserReady, windowHwnd = paneWindowHwnd, browserParked = parkedBrowser != null, selfTestRunning = selfTestActive, pending = pendingScripts.Count, error = lastError, rootFolder = Root, logPath = LogPath, version = Assembly.GetExecutingAssembly().GetName().Version.ToString() });
        }

        public string RunSelfTest(string outputPath)
        {
            try
            {
                if (String.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("A report path is required.");
                if (selfTestActive) throw new InvalidOperationException("A self-test is already running; wait for its report before starting another.");
                string normalized = Path.GetFullPath(outputPath);
                string scriptPath = Path.Combine(Root, "ui", "selftest.js");
                string script = File.ReadAllText(scriptPath, Encoding.UTF8);
                reportPath = normalized;
                selfTestActive = true;
                QueueScript(script);
                return "queued";
            }
            catch (Exception ex) { Fail("RunSelfTest", ex); return "error: " + ex.Message; }
        }

        internal void Report(string value)
        {
            if (!selfTestActive || String.IsNullOrEmpty(reportPath)) throw new InvalidOperationException("No active self-test report destination is configured.");
            json.DeserializeObject(value); // Do not turn arbitrary script output into a successful report.
            string destination = reportPath;
            string folder = Path.GetDirectoryName(destination);
            if (!String.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);
            File.WriteAllText(destination, value, new UTF8Encoding(false));
            selfTestActive = false;
            reportPath = null;
            Log("Self-test report written");
        }

        internal void Fail(string stage, Exception error)
        {
            lastError = stage + ": " + error.GetType().Name + ": " + error.Message;
            Log(lastError);
        }

        internal void Log(string text)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
                if (File.Exists(LogPath) && new FileInfo(LogPath).Length > 2 * 1024 * 1024) File.Move(LogPath, LogPath + "." + DateTime.UtcNow.Ticks + ".old");
                File.AppendAllText(LogPath, DateTime.UtcNow.ToString("o") + " " + text + Environment.NewLine, new UTF8Encoding(false));
            }
            catch { }
        }
    }

    [ComVisible(true), Guid("66A31E92-C378-4F22-A1F9-895DCDF065C9"), ProgId("Dongsidaye.PowerPointToolsPane"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class ToolboxPane : UserControl
    {
        private static readonly object LoaderConfigurationLock = new object();
        private static bool loaderConfigured;
        private WebView2 browser;
        private Connect owner;
        private bool initializing;
        private bool initialized;
        public ToolboxPane() { Dock = DockStyle.Fill; BackColor = Color.White; }

        internal async void Initialize(Connect host)
        {
            if (initializing || initialized || IsDisposed) return;
            initializing = true;
            owner = host;
            int generation = host.BrowserGeneration;
            try
            {
                owner.Log("Initializing local WebView2 toolbox");
                browser = new WebView2();
                browser.Dock = DockStyle.Fill;
                Controls.Add(browser);
                IntPtr forceHandle = Handle;
                IntPtr browserHandle = browser.Handle;
                string dataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Dongsidaye", "PowerPointTools", "WebView2");
                ConfigureLoader(owner);
                CoreWebView2Environment environment = await CoreWebView2Environment.CreateAsync(null, dataFolder, null);
                if (IsDisposed || browser == null || !owner.IsCurrentBrowser(generation)) return;
                await browser.EnsureCoreWebView2Async(environment);
                if (IsDisposed || browser == null || !owner.IsCurrentBrowser(generation)) return;
                CoreWebView2 core = browser.CoreWebView2;
                core.Settings.AreDefaultContextMenusEnabled = false;
                core.Settings.AreDevToolsEnabled = false;
                core.Settings.IsStatusBarEnabled = false;
                core.Settings.AreBrowserAcceleratorKeysEnabled = false;
                core.Settings.IsGeneralAutofillEnabled = false;
                core.Settings.IsPasswordAutosaveEnabled = false;
                core.SetVirtualHostNameToFolderMapping("ppt-toolbox.local", Path.Combine(owner.Root, "ui"), CoreWebView2HostResourceAccessKind.DenyCors);
                core.NavigationStarting += delegate(object sender, CoreWebView2NavigationStartingEventArgs args)
                {
                    Uri uri;
                    if (!Uri.TryCreate(args.Uri, UriKind.Absolute, out uri) || uri.Scheme != "https" || !String.Equals(uri.Host, "ppt-toolbox.local", StringComparison.OrdinalIgnoreCase) || !uri.IsDefaultPort)
                    {
                        args.Cancel = true;
                        owner.Log("Blocked navigation outside local toolbox");
                    }
                };
                core.FrameNavigationStarting += delegate(object sender, CoreWebView2NavigationStartingEventArgs args) { args.Cancel = true; };
                core.NewWindowRequested += delegate(object sender, CoreWebView2NewWindowRequestedEventArgs args) { args.Handled = true; };
                core.PermissionRequested += delegate(object sender, CoreWebView2PermissionRequestedEventArgs args) { args.State = CoreWebView2PermissionState.Deny; };
                core.DownloadStarting += delegate(object sender, CoreWebView2DownloadStartingEventArgs args) { args.Cancel = true; };
                core.ProcessFailed += delegate(object sender, CoreWebView2ProcessFailedEventArgs args) { owner.BrowserFailed(generation, "Browser process failed: " + args.ProcessFailedKind); };
                core.NavigationCompleted += async delegate(object sender, CoreWebView2NavigationCompletedEventArgs args)
                {
                    try
                    {
                        if (!owner.IsCurrentBrowser(generation)) return;
                        if (!args.IsSuccess) { owner.BrowserFailed(generation, "Local UI navigation failed: " + args.WebErrorStatus); return; }
                        string ready = await core.ExecuteScriptAsync("Boolean(window.__pptEngineReady===true && typeof window.__pptDispatch==='function' && typeof window.__pptEmit==='function' && window.WpsPictureReplace)");
                        if (ready == "true") owner.BrowserReady(generation);
                        else owner.BrowserFailed(generation, "Local UI loaded but the PowerPoint engine did not initialize. Check host-adapter and host-bootstrap errors.");
                    }
                    catch (Exception ex) { owner.BrowserFailed(generation, "Verify PowerPoint engine: " + ex.Message); }
                };
                core.AddHostObjectToScript("pptApplication", owner.ApplicationObject);
                core.AddHostObjectToScript("pptHost", owner.NativeBridge);
                initialized = true;
                core.Navigate("https://ppt-toolbox.local/taskpane.html#idle");
            }
            catch (Exception ex) { owner.BrowserFailed(generation, "Initialize WebView2: " + ex.Message); }
            finally { initializing = false; }
        }

        private static void ConfigureLoader(Connect host)
        {
            lock (LoaderConfigurationLock)
            {
                if (loaderConfigured) return;
                try
                {
                    CoreWebView2Environment.SetLoaderDllFolderPath(Path.Combine(host.Root, Environment.Is64BitProcess ? "x64" : "x86"));
                }
                catch (InvalidOperationException)
                {
                    // The SDK documents this exception specifically for an already
                    // loaded WebView2Loader. Another Office add-in or a previous
                    // toolbox instance may have initialized it in this process.
                    // Environment creation below still reports genuine load errors.
                    host.Log("WebView2 loader is already initialized; reusing the process loader");
                }
                loaderConfigured = true;
            }
        }

        internal Task<string> RunScript(string script)
        {
            if (browser == null || browser.CoreWebView2 == null || IsDisposed) throw new InvalidOperationException("The local toolbox browser is not ready.");
            return browser.CoreWebView2.ExecuteScriptAsync(script);
        }

        internal WebView2 DetachBrowser()
        {
            WebView2 value = browser;
            if (value != null) Controls.Remove(value);
            browser = null;
            return value;
        }

        internal void AttachBrowser(WebView2 value, Connect host)
        {
            if (value == null || value.IsDisposed) throw new InvalidOperationException("Cannot attach a disposed toolbox browser.");
            owner = host;
            browser = value;
            browser.Dock = DockStyle.Fill;
            Controls.Add(browser);
            initialized = true;
            IntPtr forceHandle = Handle;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && browser != null)
            {
                WebView2 current = browser;
                browser = null;
                if (owner == null || !owner.ParkBrowser(this, current)) current.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    [ComVisible(true), Guid("C6204290-B30E-4758-B107-74951EA827A5"), ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class Bridge
    {
        private readonly Connect owner;
        internal Bridge(Connect host) { owner = host; }
        public string RootFolder { get { return owner.Root; } }
        public string DocumentKey { get { return owner.DocumentKey; } }
        public bool FileDialogPending { get { return owner.FileDialogPending; } }
        public string BeginPictureSnapshot() { return owner.BeginPictureSnapshot(); }
        public string PollPictureSnapshot(string id) { return owner.PollPictureSnapshot(id); }
        public string BeginFileDialog(string requestJson) { return owner.BeginFileDialog(requestJson); }
        public string PollFileDialog(string requestId) { return owner.PollFileDialog(requestId); }
        public void TrackScratch(object presentation) { owner.TrackScratch(presentation); }
        public void ReleaseScratch(object presentation) { owner.ReleaseScratch(presentation); }
        public string TempFolder
        {
            get
            {
                string folder = Path.Combine(Path.GetTempPath(), "PowerPointTools", Process.GetCurrentProcess().Id.ToString(System.Globalization.CultureInfo.InvariantCulture));
                Directory.CreateDirectory(folder);
                return folder;
            }
        }
        public string UpdateRoot { get { string path = Path.Combine(RootFolder, "updates"); Directory.CreateDirectory(path); return path; } }
        // Some Office dispatch properties (notably TextRange.Text and ColorFormat.RGB)
        // are surfaced as callable proxies by WebView2. Invoke the COM PROPERTYGET /
        // PROPERTYPUT dispatch explicitly instead of relying on JavaScript inference.
        public object ReadMember(object target, string name)
        {
            ValidateComMember(target, name);
            return target.GetType().InvokeMember(name,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty | BindingFlags.OptionalParamBinding,
                null, target, new object[0], System.Globalization.CultureInfo.InvariantCulture);
        }
        public void WriteMember(object target, string name, object value)
        {
            ValidateComMember(target, name);
            target.GetType().InvokeMember(name,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty | BindingFlags.OptionalParamBinding,
                null, target, new object[] { value }, System.Globalization.CultureInfo.InvariantCulture);
        }
        public object CallMember(object target, string name, object[] args)
        {
            ValidateComMember(target, name);
            return target.GetType().InvokeMember(name,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod | BindingFlags.GetProperty | BindingFlags.OptionalParamBinding,
                null, target, args ?? new object[0], System.Globalization.CultureInfo.InvariantCulture);
        }
        private static void ValidateComMember(object target, string name)
        {
            if (target == null) throw new ArgumentNullException("target");
            if (!Marshal.IsComObject(target)) throw new ArgumentException("The target must be a native COM object.", "target");
            if (String.IsNullOrWhiteSpace(name)) throw new ArgumentException("A COM member name is required.", "name");
        }
        public bool FileExists(string path) { return File.Exists(path); }
        // WebView2's COM string projection can truncate embedded NUL bytes.
        // Binary data therefore crosses the native boundary only as base64.
        public string ReadBinaryBase64(string path) { return Convert.ToBase64String(File.ReadAllBytes(path)); }
        public string ReadPicturePreviewBase64(string path)
        {
            string root = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "DongsidayePictureSources")) + Path.DirectorySeparatorChar;
            string full = Path.GetFullPath(path);
            if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("图片预览必须来自当前快照。");
            using (Image source = Image.FromFile(full))
            {
                if ((long)source.Width * source.Height > 50000000) throw new InvalidOperationException("图片过大，改用宿主预览。");
                double scale = Math.Min(1.0, 384.0 / Math.Max(source.Width, source.Height));
                using (var bitmap = new Bitmap(Math.Max(1, (int)Math.Round(source.Width * scale)), Math.Max(1, (int)Math.Round(source.Height * scale))))
                using (Graphics graphics = Graphics.FromImage(bitmap))
                using (var stream = new MemoryStream())
                {
                    graphics.Clear(Color.Transparent);
                    graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    graphics.DrawImage(source, 0, 0, bitmap.Width, bitmap.Height);
                    bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                    return Convert.ToBase64String(stream.ToArray());
                }
            }
        }
        public void WriteBinaryBase64(string path, string data) { EnsureParent(path); File.WriteAllBytes(path, Convert.FromBase64String(data ?? "")); }
        public string ReadText(string path) { return File.ReadAllText(path, Encoding.UTF8); }
        public void WriteText(string path, string text) { EnsureParent(path); File.WriteAllText(path, text ?? "", new UTF8Encoding(false)); }
        public void RemoveFile(string path) { if (File.Exists(path)) File.Delete(path); }
        public string StatJson(string path)
        {
            FileInfo info = new FileInfo(path);
            if (!info.Exists) return "null";
            long milliseconds = (long)(info.LastWriteTimeUtc - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
            return new JavaScriptSerializer().Serialize(new { size = info.Length, mtimeMs = milliseconds });
        }
        public bool ShellExecute(string path, string args)
        {
            Uri uri;
            if (Uri.TryCreate(path, UriKind.Absolute, out uri) && (uri.Scheme == "http" || uri.Scheme == "https"))
            {
                if (!String.IsNullOrEmpty(args)) throw new InvalidOperationException("Web links do not accept process arguments.");
                Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
                return true;
            }
            if (!Path.IsPathRooted(path)) throw new InvalidOperationException("Only absolute application paths are allowed.");
            string full = Path.GetFullPath(path);
            if (!File.Exists(full)) throw new FileNotFoundException("Application was not found.", full);
            string name = Path.GetFileName(full);
            bool photoshop = String.Equals(name, "Photoshop.exe", StringComparison.OrdinalIgnoreCase);
            string updateRoot = Path.GetFullPath(UpdateRoot) + Path.DirectorySeparatorChar;
            bool updater = full.StartsWith(updateRoot, StringComparison.OrdinalIgnoreCase) && name.StartsWith("PictureReplaceTools-PPT-", StringComparison.OrdinalIgnoreCase) && name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase);
            if (!photoshop && !updater) throw new InvalidOperationException("This executable is not an approved toolbox application.");
            Process.Start(new ProcessStartInfo(full, args ?? "") { UseShellExecute = updater, WorkingDirectory = Path.GetDirectoryName(full) });
            return true;
        }
        public void ShowPane(string title) { owner.ShowPane(title); }
        public void HidePane() { owner.HidePane(); }
        public void Log(string text) { owner.Log(text ?? ""); }
        public void Report(string value) { owner.Report(value); }
        private static void EnsureParent(string path)
        {
            string folder = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!String.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);
        }
    }

    internal sealed class PictureConverter : AxHost
    {
        private PictureConverter() : base("") { }
        internal static object Convert(Image image) { return GetIPictureDispFromPicture(image); }
    }
}
