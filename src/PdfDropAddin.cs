// PdfDrop add-in for PowerPoint (Windows).
//
// Makes dragging a PDF from Explorer onto a slide insert the PDF page as a
// vector graphic (the macOS behaviour) instead of an embedded OLE object icon.
//
// How: PowerPoint registers an OLE drop target on each document window
// ("mdiClass"). This add-in replaces that drop target with a wrapper. Drops
// that consist solely of .pdf files are handled here; everything else is
// forwarded untouched to PowerPoint's original drop target.
//
// Conversion: PowerPoint cannot read PDF, but it imports SVG as a true vector
// graphic. Each page is converted PDF -> SVG with the first tool found:
//   1. pdftocairo (Poppler) — the copy bundled in tools\poppler next to this
//      DLL, else one on PATH / from MiKTeX / poppler-windows
//   2. Inkscape (its Poppler-based PDF importer)
// If neither is available (or a page's SVG is rejected) the page is rendered
// to a 300 DPI PNG with Windows' built-in PDF renderer instead.
//
// Optional settings in HKCU\Software\PdfDrop:
//   Mode        REG_SZ  "vector" (default) or "raster"
//   PdfToCairo  REG_SZ  full path to pdftocairo.exe
//   Inkscape    REG_SZ  full path to inkscape.com / inkscape.exe

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;
using ComIDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

[assembly: System.Reflection.AssemblyTitle("PdfDrop add-in for PowerPoint")]
[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]
[assembly: ComVisible(false)]

namespace PdfDrop
{
    #region Office extensibility interface (avoids needing the Extensibility PIA)

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0, ext_cm_Startup = 1, ext_cm_External = 2,
        ext_cm_CommandLine = 3, ext_cm_Solution = 4, ext_cm_UISetup = 5
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0, ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2, ext_dm_SolutionClosed = 3
    }

    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"),
     InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application,
                          [In] ext_ConnectMode connectMode,
                          [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst,
                          [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)]
        void OnDisconnection([In] ext_DisconnectMode removeMode,
                             [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)]
        void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)]
        void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)]
        void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    #endregion

    #region OLE drag & drop interop

    [StructLayout(LayoutKind.Sequential)]
    public struct POINTL { public int x; public int y; }

    [ComImport, Guid("00000122-0000-0000-C000-000000000046"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDropTarget
    {
        [PreserveSig] int DragEnter(ComIDataObject pDataObj, uint grfKeyState, POINTL pt, ref uint pdwEffect);
        [PreserveSig] int DragOver(uint grfKeyState, POINTL pt, ref uint pdwEffect);
        [PreserveSig] int DragLeave();
        [PreserveSig] int Drop(ComIDataObject pDataObj, uint grfKeyState, POINTL pt, ref uint pdwEffect);
    }

    internal static class Native
    {
        public const string DropTargetProp = "OleDropTargetInterface";
        public const uint DROPEFFECT_NONE = 0, DROPEFFECT_COPY = 1;

        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("ole32.dll")] public static extern int RegisterDragDrop(IntPtr hwnd, IDropTarget pDropTarget);
        [DllImport("ole32.dll")] public static extern int RevokeDragDrop(IntPtr hwnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr GetProp(IntPtr hwnd, string name);
        [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
        [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hwnd, EnumWindowsProc cb, IntPtr lParam);
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder sb, int max);
        [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hwnd);
        [DllImport("kernel32.dll")] public static extern uint GetCurrentProcessId();

        public static string ClassName(IntPtr hwnd)
        {
            var sb = new StringBuilder(256);
            GetClassName(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }
    }

    /// <summary>Wraps PowerPoint's own drop target; intercepts PDF-only drops.</summary>
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class DropTargetWrapper : IDropTarget
    {
        private readonly Addin owner;
        private readonly IDropTarget original;
        private readonly IntPtr hwnd;
        private bool pdfMode;
        private string[] pdfFiles;

        public DropTargetWrapper(Addin owner, IDropTarget original, IntPtr hwnd)
        {
            this.owner = owner; this.original = original; this.hwnd = hwnd;
        }

        public IDropTarget Original { get { return original; } }

        public int DragEnter(ComIDataObject pDataObj, uint grfKeyState, POINTL pt, ref uint pdwEffect)
        {
            pdfMode = false; pdfFiles = null;
            try { pdfFiles = Addin.GetPdfFiles(pDataObj); } catch (Exception ex) { Addin.Log("DragEnter: " + ex.Message); }
            Addin.Log("DragEnter on 0x" + hwnd.ToInt64().ToString("X") + (pdfFiles != null ? ": PDF drop (" + pdfFiles.Length + " file(s))" : ": forwarded to PowerPoint"));
            if (pdfFiles != null)
            {
                pdfMode = true;
                pdwEffect = Native.DROPEFFECT_COPY;
                return 0;
            }
            return original.DragEnter(pDataObj, grfKeyState, pt, ref pdwEffect);
        }

        public int DragOver(uint grfKeyState, POINTL pt, ref uint pdwEffect)
        {
            if (pdfMode) { pdwEffect = Native.DROPEFFECT_COPY; return 0; }
            return original.DragOver(grfKeyState, pt, ref pdwEffect);
        }

        public int DragLeave()
        {
            if (pdfMode) { pdfMode = false; pdfFiles = null; return 0; }
            return original.DragLeave();
        }

        public int Drop(ComIDataObject pDataObj, uint grfKeyState, POINTL pt, ref uint pdwEffect)
        {
            if (pdfMode)
            {
                var files = pdfFiles;
                pdfMode = false; pdfFiles = null;
                pdwEffect = Native.DROPEFFECT_COPY;
                // Return to the drag source immediately; do the work right after.
                owner.ScheduleInsert(files, hwnd, pt.x, pt.y);
                return 0;
            }
            return original.Drop(pDataObj, grfKeyState, pt, ref pdwEffect);
        }
    }

    #endregion

    /// <summary>Small automation surface, reachable via COMAddIns("PdfDrop.Addin").Object.</summary>
    [ComVisible(true), Guid("5E2B9D47-8C31-4F6A-A1D0-6B7E3C9F2A14"),
     InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IPdfDropApi
    {
        [DispId(1)] string Status { get; }
        [DispId(2)] int HookedWindowCount { get; }
        [DispId(3)] string LogPath { get; }
        [DispId(4)] int InsertPdf(string path, int screenX, int screenY);
        [DispId(5)] void Rescan();
        [DispId(6)] string ConverterInfo { get; }
        [DispId(7)] string LastMethod { get; }
        [DispId(8)] void InsertPdfAsync(string path, int screenX, int screenY);
    }

    [ComVisible(true), Guid("7D3F0C8E-5B1A-4E7B-9C2D-3A8F1E6B4D21"),
     ProgId("PdfDrop.Addin"), ClassInterface(ClassInterfaceType.None),
     ComDefaultInterface(typeof(IPdfDropApi))]
    public class Addin : IDTExtensibility2, IPdfDropApi
    {
        private const int ppLayoutBlank = 12;

        private dynamic app;
        private Timer scanTimer;
        // Hidden control owned by PowerPoint's UI thread; BeginInvoke on it runs code there.
        private Control ui;
        private readonly Dictionary<IntPtr, HookedWindow> hooked = new Dictionary<IntPtr, HookedWindow>();
        private string status = "Not connected";
        private string lastMethod = "";

        private class HookedWindow { public IntPtr Hwnd; public IDropTarget Original; public DropTargetWrapper Wrapper; public IntPtr OurPtr; }

        /// <summary>A drop in progress: target captured on the UI thread, conversion on a worker thread.</summary>
        private class DropJob
        {
            public string[] Files;
            public dynamic Window, Slide;
            public double SlideW, SlideH, DropX, DropY;
            public string Method;
        }

        #region IDTExtensibility2

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                app = application;
                try { ((dynamic)addInInst).Object = this; } catch (Exception ex) { Log("Set AddIn.Object: " + ex.Message); }

                ui = new Control();
                var forceHandle = ui.Handle;

                scanTimer = new Timer { Interval = 750 };
                scanTimer.Tick += (s, e) => SafeScan();
                scanTimer.Start();

                SafeScan();
                status = "Connected";
                Log("Connected (mode " + connectMode + "); converter: " + PdfConverter.Describe());
                Task.Run(() => { System.Threading.Thread.Sleep(4000); PdfConverter.Warm(); });
            }
            catch (Exception ex) { status = "OnConnection failed: " + ex.Message; Log(status); }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                if (scanTimer != null) { scanTimer.Stop(); scanTimer.Dispose(); scanTimer = null; }
                if (ui != null) { ui.Dispose(); ui = null; }
                if (removeMode != ext_DisconnectMode.ext_dm_HostShutdown) UnhookAll();
                hooked.Clear();
                app = null;
                Log("Disconnected (" + removeMode + ")");
            }
            catch (Exception ex) { Log("OnDisconnection: " + ex.Message); }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { SafeScan(); }
        public void OnBeginShutdown(ref Array custom) { }

        #endregion

        #region IPdfDropApi

        public string Status { get { return status; } }
        public int HookedWindowCount { get { return hooked.Count; } }
        public string LogPath { get { return LogFile; } }
        public string ConverterInfo { get { return PdfConverter.Describe(); } }
        public string LastMethod { get { return lastMethod; } }
        public void Rescan() { SafeScan(); }

        /// <summary>Synchronous insert (for testing/automation); same code path as a drop.</summary>
        public int InsertPdf(string path, int screenX, int screenY)
        {
            var job = CaptureTarget(new[] { path }, IntPtr.Zero, screenX, screenY);
            string method;
            var pages = PdfConverter.Convert(job.Files, out method);
            job.Method = method;
            return InsertPages(job, pages);
        }

        /// <summary>Exactly what a drop does: returns at once, inserts when the conversion finishes.</summary>
        public void InsertPdfAsync(string path, int screenX, int screenY)
        {
            // Automation calls arrive on RPC threads; a drop arrives on the UI thread. Behave like the drop.
            if (ui == null) throw new InvalidOperationException("Add-in is not connected to PowerPoint.");
            ui.BeginInvoke((Action)(() => ScheduleInsert(new[] { path }, IntPtr.Zero, screenX, screenY)));
        }

        #endregion

        #region Window hooking

        private void SafeScan()
        {
            try { Scan(); } catch (Exception ex) { Log("Scan: " + ex.Message); }
        }

        private void Scan()
        {
            var dead = new List<IntPtr>();
            foreach (var kv in hooked) if (!Native.IsWindow(kv.Key)) dead.Add(kv.Key);
            foreach (var h in dead) hooked.Remove(h);

            uint myPid = Native.GetCurrentProcessId();
            var candidates = new List<IntPtr>();
            Native.EnumWindowsProc childCb = (c, l2) => { Consider(c, candidates); return true; };
            Native.EnumWindowsProc topCb = (h, l) =>
            {
                uint pid; Native.GetWindowThreadProcessId(h, out pid);
                if (pid == myPid) { Consider(h, candidates); Native.EnumChildWindows(h, childCb, IntPtr.Zero); }
                return true;
            };
            Native.EnumWindows(topCb, IntPtr.Zero);
            GC.KeepAlive(childCb); GC.KeepAlive(topCb);

            foreach (var h in candidates)
            {
                HookedWindow hw;
                if (hooked.TryGetValue(h, out hw))
                {
                    // PowerPoint re-registered its own target over ours: hook again.
                    if (Native.GetProp(h, Native.DropTargetProp) == hw.OurPtr) continue;
                    hooked.Remove(h);
                }
                Hook(h);
            }
        }

        private static void Consider(IntPtr hwnd, List<IntPtr> list)
        {
            if (Native.GetProp(hwnd, Native.DropTargetProp) == IntPtr.Zero) return;
            if (Native.ClassName(hwnd) != "mdiClass") return;
            list.Add(hwnd);
        }

        private void Hook(IntPtr hwnd)
        {
            IntPtr origPtr = Native.GetProp(hwnd, Native.DropTargetProp);
            if (origPtr == IntPtr.Zero) return;

            IDropTarget original;
            try { original = (IDropTarget)Marshal.GetObjectForIUnknown(origPtr); }
            catch (Exception ex) { Log("Hook: cannot read original drop target: " + ex.Message); return; }

            var wrapper = new DropTargetWrapper(this, original, hwnd);
            int hr = Native.RevokeDragDrop(hwnd);
            if (hr != 0) { Log("RevokeDragDrop failed 0x" + hr.ToString("X8")); return; }
            hr = Native.RegisterDragDrop(hwnd, wrapper);
            if (hr != 0)
            {
                Log("RegisterDragDrop failed 0x" + hr.ToString("X8") + "; restoring original");
                Native.RegisterDragDrop(hwnd, original);
                return;
            }
            hooked[hwnd] = new HookedWindow
            {
                Hwnd = hwnd, Original = original, Wrapper = wrapper,
                OurPtr = Native.GetProp(hwnd, Native.DropTargetProp)
            };
            Log("Hooked window 0x" + hwnd.ToInt64().ToString("X"));
        }

        private void UnhookAll()
        {
            foreach (var hw in hooked.Values)
            {
                if (!Native.IsWindow(hw.Hwnd)) continue;
                try
                {
                    Native.RevokeDragDrop(hw.Hwnd);
                    Native.RegisterDragDrop(hw.Hwnd, hw.Original);
                }
                catch (Exception ex) { Log("Unhook: " + ex.Message); }
            }
        }

        #endregion

        #region Drop handling

        internal static string[] GetPdfFiles(ComIDataObject data)
        {
            var dobj = new DataObject(data);
            if (!dobj.GetDataPresent(DataFormats.FileDrop)) return null;
            var files = dobj.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0) return null;
            foreach (var f in files)
                if (!string.Equals(Path.GetExtension(f), ".pdf", StringComparison.OrdinalIgnoreCase)) return null;
            return files;
        }

        internal void ScheduleInsert(string[] files, IntPtr hwnd, int x, int y)
        {
            try
            {
                // Capture the target now (while the slide under the cursor is still current),
                // convert on a worker thread, then insert back on the UI thread.
                var job = CaptureTarget(files, hwnd, x, y);
                status = "Converting " + Path.GetFileName(files[0]) + (files.Length > 1 ? " (+" + (files.Length - 1) + ")" : "");
                var uiControl = ui;
                Task.Run(() =>
                {
                    string method;
                    var pages = PdfConverter.Convert(job.Files, out method);
                    job.Method = method;
                    return pages;
                }).ContinueWith(t =>
                {
                    try { uiControl.BeginInvoke((Action)(() => FinishJob(job, t))); }
                    catch (Exception ex) { Log("Cannot return to UI thread: " + ex.Message); }
                });
            }
            catch (Exception ex) { ReportFailure(ex); }
        }

        private void FinishJob(DropJob job, Task<List<PageAsset>> conversion)
        {
            try
            {
                if (conversion.IsFaulted) throw conversion.Exception.GetBaseException();
                InsertPages(job, conversion.Result);
            }
            catch (Exception ex) { ReportFailure(ex); }
        }

        private void ReportFailure(Exception ex)
        {
            status = "Insert failed: " + ex.Message; Log(status);
            try
            {
                MessageBox.Show("Could not insert the PDF:\n" + ex.Message, "PDF drop",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch { }
        }

        private dynamic FindWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return null;
            try
            {
                foreach (dynamic w in app.Windows)
                {
                    try { if ((long)(int)w.HWND == hwnd.ToInt64()) return w; } catch { }
                }
            }
            catch { }
            return null;
        }

        private DropJob CaptureTarget(string[] files, IntPtr hwnd, int screenX, int screenY)
        {
            if (app == null) throw new InvalidOperationException("Add-in is not connected to PowerPoint.");

            dynamic window = FindWindow(hwnd);
            if (window == null) window = app.ActiveWindow;
            try { window.Activate(); } catch { }
            dynamic pres = window.Presentation;

            dynamic slide = null;
            try { slide = window.View.Slide; } catch { }
            if (slide == null)
            {
                if ((int)pres.Slides.Count == 0) pres.Slides.Add(1, ppLayoutBlank);
                slide = pres.Slides.Item((int)pres.Slides.Count);
            }

            var job = new DropJob
            {
                Files = files, Window = window, Slide = slide,
                SlideW = (double)pres.PageSetup.SlideWidth,
                SlideH = (double)pres.PageSetup.SlideHeight
            };
            job.DropX = job.SlideW / 2; job.DropY = job.SlideH / 2;

            // Screen pixel -> slide points, using PowerPoint's own mapping.
            try
            {
                int px0 = (int)window.PointsToScreenPixelsX(0f), px1 = (int)window.PointsToScreenPixelsX(1000f);
                int py0 = (int)window.PointsToScreenPixelsY(0f), py1 = (int)window.PointsToScreenPixelsY(1000f);
                if (px1 != px0 && py1 != py0)
                {
                    job.DropX = (screenX - px0) * 1000.0 / (px1 - px0);
                    job.DropY = (screenY - py0) * 1000.0 / (py1 - py0);
                }
            }
            catch (Exception ex) { Log("Point mapping: " + ex.Message); }
            return job;
        }

        private int InsertPages(DropJob job, List<PageAsset> pages)
        {
            dynamic slide = job.Slide;
            dynamic pres = slide.Parent;
            double cx = job.DropX, cy = job.DropY;
            int inserted = 0;
            try
            {
                foreach (var page in pages)
                {
                    if (inserted > 0)
                    {
                        int idx = (int)slide.SlideIndex + 1;
                        slide = pres.Slides.Add(idx, ppLayoutBlank);
                        try { job.Window.View.GotoSlide(idx); } catch { }
                        cx = job.SlideW / 2; cy = job.SlideH / 2;
                    }
                    AddPicture(slide, page, job.SlideW, job.SlideH, cx, cy);
                    inserted++;
                }
                lastMethod = job.Method;
                status = "Inserted " + inserted + " page(s) (" + job.Method + ")";
                Log(status + " from " + string.Join(", ", job.Files));
            }
            finally
            {
                PdfConverter.Cleanup(pages);
            }
            return inserted;
        }

        private static void AddPicture(dynamic slide, PageAsset page, double slideW, double slideH, double cx, double cy)
        {
            // Natural page size in points, capped to 95 % of the slide.
            double scale = Math.Min(1.0, Math.Min(slideW * 0.95 / page.WidthPt, slideH * 0.95 / page.HeightPt));
            double w = page.WidthPt * scale, h = page.HeightPt * scale;

            double left = cx - w / 2, top = cy - h / 2;
            left = Math.Max(0, Math.Min(left, slideW - w));
            top = Math.Max(0, Math.Min(top, slideH - h));

            dynamic shape = null;
            // AddPicture(FileName, LinkToFile=msoFalse, SaveWithDocument=msoTrue, Left, Top, Width, Height)
            try
            {
                shape = slide.Shapes.AddPicture(page.Path, 0, -1, (float)left, (float)top, (float)w, (float)h);
            }
            catch (Exception ex)
            {
                if (!page.IsVector) throw;
                // PowerPoint rejected this SVG: fall back to a bitmap of the same page.
                Log("SVG rejected by PowerPoint (" + ex.Message + "); using bitmap for page " + page.PageIndex);
                var png = PdfConverter.RenderPageToPng(page.SourcePdf, page.PageIndex, Path.GetDirectoryName(page.Path));
                shape = slide.Shapes.AddPicture(png, 0, -1, (float)left, (float)top, (float)w, (float)h);
            }
            try { shape.LockAspectRatio = -1; } catch { }
            try { shape.Select(); } catch { }
        }

        #endregion

        #region Logging

        private static readonly string LogFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PdfDrop", "log.txt");

        internal static void Log(string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogFile));
                File.AppendAllText(LogFile, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + message + Environment.NewLine);
            }
            catch { }
        }

        #endregion
    }

    /// <summary>One converted page, ready for Shapes.AddPicture.</summary>
    internal class PageAsset
    {
        public string Path;
        public string SourcePdf;
        public uint PageIndex;
        public double WidthPt, HeightPt;
        public bool IsVector;
    }

    /// <summary>PDF -> SVG (vector) via pdftocairo or Inkscape, else PDF -> PNG via Windows.Data.Pdf.</summary>
    internal static class PdfConverter
    {
        public const int RasterDpi = 300;
        private const int ToolTimeoutMs = 120000;

        private static string Setting(string name)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\PdfDrop"))
                {
                    object value = key == null ? null : key.GetValue(name);
                    return value == null ? null : System.Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
                }
            }
            catch { return null; }
        }

        private static string FindOnPath(string exe)
        {
            var path = Environment.GetEnvironmentVariable("PATH") ?? "";
            foreach (var dir in path.Split(';'))
            {
                if (dir.Trim().Length == 0) continue;
                try
                {
                    var candidate = Path.Combine(dir.Trim(), exe);
                    if (File.Exists(candidate)) return candidate;
                }
                catch { }
            }
            return null;
        }

        private static string FirstExisting(params string[] candidates)
        {
            foreach (var c in candidates)
                if (!string.IsNullOrEmpty(c) && File.Exists(c)) return c;
            return null;
        }

        /// <summary>The bundled converter lives in tools\ next to bin\ (the DLL's folder); also accept tools\ inside it.</summary>
        private static string[] ToolsDirs()
        {
            try
            {
                var dllDir = Path.GetDirectoryName(typeof(PdfConverter).Assembly.Location);
                return new[] { Path.GetFullPath(Path.Combine(dllDir, @"..\tools")), Path.Combine(dllDir, "tools") };
            }
            catch { return new string[0]; }
        }

        public static string FindPdfToCairo()
        {
            string pf = Environment.GetEnvironmentVariable("ProgramFiles") ?? @"C:\Program Files";
            string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var candidates = new List<string> { Setting("PdfToCairo") };
            foreach (var tools in ToolsDirs())
            {
                candidates.Add(Path.Combine(tools, @"poppler\bin\pdftocairo.exe"));
                candidates.Add(Path.Combine(tools, "pdftocairo.exe"));
                candidates.Add(Path.Combine(tools, @"Library\bin\pdftocairo.exe"));
            }
            candidates.Add(FindOnPath("pdftocairo.exe"));
            candidates.Add(Path.Combine(local, @"Programs\MiKTeX\miktex\bin\x64\pdftocairo.exe"));
            candidates.Add(Path.Combine(pf, @"MiKTeX\miktex\bin\x64\pdftocairo.exe"));
            var found = FirstExisting(candidates.ToArray());
            if (found != null) return found;
            foreach (var root in new[] { pf, @"C:\" })
            {
                try
                {
                    foreach (var dir in Directory.GetDirectories(root, "poppler*"))
                    {
                        var c = FirstExisting(Path.Combine(dir, @"Library\bin\pdftocairo.exe"), Path.Combine(dir, @"bin\pdftocairo.exe"));
                        if (c != null) return c;
                    }
                }
                catch { }
            }
            return null;
        }

        public static string FindInkscape()
        {
            string pf = Environment.GetEnvironmentVariable("ProgramFiles") ?? @"C:\Program Files";
            string pf86 = Environment.GetEnvironmentVariable("ProgramFiles(x86)") ?? @"C:\Program Files (x86)";
            return FirstExisting(
                Setting("Inkscape"),
                Path.Combine(pf, @"Inkscape\bin\inkscape.com"),
                Path.Combine(pf86, @"Inkscape\bin\inkscape.com"),
                FindOnPath("inkscape.com"),
                FindOnPath("inkscape.exe"));
        }

        public static bool VectorEnabled()
        {
            var mode = Setting("Mode");
            return mode == null || !mode.Trim().Equals("raster", StringComparison.OrdinalIgnoreCase);
        }

        public static string Describe()
        {
            if (!VectorEnabled()) return "raster (Mode=raster)";
            var c = FindPdfToCairo(); if (c != null) return "vector via pdftocairo: " + c;
            var i = FindInkscape(); if (i != null) return "vector via Inkscape: " + i;
            return "raster (no pdftocairo/Inkscape found)";
        }

        public static List<PageAsset> Convert(string[] pdfs, out string method)
        {
            var result = new List<PageAsset>();
            string pdftocairo = null, inkscape = null;
            if (VectorEnabled())
            {
                pdftocairo = FindPdfToCairo();
                if (pdftocairo == null) inkscape = FindInkscape();
            }
            method = pdftocairo != null ? "vector/pdftocairo" : inkscape != null ? "vector/Inkscape" : "bitmap";

            foreach (var pdf in pdfs)
            {
                var sw = Stopwatch.StartNew();
                var outDir = Path.Combine(Path.GetTempPath(), "PdfDrop", Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(outDir);
                string pdfPath = pdf, cairo = pdftocairo, ink = inkscape;

                // Nearly every drop is a one-page figure: start on page 1 while Windows' PDF reader
                // (page count and sizes, also validates the file) is still loading the document.
                Task<string> first = cairo != null ? Task.Run(() => VectorPage(cairo, null, pdfPath, 0, outDir)) : null;
                List<double[]> sizes;
                try { sizes = GetPageSizesPt(pdf); }
                catch
                {
                    if (first != null) try { first.Wait(); } catch { }
                    try { Directory.Delete(outDir, true); } catch { }
                    throw;
                }
                long sizesMs = sw.ElapsedMilliseconds;

                var svgs = new string[sizes.Count];
                if (first != null)
                {
                    svgs[0] = first.Result;
                    if (sizes.Count > 1)
                        Parallel.For(1, sizes.Count, new ParallelOptions { MaxDegreeOfParallelism = Math.Max(1, Math.Min(4, Environment.ProcessorCount)) },
                            i => svgs[i] = VectorPage(cairo, null, pdfPath, (uint)i, outDir));
                }
                else if (ink != null)
                    for (int i = 0; i < sizes.Count; i++) svgs[i] = VectorPage(null, ink, pdfPath, (uint)i, outDir);

                for (uint i = 0; i < sizes.Count; i++)
                {
                    var page = new PageAsset
                    {
                        SourcePdf = pdf, PageIndex = i,
                        WidthPt = sizes[(int)i][0], HeightPt = sizes[(int)i][1]
                    };
                    if (svgs[i] != null) { page.Path = svgs[i]; page.IsVector = true; }
                    else
                    {
                        if (pdftocairo != null || inkscape != null) Addin.Log("Vector conversion failed for page " + (i + 1) + " of " + Path.GetFileName(pdf) + "; using bitmap");
                        page.Path = RenderPageToPng(pdf, i, outDir); page.IsVector = false;
                    }
                    result.Add(page);
                }
                Addin.Log("Converted " + Path.GetFileName(pdf) + ": " + sizes.Count + " page(s) in " + sw.ElapsedMilliseconds + " ms (page sizes " + sizesMs + " ms)");
            }
            return result;
        }

        /// <summary>One page as SVG, or null. Never throws.</summary>
        private static string VectorPage(string pdftocairo, string inkscape, string pdf, uint pageIndex, string outDir)
        {
            try
            {
                string svg = null;
                if (pdftocairo != null)
                {
                    svg = TryPdfToCairo(pdftocairo, pdf, pageIndex, outDir);
                    if (svg != null) OptimizeSvg(svg, pageIndex);
                }
                if (svg == null && inkscape != null) svg = TryInkscape(inkscape, pdf, pageIndex, outDir);
                return svg;
            }
            catch (Exception ex) { Addin.Log("Vector conversion of page " + (pageIndex + 1) + ": " + ex.Message); return null; }
        }

        private static readonly object optimizeLock = new object();

        /// <summary>
        /// PowerPoint imports SVG on its UI thread at about 0.25 ms per element, so a dense scatter plot
        /// would freeze it for a long time. See SvgOptimizer. The file is left untouched if anything goes wrong.
        /// </summary>
        private static void OptimizeSvg(string svg, uint pageIndex)
        {
            try
            {
                string mode = (Setting("DenseLayers") ?? "auto").Trim().ToLowerInvariant();
                if (mode == "off") return;
                var options = new SvgOptimizerOptions { RasterizeDenseLayers = mode != "vector" };
                int dpi;
                if (int.TryParse(Setting("DenseLayerDpi"), out dpi) && dpi >= 72 && dpi <= 2400) options.Dpi = dpi;
                string summary;
                lock (optimizeLock) summary = SvgOptimizer.Optimize(svg, options);   // one at a time: layers are large bitmaps
                if (summary != null) Addin.Log("Page " + (pageIndex + 1) + ": " + summary);
            }
            catch (Exception ex) { Addin.Log("SVG optimization skipped for page " + (pageIndex + 1) + ": " + ex.Message); }
        }

        /// <summary>Loads the converter's DLLs into the file cache so the first drop does not pay for a cold start.</summary>
        public static void Warm()
        {
            try
            {
                if (!VectorEnabled()) return;
                string exe = FindPdfToCairo();
                if (exe != null) RunTool(exe, "-v", true);
            }
            catch { }
        }

        private static string TryPdfToCairo(string exe, string pdf, uint pageIndex, string outDir)
        {
            var svg = Path.Combine(outDir, string.Format("page{0:D3}.svg", pageIndex + 1));
            var args = string.Format("-svg -f {0} -l {0} \"{1}\" \"{2}\"", pageIndex + 1, pdf, svg);
            return RunTool(exe, args) && IsUsableSvg(svg) ? svg : null;
        }

        private static string TryInkscape(string exe, string pdf, uint pageIndex, string outDir)
        {
            var svg = Path.Combine(outDir, string.Format("page{0:D3}.svg", pageIndex + 1));
            var args = string.Format("--pdf-poppler --pages={0} --export-type=svg --export-plain-svg --export-filename=\"{1}\" \"{2}\"", pageIndex + 1, svg, pdf);
            return RunTool(exe, args) && IsUsableSvg(svg) ? svg : null;
        }

        private static bool IsUsableSvg(string path)
        {
            try { return File.Exists(path) && new FileInfo(path).Length > 100; } catch { return false; }
        }

        private static bool RunTool(string exe, string args, bool quiet = false)
        {
            try
            {
                var psi = new ProcessStartInfo(exe, args)
                {
                    UseShellExecute = false, CreateNoWindow = true,
                    RedirectStandardOutput = true, RedirectStandardError = true,
                    WorkingDirectory = Path.GetTempPath()
                };
                // Bundled poppler layout: <dir>\bin\pdftocairo.exe with its encoding data in <dir>\share\poppler.
                try
                {
                    var share = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(exe), @"..\share\poppler"));
                    if (Directory.Exists(share)) psi.EnvironmentVariables["POPPLER_DATADIR"] = share;
                }
                catch { }
                using (var p = Process.Start(psi))
                {
                    // Drain both streams asynchronously so a chatty tool cannot block on a full pipe.
                    p.OutputDataReceived += (s, e) => { };
                    var err = new StringBuilder();
                    p.ErrorDataReceived += (s, e) => { if (e.Data != null && err.Length < 4000) err.AppendLine(e.Data); };
                    p.BeginOutputReadLine(); p.BeginErrorReadLine();
                    if (!p.WaitForExit(ToolTimeoutMs))
                    {
                        try { p.Kill(); } catch { }
                        Addin.Log(Path.GetFileName(exe) + " timed out");
                        return false;
                    }
                    p.WaitForExit();
                    if (p.ExitCode != 0 && !quiet) Addin.Log(Path.GetFileName(exe) + " exit " + p.ExitCode + ": " + err.ToString().Trim());
                    return p.ExitCode == 0;
                }
            }
            catch (Exception ex) { Addin.Log("Run " + Path.GetFileName(exe) + ": " + ex.Message); return false; }
        }

        public static void Cleanup(List<PageAsset> pages)
        {
            var dirs = new HashSet<string>();
            foreach (var p in pages) { try { dirs.Add(Path.GetDirectoryName(p.Path)); } catch { } }
            foreach (var d in dirs) { try { Directory.Delete(d, true); } catch { } }
        }

        #region Windows.Data.Pdf (built into Windows)

        private static PdfDocument Open(string pdfPath)
        {
            var file = StorageFile.GetFileFromPathAsync(Path.GetFullPath(pdfPath)).AsTask().Result;
            return PdfDocument.LoadFromFileAsync(file).AsTask().Result;
        }

        /// <summary>Page sizes in points. PdfPage.Size is in 1/96 inch units.</summary>
        private static List<double[]> GetPageSizesPt(string pdfPath)
        {
            var doc = Open(pdfPath);
            var sizes = new List<double[]>();
            for (uint i = 0; i < doc.PageCount; i++)
                using (var page = doc.GetPage(i))
                    sizes.Add(new[] { page.Size.Width * 72.0 / 96.0, page.Size.Height * 72.0 / 96.0 });
            return sizes;
        }

        public static string RenderPageToPng(string pdfPath, uint pageIndex, string outDir)
        {
            var doc = Open(pdfPath);
            var pngPath = Path.Combine(outDir, string.Format("page{0:D3}.png", pageIndex + 1));
            using (var page = doc.GetPage(pageIndex))
            {
                var opts = new PdfPageRenderOptions
                {
                    DestinationWidth = (uint)Math.Round(page.Size.Width * (RasterDpi / 96.0))
                };
                using (var ras = new InMemoryRandomAccessStream())
                {
                    page.RenderToStreamAsync(ras, opts).AsTask().Wait();
                    using (var src = ras.GetInputStreamAt(0).AsStreamForRead())
                    using (var dst = File.Create(pngPath))
                        src.CopyTo(dst);
                }
            }
            return pngPath;
        }

        #endregion
    }
}
