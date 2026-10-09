// Makes pdftocairo's SVG cheap for PowerPoint to import.
//
// PowerPoint's SVG importer costs roughly 0.25 ms per drawn element, and it runs on the UI
// thread: a scatter plot with 50 000 dots freezes PowerPoint for 12 s, 200 000 dots for a minute.
// Two rewrites, applied to runs of consecutive plain <path> elements:
//
//   1. Merge. Neighbouring paths with identical paint are joined into one <path> (up to 500
//      per element). Only done where the result is pixel-identical: opaque paint, and either
//      stroke-only paths or convex fills with the same winding direction. Still 100 % vector.
//
//   2. Rasterize dense layers. If the page is still too expensive (typically semi-transparent
//      dots, which cannot be merged without changing how overlaps darken), the largest runs
//      are drawn into a transparent high-resolution PNG that is embedded in the SVG at the same
//      place in the stacking order. Axes, text and everything else stay vector.
//
// Anything not fully understood (gradients, dashes, masks, unknown attributes) is left alone.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Text;

namespace PdfDrop
{
    internal class SvgOptimizerOptions
    {
        public bool RasterizeDenseLayers = true;
        public int Dpi = 600;
        /// <summary>Estimated PowerPoint import time a page may cost before dense layers are rasterized.</summary>
        public double BudgetMs = 3000;
    }

    internal static partial class SvgOptimizer
    {
        private const int MaxPathsPerElement = 500;
        private const int MinRasterRun = 500;
        private const double WorthwhileMs = 500;   // pages PowerPoint imports faster than this are left exactly as converted
        private const double MsPerElement = 0.2, MsPerPath = 0.045;
        private const long MaxPixels = 16000000, MaxTotalPixels = 24000000;
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        private sealed class Style
        {
            public bool Simple, HasFill, HasStroke, EvenOdd, Opaque, Mergeable;
            public Color Fill, Stroke;
            public float Width = 1, Miter = 4;
            public LineCap Cap = LineCap.Flat;
            public LineJoin Join = LineJoin.MiterClipped;
            public double[] M;
            public SolidBrush FillBrush, StrokeBrush;
            public Pen Pen;
        }

        private sealed class Run
        {
            public int Paths, Elements;
            public bool RasterOk, Rasterize;
            public double[] Ctx;
            public double MinX = double.MaxValue, MinY = double.MaxValue, MaxX = double.MinValue, MaxY = double.MinValue;
            public double X, Y, W, H, ExtW, ExtH, Ppu;
            public int PxW, PxH;
            public double Est { get { return MsPerElement * Elements + MsPerPath * Paths; } }
        }

        private sealed class Ctx { public bool Supported; public double[] M; }

        private sealed class Clip { public PointF[] Pts; public byte[] Types; public bool EvenOdd; }

        /// <summary>Path data parsed into GDI+ form: points plus PathPointType bytes.</summary>
        private sealed class Geometry
        {
            public PointF[] Pts = new PointF[64];
            public byte[] Types = new byte[64];
            public int N;
            public int Subpaths;

            private void Add(float x, float y, byte type)
            {
                if (N == Pts.Length) { Array.Resize(ref Pts, N * 2); Array.Resize(ref Types, N * 2); }
                Pts[N] = new PointF(x, y); Types[N] = type; N++;
            }

            public bool Parse(string s, int i, int end)
            {
                N = 0; Subpaths = 0;
                char cmd = '\0';
                bool afterMove = false;
                
                while (true)
                {
                    while (i < end && (s[i] == ' ' || s[i] == ',' || s[i] == '\t')) i++;
                    if (i >= end) break;
                    char c = s[i];
                    if (c == 'M' || c == 'L' || c == 'C') { cmd = c; afterMove = false; i++; continue; }
                    if (c == 'Z' || c == 'z')
                    {
                        if (N == 0) return false;
                        Types[N - 1] |= 0x80;
                        // After a close the current point is the subpath start; cairo always follows with M.
                        cmd = '\0'; i++; continue;
                    }
                    if (!(c == '-' || c == '+' || c == '.' || (c >= '0' && c <= '9'))) return false;
                    float a, b;
                    if (cmd == 'M')
                    {
                        if (!Num(s, ref i, end, out a) || !Num(s, ref i, end, out b)) return false;
                        if (afterMove) Add(a, b, 1);
                        else
                        {
                            // A move with nothing drawn after it (cairo's trailing "M x y") is dropped below.
                            if (N > 0 && Types[N - 1] == 0) N--; else Subpaths++;
                            Add(a, b, 0); afterMove = true;
                        }
                    }
                    else if (cmd == 'L')
                    {
                        if (N == 0) return false;
                        if (!Num(s, ref i, end, out a) || !Num(s, ref i, end, out b)) return false;
                        Add(a, b, 1);
                    }
                    else if (cmd == 'C')
                    {
                        if (N == 0) return false;
                        for (int k = 0; k < 3; k++)
                        {
                            if (!Num(s, ref i, end, out a) || !Num(s, ref i, end, out b)) return false;
                            Add(a, b, 3);
                        }
                    }
                    else return false;
                }
                if (N > 0 && Types[N - 1] == 0) { N--; Subpaths--; }
                return N > 0;
            }

            private static bool Num(string s, ref int i, int end, out float value)
            {
                value = 0;
                while (i < end && (s[i] == ' ' || s[i] == ',')) i++;
                int start = i;
                bool neg = false;
                if (i < end && (s[i] == '-' || s[i] == '+')) { neg = s[i] == '-'; i++; }
                double v = 0; int digits = 0;
                while (i < end && s[i] >= '0' && s[i] <= '9') { v = v * 10 + (s[i] - '0'); i++; digits++; }
                if (i < end && s[i] == '.')
                {
                    i++;
                    double scale = 0.1;
                    while (i < end && s[i] >= '0' && s[i] <= '9') { v += (s[i] - '0') * scale; scale *= 0.1; i++; digits++; }
                }
                if (digits == 0) return false;
                if (i < end && (s[i] == 'e' || s[i] == 'E'))
                {
                    i++;
                    if (i < end && (s[i] == '-' || s[i] == '+')) i++;
                    while (i < end && s[i] >= '0' && s[i] <= '9') i++;
                    double parsed;
                    if (!double.TryParse(s.Substring(start, i - start), NumberStyles.Float, Inv, out parsed)) return false;
                    value = (float)parsed; return true;
                }
                value = (float)(neg ? -v : v);
                return true;
            }

            /// <summary>+1 / -1 if this is one convex closed outline (so its winding number is 0 or that sign everywhere), else 0.</summary>
            public int ConvexSign()
            {
                if (Subpaths != 1 || N < 3) return 0;
                double area = 0;
                int sign = 0, dxChanges = 0, lastDx = 0, firstDx = 0;
                for (int k = 0; k < N; k++)
                {
                    PointF p0 = Pts[k], p1 = Pts[(k + 1) % N], p2 = Pts[(k + 2) % N];
                    area += (double)p0.X * p1.Y - (double)p1.X * p0.Y;
                    double ex = p1.X - p0.X, ey = p1.Y - p0.Y, fx = p2.X - p1.X, fy = p2.Y - p1.Y;
                    double cross = ex * fy - ey * fx;
                    double tol = 1e-7 * (Math.Abs(ex) + Math.Abs(ey)) * (Math.Abs(fx) + Math.Abs(fy));
                    if (cross > tol) { if (sign < 0) return 0; sign = 1; }
                    else if (cross < -tol) { if (sign > 0) return 0; sign = -1; }
                    int dx = ex > 1e-9 ? 1 : ex < -1e-9 ? -1 : 0;
                    if (dx != 0)
                    {
                        if (lastDx != 0 && dx != lastDx) dxChanges++;
                        if (firstDx == 0) firstDx = dx;
                        lastDx = dx;
                    }
                }
                if (firstDx != 0 && lastDx != 0 && firstDx != lastDx) dxChanges++;
                // A convex polygon turns around exactly once: its edges change x-direction twice.
                if (sign == 0 || dxChanges > 2 || Math.Abs(area) < 1e-12) return 0;
                return area > 0 ? 1 : -1;
            }
        }

        /// <summary>Rewrites the SVG in place when that helps. Returns a one-line summary for the log (null: page too simple to bother).</summary>
        public static string Optimize(string svgPath, SvgOptimizerOptions opt)
        {
            var sw = Stopwatch.StartNew();
            var runs = new List<Run>();
            double[] page;
            using (var pass = new Pass(svgPath, null, runs)) page = pass.Execute();
            long analysisMs = sw.ElapsedMilliseconds, drawMs = 0, pngMs = 0;

            int paths = 0, elements = 0;
            double est = 0;
            foreach (var r in runs) { paths += r.Paths; elements += r.Elements; est += r.Est; }

            if (est < WorthwhileMs) return null;

            int layers = 0; long pixels = 0;
            if (opt.RasterizeDenseLayers && est > opt.BudgetMs && page != null)
            {
                var candidates = runs.FindAll(r => r.RasterOk && r.Paths >= MinRasterRun);
                candidates.Sort((a, b) => b.Est.CompareTo(a.Est));
                foreach (var r in candidates)
                {
                    if (est <= opt.BudgetMs) break;
                    if (!PlanRaster(r, page, opt.Dpi)) continue;
                    r.Rasterize = true; est -= r.Est; layers++; pixels += (long)r.PxW * r.PxH;
                }
                if (pixels > MaxTotalPixels)
                {
                    double shrink = Math.Sqrt((double)MaxTotalPixels / pixels);
                    pixels = 0;
                    foreach (var r in runs)
                        if (r.Rasterize) { SetResolution(r, r.Ppu * shrink); pixels += (long)r.PxW * r.PxH; }
                }
            }
            if (layers == 0 && elements >= paths) return "SVG: " + paths + " plain paths, nothing to optimize";

            string tmp = svgPath + ".opt";
            try
            {
                using (var writer = new StreamWriter(tmp, false, new UTF8Encoding(false), 1 << 16))
                using (var pass = new Pass(svgPath, writer, runs))
                {
                    pass.Execute();
                    drawMs = pass.DrawTicks * 1000 / Stopwatch.Frequency; pngMs = pass.PngTicks * 1000 / Stopwatch.Frequency;
                }
                File.Delete(svgPath);
                File.Move(tmp, svgPath);
            }
            finally { try { if (File.Exists(tmp)) File.Delete(tmp); } catch { } }

            int rasterPaths = 0, vectorElements = 0;
            foreach (var r in runs) { if (r.Rasterize) rasterPaths += r.Paths; else vectorElements += r.Elements; }
            return string.Format(Inv, "SVG optimized in {0} ms: {1} paths -> {2} vector elements{3}",
                sw.ElapsedMilliseconds, paths, vectorElements,
                layers == 0 ? "" : string.Format(Inv, " + {0} dense layer(s) as image ({1} paths, {2:0.0} MP; analysis {3}, draw {4}, png {5} ms)", layers, rasterPaths, pixels / 1e6, analysisMs, drawMs, pngMs));
        }

        /// <summary>Chooses the image rectangle and resolution for a run; false if it cannot be rasterized sensibly.</summary>
        private static bool PlanRaster(Run r, double[] page, int dpi)
        {
            double[] inv = Invert(r.Ctx);
            if (inv == null) return false;
            // Page rectangle in the run's coordinate system: nothing outside it is visible.
            double px0 = double.MaxValue, py0 = double.MaxValue, px1 = double.MinValue, py1 = double.MinValue;
            for (int k = 0; k < 4; k++)
            {
                double x = (k & 1) == 0 ? page[0] : page[0] + page[2], y = (k & 2) == 0 ? page[1] : page[1] + page[3];
                double lx = inv[0] * x + inv[2] * y + inv[4], ly = inv[1] * x + inv[3] * y + inv[5];
                px0 = Math.Min(px0, lx); px1 = Math.Max(px1, lx); py0 = Math.Min(py0, ly); py1 = Math.Max(py1, ly);
            }
            double x0 = Math.Max(r.MinX, px0), y0 = Math.Max(r.MinY, py0), x1 = Math.Min(r.MaxX, px1), y1 = Math.Min(r.MaxY, py1);
            if (!(x1 > x0) || !(y1 > y0)) return false;

            double scale = Math.Sqrt(Math.Abs(r.Ctx[0] * r.Ctx[3] - r.Ctx[1] * r.Ctx[2]));
            if (!(scale > 1e-9)) return false;
            double ppu = scale * dpi / 72.0;
            double area = (x1 - x0) * (y1 - y0) * ppu * ppu;
            if (area > MaxPixels) ppu *= Math.Sqrt(MaxPixels / area);
            r.X = x0; r.Y = y0; r.ExtW = x1 - x0; r.ExtH = y1 - y0;
            SetResolution(r, ppu);
            return true;
        }

        private static void SetResolution(Run r, double ppu)
        {
            r.Ppu = ppu;
            r.PxW = Math.Max(1, (int)Math.Ceiling(r.ExtW * ppu));
            r.PxH = Math.Max(1, (int)Math.Ceiling(r.ExtH * ppu));
            r.W = r.PxW / ppu; r.H = r.PxH / ppu;
        }

        /// <summary>
        /// One streaming pass over the file. Without a writer it only analyses (fills the run list);
        /// with one it rewrites the file using the decisions stored in that list. Both walk the file
        /// identically, so the n-th run of one pass is the n-th run of the other.
        /// </summary>
        private sealed class Pass : IDisposable
        {
            private readonly TextWriter writer;
            private readonly List<Run> runs;
            private readonly bool analysis;
            private readonly StreamReader reader;
            private Queue<string> pending = new Queue<string>();

            private readonly Dictionary<string, Style> styles = new Dictionary<string, Style>();
            private string cacheLine; private int cachePrefix, cacheSuffix; private Style cacheStyle;
            private readonly Geometry geo = new Geometry();
            private readonly Stack<Ctx> stack = new Stack<Ctx>();
            private readonly Dictionary<string, Clip> clips = new Dictionary<string, Clip>();
            private double[] page;

            private int runIndex = -1;
            private Run cur;
            private Style lastStyle; private int lastSign, inChunk;
            private readonly StringBuilder chunk = new StringBuilder();
            private string chunkPrefix, chunkSuffix;
            private Canvas canvas;
            public long DrawTicks, PngTicks;

            public Pass(string svgPath, TextWriter writer, List<Run> runs)
            {
                this.writer = writer; this.runs = runs; analysis = writer == null;
                reader = new StreamReader(svgPath, Encoding.UTF8, true, 1 << 16);
                stack.Push(new Ctx { Supported = true, M = new double[] { 1, 0, 0, 1, 0, 0 } });
            }

            public void Dispose()
            {
                reader.Dispose();
                if (canvas != null) canvas.Dispose();
                foreach (var s in styles.Values)
                {
                    if (s.FillBrush != null) s.FillBrush.Dispose();
                    if (s.StrokeBrush != null) s.StrokeBrush.Dispose();
                    if (s.Pen != null) s.Pen.Dispose();
                }
            }

            private string Next() { return pending.Count > 0 ? pending.Dequeue() : reader.ReadLine(); }

            private void PushBack(List<string> lines)
            {
                var q = new Queue<string>(lines);
                foreach (var l in pending) q.Enqueue(l);
                pending = q;
            }

            private static string Trim(string raw)
            {
                return raw.Length > 0 && (raw[0] == ' ' || raw[0] == '\t') ? raw.TrimStart() : raw;
            }

            /// <summary>Style of a plain painting path (geometry left in geo), or null for anything else.</summary>
            private Style PathStyle(string line, out int dStart, out int dEnd)
            {
                dStart = dEnd = 0;
                if (!line.StartsWith("<path ", StringComparison.Ordinal) || !line.EndsWith("/>", StringComparison.Ordinal)) return null;
                int at = line.IndexOf(" d=\"", StringComparison.Ordinal);
                if (at <= 0) return null;
                dStart = at + 4;
                dEnd = line.IndexOf('"', dStart);
                if (dEnd <= dStart) return null;

                // Neighbouring paths nearly always share their attributes: compare with the previous line first.
                Style st;
                int suffix = line.Length - dEnd;
                if (cacheLine != null && cachePrefix == dStart && cacheSuffix == suffix
                    && string.CompareOrdinal(line, 0, cacheLine, 0, dStart) == 0
                    && string.CompareOrdinal(line, dEnd, cacheLine, cacheLine.Length - suffix, suffix) == 0)
                    st = cacheStyle;
                else
                {
                    string key = line.Substring(0, dStart) + line.Substring(dEnd);
                    if (!styles.TryGetValue(key, out st)) { st = ParseStyle(line, at, dEnd + 1); styles[key] = st; }
                    cacheLine = line; cachePrefix = dStart; cacheSuffix = suffix; cacheStyle = st;
                }
                return st.Simple && geo.Parse(line, dStart, dEnd) ? st : null;
            }

            public double[] Execute()
            {
                var look = new List<string>();
                string raw;
                while ((raw = Next()) != null)
                {
                    string line = Trim(raw);
                    int dStart, dEnd;
                    Style st = PathStyle(line, out dStart, out dEnd);
                    if (st != null) { BeginRun(); AddPath(raw, line, st, dStart, dEnd); continue; }

                    // cairo wraps each shape that crosses a clip edge in its own tiny clip group:
                    //   <g clip-path="url(#clip-7)"> <path .../> </g>
                    // Those must not break a run, or a clipped scatter plot falls apart into hundreds of short ones.
                    const string clipOpen = "<g clip-path=\"url(#";
                    if (line.StartsWith(clipOpen, StringComparison.Ordinal) && line.EndsWith(")\">", StringComparison.Ordinal))
                    {
                        string id = line.Substring(clipOpen.Length, line.Length - clipOpen.Length - 3);
                        Clip clip;
                        if (clips.TryGetValue(id, out clip) && clip != null)
                        {
                            look.Clear();
                            bool match = false;
                            while (look.Count < 9)
                            {
                                string l = Next();
                                if (l == null) break;
                                look.Add(l);
                                string t = Trim(l);
                                if (t == "</g>") { match = look.Count > 1; break; }
                                if (PathStyle(t, out dStart, out dEnd) == null) break;
                            }
                            if (match) { BeginRun(); AddClipped(raw, look, clip); continue; }
                            PushBack(look);
                        }
                    }

                    EndRun();
                    const string clipDef = "<clipPath id=\"";
                    if (line.StartsWith(clipDef, StringComparison.Ordinal) && line.EndsWith("\">", StringComparison.Ordinal)
                        && line.IndexOf('"', clipDef.Length) == line.Length - 2)
                    {
                        ReadClip(raw, line.Substring(clipDef.Length, line.Length - clipDef.Length - 2));
                        continue;
                    }
                    if (line.StartsWith("<g", StringComparison.Ordinal) && line.Length > 2 && (line[2] == ' ' || line[2] == '>'))
                    {
                        if (!line.EndsWith("/>", StringComparison.Ordinal) && line.IndexOf("</g>", StringComparison.Ordinal) < 0)
                            stack.Push(ParseGroup(line, stack.Peek()));
                    }
                    else
                    {
                        for (int k = line.IndexOf("</g>", StringComparison.Ordinal); k >= 0; k = line.IndexOf("</g>", k + 4, StringComparison.Ordinal))
                        {
                            if (stack.Count <= 1) throw new InvalidDataException("unbalanced groups");
                            stack.Pop();
                        }
                        if (page == null && line.StartsWith("<svg", StringComparison.Ordinal)) page = ParseViewBox(line);
                    }
                    if (!analysis) writer.WriteLine(raw);
                }
                EndRun();
                return page;
            }

            /// <summary>Remembers a clip path made only of plain outlines, so clipped shapes can be drawn into a layer image.</summary>
            private void ReadClip(string openLine, string id)
            {
                if (!analysis) writer.WriteLine(openLine);
                var pts = new List<PointF>(); var types = new List<byte>();
                bool ok = true, evenOdd = false; int children = 0;
                string raw;
                while ((raw = Next()) != null)
                {
                    if (!analysis) writer.WriteLine(raw);
                    string line = Trim(raw);
                    if (line.StartsWith("</clipPath>", StringComparison.Ordinal)) break;
                    int at = line.IndexOf(" d=\"", StringComparison.Ordinal);
                    int dEnd = at > 0 ? line.IndexOf('"', at + 4) : -1;
                    if (!ok || !line.StartsWith("<path ", StringComparison.Ordinal) || !line.EndsWith("/>", StringComparison.Ordinal)
                        || dEnd < 0 || !geo.Parse(line, at + 4, dEnd)) { ok = false; continue; }
                    double[] m = null; bool childEvenOdd = false;
                    for (int part = 0; part < 2; part++)
                        foreach (var a in Attributes(line, part == 0 ? 6 : dEnd + 1, part == 0 ? at : line.Length - 2))
                        {
                            if (a.Key == "clip-rule") { if (a.Value == "evenodd") childEvenOdd = true; else ok &= a.Value == "nonzero"; }
                            else if (a.Key == "transform") { m = ParseMatrix(a.Value.Trim()); ok &= m != null; }
                            else ok = false;
                        }
                    if (children++ > 0 && childEvenOdd != evenOdd) ok = false;
                    evenOdd = childEvenOdd;
                    for (int k = 0; k < geo.N && ok; k++)
                    {
                        double x = geo.Pts[k].X, y = geo.Pts[k].Y;
                        if (m != null) { double tx = m[0] * x + m[2] * y + m[4]; y = m[1] * x + m[3] * y + m[5]; x = tx; }
                        pts.Add(new PointF((float)x, (float)y)); types.Add(geo.Types[k]);
                    }
                }
                clips[id] = ok && pts.Count > 0 ? new Clip { Pts = pts.ToArray(), Types = types.ToArray(), EvenOdd = evenOdd } : null;
            }

            private void BeginRun()
            {
                if (cur != null) return;
                runIndex++;
                if (analysis) { cur = new Run { RasterOk = stack.Peek().Supported, Ctx = stack.Peek().M }; return; }
                cur = runs[runIndex];
                if (!cur.Rasterize) return;
                canvas = new Canvas(cur);
            }

            private void AddPath(string raw, string line, Style st, int dStart, int dEnd)
            {
                if (!analysis && cur.Rasterize) { long t0 = Stopwatch.GetTimestamp(); canvas.Draw(st, geo, null); DrawTicks += Stopwatch.GetTimestamp() - t0; return; }

                // Vector: start a new element unless this path can join the current one unchanged.
                int sign = 0;
                bool mergeable = st.Mergeable && (!st.HasFill || (sign = geo.ConvexSign()) != 0);
                bool joins = mergeable && st == lastStyle && sign == lastSign && inChunk < MaxPathsPerElement;
                if (analysis)
                {
                    cur.Paths++;
                    if (!joins) cur.Elements++;
                    Bound(cur, st, geo);
                }
                else if (joins) chunk.Append(' ').Append(line, dStart, dEnd - dStart);
                else
                {
                    FlushChunk();
                    if (mergeable)
                    {
                        chunkPrefix = line.Substring(0, dStart); chunkSuffix = line.Substring(dEnd);
                        chunk.Append(line, dStart, dEnd - dStart);
                    }
                    else writer.WriteLine(raw);
                }
                if (joins) inChunk++;
                else { inChunk = 1; lastStyle = mergeable ? st : null; lastSign = sign; }
            }

            /// <summary>look = the group's paths followed by its closing tag.</summary>
            private void AddClipped(string openLine, List<string> look, Clip clip)
            {
                int n = look.Count - 1, dStart, dEnd;
                if (analysis)
                {
                    for (int k = 0; k < n; k++)
                    {
                        Style st = PathStyle(Trim(look[k]), out dStart, out dEnd);
                        cur.Paths++; cur.Elements++;
                        Bound(cur, st, geo);
                    }
                    cur.Elements++;
                }
                else if (cur.Rasterize)
                {
                    for (int k = 0; k < n; k++)
                        canvas.Draw(PathStyle(Trim(look[k]), out dStart, out dEnd), geo, clip);
                }
                else
                {
                    FlushChunk();
                    writer.WriteLine(openLine);
                    foreach (var l in look) writer.WriteLine(l);
                }
                lastStyle = null; inChunk = 0;
            }

            private void FlushChunk()
            {
                if (chunkPrefix != null) { writer.Write(chunkPrefix); writer.Write(chunk.ToString()); writer.WriteLine(chunkSuffix); }
                chunkPrefix = null; chunk.Length = 0;
            }

            private void EndRun()
            {
                if (cur == null) return;
                if (analysis) runs.Add(cur);
                else if (cur.Rasterize)
                {
                    writer.Write(string.Format(Inv,
                        "<image x=\"{0:0.####}\" y=\"{1:0.####}\" width=\"{2:0.####}\" height=\"{3:0.####}\" preserveAspectRatio=\"none\" xlink:href=\"data:image/png;base64,",
                        cur.X, cur.Y, cur.W, cur.H));
                    long t0 = Stopwatch.GetTimestamp();
                    writer.Write(canvas.ToPngBase64());
                    PngTicks += Stopwatch.GetTimestamp() - t0;
                    writer.WriteLine("\"/>");
                    canvas.Dispose(); canvas = null;
                }
                else FlushChunk();
                cur = null; lastStyle = null; inChunk = 0;
            }
        }

        private static void Bound(Run r, Style st, Geometry geo)
        {
            double pad = 0;
            if (st.HasStroke)
            {
                pad = st.Width / 2.0 * (st.Join == LineJoin.MiterClipped ? Math.Max(1.0, st.Miter) : 1.0);
                if (st.Cap == LineCap.Square) pad = Math.Max(pad, st.Width * 0.75);
                if (st.M != null) pad *= Math.Sqrt(Math.Max(st.M[0] * st.M[0] + st.M[1] * st.M[1], st.M[2] * st.M[2] + st.M[3] * st.M[3]));
            }
            for (int k = 0; k < geo.N; k++)
            {
                double x = geo.Pts[k].X, y = geo.Pts[k].Y;
                if (st.M != null) { double tx = st.M[0] * x + st.M[2] * y + st.M[4]; y = st.M[1] * x + st.M[3] * y + st.M[5]; x = tx; }
                if (x - pad < r.MinX) r.MinX = x - pad;
                if (x + pad > r.MaxX) r.MaxX = x + pad;
                if (y - pad < r.MinY) r.MinY = y - pad;
                if (y + pad > r.MaxY) r.MaxY = y + pad;
            }
        }

        #region Attribute parsing

        private static IEnumerable<KeyValuePair<string, string>> Attributes(string s, int from, int to)
        {
            int i = from;
            while (i < to)
            {
                while (i < to && (s[i] == ' ' || s[i] == '\t')) i++;
                int nameStart = i;
                while (i < to && s[i] != '=' && s[i] != ' ' && s[i] != '/' && s[i] != '>') i++;
                if (i >= to || s[i] != '=' || i + 1 >= to || (s[i + 1] != '"' && s[i + 1] != '\'')) yield break;
                string name = s.Substring(nameStart, i - nameStart);
                char quote = s[i + 1];
                int valueStart = i + 2, valueEnd = s.IndexOf(quote, valueStart);
                if (valueEnd < 0 || valueEnd > to) yield break;
                yield return new KeyValuePair<string, string>(name, s.Substring(valueStart, valueEnd - valueStart));
                i = valueEnd + 1;
            }
        }

        /// <summary>Style of a path line whose d attribute spans [dAttr, dAfter). Simple == every attribute is understood.</summary>
        private static Style ParseStyle(string line, int dAttr, int dAfter)
        {
            var st = new Style();
            bool fillSeen = false, ok = true;
            double fillAlpha = 1, strokeAlpha = 1;
            Color fill = Color.Black, stroke = Color.Black;
            string joinName = "miter";
            for (int part = 0; part < 2 && ok; part++)
            {
                int from = part == 0 ? 6 : dAfter, to = part == 0 ? dAttr : line.Length - 2;
                foreach (var a in Attributes(line, from, to))
                {
                    string v = a.Value.Trim();
                    switch (a.Key)
                    {
                        case "fill": fillSeen = true; if (v == "none") st.HasFill = false; else { st.HasFill = true; ok &= ParseColor(v, out fill); } break;
                        case "stroke": if (v == "none") st.HasStroke = false; else { st.HasStroke = true; ok &= ParseColor(v, out stroke); } break;
                        case "fill-rule": if (v == "evenodd") st.EvenOdd = true; else ok &= v == "nonzero"; break;
                        case "fill-opacity": ok &= double.TryParse(v, NumberStyles.Float, Inv, out fillAlpha); break;
                        case "stroke-opacity": ok &= double.TryParse(v, NumberStyles.Float, Inv, out strokeAlpha); break;
                        case "stroke-width": ok &= float.TryParse(v, NumberStyles.Float, Inv, out st.Width) && st.Width >= 0; break;
                        case "stroke-miterlimit": ok &= float.TryParse(v, NumberStyles.Float, Inv, out st.Miter) && st.Miter >= 1; break;
                        case "stroke-linecap":
                            if (v == "butt") st.Cap = LineCap.Flat; else if (v == "round") st.Cap = LineCap.Round; else if (v == "square") st.Cap = LineCap.Square; else ok = false;
                            break;
                        case "stroke-linejoin": joinName = v; ok &= v == "miter" || v == "round" || v == "bevel"; break;
                        case "transform": st.M = ParseMatrix(v); ok &= st.M != null; break;
                        default: ok = false; break;
                    }
                    if (!ok) break;
                }
            }
            // MiterClipped is GDI+'s name for SVG's behaviour: bevel when the miter limit is exceeded.
            st.Join = joinName == "round" ? LineJoin.Round : joinName == "bevel" ? LineJoin.Bevel : LineJoin.MiterClipped;
            if (st.HasStroke && st.Width <= 0) st.HasStroke = false;
            fillAlpha = Math.Max(0, Math.Min(1, fillAlpha)); strokeAlpha = Math.Max(0, Math.Min(1, strokeAlpha));
            st.Fill = Color.FromArgb((int)Math.Round(fillAlpha * 255), fill);
            st.Stroke = Color.FromArgb((int)Math.Round(strokeAlpha * 255), stroke);
            st.Simple = ok && fillSeen && (st.HasFill || st.HasStroke);
            st.Opaque = (!st.HasFill || fillAlpha >= 0.999) && (!st.HasStroke || strokeAlpha >= 0.999);
            // Fill + stroke on one element paints in an order that merging would change.
            st.Mergeable = st.Simple && st.Opaque && !(st.HasFill && st.HasStroke) && !(st.HasFill && st.EvenOdd);
            return st;
        }

        private static Ctx ParseGroup(string line, Ctx parent)
        {
            var c = new Ctx { Supported = parent.Supported, M = parent.M };
            int end = line.LastIndexOf('>');
            foreach (var a in Attributes(line, 2, end < 0 ? line.Length : end))
            {
                if (a.Key == "clip-path" || a.Key == "id") continue;   // the embedded image is clipped like the paths were
                if (a.Key == "transform")
                {
                    double[] m = ParseMatrix(a.Value.Trim());
                    if (m == null) c.Supported = false; else c.M = Multiply(parent.M, m);
                }
                else c.Supported = false;
            }
            return c;
        }

        private static bool ParseColor(string v, out Color color)
        {
            color = Color.Black;
            if (v.StartsWith("rgb(", StringComparison.Ordinal) && v.EndsWith(")", StringComparison.Ordinal))
            {
                var parts = v.Substring(4, v.Length - 5).Split(',');
                if (parts.Length != 3) return false;
                var c = new int[3];
                for (int k = 0; k < 3; k++)
                {
                    string p = parts[k].Trim();
                    bool percent = p.EndsWith("%", StringComparison.Ordinal);
                    double d;
                    if (!double.TryParse(percent ? p.Substring(0, p.Length - 1) : p, NumberStyles.Float, Inv, out d)) return false;
                    c[k] = (int)Math.Round(Math.Max(0, Math.Min(255, percent ? d * 2.55 : d)));
                }
                color = Color.FromArgb(c[0], c[1], c[2]);
                return true;
            }
            if (v.Length == 7 && v[0] == '#')
            {
                int rgb;
                if (!int.TryParse(v.Substring(1), NumberStyles.HexNumber, Inv, out rgb)) return false;
                color = Color.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255);
                return true;
            }
            return false;
        }

        private static double[] ParseMatrix(string v)
        {
            if (!v.StartsWith("matrix(", StringComparison.Ordinal) || !v.EndsWith(")", StringComparison.Ordinal)) return null;
            var parts = v.Substring(7, v.Length - 8).Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 6) return null;
            var m = new double[6];
            for (int k = 0; k < 6; k++)
                if (!double.TryParse(parts[k], NumberStyles.Float, Inv, out m[k])) return null;
            return m;
        }

        private static double[] ParseViewBox(string line)
        {
            foreach (var a in Attributes(line, 4, line.LastIndexOf('>')))
            {
                if (a.Key != "viewBox") continue;
                var parts = a.Value.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                var box = new double[4];
                if (parts.Length != 4) return null;
                for (int k = 0; k < 4; k++)
                    if (!double.TryParse(parts[k], NumberStyles.Float, Inv, out box[k])) return null;
                return box[2] > 0 && box[3] > 0 ? box : null;
            }
            return null;
        }

        // Matrices are SVG's [a b c d e f]: x' = a x + c y + e, y' = b x + d y + f.
        private static double[] Multiply(double[] p, double[] l)
        {
            return new[]
            {
                p[0] * l[0] + p[2] * l[1], p[1] * l[0] + p[3] * l[1],
                p[0] * l[2] + p[2] * l[3], p[1] * l[2] + p[3] * l[3],
                p[0] * l[4] + p[2] * l[5] + p[4], p[1] * l[4] + p[3] * l[5] + p[5]
            };
        }

        private static double[] Invert(double[] m)
        {
            double det = m[0] * m[3] - m[1] * m[2];
            if (Math.Abs(det) < 1e-12) return null;
            return new[]
            {
                m[3] / det, -m[1] / det, -m[2] / det, m[0] / det,
                (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det
            };
        }

        #endregion
    }
}
