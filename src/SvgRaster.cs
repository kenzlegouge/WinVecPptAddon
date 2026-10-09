// Draws a dense layer (see SvgOptimizer.cs) into a transparent bitmap with GDI+ and encodes it as PNG.

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace PdfDrop
{
    internal static partial class SvgOptimizer
    {
        private sealed class Canvas : IDisposable
        {
            private Bitmap bmp;
            private Graphics g;

            public Canvas(Run r)
            {
                bmp = new Bitmap(r.PxW, r.PxH, PixelFormat.Format32bppPArgb);
                try
                {
                    g = Graphics.FromImage(bmp);
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.PixelOffsetMode = PixelOffsetMode.Half;
                    g.ScaleTransform((float)r.Ppu, (float)r.Ppu);
                    g.TranslateTransform((float)-r.X, (float)-r.Y);
                }
                catch { Dispose(); throw; }
            }

            public void Dispose()
            {
                if (g != null) { g.Dispose(); g = null; }
                if (bmp != null) { bmp.Dispose(); bmp = null; }
            }

            /// <summary>Paints one path; clip (optional) is in the layer's coordinate system.</summary>
            public void Draw(Style st, Geometry geo, Clip clip)
            {
                if (clip != null)
                    using (var cp = new GraphicsPath(clip.Pts, clip.Types, clip.EvenOdd ? FillMode.Alternate : FillMode.Winding))
                        g.SetClip(cp);
                GraphicsState saved = null;
                if (st.M != null)
                {
                    saved = g.Save();
                    using (var m = new Matrix((float)st.M[0], (float)st.M[1], (float)st.M[2], (float)st.M[3], (float)st.M[4], (float)st.M[5]))
                        g.MultiplyTransform(m);
                }
                var pts = new PointF[geo.N]; var types = new byte[geo.N];
                Array.Copy(geo.Pts, pts, geo.N); Array.Copy(geo.Types, types, geo.N);
                using (var gp = new GraphicsPath(pts, types, st.EvenOdd ? FillMode.Alternate : FillMode.Winding))
                {
                    if (st.HasFill)
                    {
                        if (st.FillBrush == null) st.FillBrush = new SolidBrush(st.Fill);
                        g.FillPath(st.FillBrush, gp);
                    }
                    if (st.HasStroke)
                    {
                        if (st.Pen == null)
                        {
                            st.Pen = new Pen(st.Stroke, st.Width) { StartCap = st.Cap, EndCap = st.Cap, LineJoin = st.Join, MiterLimit = st.Miter };
                            st.StrokeBrush = new SolidBrush(st.Stroke);
                        }
                        g.DrawPath(st.Pen, gp);
                        if (st.Cap != LineCap.Flat) DrawDots(st, geo);
                    }
                }
                if (saved != null) g.Restore(saved);
                if (clip != null) g.ResetClip();
            }

            /// <summary>Zero-length subpaths with round or square caps are dots in SVG and PDF; GDI+ draws nothing for them.</summary>
            private void DrawDots(Style st, Geometry geo)
            {
                int start = 0;
                for (int k = 1; k <= geo.N; k++)
                {
                    if (k < geo.N && (geo.Types[k] & 7) != 0) continue;
                    bool degenerate = k - start >= 2;
                    for (int j = start + 1; j < k && degenerate; j++)
                        if (geo.Pts[j].X != geo.Pts[start].X || geo.Pts[j].Y != geo.Pts[start].Y) degenerate = false;
                    if (degenerate)
                    {
                        float x = geo.Pts[start].X - st.Width / 2, y = geo.Pts[start].Y - st.Width / 2;
                        if (st.Cap == LineCap.Round) g.FillEllipse(st.StrokeBrush, x, y, st.Width, st.Width);
                        else g.FillRectangle(st.StrokeBrush, x, y, st.Width, st.Width);
                    }
                    start = k;
                }
            }

            public string ToPngBase64()
            {
                g.Flush();
                using (var ms = new MemoryStream())
                {
                    bmp.Save(ms, ImageFormat.Png);
                    return Convert.ToBase64String(ms.GetBuffer(), 0, (int)ms.Length);
                }
            }
        }
    }
}
