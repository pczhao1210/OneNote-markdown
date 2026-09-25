using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using Media = System.Windows.Media;

namespace OneNoteMarkdown.Rendering
{
    internal static class EmfImageRenderer
    {
        internal static bool TryCreate(
            int width,
            int height,
            Action<Graphics> draw,
            out byte[] emfBytes,
            out string error)
        {
            emfBytes = null;
            error = string.Empty;
            if (width <= 0 || height <= 0)
            {
                error = "The EMF dimensions are invalid.";
                return false;
            }
            if (draw == null)
            {
                error = "The EMF drawing callback is missing.";
                return false;
            }

            try
            {
                using (MemoryStream stream = new MemoryStream())
                using (Bitmap referenceBitmap = new Bitmap(1, 1))
                using (Graphics referenceGraphics = Graphics.FromImage(referenceBitmap))
                {
                    IntPtr hdc = referenceGraphics.GetHdc();
                    Metafile metafile;
                    try
                    {
                        metafile = new Metafile(
                            stream,
                            hdc,
                            new Rectangle(0, 0, width, height),
                            MetafileFrameUnit.Pixel,
                            EmfType.EmfPlusDual);
                    }
                    finally
                    {
                        referenceGraphics.ReleaseHdc(hdc);
                    }

                    using (metafile)
                    using (Graphics graphics = Graphics.FromImage(metafile))
                    {
                        graphics.PageUnit = GraphicsUnit.Pixel;
                        graphics.SmoothingMode = SmoothingMode.AntiAlias;
                        graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                        draw(graphics);
                    }

                    emfBytes = stream.ToArray();
                }

                if (!IsEmf(emfBytes))
                {
                    emfBytes = null;
                    error = "The vector renderer returned an invalid EMF.";
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                emfBytes = null;
                error = ex.Message;
                return false;
            }
        }

        internal static void DrawGeometry(
            Graphics graphics,
            Media.Geometry geometry,
            float offsetX,
            float offsetY,
            Color color)
        {
            if (graphics == null) throw new ArgumentNullException(nameof(graphics));
            if (geometry == null) throw new ArgumentNullException(nameof(geometry));

            Media.PathGeometry flattened = geometry.GetFlattenedPathGeometry(0.1, Media.ToleranceType.Absolute);
            using (GraphicsPath closedPath = new GraphicsPath(
                flattened.FillRule == Media.FillRule.EvenOdd ? FillMode.Alternate : FillMode.Winding))
            using (GraphicsPath openPath = new GraphicsPath())
            {
                for (int i = 0; i < flattened.Figures.Count; i++)
                {
                    Media.PathFigure figure = flattened.Figures[i];
                    GraphicsPath target = figure.IsClosed ? closedPath : openPath;
                    AppendFigure(target, figure);
                }

                GraphicsState state = graphics.Save();
                try
                {
                    graphics.TranslateTransform(offsetX, offsetY);
                    using (SolidBrush brush = new SolidBrush(color))
                    using (Pen pen = new Pen(color, 1f))
                    {
                        if (closedPath.PointCount > 0) graphics.FillPath(brush, closedPath);
                        if (openPath.PointCount > 0) graphics.DrawPath(pen, openPath);
                    }
                }
                finally
                {
                    graphics.Restore(state);
                }
            }
        }

        internal static bool IsEmf(byte[] bytes)
        {
            return bytes != null &&
                bytes.Length > 44 &&
                bytes[40] == 0x20 &&
                bytes[41] == 0x45 &&
                bytes[42] == 0x4D &&
                bytes[43] == 0x46;
        }

        private static void AppendFigure(GraphicsPath path, Media.PathFigure figure)
        {
            PointF current = ToPointF(figure.StartPoint);
            path.StartFigure();
            for (int i = 0; i < figure.Segments.Count; i++)
            {
                Media.PathSegment segment = figure.Segments[i];
                Media.LineSegment line = segment as Media.LineSegment;
                if (line != null)
                {
                    PointF next = ToPointF(line.Point);
                    path.AddLine(current, next);
                    current = next;
                    continue;
                }

                Media.PolyLineSegment polyLine = segment as Media.PolyLineSegment;
                if (polyLine != null)
                {
                    for (int p = 0; p < polyLine.Points.Count; p++)
                    {
                        PointF next = ToPointF(polyLine.Points[p]);
                        path.AddLine(current, next);
                        current = next;
                    }
                    continue;
                }

                Media.BezierSegment bezier = segment as Media.BezierSegment;
                if (bezier != null)
                {
                    PointF end = ToPointF(bezier.Point3);
                    path.AddBezier(
                        current,
                        ToPointF(bezier.Point1),
                        ToPointF(bezier.Point2),
                        end);
                    current = end;
                    continue;
                }

                Media.PolyBezierSegment polyBezier = segment as Media.PolyBezierSegment;
                if (polyBezier != null)
                {
                    for (int p = 0; p + 2 < polyBezier.Points.Count; p += 3)
                    {
                        PointF end = ToPointF(polyBezier.Points[p + 2]);
                        path.AddBezier(
                            current,
                            ToPointF(polyBezier.Points[p]),
                            ToPointF(polyBezier.Points[p + 1]),
                            end);
                        current = end;
                    }
                }
            }
            if (figure.IsClosed) path.CloseFigure();
        }

        private static PointF ToPointF(System.Windows.Point point)
        {
            return new PointF((float)point.X, (float)point.Y);
        }
    }
}
