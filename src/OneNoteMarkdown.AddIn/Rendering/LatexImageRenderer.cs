using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using OneNoteMarkdown.Logging;
using WpfMath;
using WpfMath.Parsers;
using WpfMath.Rendering;
using XamlMath;
using XamlMath.Exceptions;
using Media = System.Windows.Media;

namespace OneNoteMarkdown.Rendering
{
    internal sealed class LatexImageRenderer
    {
        private const double RenderScale = 20.0;
        private const int MaxCacheEntries = 128;
        private static readonly object ParserLock = new object();
        private static readonly object CacheLock = new object();
        private static readonly Dictionary<string, CacheEntry> Cache =
            new Dictionary<string, CacheEntry>(StringComparer.Ordinal);
        private static readonly Queue<string> CacheOrder = new Queue<string>();
        private static readonly Regex InlineFormulaRegex = new Regex(
            "(?<!\\\\)(?<!\\$)\\$([^$\\r\\n]+?)\\$(?!\\$)",
            RegexOptions.Compiled);

        public bool TryRenderToPng(
            string latex,
            string textFontFamily,
            out byte[] pngBytes,
            out int pixelWidth,
            out int pixelHeight)
        {
            string error;
            return TryRenderToPng(
                latex,
                textFontFamily,
                out pngBytes,
                out pixelWidth,
                out pixelHeight,
                out error);
        }

        public bool TryRenderToPng(
            string latex,
            string textFontFamily,
            out byte[] pngBytes,
            out int pixelWidth,
            out int pixelHeight,
            out string error)
        {
            pngBytes = null;
            pixelWidth = 0;
            pixelHeight = 0;
            error = string.Empty;

            string text = Normalize(latex);
            if (text.Length == 0)
            {
                error = "The formula is empty.";
                return false;
            }

            string font = string.IsNullOrWhiteSpace(textFontFamily) ? "Cambria Math" : textFontFamily.Trim();
            string cacheKey = ComputeKey("png\n" + font + "\n" + text);
            CacheEntry cached;
            lock (CacheLock)
            {
                if (Cache.TryGetValue(cacheKey, out cached))
                {
                    pngBytes = (byte[])cached.Bytes.Clone();
                    pixelWidth = cached.Width;
                    pixelHeight = cached.Height;
                    return true;
                }
            }

            try
            {
                XamlMath.TexFormula formula;
                lock (ParserLock)
                {
                    formula = WpfTeXFormulaParser.Instance.Parse(text, "text");
                }

                byte[] bytes = Extensions.RenderToPng(formula, RenderScale, 0.0, 0.0, font);
                if (bytes == null || bytes.Length == 0)
                {
                    error = "The formula renderer returned no image.";
                    return false;
                }

                int width;
                int height;
                if (!TryReadPngSize(bytes, out width, out height))
                {
                    error = "The formula renderer returned an invalid PNG.";
                    return false;
                }

                pngBytes = bytes;
                pixelWidth = width;
                pixelHeight = height;
                lock (CacheLock)
                {
                    Cache[cacheKey] = new CacheEntry(bytes, width, height);
                    CacheOrder.Enqueue(cacheKey);
                    while (CacheOrder.Count > MaxCacheEntries)
                    {
                        Cache.Remove(CacheOrder.Dequeue());
                    }
                }
                return true;
            }
            catch (TexException ex)
            {
                error = ex.Message;
                return false;
            }
            catch (Exception ex)
            {
                Logger.Error("LatexImageRenderer unexpected error", ex);
                error = ex.Message;
                return false;
            }
        }

        public bool TryRenderToEmf(
            string latex,
            string textFontFamily,
            out byte[] emfBytes,
            out int pixelWidth,
            out int pixelHeight,
            out string error)
        {
            emfBytes = null;
            pixelWidth = 0;
            pixelHeight = 0;
            error = string.Empty;

            string text = Normalize(latex);
            if (text.Length == 0)
            {
                error = "The formula is empty.";
                return false;
            }

            string font = string.IsNullOrWhiteSpace(textFontFamily) ? "Cambria Math" : textFontFamily.Trim();
            string cacheKey = ComputeKey("emf\n" + font + "\n" + text);
            CacheEntry cached;
            lock (CacheLock)
            {
                if (Cache.TryGetValue(cacheKey, out cached))
                {
                    emfBytes = (byte[])cached.Bytes.Clone();
                    pixelWidth = cached.Width;
                    pixelHeight = cached.Height;
                    return true;
                }
            }

            Media.Geometry geometry;
            System.Windows.Rect bounds;
            if (!TryBuildGeometry(text, font, TexStyle.Display, out geometry, out bounds, out error))
            {
                return false;
            }

            int width = Math.Max(1, (int)Math.Ceiling(bounds.Width) + 4);
            int height = Math.Max(1, (int)Math.Ceiling(bounds.Height) + 4);
            if (width > 8192 || height > 8192)
            {
                error = "The formula exceeds the EMF size limit.";
                return false;
            }

            if (!EmfImageRenderer.TryCreate(
                width,
                height,
                delegate(Graphics graphics)
                {
                    EmfImageRenderer.DrawGeometry(
                        graphics,
                        geometry,
                        (float)(2d - bounds.X),
                        (float)(2d - bounds.Y),
                        Color.Black);
                },
                out emfBytes,
                out error))
            {
                return false;
            }

            pixelWidth = width;
            pixelHeight = height;
            lock (CacheLock)
            {
                Cache[cacheKey] = new CacheEntry(emfBytes, width, height);
                CacheOrder.Enqueue(cacheKey);
                while (CacheOrder.Count > MaxCacheEntries)
                {
                    Cache.Remove(CacheOrder.Dequeue());
                }
            }
            return true;
        }

        public bool TryRenderInlineTextToPng(
            string markdown,
            string textFontFamily,
            double textFontSize,
            out byte[] pngBytes,
            out int pixelWidth,
            out int pixelHeight,
            out string error)
        {
            pngBytes = null;
            pixelWidth = 0;
            pixelHeight = 0;
            error = string.Empty;

            string source = markdown ?? string.Empty;
            MatchCollection matches = InlineFormulaRegex.Matches(source);
            if (matches.Count == 0)
            {
                error = "No complete inline formula was found.";
                return false;
            }

            List<InlinePart> parts = new List<InlinePart>();
            try
            {
                int offset = 0;
                for (int i = 0; i < matches.Count; i++)
                {
                    Match match = matches[i];
                    if (match.Index > offset)
                    {
                        parts.Add(InlinePart.FromText(source.Substring(offset, match.Index - offset)));
                    }

                    byte[] formulaBytes;
                    int formulaWidth;
                    int formulaHeight;
                    string formulaError;
                    if (!TryRenderToPng(
                        match.Groups[1].Value,
                        textFontFamily,
                        out formulaBytes,
                        out formulaWidth,
                        out formulaHeight,
                        out formulaError))
                    {
                        error = formulaError;
                        return false;
                    }

                    using (MemoryStream formulaStream = new MemoryStream(formulaBytes))
                    using (Image formulaImage = Image.FromStream(formulaStream, true, true))
                    {
                        parts.Add(InlinePart.FromImage(new Bitmap(formulaImage)));
                    }
                    offset = match.Index + match.Length;
                }

                if (offset < source.Length)
                {
                    parts.Add(InlinePart.FromText(source.Substring(offset)));
                }

                string fontName = string.IsNullOrWhiteSpace(textFontFamily) ? "Calibri" : textFontFamily.Trim();
                float fontSize = (float)(textFontSize <= 0d ? 11d : textFontSize);
                using (Font font = new Font(fontName, fontSize, FontStyle.Regular, GraphicsUnit.Point))
                using (Bitmap measureBitmap = new Bitmap(1, 1, PixelFormat.Format32bppArgb))
                using (Graphics measure = Graphics.FromImage(measureBitmap))
                {
                    measure.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                    int width = 4;
                    int height = 1;
                    for (int i = 0; i < parts.Count; i++)
                    {
                        InlinePart part = parts[i];
                        if (part.Image != null)
                        {
                            part.Width = part.Image.Width;
                            part.Height = part.Image.Height;
                        }
                        else
                        {
                            SizeF size = measure.MeasureString(
                                part.Text,
                                font,
                                int.MaxValue,
                                StringFormat.GenericTypographic);
                            part.Width = Math.Max(1, (int)Math.Ceiling(size.Width));
                            part.Height = Math.Max(1, (int)Math.Ceiling(size.Height));
                        }
                        width += part.Width;
                        height = Math.Max(height, part.Height);
                    }

                    width = Math.Min(width + 4, 4096);
                    height += 8;
                    using (Bitmap output = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                    using (Graphics graphics = Graphics.FromImage(output))
                    {
                        graphics.Clear(Color.Transparent);
                        graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                        int x = 4;
                        for (int i = 0; i < parts.Count && x < width - 4; i++)
                        {
                            InlinePart part = parts[i];
                            int y = Math.Max(2, (height - part.Height) / 2);
                            if (part.Image != null)
                            {
                                graphics.DrawImageUnscaled(part.Image, x, y);
                            }
                            else
                            {
                                graphics.DrawString(
                                    part.Text,
                                    font,
                                    Brushes.Black,
                                    new PointF(x, y),
                                    StringFormat.GenericTypographic);
                            }
                            x += part.Width;
                        }

                        using (MemoryStream outputStream = new MemoryStream())
                        {
                            output.Save(outputStream, ImageFormat.Png);
                            pngBytes = outputStream.ToArray();
                        }
                    }
                    pixelWidth = width;
                    pixelHeight = height;
                    return pngBytes.Length > 0;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Inline LaTeX rendering failed", ex);
                error = ex.Message;
                return false;
            }
            finally
            {
                for (int i = 0; i < parts.Count; i++)
                {
                    if (parts[i].Image != null) parts[i].Image.Dispose();
                }
            }
        }

        public bool TryRenderInlineTextToEmf(
            string markdown,
            string textFontFamily,
            double textFontSize,
            out byte[] emfBytes,
            out int pixelWidth,
            out int pixelHeight,
            out string error)
        {
            emfBytes = null;
            pixelWidth = 0;
            pixelHeight = 0;
            error = string.Empty;

            string source = markdown ?? string.Empty;
            MatchCollection matches = InlineFormulaRegex.Matches(source);
            if (matches.Count == 0)
            {
                error = "No complete inline formula was found.";
                return false;
            }

            List<InlineVectorPart> parts = new List<InlineVectorPart>();
            string formulaFont = string.IsNullOrWhiteSpace(textFontFamily) ? "Cambria Math" : textFontFamily.Trim();
            int offset = 0;
            for (int i = 0; i < matches.Count; i++)
            {
                Match match = matches[i];
                if (match.Index > offset)
                {
                    parts.Add(InlineVectorPart.FromText(source.Substring(offset, match.Index - offset)));
                }

                Media.Geometry geometry;
                System.Windows.Rect bounds;
                if (!TryBuildGeometry(
                    match.Groups[1].Value,
                    formulaFont,
                    TexStyle.Text,
                    out geometry,
                    out bounds,
                    out error))
                {
                    return false;
                }
                parts.Add(InlineVectorPart.FromGeometry(geometry, bounds));
                offset = match.Index + match.Length;
            }
            if (offset < source.Length)
            {
                parts.Add(InlineVectorPart.FromText(source.Substring(offset)));
            }

            string fontName = string.IsNullOrWhiteSpace(textFontFamily) ? "Calibri" : textFontFamily.Trim();
            float fontSize = (float)(textFontSize <= 0d ? 11d : textFontSize);
            using (Font font = new Font(fontName, fontSize, FontStyle.Regular, GraphicsUnit.Point))
            using (Bitmap measureBitmap = new Bitmap(1, 1, PixelFormat.Format32bppArgb))
            using (Graphics measure = Graphics.FromImage(measureBitmap))
            {
                int width = 8;
                int height = 1;
                for (int i = 0; i < parts.Count; i++)
                {
                    InlineVectorPart part = parts[i];
                    if (part.Geometry != null)
                    {
                        part.Width = Math.Max(1, (int)Math.Ceiling(part.Bounds.Width));
                        part.Height = Math.Max(1, (int)Math.Ceiling(part.Bounds.Height));
                    }
                    else
                    {
                        SizeF size = measure.MeasureString(
                            part.Text,
                            font,
                            int.MaxValue,
                            StringFormat.GenericTypographic);
                        part.Width = Math.Max(1, (int)Math.Ceiling(size.Width));
                        part.Height = Math.Max(1, (int)Math.Ceiling(size.Height));
                    }
                    width += part.Width;
                    height = Math.Max(height, part.Height);
                }

                width = Math.Min(width, 8192);
                height = Math.Min(height + 8, 8192);
                if (!EmfImageRenderer.TryCreate(
                    width,
                    height,
                    delegate(Graphics graphics)
                    {
                        int x = 4;
                        for (int i = 0; i < parts.Count && x < width - 4; i++)
                        {
                            InlineVectorPart part = parts[i];
                            int y = Math.Max(2, (height - part.Height) / 2);
                            if (part.Geometry != null)
                            {
                                EmfImageRenderer.DrawGeometry(
                                    graphics,
                                    part.Geometry,
                                    (float)(x - part.Bounds.X),
                                    (float)(y - part.Bounds.Y),
                                    Color.Black);
                            }
                            else
                            {
                                graphics.DrawString(
                                    part.Text,
                                    font,
                                    Brushes.Black,
                                    new PointF(x, y),
                                    StringFormat.GenericTypographic);
                            }
                            x += part.Width;
                        }
                    },
                    out emfBytes,
                    out error))
                {
                    return false;
                }

                pixelWidth = width;
                pixelHeight = height;
                return true;
            }
        }

        internal static void ClearCache()
        {
            lock (CacheLock)
            {
                Cache.Clear();
                CacheOrder.Clear();
            }
        }

        private static string ComputeKey(string value)
        {
            using (SHA256 sha = SHA256.Create())
            {
                return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty)));
            }
        }

        private static string Normalize(string latex)
        {
            string text = (latex ?? string.Empty).Trim();
            if (text.Length == 0)
            {
                return string.Empty;
            }

            if (text.StartsWith("$$", StringComparison.Ordinal) &&
                text.EndsWith("$$", StringComparison.Ordinal) &&
                text.Length > 4)
            {
                text = text.Substring(2, text.Length - 4).Trim();
            }

            return text.Replace("\r\n", "\n").Replace('\r', '\n').Trim();
        }

        private static bool TryBuildGeometry(
            string latex,
            string font,
            TexStyle style,
            out Media.Geometry geometry,
            out System.Windows.Rect bounds,
            out string error)
        {
            geometry = null;
            bounds = System.Windows.Rect.Empty;
            error = string.Empty;
            try
            {
                TexFormula formula;
                lock (ParserLock)
                {
                    formula = WpfTeXFormulaParser.Instance.Parse(Normalize(latex), "text");
                }
                TexEnvironment environment = WpfTeXEnvironment.Create(
                    style,
                    RenderScale,
                    font,
                    Media.Brushes.Transparent,
                    Media.Brushes.Black);
                geometry = WpfTeXFormulaExtensions.RenderToGeometry(
                    formula,
                    environment,
                    RenderScale,
                    0d,
                    0d);
                bounds = geometry == null ? System.Windows.Rect.Empty : geometry.Bounds;
                if (geometry == null || bounds.IsEmpty || bounds.Width <= 0d || bounds.Height <= 0d)
                {
                    geometry = null;
                    error = "The formula renderer returned empty vector geometry.";
                    return false;
                }
                return true;
            }
            catch (TexException ex)
            {
                error = ex.Message;
                return false;
            }
            catch (Exception ex)
            {
                Logger.Error("LaTeX vector rendering failed", ex);
                error = ex.Message;
                return false;
            }
        }

        private static bool TryReadPngSize(byte[] png, out int width, out int height)
        {
            width = 0;
            height = 0;
            if (png == null || png.Length < 24) return false;
            if (!(png[0] == 0x89 && png[1] == 0x50 && png[2] == 0x4E && png[3] == 0x47)) return false;
            if (!(png[12] == 0x49 && png[13] == 0x48 && png[14] == 0x44 && png[15] == 0x52)) return false;

            width = (png[16] << 24) | (png[17] << 16) | (png[18] << 8) | png[19];
            height = (png[20] << 24) | (png[21] << 16) | (png[22] << 8) | png[23];
            return width > 0 && height > 0;
        }

        private sealed class CacheEntry
        {
            public CacheEntry(byte[] bytes, int width, int height)
            {
                Bytes = (byte[])bytes.Clone();
                Width = width;
                Height = height;
            }

            public byte[] Bytes;
            public int Width;
            public int Height;
        }

        private sealed class InlinePart
        {
            public string Text;
            public Bitmap Image;
            public int Width;
            public int Height;

            public static InlinePart FromText(string value)
            {
                return new InlinePart { Text = value ?? string.Empty };
            }

            public static InlinePart FromImage(Bitmap value)
            {
                return new InlinePart { Image = value };
            }
        }

        private sealed class InlineVectorPart
        {
            public string Text;
            public Media.Geometry Geometry;
            public System.Windows.Rect Bounds;
            public int Width;
            public int Height;

            public static InlineVectorPart FromText(string value)
            {
                return new InlineVectorPart { Text = value ?? string.Empty };
            }

            public static InlineVectorPart FromGeometry(Media.Geometry geometry, System.Windows.Rect bounds)
            {
                return new InlineVectorPart { Geometry = geometry, Bounds = bounds };
            }
        }
    }
}
