using System;
using System.IO;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using OneNoteMarkdown.Logging;
using WpfMath;
using WpfMath.Parsers;
using XamlMath.Exceptions;

namespace OneNoteMarkdown.Rendering
{
    internal sealed class LatexImageRenderer
    {
        private const double RenderScale = 20.0;
        private static readonly object ParserLock = new object();
        private static readonly object CacheLock = new object();
        private static readonly Dictionary<string, CacheEntry> Cache = new Dictionary<string, CacheEntry>(StringComparer.Ordinal);
        private static readonly Queue<string> CacheOrder = new Queue<string>();
        private const int MaxCacheEntries = 128;

        public bool TryRenderToPng(string latex, string textFontFamily, out byte[] pngBytes, out int pixelWidth, out int pixelHeight)
        {
            string error;
            return TryRenderToPng(latex, textFontFamily, out pngBytes, out pixelWidth, out pixelHeight, out error);
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
            string cacheKey = ComputeKey(font + "\n" + text);
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
                    return false;
                }

                int width;
                int height;
                if (!TryReadPngSize(bytes, out width, out height))
                {
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

        private static string Normalize(string latex)
        {
            string text = (latex ?? string.Empty).Trim();
            if (text.Length == 0)
            {
                return string.Empty;
            }

            if (text.StartsWith("$$", StringComparison.Ordinal) && text.EndsWith("$$", StringComparison.Ordinal) && text.Length > 4)
            {
                text = text.Substring(2, text.Length - 4).Trim();
            }

            return text.Replace("\r\n", "\n").Replace('\r', '\n').Trim();
        }

        private static bool TryReadPngSize(byte[] png, out int width, out int height)
        {
            width = 0;
            height = 0;
            if (png == null || png.Length < 24) return false;

            // PNG signature + first chunk should contain IHDR width/height.
            if (!(png[0] == 0x89 && png[1] == 0x50 && png[2] == 0x4E && png[3] == 0x47)) return false;
            if (!(png[12] == 0x49 && png[13] == 0x48 && png[14] == 0x44 && png[15] == 0x52)) return false;

            width = (png[16] << 24) | (png[17] << 16) | (png[18] << 8) | png[19];
            height = (png[20] << 24) | (png[21] << 16) | (png[22] << 8) | png[23];
            return width > 0 && height > 0;
        }
    }
}
