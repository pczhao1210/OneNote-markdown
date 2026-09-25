using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using OneNoteMarkdown.Logging;
using OneNoteMarkdown.OneNote.Models;

namespace OneNoteMarkdown.Settings
{
    internal sealed class ThemeSettings
    {
        private static readonly string SettingsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OneNoteMarkdown",
            "settings");
        private static readonly string SettingsPath = Path.Combine(SettingsDir, "theme.ini");

        public string DefaultFontFamily { get; private set; }
        public string MonospaceFontFamily { get; private set; }
        public string MathFontFamily { get; private set; }
        public double ParagraphFontSize { get; private set; }
        public double CodeFontSize { get; private set; }
        public bool EnableLatexToImage { get; private set; }
        public bool EnableCodeLineNumber { get; private set; }
        public string Language { get; private set; }
        public string ThemePreset { get; private set; }
        public bool PreviewShowTitle { get; private set; }
        public string PreviewTitle { get; private set; }
        public PreviewPlacement PreviewPlacement { get; private set; }
        public double PreviewGap { get; private set; }
        public double PreviewWidth { get; private set; }
        public bool AutoRefreshEnabled { get; private set; }
        public int AutoRefreshDelayMilliseconds { get; private set; }
        public bool AllowRemoteImages { get; private set; }
        public bool ImportKeepSource { get; private set; }
        public int DiagramTimeoutMilliseconds { get; private set; }
        public string HeadingColor { get; private set; }
        public string CodeBackgroundColor { get; private set; }
        public string QuotePrefix { get; private set; }

        internal string RenderFingerprint
        {
            get
            {
                return string.Join("|",
                    ThemePreset,
                    DefaultFontFamily,
                    MonospaceFontFamily,
                    MathFontFamily,
                    ParagraphFontSize.ToString(CultureInfo.InvariantCulture),
                    CodeFontSize.ToString(CultureInfo.InvariantCulture),
                    EnableLatexToImage ? "1" : "0",
                    EnableCodeLineNumber ? "1" : "0",
                    AllowRemoteImages ? "1" : "0",
                    HeadingColor,
                    CodeBackgroundColor,
                    QuotePrefix);
            }
        }

        private ThemeSettings()
        {
            DefaultFontFamily = "Calibri";
            MonospaceFontFamily = "Consolas";
            MathFontFamily = "Cambria Math";
            ParagraphFontSize = 11.0;
            CodeFontSize = 10.0;
            EnableLatexToImage = true;
            EnableCodeLineNumber = false;
            Language = "auto";
            ThemePreset = "technical";
            PreviewShowTitle = true;
            PreviewTitle = "Markdown Render";
            PreviewPlacement = PreviewPlacement.Right;
            PreviewGap = 40.0;
            PreviewWidth = 520.0;
            AutoRefreshEnabled = false;
            AutoRefreshDelayMilliseconds = 700;
            AllowRemoteImages = false;
            ImportKeepSource = false;
            DiagramTimeoutMilliseconds = 3000;
            HeadingColor = "#3f3f46";
            CodeBackgroundColor = "#f5f5f5";
            QuotePrefix = "▌ ";
        }

        public static ThemeSettings Load()
        {
            ThemeSettings s = new ThemeSettings();
            EnsureDefaultFile();
            if (!File.Exists(SettingsPath)) return s;

            string[] lines;
            try { lines = File.ReadAllLines(SettingsPath); }
            catch { return s; }

            Dictionary<string, string> kv = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < lines.Length; i++)
            {
                string line = (lines[i] ?? string.Empty).Trim();
                if (line.Length == 0 || line.StartsWith("#") || line.StartsWith(";")) continue;
                int eq = line.IndexOf('=');
                if (eq <= 0) continue;
                string key = line.Substring(0, eq).Trim();
                string val = line.Substring(eq + 1).Trim();
                if (key.Length == 0) continue;
                kv[key] = val;
            }

            string value;
            if (kv.TryGetValue("theme.preset", out value) && !string.IsNullOrWhiteSpace(value))
            {
                s.ApplyPreset(value);
            }
            if (kv.TryGetValue("font.family", out value) && !string.IsNullOrWhiteSpace(value)) s.DefaultFontFamily = value;
            if (kv.TryGetValue("font.monospace", out value) && !string.IsNullOrWhiteSpace(value)) s.MonospaceFontFamily = value;
            if (kv.TryGetValue("font.math", out value) && !string.IsNullOrWhiteSpace(value)) s.MathFontFamily = value;
            if (kv.TryGetValue("font.size.paragraph", out value)) s.ParagraphFontSize = ParseDouble(value, s.ParagraphFontSize);
            if (kv.TryGetValue("font.size.code", out value)) s.CodeFontSize = ParseDouble(value, s.CodeFontSize);
            if (kv.TryGetValue("enable.latex.image", out value)) s.EnableLatexToImage = ParseBool(value, s.EnableLatexToImage);
            if (kv.TryGetValue("enable.code.lineNumber", out value)) s.EnableCodeLineNumber = ParseBool(value, s.EnableCodeLineNumber);
            if (kv.TryGetValue("language", out value) && !string.IsNullOrWhiteSpace(value)) s.Language = value.Trim().ToLowerInvariant();
            if (kv.TryGetValue("preview.title.show", out value)) s.PreviewShowTitle = ParseBool(value, s.PreviewShowTitle);
            if (kv.TryGetValue("preview.title.text", out value)) s.PreviewTitle = value ?? string.Empty;
            if (kv.TryGetValue("preview.position", out value))
            {
                s.PreviewPlacement = string.Equals(value, "below", StringComparison.OrdinalIgnoreCase)
                    ? PreviewPlacement.Below
                    : PreviewPlacement.Right;
            }
            if (kv.TryGetValue("preview.gap", out value)) s.PreviewGap = ParseRange(value, s.PreviewGap, 0.0, 500.0);
            if (kv.TryGetValue("preview.width", out value)) s.PreviewWidth = ParseRange(value, s.PreviewWidth, 160.0, 2000.0);
            if (kv.TryGetValue("preview.autoRefresh", out value)) s.AutoRefreshEnabled = ParseBool(value, s.AutoRefreshEnabled);
            if (kv.TryGetValue("preview.autoRefresh.delayMs", out value))
            {
                s.AutoRefreshDelayMilliseconds = ParseInt(value, s.AutoRefreshDelayMilliseconds, 250, 10000);
            }
            if (kv.TryGetValue("image.allowRemote", out value)) s.AllowRemoteImages = ParseBool(value, s.AllowRemoteImages);
            if (kv.TryGetValue("import.keepSource", out value)) s.ImportKeepSource = ParseBool(value, s.ImportKeepSource);
            if (kv.TryGetValue("diagram.timeoutMs", out value))
            {
                s.DiagramTimeoutMilliseconds = ParseInt(value, s.DiagramTimeoutMilliseconds, 250, 30000);
            }

            return s;
        }

        public static string EnsureDefaultFile()
        {
            try
            {
                Directory.CreateDirectory(SettingsDir);
                if (!File.Exists(SettingsPath))
                {
                    File.WriteAllText(SettingsPath,
                        "# OneNote Markdown theme settings\r\n" +
                        "# Edit values and re-render page to apply\r\n" +
                        "font.family=Calibri\r\n" +
                        "font.monospace=Consolas\r\n" +
                        "font.math=Cambria Math\r\n" +
                        "font.size.paragraph=11\r\n" +
                        "font.size.code=10\r\n" +
                        "enable.latex.image=true\r\n" +
                        "enable.code.lineNumber=false\r\n" +
                        "theme.preset=technical\r\n" +
                        "preview.title.show=true\r\n" +
                        "preview.title.text=Markdown Render\r\n" +
                        "preview.position=right\r\n" +
                        "preview.gap=40\r\n" +
                        "preview.width=520\r\n" +
                        "preview.autoRefresh=false\r\n" +
                        "preview.autoRefresh.delayMs=700\r\n" +
                        "image.allowRemote=false\r\n" +
                        "import.keepSource=false\r\n" +
                        "diagram.timeoutMs=3000\r\n" +
                        "language=auto\r\n");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("ThemeSettings: failed to create default file", ex);
            }
            return SettingsPath;
        }

        private static double ParseDouble(string value, double fallback)
        {
            double parsed;
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed) && parsed > 0.0)
            {
                return parsed;
            }
            return fallback;
        }

        private static bool ParseBool(string value, bool fallback)
        {
            if (string.IsNullOrWhiteSpace(value)) return fallback;
            bool parsed;
            if (bool.TryParse(value, out parsed)) return parsed;
            string v = value.Trim();
            if (string.Equals(v, "1", StringComparison.OrdinalIgnoreCase)) return true;
            if (string.Equals(v, "0", StringComparison.OrdinalIgnoreCase)) return false;
            if (string.Equals(v, "yes", StringComparison.OrdinalIgnoreCase)) return true;
            if (string.Equals(v, "no", StringComparison.OrdinalIgnoreCase)) return false;
            return fallback;
        }

        private void ApplyPreset(string value)
        {
            string preset = (value ?? string.Empty).Trim().ToLowerInvariant();
            if (preset == "study")
            {
                ThemePreset = preset;
                DefaultFontFamily = "Calibri";
                ParagraphFontSize = 12.0;
                CodeFontSize = 10.0;
                HeadingColor = "#6b21a8";
                CodeBackgroundColor = "#faf5ff";
                QuotePrefix = "✦ ";
                return;
            }
            if (preset == "minimal")
            {
                ThemePreset = preset;
                DefaultFontFamily = "Segoe UI";
                ParagraphFontSize = 10.5;
                CodeFontSize = 9.5;
                HeadingColor = "#374151";
                CodeBackgroundColor = "automatic";
                QuotePrefix = "│ ";
                return;
            }

            ThemePreset = "technical";
            DefaultFontFamily = "Calibri";
            ParagraphFontSize = 11.0;
            CodeFontSize = 10.0;
            HeadingColor = "#1f4e79";
            CodeBackgroundColor = "#f5f5f5";
            QuotePrefix = "▌ ";
        }

        private static double ParseRange(string value, double fallback, double minimum, double maximum)
        {
            double parsed = ParseDouble(value, fallback);
            if (parsed < minimum || parsed > maximum) return fallback;
            return parsed;
        }

        private static int ParseInt(string value, int fallback, int minimum, int maximum)
        {
            int parsed;
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed)) return fallback;
            return parsed < minimum || parsed > maximum ? fallback : parsed;
        }
    }
}
