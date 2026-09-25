using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OneNoteMarkdown.Localization;
using OneNoteMarkdown.Logging;
using OneNoteMarkdown.Markdown;
using OneNoteMarkdown.OneNote;
using OneNoteMarkdown.OneNote.Models;
using OneNoteMarkdown.UI;

namespace OneNoteMarkdown.Features
{
    public static class ImportMarkdownCommand
    {
        public static void Execute()
        {
            try
            {
                Logger.Info("ImportMarkdownCommand started");
                OneNoteProvider provider = new OneNoteProvider();

                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    dialog.Filter = "Markdown|*.md;*.markdown|*.*|*.*";
                    DialogResult result = FileDialogHost.ShowOpen(dialog);
                    Logger.Info("ImportMarkdownCommand file dialog result=" + result);
                    if (result != DialogResult.OK)
                    {
                        return;
                    }

                    string filePath = NormalizeSelectedPath(dialog.FileName);
                    string pageId = provider.GetCurrentPageId();
                    if (string.IsNullOrWhiteSpace(pageId))
                    {
                        Msg.Show(Loc.S("Msg.NoPage"), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    FileInfo fileInfo = new FileInfo(filePath);
                    if (fileInfo.Length > 10 * 1024 * 1024)
                    {
                        Msg.Show(Loc.S("Msg.FileTooLarge"), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    string markdown = File.ReadAllText(filePath, Encoding.UTF8);
                    if (string.IsNullOrWhiteSpace(markdown))
                    {
                        Msg.Show(Loc.S("Msg.FileEmpty"), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    PreviewSource source = new PreviewSource
                    {
                        PageId = pageId,
                        SourceKey = "import:" + filePath.ToUpperInvariant(),
                        Markdown = markdown,
                        BaseDirectory = fileInfo.DirectoryName
                    };

                    PageWriter writer = new PageWriter();
                    writer.UpsertImportedMarkdownSource(source);
                    PreviewSource pageSource = provider.GetCurrentPagePreviewSource();
                    if (pageSource == null || string.IsNullOrWhiteSpace(pageSource.Markdown))
                    {
                        Msg.Show(Loc.S("Msg.ParseEmpty"), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    PreviewManager.Render(pageSource, "PagePreview", false, true);

                    Logger.Info("ImportMarkdownCommand completed");
                    Msg.Show(Loc.S("Msg.ImportSuccess"), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Import markdown failed", ex);
                Msg.Show(Loc.S("Msg.ImportFailed", ex.Message), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        internal static string NormalizeSelectedPath(string selectedPath)
        {
            string path = (selectedPath ?? string.Empty).Trim().Trim('"');
            if (path.Length == 0)
            {
                throw new ArgumentException("No Markdown file was selected.", nameof(selectedPath));
            }
            if (path.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
            {
                throw new ArgumentException("The selected Markdown path contains invalid characters.", nameof(selectedPath));
            }

            return Path.GetFullPath(path);
        }
    }
}
