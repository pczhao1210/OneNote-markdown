using System.Linq;
using System.Windows.Forms;
using OneNoteMarkdown.Localization;
using OneNoteMarkdown.Markdown;
using OneNoteMarkdown.OneNote;
using OneNoteMarkdown.OneNote.Models;
using OneNoteMarkdown.UI;

namespace OneNoteMarkdown.Features
{
    internal static class MarkdownExportService
    {
        public static string ExportCurrentPageMarkdown(out string pageTitle)
        {
            OneNoteProvider provider = new OneNoteProvider();
            var page = provider.GetCurrentPage();
            pageTitle = page == null ? string.Empty : (page.Title ?? string.Empty);
            if (page != null && page.Outlines.Any(delegate(OutlineContent outline)
            {
                return outline != null && outline.IsManagedPreviewModified;
            }))
            {
                DialogResult result = Msg.Show(
                    Loc.S("Msg.ExportPreviewConflict"),
                    Loc.S("Common.AppTitle"),
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result != DialogResult.Yes) return string.Empty;
            }
            if (page != null && page.UnsupportedObjectCount > 0)
            {
                DialogResult result = Msg.Show(
                    Loc.S("Msg.ExportUnsupportedObjects", page.UnsupportedObjectCount),
                    Loc.S("Common.AppTitle"),
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result != DialogResult.Yes) return string.Empty;
            }
            return MarkdownExporter.Export(page);
        }
    }
}
