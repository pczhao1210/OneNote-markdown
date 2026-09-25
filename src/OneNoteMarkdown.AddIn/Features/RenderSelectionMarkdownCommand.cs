using System;
using System.Windows.Forms;
using OneNoteMarkdown.Localization;
using OneNoteMarkdown.Logging;
using OneNoteMarkdown.OneNote;
using OneNoteMarkdown.OneNote.Models;
using OneNoteMarkdown.UI;

namespace OneNoteMarkdown.Features
{
    public static class RenderSelectionMarkdownCommand
    {
        public static void Execute()
        {
            try
            {
                OneNoteProvider provider = new OneNoteProvider();
                string pageId = provider.GetCurrentPageId();
                if (string.IsNullOrWhiteSpace(pageId))
                {
                    Msg.Show(Loc.S("Msg.NoPage"), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                PreviewSource source = provider.GetCurrentSelectionPreviewSource();
                if (source == null || string.IsNullOrWhiteSpace(source.Markdown))
                {
                    Msg.Show(Loc.S("Msg.NoSelection"), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                PreviewUpdateStatus status = PreviewManager.Render(
                    source,
                    "SelectionPreview",
                    false,
                    true,
                    false,
                    true);
                if (status == PreviewUpdateStatus.Unchanged && string.IsNullOrWhiteSpace(source.Markdown))
                {
                    Msg.Show(Loc.S("Msg.RenderSelEmpty"), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

            }
            catch (Exception ex)
            {
                Logger.Error("Render selection failed", ex);
                Msg.Show(Loc.S("Msg.RenderSelFailed", ex.Message), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
