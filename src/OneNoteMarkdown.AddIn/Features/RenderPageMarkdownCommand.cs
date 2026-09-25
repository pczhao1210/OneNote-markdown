using System;
using System.Windows.Forms;
using OneNoteMarkdown.Localization;
using OneNoteMarkdown.Logging;
using OneNoteMarkdown.OneNote;
using OneNoteMarkdown.OneNote.Models;
using OneNoteMarkdown.UI;

namespace OneNoteMarkdown.Features
{
    public static class RenderPageMarkdownCommand
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
                PreviewSource source = provider.GetCurrentPagePreviewSource();
                if (source == null || string.IsNullOrWhiteSpace(source.Markdown))
                {
                    Msg.Show(Loc.S("Msg.PageContentEmpty"), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                PreviewUpdateStatus status = PreviewManager.Render(source, "PagePreview", false, true);
                if (status == PreviewUpdateStatus.Unchanged && string.IsNullOrWhiteSpace(source.Markdown))
                {
                    Msg.Show(Loc.S("Msg.RenderPageEmpty"), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

            }
            catch (Exception ex)
            {
                Logger.Error("Render page failed", ex);
                Msg.Show(Loc.S("Msg.RenderPageFailed", ex.Message), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
