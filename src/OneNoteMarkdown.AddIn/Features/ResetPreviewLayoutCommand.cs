using System;
using System.Windows.Forms;
using OneNoteMarkdown.Localization;
using OneNoteMarkdown.Logging;
using OneNoteMarkdown.OneNote;
using OneNoteMarkdown.OneNote.Models;
using OneNoteMarkdown.UI;

namespace OneNoteMarkdown.Features
{
    public static class ResetPreviewLayoutCommand
    {
        public static void Execute()
        {
            try
            {
                OneNoteProvider provider = new OneNoteProvider();
                PreviewSource source = provider.GetCurrentPagePreviewSource();
                if (source == null)
                {
                    Msg.Show(Loc.S("Msg.PageContentEmpty"), Loc.S("Common.AppTitle"),
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                PreviewManager.Render(source, "PagePreview", true, true);
            }
            catch (Exception ex)
            {
                Logger.Error("Reset preview layout failed", ex);
                Msg.Show(Loc.S("Msg.ResetLayoutFailed", ex.Message), Loc.S("Common.AppTitle"),
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
