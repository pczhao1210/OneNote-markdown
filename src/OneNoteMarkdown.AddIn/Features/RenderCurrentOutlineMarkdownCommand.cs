using System;
using System.Windows.Forms;
using OneNoteMarkdown.Localization;
using OneNoteMarkdown.Logging;
using OneNoteMarkdown.OneNote;
using OneNoteMarkdown.OneNote.Models;
using OneNoteMarkdown.UI;

namespace OneNoteMarkdown.Features
{
    public static class RenderCurrentOutlineMarkdownCommand
    {
        public static void Execute()
        {
            try
            {
                OneNoteProvider provider = new OneNoteProvider();
                PreviewSource source = provider.GetCurrentOutlinePreviewSource();
                if (source == null || string.IsNullOrWhiteSpace(source.Markdown))
                {
                    Msg.Show(Loc.S("Msg.NoOutline"), Loc.S("Common.AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                PreviewManager.Render(source, "OutlinePreview", false, true);
            }
            catch (Exception ex)
            {
                Logger.Error("Render current outline failed", ex);
                Msg.Show(Loc.S("Msg.RenderOutlineFailed", ex.Message), Loc.S("Common.AppTitle"),
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
