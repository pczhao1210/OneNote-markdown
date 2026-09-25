using System;
using System.Windows.Forms;
using OneNoteMarkdown.Localization;
using OneNoteMarkdown.Logging;
using OneNoteMarkdown.OneNote;
using OneNoteMarkdown.UI;

namespace OneNoteMarkdown.Features
{
    public static class JumpToPreviewSourceCommand
    {
        public static void Execute()
        {
            try
            {
                OneNoteProvider provider = new OneNoteProvider();
                if (!provider.NavigateToCurrentPreviewSource())
                {
                    Msg.Show(Loc.S("Msg.NoManagedPreview"), Loc.S("Common.AppTitle"),
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Jump to preview source failed", ex);
                Msg.Show(Loc.S("Msg.JumpSourceFailed", ex.Message), Loc.S("Common.AppTitle"),
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
