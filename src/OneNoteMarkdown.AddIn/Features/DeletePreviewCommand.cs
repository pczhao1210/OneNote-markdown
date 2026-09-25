using System;
using System.Windows.Forms;
using OneNoteMarkdown.Localization;
using OneNoteMarkdown.Logging;
using OneNoteMarkdown.OneNote;
using OneNoteMarkdown.OneNote.Models;
using OneNoteMarkdown.UI;

namespace OneNoteMarkdown.Features
{
    public static class DeletePreviewCommand
    {
        public static void Execute()
        {
            try
            {
                OneNoteProvider provider = new OneNoteProvider();
                OeInfo current = provider.GetCurrentOeInfo();
                if (current == null || string.IsNullOrWhiteSpace(current.ObjectId))
                {
                    Msg.Show(Loc.S("Msg.NoManagedPreview"), Loc.S("Common.AppTitle"),
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                PageWriter writer = new PageWriter();
                if (!writer.DeleteManagedPreview(current.PageId, current.ObjectId))
                {
                    Msg.Show(Loc.S("Msg.NoManagedPreview"), Loc.S("Common.AppTitle"),
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Delete preview failed", ex);
                Msg.Show(Loc.S("Msg.DeletePreviewFailed", ex.Message), Loc.S("Common.AppTitle"),
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
