using System;
using OneNoteMarkdown.Logging;

namespace OneNoteMarkdown.Features
{
    public static class RefreshPreviewCommand
    {
        public static void Execute()
        {
            try
            {
                PreviewManager.RefreshCurrentPage(true);
            }
            catch (Exception ex)
            {
                Logger.Error("Refresh preview failed", ex);
            }
        }
    }
}
