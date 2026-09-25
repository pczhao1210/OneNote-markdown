using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OneNoteMarkdown.Localization;
using OneNoteMarkdown.Markdown;
using OneNoteMarkdown.OneNote;
using OneNoteMarkdown.OneNote.Models;
using OneNoteMarkdown.Settings;
using OneNoteMarkdown.UI;

namespace OneNoteMarkdown.Features
{
    internal static class PreviewManager
    {
        internal static PreviewUpdateStatus Render(
            PreviewSource source,
            string role,
            bool forceLayout,
            bool interactive,
            bool applyTitleDefaults = false,
            bool allowCreateNew = false)
        {
            if (source == null || string.IsNullOrWhiteSpace(source.Markdown))
            {
                return PreviewUpdateStatus.Unchanged;
            }

            List<MarkdownBlock> blocks = MarkdownRenderer.RenderToBlocks(source.Markdown, source.BaseDirectory);
            if (blocks == null || blocks.Count == 0)
            {
                return PreviewUpdateStatus.Unchanged;
            }

            ThemeSettings settings = ThemeSettings.Load();
            PreviewWriteOptions options = new PreviewWriteOptions
            {
                Role = role,
                Title = settings.PreviewTitle,
                ShowTitle = settings.PreviewShowTitle,
                Placement = settings.PreviewPlacement,
                Gap = settings.PreviewGap,
                Width = settings.PreviewWidth,
                ForceLayout = forceLayout,
                ApplyTitleDefaults = applyTitleDefaults,
                OverwriteUserChanges = !settings.PreviewConflictNotificationEnabled,
                CreateNewPreview = allowCreateNew && settings.PreviewCreateNewOnRefresh
            };

            PageWriter writer = new PageWriter();
            PreviewUpdateStatus status = writer.UpsertManagedPreview(source, blocks, options);
            if (status != PreviewUpdateStatus.Conflict || !interactive)
            {
                return status;
            }

            DialogResult result = Msg.Show(
                Loc.S("Msg.PreviewConflict"),
                Loc.S("Common.AppTitle"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (result != DialogResult.Yes)
            {
                return PreviewUpdateStatus.Conflict;
            }

            options.OverwriteUserChanges = true;
            return writer.UpsertManagedPreview(source, blocks, options);
        }

        internal static PreviewUpdateStatus RefreshCurrentPage(bool interactive, bool allowCreateNew = false)
        {
            OneNoteProvider provider = new OneNoteProvider();
            PreviewSource source = provider.GetCurrentPagePreviewSource();
            return Render(source, "PagePreview", false, interactive, false, allowCreateNew);
        }
    }
}
