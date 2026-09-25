using System;
using System.Collections.Generic;
using OneNoteMarkdown.Logging;
using OneNoteMarkdown.Markdown;
using OneNoteMarkdown.OneNote;
using OneNoteMarkdown.OneNote.Models;

namespace OneNoteMarkdown.Features
{
    /// <summary>
    /// Renders the Markdown on the current line (the OE under the cursor).
    /// Triggered by the Enter key hook when live mode is active.
    ///
    /// Flow:
    ///   1. GetCurrentOeInfo()  — find active OE via piSelection
    ///   2. MarkdownRenderer.RenderToBlocks(mdSrc)
    ///   3. PageWriter.ReplaceOeWithRenderedBlocks(...)
    ///
    /// Must be called on the UI/STA thread.
    /// </summary>
    internal static class RenderCurrentLineCommand
    {
        // Guard against re-entrancy: rapid Enter presses queue multiple Execute()
        // calls on the UiThread. Only one render should run at a time.
        private static int _running;

        public static void Execute()
        {
            Execute(null);
        }

        internal static void Execute(OeInfo expected)
        {
            if (System.Threading.Interlocked.CompareExchange(ref _running, 1, 0) != 0)
            {
                return; // a render is already in progress
            }
            try
            {
                OneNoteProvider provider = new OneNoteProvider();
                if (expected != null &&
                    !string.Equals(provider.GetCurrentPageId(), expected.PageId, StringComparison.Ordinal))
                {
                    Logger.Info("RenderCurrentLine: cancelled after page change");
                    return;
                }
                OeInfo oe = expected ?? provider.GetCurrentOeInfo();
                if (oe == null)
                {
                    Logger.Info("RenderCurrentLine: no active OE found");
                    return;
                }

                string src = (oe.MarkdownSource ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(src))
                {
                    Logger.Info("RenderCurrentLine: OE source is empty, skipping");
                    return;
                }

                if (!OneNoteProvider.LooksLikeMarkdown(src))
                {
                    Logger.Info("RenderCurrentLine: source is not complete Markdown, skipping");
                    return;
                }

                Logger.Info("RenderCurrentLine: rendering source length=" + src.Length);

                List<MarkdownBlock> blocks = MarkdownRenderer.RenderToBlocks(src);
                if (blocks == null || blocks.Count == 0)
                {
                    blocks = new List<MarkdownBlock> { MarkdownBlock.Blank() };
                }

                PageWriter writer = new PageWriter();
                writer.ReplaceOeWithRenderedBlocks(oe.PageId, oe.ObjectId, src, blocks);
            }
            catch (Exception ex)
            {
                Logger.Error("RenderCurrentLine failed", ex);
            }
            finally
            {
                System.Threading.Interlocked.Exchange(ref _running, 0);
            }
        }
    }
}
