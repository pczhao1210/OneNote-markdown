using System;
using System.Text;
using OneNoteMarkdown.OneNote.Models;

namespace OneNoteMarkdown.Markdown
{
    internal static class MarkdownExporter
    {
        public static string Export(PageContent page)
        {
            if (page == null)
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(page.Title))
            {
                builder.Append("# ");
                builder.Append(page.Title.Trim());
                builder.AppendLine();
                builder.AppendLine();
            }

            bool firstLine = true;
            foreach (OutlineContent outline in page.Outlines)
            {
                if (outline == null || outline.TextBlocks == null)
                {
                    continue;
                }

                if (outline.IsManagedPreview)
                {
                    if (string.Equals(outline.ManagedRole, "ImportPreview", System.StringComparison.Ordinal) &&
                        !string.IsNullOrEmpty(outline.ManagedSource))
                    {
                        AppendSeparated(builder, outline.ManagedSource, ref firstLine);
                    }
                    continue;
                }

                foreach (TextBlock block in outline.TextBlocks)
                {
                    if (block == null || block.IsMarkdownContinuation)
                    {
                        continue;
                    }

                    string text = !string.IsNullOrEmpty(block.MarkdownSource)
                        ? block.MarkdownSource
                        : (block.Text ?? string.Empty).Trim();
                    if (text.Length == 0)
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(block.MarkdownSource))
                    {
                        AppendSeparated(builder, text, ref firstLine);
                        continue;
                    }

                    if (!firstLine) builder.AppendLine();
                    firstLine = false;

                    int indent = block.IndentLevel < 0 ? 0 : block.IndentLevel;
                    for (int i = 0; i < indent; i++)
                    {
                        builder.Append("  ");
                    }

                    if (LooksLikeTask(text))
                    {
                        builder.Append(text.StartsWith("☑") ? "- [x] " : "- [ ] ");
                        builder.Append(text.Substring(1).Trim());
                    }
                    else if (LooksLikeBullet(text))
                    {
                        builder.Append("- ");
                        builder.Append(text.Substring(1).Trim());
                    }
                    else
                    {
                        builder.Append(text);
                    }
                }
            }

            return builder.ToString().TrimEnd('\r', '\n');
        }

        private static void AppendSeparated(StringBuilder builder, string markdown, ref bool firstLine)
        {
            if (!firstLine && builder.Length > 0)
            {
                builder.AppendLine();
            }
            builder.Append((markdown ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n'));
            firstLine = false;
        }

        private static bool LooksLikeTask(string text)
        {
            return text.StartsWith("☐") || text.StartsWith("☑");
        }

        private static bool LooksLikeBullet(string text)
        {
            if (text.Length < 2)
            {
                return false;
            }

            char c = text[0];
            return (c == '•' || c == '-' || c == '*') && char.IsWhiteSpace(text[1]);
        }
    }
}
