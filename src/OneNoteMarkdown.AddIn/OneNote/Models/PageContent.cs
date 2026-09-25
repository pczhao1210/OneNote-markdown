using System;
using System.Collections.Generic;

namespace OneNoteMarkdown.OneNote.Models
{
    public class PageContent
    {
        public string PageId { get; set; }
        public string Title { get; set; }
        public DateTime DateCreated { get; set; }
        public DateTime DateModified { get; set; }
        public int UnsupportedObjectCount { get; set; }
        public List<OutlineContent> Outlines { get; set; } = new List<OutlineContent>();
    }

    public class OutlineContent
    {
        public string OutlineId { get; set; }
        public List<TextBlock> TextBlocks { get; set; } = new List<TextBlock>();
        public bool IsManagedPreview { get; set; }
        public bool IsManagedPreviewModified { get; set; }
        public string ManagedRole { get; set; }
        public string ManagedSource { get; set; }
    }

    public class TextBlock
    {
        public string ElementId { get; set; }
        public string Text { get; set; }
        public string RawHtml { get; set; }
        public string MarkdownSource { get; set; }
        public bool IsMarkdownContinuation { get; set; }
        public int IndentLevel { get; set; }
    }
}
