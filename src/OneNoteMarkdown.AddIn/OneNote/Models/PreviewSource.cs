namespace OneNoteMarkdown.OneNote.Models
{
    internal enum PreviewPlacement
    {
        Right = 0,
        Below = 1
    }

    internal sealed class PreviewSource
    {
        public string PageId { get; set; }
        public string SourceKey { get; set; }
        public string Markdown { get; set; }
        public string BaseDirectory { get; set; }
        public bool HasBounds { get; set; }
        public double Left { get; set; }
        public double Top { get; set; }
        public double Right { get; set; }
        public double Bottom { get; set; }
    }

    internal sealed class PreviewWriteOptions
    {
        public string Role { get; set; }
        public string Title { get; set; }
        public bool ShowTitle { get; set; }
        public PreviewPlacement Placement { get; set; }
        public double Gap { get; set; }
        public double Width { get; set; }
        public bool ForceLayout { get; set; }
        public bool ApplyTitleDefaults { get; set; }
        public bool OverwriteUserChanges { get; set; }
    }

    internal enum PreviewUpdateStatus
    {
        Created = 0,
        Updated = 1,
        Unchanged = 2,
        Conflict = 3
    }
}
