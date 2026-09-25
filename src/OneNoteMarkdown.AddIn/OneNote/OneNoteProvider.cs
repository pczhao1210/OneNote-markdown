using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using OneNoteMarkdown.OneNote.Interop;
using OneNoteMarkdown.OneNote.Models;

namespace OneNoteMarkdown.OneNote
{
    public class OneNoteProvider
    {
        private static readonly XNamespace OneNs = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        private static readonly Regex BreakRegex = new Regex("<br\\s*/?>", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex BlockCloseRegex = new Regex("</(p|div|li|h[1-6])>", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex TagRegex = new Regex("<[^>]+>", RegexOptions.Compiled);
        private readonly IOneNoteApplication _app;

        public OneNoteProvider()
        {
            object raw;
            try { raw = System.Runtime.InteropServices.Marshal.GetActiveObject("OneNote.Application"); }
            catch { raw = AddIn.Connect.Instance?.OneNoteApp; }
            if (raw == null) throw new InvalidOperationException("OneNote application not available.");
            _app = raw as IOneNoteApplication ?? throw new InvalidOperationException("Cast to IOneNoteApplication failed.");
        }

        public PageContent GetCurrentPage()
        {
            string pageId = GetCurrentPageId();
            if (string.IsNullOrWhiteSpace(pageId)) throw new InvalidOperationException("Current page is not available.");
            return GetPage(pageId);
        }

        public string GetCurrentPageId()
        {
            Microsoft.Office.Interop.OneNote.Windows windows = _app.GetWindows();
            if (windows == null) return string.Empty;
            Window currentWindow = windows.CurrentWindow;
            if (currentWindow == null) return string.Empty;
            return currentWindow.CurrentPageId ?? string.Empty;
        }

        public PageContent GetPage(string pageId)
        {
            if (string.IsNullOrWhiteSpace(pageId)) throw new ArgumentException("Page ID cannot be null or empty.", nameof(pageId));
            return PageParser.Parse(GetPageXml(pageId));
        }

        public string GetPageXml(string pageId)
        {
            if (string.IsNullOrWhiteSpace(pageId)) throw new ArgumentException("Page ID cannot be null or empty.", nameof(pageId));
            string xml;
            _app.GetPageContent(pageId, out xml, PageInfo.piBasic, XMLSchema.xs2013);
            return xml;
        }

        public string GetCurrentSelectionText()
        {
            PreviewSource source = GetCurrentSelectionPreviewSource();
            return source == null ? string.Empty : source.Markdown;
        }

        internal PreviewSource GetCurrentSelectionPreviewSource()
        {
            string pageId = GetCurrentPageId();
            if (string.IsNullOrWhiteSpace(pageId)) return null;

            string xml;
            _app.GetPageContent(pageId, out xml, PageInfo.piSelection, XMLSchema.xs2013);
            if (string.IsNullOrWhiteSpace(xml)) return null;

            XDocument doc;
            try { doc = XDocument.Parse(xml); }
            catch { return null; }

            List<string> texts = new List<string>();
            List<string> objectIds = new List<string>();
            List<XElement> selectedTextNodes = doc.Descendants(OneNs + "T")
                .Where(IsInsideSelectedSubtree)
                .ToList();
            if (selectedTextNodes.Count == 0)
            {
                return null;
            }

            foreach (XElement t in selectedTextNodes)
            {
                string plain = HtmlToPlainText(t.Value, true);
                if (!string.IsNullOrWhiteSpace(plain)) texts.Add(plain);
                XElement owner = t.Ancestors(OneNs + "OE").FirstOrDefault();
                string objectId = owner == null ? null : (string)owner.Attribute("objectID");
                if (!string.IsNullOrWhiteSpace(objectId) && !objectIds.Contains(objectId))
                {
                    objectIds.Add(objectId);
                }
            }

            string markdown = string.Join("\n", texts).Trim();
            if (markdown.Length == 0) return null;
            objectIds.Sort(StringComparer.Ordinal);
            PreviewSource result = new PreviewSource
            {
                PageId = pageId,
                SourceKey = "selection:" + string.Join("|", objectIds),
                Markdown = markdown
            };
            ApplyBounds(result, selectedTextNodes.Select(delegate(XElement node)
            {
                return node.Ancestors(OneNs + "Outline").FirstOrDefault();
            }));
            return result;
        }

        public string GetCurrentPageTextForRender()
        {
            PreviewSource source = GetCurrentPagePreviewSource();
            return source == null ? string.Empty : source.Markdown;
        }

        internal PreviewSource GetCurrentPagePreviewSource()
        {
            string pageId = GetCurrentPageId();
            if (string.IsNullOrWhiteSpace(pageId)) return null;

            string xml;
            _app.GetPageContent(pageId, out xml, PageInfo.piAll, XMLSchema.xs2013);
            if (string.IsNullOrWhiteSpace(xml)) return null;

            XDocument doc;
            try { doc = XDocument.Parse(xml); }
            catch { return null; }

            List<string> texts = new List<string>();
            List<string> baseDirectories = new List<string>();
            List<XElement> outlines = doc.Root == null
                ? new List<XElement>()
                : doc.Root.Elements(OneNs + "Outline").ToList();
            for (int oi = 0; oi < outlines.Count; oi++)
            {
                XElement outline = outlines[oi];
                string baseDirectory = ReadSourceBaseDirectory(outline);
                if (!string.IsNullOrWhiteSpace(baseDirectory) &&
                    !baseDirectories.Contains(baseDirectory, StringComparer.OrdinalIgnoreCase))
                {
                    baseDirectories.Add(baseDirectory);
                }
                if (IsManagedOutline(outline)) continue;
                string text = ExtractOutlineText(outline);
                if (!string.IsNullOrWhiteSpace(text)) texts.Add(text);
            }

            string markdown = string.Join("\n\n", texts).Trim();
            if (markdown.Length == 0) return null;
            PreviewSource result = new PreviewSource
            {
                PageId = pageId,
                SourceKey = "page:" + pageId,
                Markdown = markdown,
                BaseDirectory = baseDirectories.Count == 1 ? baseDirectories[0] : null
            };
            ApplyBounds(result, outlines.Where(delegate(XElement outline) { return !IsManagedOutline(outline); }));
            return result;
        }

        internal PreviewSource GetCurrentOutlinePreviewSource()
        {
            string pageId = GetCurrentPageId();
            if (string.IsNullOrWhiteSpace(pageId)) return null;

            string selectionXml;
            _app.GetPageContent(pageId, out selectionXml, PageInfo.piSelection, XMLSchema.xs2013);
            if (string.IsNullOrWhiteSpace(selectionXml)) return null;

            XDocument selection;
            try { selection = XDocument.Parse(selectionXml); }
            catch { return null; }
            XElement active = FindDeepestSelectedOe(selection);
            if (active == null) return null;
            XElement selectedOutline = active.Ancestors(OneNs + "Outline").FirstOrDefault();
            string outlineId = selectedOutline == null ? null :
                ((string)selectedOutline.Attribute("objectID") ?? (string)selectedOutline.Attribute("ID"));

            string fullXml;
            _app.GetPageContent(pageId, out fullXml, PageInfo.piAll, XMLSchema.xs2013);
            XDocument full;
            try { full = XDocument.Parse(fullXml); }
            catch { return null; }

            XElement outline = null;
            if (!string.IsNullOrWhiteSpace(outlineId))
            {
                outline = full.Descendants(OneNs + "Outline").FirstOrDefault(delegate(XElement candidate)
                {
                    return string.Equals((string)candidate.Attribute("objectID"), outlineId, StringComparison.Ordinal)
                        || string.Equals((string)candidate.Attribute("ID"), outlineId, StringComparison.Ordinal);
                });
            }
            if (outline == null)
            {
                string objectId = (string)active.Attribute("objectID");
                XElement fullOe = full.Descendants(OneNs + "OE").FirstOrDefault(delegate(XElement candidate)
                {
                    return string.Equals((string)candidate.Attribute("objectID"), objectId, StringComparison.Ordinal);
                });
                outline = fullOe == null ? null : fullOe.Ancestors(OneNs + "Outline").FirstOrDefault();
            }
            if (outline == null || IsManagedOutline(outline)) return null;

            string markdown = ExtractOutlineText(outline);
            if (string.IsNullOrWhiteSpace(markdown)) return null;
            string key = (string)outline.Attribute("objectID") ?? (string)outline.Attribute("ID")
                ?? (string)active.Attribute("objectID");
            PreviewSource result = new PreviewSource
            {
                PageId = pageId,
                SourceKey = "outline:" + (key ?? pageId),
                Markdown = markdown
            };
            ApplyBounds(result, new[] { outline });
            return result;
        }

        public string GetManagedOutlineText(string pageId, string role)
        {
            if (string.IsNullOrWhiteSpace(pageId) || string.IsNullOrWhiteSpace(role)) return string.Empty;

            string xml;
            _app.GetPageContent(pageId, out xml, PageInfo.piAll, XMLSchema.xs2013);
            if (string.IsNullOrWhiteSpace(xml)) return string.Empty;

            XDocument doc;
            try { doc = XDocument.Parse(xml); }
            catch { return string.Empty; }

            string heading = ResolveManagedHeading(role);
            if (string.IsNullOrWhiteSpace(heading)) return string.Empty;

            XElement outline = doc.Root == null
                ? null
                : doc.Root.Elements(OneNs + "Outline").FirstOrDefault(delegate(XElement o)
                {
                    XElement firstT = o.Descendants(OneNs + "T").FirstOrDefault();
                    if (firstT == null) return false;
                    string plain = HtmlToPlainText(firstT.Value, false);
                    return string.Equals((plain ?? string.Empty).Trim(), heading, StringComparison.Ordinal);
                });
            if (outline == null) return string.Empty;

            List<string> lines = new List<string>();
            List<XElement> nodes = outline.Descendants(OneNs + "T").ToList();
            for (int i = 1; i < nodes.Count; i++)
            {
                string plain = HtmlToPlainText(nodes[i].Value, true);
                if (!string.IsNullOrWhiteSpace(plain)) lines.Add(plain);
            }

            return string.Join("\n", lines).Trim();
        }

        private static bool IsInsideSelectedSubtree(XElement node)
        {
            bool isTextNode = node != null && node.Name == OneNs + "T";
            for (XElement cur = node; cur != null; cur = cur.Parent)
            {
                string sel = (string)cur.Attribute("selected");
                if (string.Equals(sel, "all", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
                if (isTextNode && ReferenceEquals(cur, node) &&
                    string.Equals(sel, "partial", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsManagedOutline(XElement outline)
        {
            if (outline == null) return false;
            return outline.Descendants(OneNs + "Meta").Any(delegate(XElement meta)
            {
                string name = (string)meta.Attribute("name");
                return string.Equals(name, "md-preview-id", StringComparison.Ordinal)
                    || string.Equals(name, "md-preview-role", StringComparison.Ordinal);
            });
        }

        private static string ReadSourceBaseDirectory(XElement outline)
        {
            if (outline == null) return string.Empty;
            XElement meta = outline.Descendants(OneNs + "Meta").FirstOrDefault(delegate(XElement candidate)
            {
                string name = (string)candidate.Attribute("name");
                return string.Equals(name, "md-import-base-directory", StringComparison.Ordinal)
                    || string.Equals(name, "md-preview-base-directory", StringComparison.Ordinal);
            });
            string encoded = meta == null ? null : (string)meta.Attribute("content");
            if (string.IsNullOrWhiteSpace(encoded)) return string.Empty;
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(encoded));
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ResolveManagedHeading(string role)
        {
            if (string.Equals(role, "LiveSource", StringComparison.OrdinalIgnoreCase)) return "Markdown 源码（实时）";
            if (string.Equals(role, "LivePreview", StringComparison.OrdinalIgnoreCase)) return "Markdown 预览（实时）";
            return string.Empty;
        }

        private static string ExtractOutlineText(XElement outline)
        {
            if (outline == null) return string.Empty;
            List<string> lines = new List<string>();
            XElement children = outline.Element(OneNs + "OEChildren");
            if (children != null)
            {
                AppendOutlineSourceLines(children, 0, lines);
            }
            else
            {
                foreach (XElement text in outline.Descendants(OneNs + "T"))
                {
                    string plain = HtmlToPlainText(text.Value, true);
                    if (!string.IsNullOrWhiteSpace(plain)) lines.Add(plain);
                }
            }
            return string.Join("\n", lines).Trim();
        }

        private static void AppendOutlineSourceLines(XElement children, int depth, List<string> lines)
        {
            foreach (XElement oe in children.Elements(OneNs + "OE"))
            {
                foreach (XElement text in oe.Elements(OneNs + "T"))
                {
                    string plain = HtmlToPlainText(text.Value, true);
                    if (string.IsNullOrWhiteSpace(plain)) continue;
                    lines.Add(RestoreStructuralIndent(plain, depth));
                }

                XElement nested = oe.Element(OneNs + "OEChildren");
                if (nested != null)
                {
                    AppendOutlineSourceLines(nested, depth + 1, lines);
                }
            }
        }

        private static string RestoreStructuralIndent(string text, int depth)
        {
            if (depth <= 0 || string.IsNullOrEmpty(text)) return text ?? string.Empty;
            string indent = new string(' ', depth * 2);
            string[] parts = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i].Length > 0 && !char.IsWhiteSpace(parts[i][0]))
                {
                    parts[i] = indent + parts[i];
                }
            }
            return string.Join("\n", parts);
        }

        private static void ApplyBounds(PreviewSource source, IEnumerable<XElement> outlines)
        {
            if (source == null || outlines == null) return;
            bool found = false;
            double left = 0.0;
            double top = 0.0;
            double right = 0.0;
            double bottom = 0.0;
            foreach (XElement outline in outlines.Where(delegate(XElement item) { return item != null; }).Distinct())
            {
                XElement position = outline.Element(OneNs + "Position");
                if (position == null) continue;
                double x;
                double y;
                if (!double.TryParse((string)position.Attribute("x"), System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out x)) continue;
                if (!double.TryParse((string)position.Attribute("y"), System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out y)) continue;
                XElement size = outline.Element(OneNs + "Size");
                double width = 500.0;
                double height = 80.0;
                if (size != null)
                {
                    double parsed;
                    if (double.TryParse((string)size.Attribute("width"), System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out parsed) && parsed > 0.0) width = parsed;
                    if (double.TryParse((string)size.Attribute("height"), System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out parsed) && parsed > 0.0) height = parsed;
                }
                if (!found)
                {
                    left = x;
                    top = y;
                    right = x + width;
                    bottom = y + height;
                    found = true;
                }
                else
                {
                    left = Math.Min(left, x);
                    top = Math.Min(top, y);
                    right = Math.Max(right, x + width);
                    bottom = Math.Max(bottom, y + height);
                }
            }
            source.HasBounds = found;
            source.Left = left;
            source.Top = top;
            source.Right = right;
            source.Bottom = bottom;
        }

        /// <summary>
        /// Returns info about the OE (Outline Element) that currently contains the cursor.
        /// Uses piSelection to find the selected/active OE on the current page.
        /// Returns null if no active OE is found or the page cannot be read.
        /// </summary>
        public OeInfo GetCurrentOeInfo()
        {
            string pageId = GetCurrentPageId();
            if (string.IsNullOrWhiteSpace(pageId)) return null;

            string xml;
            _app.GetPageContent(pageId, out xml, PageInfo.piSelection, XMLSchema.xs2013);
            if (string.IsNullOrWhiteSpace(xml)) return null;

            XDocument doc;
            try { doc = XDocument.Parse(xml); }
            catch { return null; }

            // Complex content may mark both a parent container and its active
            // child as selected. Choosing the first ancestor can replace an
            // entire outline, so always choose the deepest selected OE.
            XElement activeOe = FindDeepestSelectedOe(doc);

            if (activeOe == null) return null;

            activeOe = ResolveRenderedGroupAnchor(pageId, activeOe) ?? activeOe;
            string objectId = (string)activeOe.Attribute("objectID");
            string mdSrc = ExtractOeMarkdown(activeOe);

            // When the user presses Enter at the end of a line, OneNote first
            // inserts a NEW empty OE and moves the cursor into it, so piSelection
            // reports that empty OE as active. In that case the line we actually
            // want to render is the PREVIOUS sibling OE (the one just committed).
            // Resolve it against the full page XML, since piSelection may not
            // include sibling content.
            if (string.IsNullOrWhiteSpace(mdSrc) && !string.IsNullOrWhiteSpace(objectId))
            {
                XElement prevOe = FindPreviousSiblingOe(pageId, objectId);
                if (prevOe != null)
                {
                    // Skip if the previous OE is already rendered (has md-src Meta);
                    // re-rendering it would be redundant and could disturb the cursor.
                    bool alreadyRendered = prevOe.Elements(OneNs + "Meta")
                        .Any(delegate(XElement m)
                        {
                            return string.Equals((string)m.Attribute("name"), "md-src", StringComparison.Ordinal);
                        });
                    string prevSrc = alreadyRendered ? null : ExtractOeMarkdown(prevOe);
                    if (!string.IsNullOrWhiteSpace(prevSrc) && LooksLikeMarkdown(prevSrc))
                    {
                        return new OeInfo
                        {
                            PageId = pageId,
                            ObjectId = (string)prevOe.Attribute("objectID") ?? string.Empty,
                            MarkdownSource = prevSrc
                        };
                    }
                }
            }

            return new OeInfo
            {
                PageId   = pageId,
                ObjectId = objectId ?? string.Empty,
                MarkdownSource = mdSrc ?? string.Empty
            };
        }

        /// <summary>
        /// Heuristic: returns true if the line contains Markdown syntax worth
        /// rendering on Enter. Plain prose (no markup) returns false so that
        /// ordinary line breaks are left untouched.
        /// </summary>
        internal static bool LooksLikeMarkdown(string src)
        {
            if (string.IsNullOrWhiteSpace(src)) return false;
            string t = src.TrimStart();

            // Block-level markers.
            if (t.StartsWith("#")) return true;                 // headings
            if (Regex.IsMatch(t, @"^[-+*]\s+\S")) return true;   // - + * list
            if (Regex.IsMatch(t, @"^\d+[.)]\s+\S")) return true; // 1. ordered list
            if (Regex.IsMatch(t, @"^\[( |x|X)\]\s+\S")) return true; // [ ] task
            if (t.StartsWith(">")) return true;                  // blockquote
            // Multi-line fences are rendered by F5 after the block is complete.
            // A lone opening marker is incomplete and must not be rendered yet.
            if (Regex.IsMatch(t, @"^\$\$.+\$\$$")) return true; // one-line latex block
            if (t.StartsWith("---") || t.StartsWith("***") || t.StartsWith("___")) return true; // hr
            if (t.StartsWith("|")) return true;                  // table row

            // Inline emphasis / code / link / image.
            if (Regex.IsMatch(src, @"\*\*[^*]+\*\*")) return true;     // bold
            if (Regex.IsMatch(src, @"(?<!\*)\*[^*]+\*(?!\*)")) return true; // italic
            if (Regex.IsMatch(src, @"__[^_]+__")) return true;          // bold
            if (Regex.IsMatch(src, @"~~[^~]+~~")) return true;          // strikethrough
            if (Regex.IsMatch(src, @"==[^=]+==")) return true;          // highlight
            if (Regex.IsMatch(src, @"\+\+[^+]+\+\+")) return true;      // underline
            if (Regex.IsMatch(src, @"`[^`]+`")) return true;            // inline code
            if (Regex.IsMatch(src, @"\[[^\]]+\]\([^)]+\)")) return true; // link/image
            if (Regex.IsMatch(src, @"(?<!\\)\$(?!\$)(?:\\.|[^$\r\n])+\$")) return true; // inline latex

            return false;
        }

        /// <summary>
        /// Reads the Markdown source of an OE: first the md-src Meta tag, then the
        /// visible text as fallback.
        /// </summary>
        internal static XElement FindDeepestSelectedOe(XDocument doc)
        {
            if (doc == null) return null;

            return doc.Descendants(OneNs + "OE")
                .Where(IsSelectedOe)
                .OrderByDescending(delegate(XElement oe)
                {
                    return oe.Ancestors(OneNs + "OE").Count();
                })
                .FirstOrDefault();
        }

        private static bool IsSelectedOe(XElement oe)
        {
            if (oe == null) return false;

            string oeSelection = (string)oe.Attribute("selected");
            if (string.Equals(oeSelection, "all", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(oeSelection, "partial", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return oe.Elements(OneNs + "T").Any(delegate(XElement text)
            {
                string textSelection = (string)text.Attribute("selected");
                return string.Equals(textSelection, "all", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(textSelection, "partial", StringComparison.OrdinalIgnoreCase);
            });
        }

        internal static string ExtractOeMarkdown(XElement oe)
        {
            if (oe == null) return string.Empty;

            XElement meta = oe.Elements(OneNs + "Meta")
                .FirstOrDefault(delegate(XElement m)
                {
                    return string.Equals((string)m.Attribute("name"), "md-src", StringComparison.Ordinal);
                });
            if (meta != null)
            {
                string encoded = (string)meta.Attribute("content");
                if (!string.IsNullOrEmpty(encoded))
                {
                    try { return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(encoded)); }
                    catch { /* fall through to plain text */ }
                }
            }

            List<string> parts = new List<string>();
            // Read only text owned by this OE. Descendant T elements can belong
            // to nested paragraphs, images, diagrams, or other rich content.
            foreach (XElement t in oe.Elements(OneNs + "T"))
            {
                string plain = HtmlToPlainText(t.Value, true);
                if (plain.Length > 0) parts.Add(plain);
            }
            return string.Join("", parts).Trim();
        }

        /// <summary>
        /// Locates the OE identified by objectId in the full page XML and returns
        /// its immediately preceding sibling OE (skipping non-OE elements such as
        /// Meta). Returns null if not found or there is no previous sibling.
        /// </summary>
        private XElement FindPreviousSiblingOe(string pageId, string objectId)
        {
            string xml;
            _app.GetPageContent(pageId, out xml, PageInfo.piAll, XMLSchema.xs2013);
            if (string.IsNullOrWhiteSpace(xml)) return null;

            XDocument doc;
            try { doc = XDocument.Parse(xml); }
            catch { return null; }

            XElement target = doc.Descendants(OneNs + "OE")
                .FirstOrDefault(delegate(XElement e)
                {
                    return string.Equals((string)e.Attribute("objectID"), objectId, StringComparison.Ordinal);
                });
            if (target == null) return null;

            XElement prev = null;
            for (XElement sib = target.PreviousNode as XElement; sib != null; sib = sib.PreviousNode as XElement)
            {
                if (sib.Name == OneNs + "OE") { prev = sib; break; }
            }
            return prev;
        }

        private XElement ResolveRenderedGroupAnchor(string pageId, XElement selectedOe)
        {
            if (selectedOe == null) return null;
            XElement groupMeta = selectedOe.Elements(OneNs + "Meta").FirstOrDefault(delegate(XElement meta)
            {
                return string.Equals((string)meta.Attribute("name"), "md-render-group", StringComparison.Ordinal);
            });
            string groupId = groupMeta == null ? null : (string)groupMeta.Attribute("content");
            if (string.IsNullOrWhiteSpace(groupId)) return null;

            string xml;
            _app.GetPageContent(pageId, out xml, PageInfo.piAll, XMLSchema.xs2013);
            XDocument full;
            try { full = XDocument.Parse(xml); }
            catch { return null; }

            return full.Descendants(OneNs + "OE").FirstOrDefault(delegate(XElement oe)
            {
                bool inGroup = oe.Elements(OneNs + "Meta").Any(delegate(XElement meta)
                {
                    return string.Equals((string)meta.Attribute("name"), "md-render-group", StringComparison.Ordinal)
                        && string.Equals((string)meta.Attribute("content"), groupId, StringComparison.Ordinal);
                });
                if (!inGroup) return false;
                return oe.Elements(OneNs + "Meta").Any(delegate(XElement meta)
                {
                    return string.Equals((string)meta.Attribute("name"), "md-src", StringComparison.Ordinal);
                });
            });
        }

        /// <summary>
        /// Returns true if the OE identified by objectId has a &lt;one:Meta name="md-src"&gt; element,
        /// indicating it is in rendered (not raw source) state.
        /// </summary>
        public bool CurrentOeHasMdMeta(string pageId, string objectId)
        {
            if (string.IsNullOrWhiteSpace(pageId) || string.IsNullOrWhiteSpace(objectId)) return false;
            string xml;
            _app.GetPageContent(pageId, out xml, PageInfo.piAll, XMLSchema.xs2013);
            if (string.IsNullOrWhiteSpace(xml)) return false;
            XDocument doc;
            try { doc = XDocument.Parse(xml); }
            catch { return false; }
            XElement oe = doc.Descendants(OneNs + "OE")
                .FirstOrDefault(delegate(XElement e)
                {
                    return string.Equals((string)e.Attribute("objectID"), objectId, StringComparison.Ordinal);
                });
            if (oe == null) return false;
            return oe.Elements(OneNs + "Meta")
                .Any(delegate(XElement m)
                {
                    return string.Equals((string)m.Attribute("name"), "md-src", StringComparison.Ordinal);
                });
        }

        internal bool NavigateToCurrentPreviewSource()
        {
            string pageId = GetCurrentPageId();
            if (string.IsNullOrWhiteSpace(pageId)) return false;

            string selectionXml;
            _app.GetPageContent(pageId, out selectionXml, PageInfo.piSelection, XMLSchema.xs2013);
            XDocument selection;
            try { selection = XDocument.Parse(selectionXml); }
            catch { return false; }
            XElement selectedOe = FindDeepestSelectedOe(selection);
            if (selectedOe == null) return false;

            string objectId = (string)selectedOe.Attribute("objectID");
            string fullXml;
            _app.GetPageContent(pageId, out fullXml, PageInfo.piAll, XMLSchema.xs2013);
            XDocument full;
            try { full = XDocument.Parse(fullXml); }
            catch { return false; }
            XElement fullOe = full.Descendants(OneNs + "OE").FirstOrDefault(delegate(XElement oe)
            {
                return string.Equals((string)oe.Attribute("objectID"), objectId, StringComparison.Ordinal);
            });
            XElement outline = fullOe == null ? null : fullOe.Ancestors(OneNs + "Outline").FirstOrDefault();
            if (outline == null) return false;

            XElement sourceMeta = outline.Descendants(OneNs + "Meta").FirstOrDefault(delegate(XElement meta)
            {
                return string.Equals((string)meta.Attribute("name"), "md-preview-source-key", StringComparison.Ordinal);
            });
            string encoded = sourceMeta == null ? null : (string)sourceMeta.Attribute("content");
            if (string.IsNullOrWhiteSpace(encoded)) return false;

            string sourceKey;
            try { sourceKey = Encoding.UTF8.GetString(Convert.FromBase64String(encoded)); }
            catch { return false; }

            string sourceObjectId = string.Empty;
            if (sourceKey.StartsWith("outline:", StringComparison.Ordinal))
            {
                sourceObjectId = sourceKey.Substring("outline:".Length);
            }
            else if (sourceKey.StartsWith("selection:", StringComparison.Ordinal))
            {
                string[] ids = sourceKey.Substring("selection:".Length).Split('|');
                if (ids.Length > 0) sourceObjectId = ids[0];
            }

            _app.NavigateTo(pageId, sourceObjectId, false);
            return true;
        }

        private static string HtmlToPlainText(string html, bool preserveBreaks)
        {
            if (string.IsNullOrWhiteSpace(html)) return string.Empty;
            string value = html;
            if (preserveBreaks)
            {
                value = BreakRegex.Replace(value, "\n");
                value = BlockCloseRegex.Replace(value, "\n");
            }
            value = TagRegex.Replace(value, string.Empty);
            value = WebUtility.HtmlDecode(value ?? string.Empty);
            value = value.Replace("\r\n", "\n").Replace('\r', '\n');
            value = value.Replace('\u00a0', ' ');
            return preserveBreaks ? value.Trim('\r', '\n') : value.Trim();
        }
    }
}
