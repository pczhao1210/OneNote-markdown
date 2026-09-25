using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Net;
using System.Text;
using System.Xml.Linq;
using OneNoteMarkdown.Markdown;
using OneNoteMarkdown.OneNote;
using OneNoteMarkdown.OneNote.Models;
using OneNoteMarkdown.Rendering;

namespace OneNoteMarkdown.Tests
{
    internal static class Program
    {
        private const string OneNs = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        private static int _failures;

        private static int Main()
        {
            Run("Markdown renderer", TestMarkdownRenderer);
            Run("HTML line breaks", TestHtmlLineBreaks);
            Run("Markdown source export", TestMarkdownSourceExport);
            Run("Enter-render Markdown detection", TestEnterRenderMarkdownDetection);
            Run("Nested OneNote selection safety", TestNestedOneNoteSelectionSafety);
            Run("Markdown parser boundaries", TestMarkdownParserBoundaries);
            Run("Table and image parsing", TestTableAndImageParsing);
            Run("Import path normalization", TestImportPathNormalization);
            Run("Imported source metadata", TestImportedSourceMetadata);
            Run("Nested source structure", TestNestedSourceStructure);
            Run("Inline escaping", TestInlineEscaping);
            Run("Inline LaTeX rendering", TestInlineLatexRendering);
            Run("Remote image request", TestRemoteImageRequest);
            Run("DPI-aware settings", TestDpiAwareSettings);
            Run("Managed preview reuse", TestManagedPreviewReuse);
            Run("Hidden UI anchor", TestHiddenUiAnchor);
            Run("Managed preview export", TestManagedPreviewExport);
            Run("Offline Mermaid rendering", TestOfflineMermaidRendering);

            Console.WriteLine(_failures == 0
                ? "All regression tests passed."
                : _failures + " regression test(s) failed.");
            return _failures == 0 ? 0 : 1;
        }

        private static void TestMarkdownRenderer()
        {
            var blocks = MarkdownRenderer.RenderToBlocks("# Title\n- [x] done\n```csharp\nvar x = 1;\n```");
            Assert(blocks.Count == 3, "Expected three blocks.");
            Assert(blocks[0].Kind == MarkdownBlockKind.Heading && blocks[0].Level == 1, "Heading was not parsed.");
            Assert(blocks[1].Kind == MarkdownBlockKind.ListItem && blocks[1].IsTaskChecked, "Task item was not parsed.");
            Assert(blocks[2].Kind == MarkdownBlockKind.CodeBlock && blocks[2].CodeLanguage == "csharp", "Code fence was not parsed.");
        }

        private static void TestHtmlLineBreaks()
        {
            string xml = PageXml(
                "<one:OE objectID=\"oe1\"><one:T><![CDATA[first<br>second]]></one:T></one:OE>");
            var page = PageParser.Parse(xml);
            Assert(page.Outlines[0].TextBlocks[0].Text == "first\nsecond", "HTML break was lost.");
        }

        private static void TestMarkdownSourceExport()
        {
            string source = "```csharp\nvar x = 1;\n```";
            string encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(source));
            string xml = PageXml(
                "<one:OE objectID=\"oe1\">" +
                "<one:Meta name=\"md-src\" content=\"" + encoded + "\"/>" +
                "<one:T><![CDATA[var x = 1;]]></one:T></one:OE>" +
                "<one:OE objectID=\"oe2\">" +
                "<one:Meta name=\"md-continuation\" content=\"true\"/>" +
                "<one:T><![CDATA[duplicate continuation]]></one:T></one:OE>");

            var page = PageParser.Parse(xml);
            string markdown = MarkdownExporter.Export(page);
            Assert(markdown.Contains(source), "Stored Markdown source was not restored.");
            Assert(!markdown.Contains("duplicate continuation"), "Continuation text was exported twice.");
        }

        private static void TestEnterRenderMarkdownDetection()
        {
            MethodInfo method = typeof(OneNoteProvider).GetMethod(
                "LooksLikeMarkdown",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert(method != null, "Markdown detection method was not found.");

            Func<string, bool> looksLikeMarkdown = value =>
                (bool)method.Invoke(null, new object[] { value });

            Assert(looksLikeMarkdown("欧拉公式 $e^{i\\pi}+1=0$"), "Complete inline LaTeX was not detected.");
            Assert(looksLikeMarkdown("$$E=mc^2$$"), "Complete one-line block LaTeX was not detected.");
            Assert(!looksLikeMarkdown("$$"), "An incomplete block LaTeX marker must not render on Enter.");
            Assert(!looksLikeMarkdown("```csharp"), "An incomplete code fence must not render on Enter.");
            Assert(!looksLikeMarkdown("x+y=z"), "Plain or native equation text must not be rewritten.");
        }

        private static void TestNestedOneNoteSelectionSafety()
        {
            XNamespace one = OneNs;
            XDocument document = XDocument.Parse(
                "<one:Page xmlns:one=\"" + OneNs + "\" ID=\"page1\">" +
                "<one:Outline><one:OEChildren>" +
                "<one:OE objectID=\"parent\" selected=\"partial\">" +
                "<one:T><![CDATA[parent summary]]></one:T>" +
                "<one:OEChildren><one:OE objectID=\"child\">" +
                "<one:T selected=\"partial\"><![CDATA[# Child heading]]></one:T>" +
                "</one:OE></one:OEChildren></one:OE>" +
                "</one:OEChildren></one:Outline></one:Page>");

            MethodInfo findMethod = typeof(OneNoteProvider).GetMethod(
                "FindDeepestSelectedOe",
                BindingFlags.Static | BindingFlags.NonPublic);
            MethodInfo extractMethod = typeof(OneNoteProvider).GetMethod(
                "ExtractOeMarkdown",
                BindingFlags.Static | BindingFlags.NonPublic);
            MethodInfo safetyMethod = typeof(PageWriter).GetMethod(
                "IsSafeTextOeForInPlaceRender",
                BindingFlags.Static | BindingFlags.NonPublic);

            Assert(findMethod != null && extractMethod != null && safetyMethod != null,
                "OneNote selection safety methods were not found.");

            XElement selected = (XElement)findMethod.Invoke(null, new object[] { document });
            Assert((string)selected.Attribute("objectID") == "child",
                "The deepest selected OE was not chosen.");
            Assert((string)extractMethod.Invoke(null, new object[] { selected }) == "# Child heading",
                "Nested content leaked into the selected OE source.");

            XElement parent = document.Descendants(one + "OE")
                .First(element => (string)element.Attribute("objectID") == "parent");
            Assert(!(bool)safetyMethod.Invoke(null, new object[] { parent }),
                "A structural parent OE must never be replaced in place.");
            Assert((bool)safetyMethod.Invoke(null, new object[] { selected }),
                "A leaf text OE should remain renderable.");

            MethodInfo selectionMethod = typeof(OneNoteProvider).GetMethod(
                "IsInsideSelectedSubtree",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert(selectionMethod != null, "Selection boundary method was not found.");
            XElement unselectedSiblingText = new XElement(one + "T", "not selected");
            parent.Add(unselectedSiblingText);
            Assert(!(bool)selectionMethod.Invoke(null, new object[] { unselectedSiblingText }),
                "A partially selected parent leaked an unselected text node.");

            XElement imageOe = XElement.Parse(
                "<one:OE xmlns:one=\"" + OneNs + "\"><one:Image/></one:OE>");
            Assert(!(bool)safetyMethod.Invoke(null, new object[] { imageOe }),
                "An image OE must never be replaced in place.");
        }

        private static void TestMarkdownParserBoundaries()
        {
            var blocks = MarkdownRenderer.RenderToBlocks(
                "* * *\n" +
                "$$E=mc^2$$\n" +
                "````csharp\n" +
                "```\n" +
                "var x = 1;\n" +
                "````");

            Assert(blocks.Count == 3, "Expected horizontal rule, formula, and code block.");
            Assert(blocks[0].Kind == MarkdownBlockKind.HorizontalRule,
                "A spaced horizontal rule was parsed as a list item.");
            Assert(blocks[1].Kind == MarkdownBlockKind.LatexBlock && blocks[1].Text == "E=mc^2",
                "One-line block LaTeX was not parsed.");
            Assert(blocks[2].Kind == MarkdownBlockKind.CodeBlock &&
                blocks[2].Text.Contains("```") &&
                blocks[2].Text.Contains("var x = 1;"),
                "A shorter code fence incorrectly closed a longer fence.");

            var quotes = MarkdownRenderer.RenderToBlocks(
                "> outer\n" +
                "> > inner\n" +
                "> > - nested item");
            Assert(quotes.Count == 3 &&
                quotes[0].Kind == MarkdownBlockKind.Blockquote && quotes[0].Level == 1 &&
                quotes[1].Level == 2 && quotes[1].Text == "inner" &&
                quotes[2].Level == 2 && quotes[2].Text == "- nested item",
                "Nested blockquote depth was not retained.");

            var incomplete = MarkdownRenderer.RenderToBlocks("```csharp\nvar value = 1;");
            Assert(incomplete.Count == 1 &&
                incomplete[0].Kind == MarkdownBlockKind.Paragraph &&
                incomplete[0].Text.StartsWith("```csharp", StringComparison.Ordinal),
                "An incomplete fence did not preserve its source.");
        }

        private static void TestTableAndImageParsing()
        {
            var blocks = MarkdownRenderer.RenderToBlocks(
                "| Name | Value |\n" +
                "| --- | ---: |\n" +
                "| alpha | 1 |\n\n" +
                "![diagram](images/demo.png \"Preview image\")",
                @"C:\notes");

            Assert(blocks.Count == 3, "Expected table, blank line, and image.");
            Assert(blocks[0].Kind == MarkdownBlockKind.Table &&
                blocks[0].TableRows.Count == 2 &&
                blocks[0].TableRows[1][0] == "alpha",
                "Markdown table rows were not parsed.");
            Assert(blocks[2].Kind == MarkdownBlockKind.Image &&
                string.Equals(blocks[2].Target, @"C:\notes\images\demo.png", StringComparison.OrdinalIgnoreCase),
                "Relative image path was not resolved against the import directory.");
        }

        private static void TestImportPathNormalization()
        {
            string expected = System.IO.Path.GetFullPath(@"C:\notes\test.md");
            string actual = Features.ImportMarkdownCommand.NormalizeSelectedPath("  \"" + expected + "\"  ");
            Assert(string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase),
                "A quoted file-dialog path was not normalized.");
        }

        private static void TestImportedSourceMetadata()
        {
            Type sourceType = typeof(PageWriter).Assembly.GetType(
                "OneNoteMarkdown.OneNote.Models.PreviewSource",
                true);
            object source = Activator.CreateInstance(sourceType, true);
            sourceType.GetProperty("Markdown").SetValue(source, "# Title\n\n  indented &lt;");
            sourceType.GetProperty("BaseDirectory").SetValue(source, @"C:\notes");

            MethodInfo createMethod = typeof(PageWriter).GetMethod(
                "CreateImportedSourceOutline",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert(createMethod != null, "Imported source outline builder was not found.");
            string sourceKey = Convert.ToBase64String(Encoding.UTF8.GetBytes(@"import:C:\NOTES\TEST.MD"));
            XElement outline = (XElement)createMethod.Invoke(
                null,
                new object[] { source, sourceKey, 36d, 200d, 520d, 100d, 0 });

            XNamespace one = OneNs;
            XElement sourceOe = outline.Descendants(one + "OE").Single();
            string storedKey = sourceOe.Elements(one + "Meta")
                .Single(meta => (string)meta.Attribute("name") == "md-import-source-key")
                .Attribute("content").Value;
            string storedBase = sourceOe.Elements(one + "Meta")
                .Single(meta => (string)meta.Attribute("name") == "md-import-base-directory")
                .Attribute("content").Value;
            Assert(storedKey == sourceKey, "Imported source identity was not persisted.");
            Assert(Encoding.UTF8.GetString(Convert.FromBase64String(storedBase)) == @"C:\notes",
                "Imported source base directory was not persisted.");
            string html = sourceOe.Element(one + "T").Value;
            Assert(html.Contains("<br><br>") && html.Contains("&nbsp;&nbsp;indented&nbsp;&amp;lt;"),
                "Imported Markdown was not preserved as editable raw text.");

            MethodInfo readBaseMethod = typeof(OneNoteProvider).GetMethod(
                "ReadSourceBaseDirectory",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert(readBaseMethod != null, "Imported base-directory reader was not found.");
            Assert((string)readBaseMethod.Invoke(null, new object[] { outline }) == @"C:\notes",
                "Imported base directory could not be restored for relative images.");
        }

        private static void TestNestedSourceStructure()
        {
            XDocument document = XDocument.Parse(
                "<one:Outline xmlns:one=\"" + OneNs + "\"><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[- parent]]></one:T><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[- child]]></one:T></one:OE>" +
                "</one:OEChildren></one:OE></one:OEChildren></one:Outline>");
            MethodInfo method = typeof(OneNoteProvider).GetMethod(
                "ExtractOutlineText",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert(method != null, "Outline source extraction method was not found.");
            string source = (string)method.Invoke(null, new object[] { document.Root });
            Assert(source == "- parent\n  - child",
                "Nested OneNote structure was flattened while rebuilding Markdown source.");
        }

        private static void TestInlineEscaping()
        {
            MethodInfo method = typeof(PageWriter).GetMethod(
                "ApplyInlineMarkdownStyles",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert(method != null, "Inline Markdown formatter was not found.");
            string html = (string)method.Invoke(null, new object[] { @"\*literal\* and \[text\]" });
            Assert(html == "*literal* and [text]",
                "Escaped Markdown punctuation was not restored as literal text.");
        }

        private static void TestInlineLatexRendering()
        {
            LatexImageRenderer renderer = new LatexImageRenderer();
            byte[] png;
            int width;
            int height;
            string error;
            bool rendered = renderer.TryRenderInlineTextToPng(
                "Energy $E=mc^2$ and $a^2+b^2=c^2$.",
                "Calibri",
                11d,
                out png,
                out width,
                out height,
                out error);

            Assert(rendered, "Inline LaTeX did not render: " + error);
            Assert(png != null && png.Length > 24 && png[0] == 0x89 && png[1] == 0x50,
                "Inline LaTeX renderer did not return a PNG.");
            Assert(width > 100 && height > 10, "Inline LaTeX image dimensions were invalid.");
        }

        private static void TestRemoteImageRequest()
        {
            MethodInfo method = typeof(PageWriter).GetMethod(
                "CreateRemoteImageRequest",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert(method != null, "Remote image request factory was not found.");
            HttpWebRequest request = (HttpWebRequest)method.Invoke(
                null,
                new object[] { new Uri("https://example.com/image.png") });
            Assert(request.AllowAutoRedirect && request.MaximumAutomaticRedirections == 5,
                "Remote image redirects were not enabled.");
            Assert(request.UserAgent.StartsWith("OneNoteMarkdown/", StringComparison.Ordinal) &&
                request.Accept == "image/*",
                "Remote image request headers were not configured.");
            SecurityProtocolType protocols = ServicePointManager.SecurityProtocol;
            Assert((request.AutomaticDecompression & DecompressionMethods.GZip) != 0 &&
                (protocols == SecurityProtocolType.SystemDefault ||
                 (protocols & SecurityProtocolType.Tls12) != 0),
                "Remote image TLS or compression support was not configured.");
        }

        private static void TestDpiAwareSettings()
        {
            Type settingsType = typeof(PageWriter).Assembly.GetType(
                "OneNoteMarkdown.UI.SettingsDialog",
                true);
            using (System.Windows.Forms.Form dialog =
                (System.Windows.Forms.Form)Activator.CreateInstance(settingsType, true))
            {
                Assert(dialog.AutoScaleMode == System.Windows.Forms.AutoScaleMode.Dpi,
                    "Settings dialog does not use DPI-based scaling.");
                Assert(dialog.AutoScaleDimensions.Width >= 96f &&
                    dialog.AutoScaleDimensions.Height >= 96f,
                    "Settings dialog reported invalid DPI scaling dimensions.");
            }
        }

        private static void TestManagedPreviewReuse()
        {
            XNamespace one = OneNs;
            string oldKey = Convert.ToBase64String(Encoding.UTF8.GetBytes("page:old"));
            string currentKey = Convert.ToBase64String(Encoding.UTF8.GetBytes("page:current"));
            XElement page = XElement.Parse(
                "<one:Page xmlns:one=\"" + OneNs + "\">" +
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:Meta name=\"md-preview-role\" content=\"PagePreview\"/>" +
                "<one:Meta name=\"md-preview-source-key\" content=\"" + oldKey + "\"/>" +
                "<one:T><![CDATA[existing]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline></one:Page>");
            MethodInfo findMethod = typeof(PageWriter).GetMethod(
                "FindManagedOutline",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert(findMethod != null, "Managed preview lookup was not found.");
            XElement reused = (XElement)findMethod.Invoke(
                null,
                new object[] { page, "PagePreview", currentKey });
            Assert(reused != null,
                "A page preview with a stale source key was not reused.");

            XElement outline = new XElement(one + "Outline",
                new XElement(one + "OEChildren",
                    new XElement(one + "OE", new XElement(one + "T", "first")),
                    new XElement(one + "OE", new XElement(one + "T", "second"))));
            MethodInfo markMethod = typeof(PageWriter).GetMethod(
                "MarkManagedPreview",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert(markMethod != null, "Managed preview marker was not found.");
            markMethod.Invoke(
                null,
                new object[] { outline, "preview-id", "PagePreview", currentKey, "source-hash", "# source", @"C:\notes", false });
            Assert(outline.Descendants(one + "OE").All(oe =>
                oe.Elements(one + "Meta").Any(meta =>
                    (string)meta.Attribute("name") == "md-preview-source-key")),
                "Managed preview identity was not replicated across all preview paragraphs.");
            XElement baseMeta = outline.Descendants(one + "Meta").Single(meta =>
                (string)meta.Attribute("name") == "md-preview-base-directory");
            Assert(Encoding.UTF8.GetString(Convert.FromBase64String((string)baseMeta.Attribute("content"))) == @"C:\notes",
                "Managed preview did not retain the relative-image base directory.");
        }

        private static void TestHiddenUiAnchor()
        {
            System.Windows.Forms.IWin32Window anchor = UI.UiThread.Anchor;
            Assert(anchor != null && anchor.Handle != IntPtr.Zero, "UI anchor window was not created.");
            Assert(!IsWindowVisible(anchor.Handle), "UI anchor window must never be visible.");
            const int gwlExStyle = -20;
            const int wsExToolWindow = 0x00000080;
            int exStyle = GetWindowLong(anchor.Handle, gwlExStyle);
            Assert((exStyle & wsExToolWindow) != 0,
                "UI anchor window must be a tool window so it cannot appear in Alt+Tab.");
            UI.UiThread.Shutdown();
        }

        private static void TestManagedPreviewExport()
        {
            string source = "# Imported\n\nOriginal";
            string encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(source));
            string xml =
                "<one:Page xmlns:one=\"" + OneNs + "\" ID=\"page1\" name=\"Test\">" +
                "<one:Outline><one:OEChildren>" +
                "<one:OE><one:T><![CDATA[Markdown Render]]></one:T></one:OE>" +
                "<one:OE><one:T><![CDATA[ordinary user note]]></one:T></one:OE>" +
                "</one:OEChildren></one:Outline>" +
                "<one:Outline><one:OEChildren><one:OE>" +
                "<one:Meta name=\"md-preview-role\" content=\"ImportPreview\"/>" +
                "<one:Meta name=\"md-preview-id\" content=\"preview1\"/>" +
                "<one:Meta name=\"md-preview-source\" content=\"" + encoded + "\"/>" +
                "<one:Meta name=\"md-preview-group\" content=\"preview1\"/>" +
                "<one:T><![CDATA[rendered duplicate]]></one:T>" +
                "</one:OE></one:OEChildren></one:Outline>" +
                "</one:Page>";

            PageContent page = PageParser.Parse(xml);
            string markdown = MarkdownExporter.Export(page);
            Assert(markdown.Contains(source), "Managed import source was not exported.");
            Assert(!markdown.Contains("rendered duplicate"), "Managed preview body was exported as a duplicate.");
            Assert(markdown.Contains("ordinary user note"),
                "A title-only legacy block was incorrectly claimed as a managed preview.");
        }

        private static void TestOfflineMermaidRendering()
        {
            DiagramImageRenderer renderer = new DiagramImageRenderer();
            byte[] png;
            int width;
            int height;
            string error;
            bool rendered = renderer.TryRenderToPng(
                "mermaid",
                "flowchart TD\nA[Start] --> B{Ready?}\nB -->|Yes| C[Done]",
                3000,
                out png,
                out width,
                out height,
                out error);

            Assert(rendered, "Mermaid flowchart did not render: " + error);
            Assert(png != null && png.Length > 24 && png[0] == 0x89 && png[1] == 0x50,
                "Mermaid renderer did not return a PNG.");
            Assert(width >= 320 && height >= 160, "Mermaid image dimensions were invalid.");
        }

        private static string PageXml(string oeXml)
        {
            return "<one:Page xmlns:one=\"" + OneNs + "\" ID=\"page1\" name=\"Test\">" +
                "<one:Outline><one:OEChildren>" + oeXml +
                "</one:OEChildren></one:Outline></one:Page>";
        }

        private static void Run(string name, Action test)
        {
            try
            {
                test();
                Console.WriteLine("PASS " + name);
            }
            catch (Exception ex)
            {
                _failures++;
                Console.WriteLine("FAIL " + name + ": " + ex.Message);
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
        private static extern int GetWindowLong(IntPtr hWnd, int index);
    }
}
