using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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
            Run("Nested source structure", TestNestedSourceStructure);
            Run("Inline escaping", TestInlineEscaping);
            Run("Inline LaTeX rendering", TestInlineLatexRendering);
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
