using System;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using OneNoteMarkdown.Markdown;
using OneNoteMarkdown.OneNote;

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

            XElement imageOe = XElement.Parse(
                "<one:OE xmlns:one=\"" + OneNs + "\"><one:Image/></one:OE>");
            Assert(!(bool)safetyMethod.Invoke(null, new object[] { imageOe }),
                "An image OE must never be replaced in place.");
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
    }
}
