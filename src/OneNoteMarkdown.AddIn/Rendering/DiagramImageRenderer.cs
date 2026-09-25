using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace OneNoteMarkdown.Rendering
{
    internal sealed class DiagramImageRenderer
    {
        private const int MaxCacheEntries = 64;
        private const int MaxNodes = 100;
        private const int NodeWidth = 170;
        private const int NodeHeight = 52;
        private const int HorizontalGap = 45;
        private const int VerticalGap = 70;
        private static readonly Regex EdgeRegex = new Regex(
            "^\\s*([A-Za-z0-9_-]+)(?:\\[([^\\]]+)\\]|\\(([^)]+)\\)|\\{([^}]+)\\})?\\s*(?:--+>|==+>)\\s*(?:\\|([^|]+)\\|\\s*)?([A-Za-z0-9_-]+)(?:\\[([^\\]]+)\\]|\\(([^)]+)\\)|\\{([^}]+)\\})?\\s*$",
            RegexOptions.Compiled);
        private static readonly Regex NodeRegex = new Regex(
            "^\\s*([A-Za-z0-9_-]+)(?:\\[([^\\]]+)\\]|\\(([^)]+)\\)|\\{([^}]+)\\})\\s*$",
            RegexOptions.Compiled);
        private static readonly object CacheGate = new object();
        private static readonly Dictionary<string, CacheEntry> Cache = new Dictionary<string, CacheEntry>(StringComparer.Ordinal);
        private static readonly Queue<string> CacheOrder = new Queue<string>();

        public bool TryRenderToPng(
            string diagramType,
            string source,
            int timeoutMilliseconds,
            out byte[] pngBytes,
            out int pixelWidth,
            out int pixelHeight,
            out string error)
        {
            return TryRender(
                diagramType,
                source,
                timeoutMilliseconds,
                false,
                out pngBytes,
                out pixelWidth,
                out pixelHeight,
                out error);
        }

        public bool TryRenderToEmf(
            string diagramType,
            string source,
            int timeoutMilliseconds,
            out byte[] emfBytes,
            out int pixelWidth,
            out int pixelHeight,
            out string error)
        {
            return TryRender(
                diagramType,
                source,
                timeoutMilliseconds,
                true,
                out emfBytes,
                out pixelWidth,
                out pixelHeight,
                out error);
        }

        private bool TryRender(
            string diagramType,
            string source,
            int timeoutMilliseconds,
            bool emf,
            out byte[] imageBytes,
            out int pixelWidth,
            out int pixelHeight,
            out string error)
        {
            imageBytes = null;
            pixelWidth = 0;
            pixelHeight = 0;
            error = string.Empty;

            string type = (diagramType ?? string.Empty).Trim().ToLowerInvariant();
            string body = (source ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Trim();
            if (!string.Equals(type, "mermaid", StringComparison.OrdinalIgnoreCase))
            {
                error = "Only Mermaid flowcharts are rendered as images.";
                return false;
            }
            if (body.Length == 0)
            {
                error = "The Mermaid source is empty.";
                return false;
            }
            if (body.Length > 100000)
            {
                error = "The Mermaid source exceeds the 100 KB limit.";
                return false;
            }

            string cacheKey = ComputeKey((emf ? "emf" : "png") + "\n" + type + "\n" + body);
            CacheEntry cached;
            lock (CacheGate)
            {
                if (Cache.TryGetValue(cacheKey, out cached))
                {
                    imageBytes = (byte[])cached.Bytes.Clone();
                    pixelWidth = cached.Width;
                    pixelHeight = cached.Height;
                    return true;
                }
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            List<Node> nodes;
            List<Edge> edges;
            if (!TryParseFlowchart(body, out nodes, out edges, out error)) return false;
            if (stopwatch.ElapsedMilliseconds > timeoutMilliseconds)
            {
                error = "Mermaid rendering timed out during parsing.";
                return false;
            }

            AssignLevels(nodes, edges);
            int maxLevel = nodes.Count == 0 ? 0 : nodes.Max(delegate(Node node) { return node.Level; });
            int widestLevel = 1;
            for (int level = 0; level <= maxLevel; level++)
            {
                widestLevel = Math.Max(widestLevel, nodes.Count(delegate(Node node) { return node.Level == level; }));
            }
            int width = Math.Min(2400, Math.Max(320, 40 + widestLevel * NodeWidth + (widestLevel - 1) * HorizontalGap));
            int height = Math.Min(2400, Math.Max(160, 40 + (maxLevel + 1) * NodeHeight + maxLevel * VerticalGap));
            PositionNodes(nodes, width);

            if (emf)
            {
                if (!EmfImageRenderer.TryCreate(
                    width,
                    height,
                    delegate(Graphics graphics) { DrawDiagram(graphics, nodes, edges); },
                    out imageBytes,
                    out error))
                {
                    return false;
                }
                pixelWidth = width;
                pixelHeight = height;
            }
            else
            {
                using (Bitmap bitmap = new Bitmap(width, height))
                using (Graphics graphics = Graphics.FromImage(bitmap))
                using (MemoryStream output = new MemoryStream())
                {
                    DrawDiagram(graphics, nodes, edges);
                    bitmap.Save(output, ImageFormat.Png);
                    imageBytes = output.ToArray();
                    pixelWidth = width;
                    pixelHeight = height;
                }
            }

            if (stopwatch.ElapsedMilliseconds > timeoutMilliseconds)
            {
                imageBytes = null;
                error = "Mermaid rendering timed out.";
                return false;
            }

            lock (CacheGate)
            {
                Cache[cacheKey] = new CacheEntry(imageBytes, pixelWidth, pixelHeight);
                CacheOrder.Enqueue(cacheKey);
                while (CacheOrder.Count > MaxCacheEntries)
                {
                    Cache.Remove(CacheOrder.Dequeue());
                }
            }
            return true;
        }

        private static void DrawDiagram(Graphics graphics, List<Node> nodes, List<Edge> edges)
        {
            using (Font font = new Font("Segoe UI", 10f, FontStyle.Regular, GraphicsUnit.Point))
            using (Pen nodePen = new Pen(Color.FromArgb(91, 33, 182), 1.6f))
            using (Pen edgePen = new Pen(Color.FromArgb(75, 85, 99), 1.5f))
            using (Brush nodeBrush = new SolidBrush(Color.FromArgb(245, 243, 255)))
            using (Brush textBrush = new SolidBrush(Color.FromArgb(31, 41, 55)))
            using (Brush edgeTextBrush = new SolidBrush(Color.FromArgb(75, 85, 99)))
            {
                graphics.Clear(Color.White);
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                edgePen.CustomEndCap = new AdjustableArrowCap(4, 6);
                for (int i = 0; i < edges.Count; i++)
                {
                    DrawEdge(graphics, edgePen, edgeTextBrush, font, edges[i]);
                }
                for (int i = 0; i < nodes.Count; i++)
                {
                    DrawNode(graphics, nodePen, nodeBrush, textBrush, font, nodes[i]);
                }
            }
        }

        internal static void ClearCache()
        {
            lock (CacheGate)
            {
                Cache.Clear();
                CacheOrder.Clear();
            }
        }

        private static bool TryParseFlowchart(string source, out List<Node> nodes, out List<Edge> edges, out string error)
        {
            nodes = new List<Node>();
            edges = new List<Edge>();
            error = string.Empty;
            Dictionary<string, Node> byId = new Dictionary<string, Node>(StringComparer.Ordinal);
            string[] lines = source.Split('\n');
            int start = 0;
            if (lines.Length > 0)
            {
                string first = lines[0].Trim();
                if (first.StartsWith("flowchart ", StringComparison.OrdinalIgnoreCase) ||
                    first.StartsWith("graph ", StringComparison.OrdinalIgnoreCase))
                {
                    start = 1;
                }
            }

            for (int i = start; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.Length == 0 || line.StartsWith("%%", StringComparison.Ordinal)) continue;
                Match edgeMatch = EdgeRegex.Match(line);
                if (edgeMatch.Success)
                {
                    Node from = GetOrAddNode(byId, nodes, edgeMatch.Groups[1].Value,
                        FirstValue(edgeMatch.Groups[2], edgeMatch.Groups[3], edgeMatch.Groups[4]),
                        ShapeFromGroups(edgeMatch.Groups[3], edgeMatch.Groups[4]));
                    Node to = GetOrAddNode(byId, nodes, edgeMatch.Groups[6].Value,
                        FirstValue(edgeMatch.Groups[7], edgeMatch.Groups[8], edgeMatch.Groups[9]),
                        ShapeFromGroups(edgeMatch.Groups[8], edgeMatch.Groups[9]));
                    edges.Add(new Edge { From = from, To = to, Label = edgeMatch.Groups[5].Value.Trim() });
                    continue;
                }

                Match nodeMatch = NodeRegex.Match(line);
                if (nodeMatch.Success)
                {
                    GetOrAddNode(byId, nodes, nodeMatch.Groups[1].Value,
                        FirstValue(nodeMatch.Groups[2], nodeMatch.Groups[3], nodeMatch.Groups[4]),
                        ShapeFromGroups(nodeMatch.Groups[3], nodeMatch.Groups[4]));
                    continue;
                }

                error = "Unsupported Mermaid statement at line " + (i + 1) + ": " + line;
                return false;
            }

            if (nodes.Count == 0)
            {
                error = "No flowchart nodes were found.";
                return false;
            }
            if (nodes.Count > MaxNodes)
            {
                error = "The Mermaid diagram exceeds the 100-node limit.";
                return false;
            }
            return true;
        }

        private static Node GetOrAddNode(
            Dictionary<string, Node> byId,
            List<Node> nodes,
            string id,
            string label,
            NodeShape shape)
        {
            Node node;
            if (!byId.TryGetValue(id, out node))
            {
                node = new Node { Id = id, Label = string.IsNullOrWhiteSpace(label) ? id : label.Trim(), Shape = shape };
                byId[id] = node;
                nodes.Add(node);
            }
            else if (!string.IsNullOrWhiteSpace(label))
            {
                node.Label = label.Trim();
                node.Shape = shape;
            }
            return node;
        }

        private static string FirstValue(params Group[] groups)
        {
            for (int i = 0; i < groups.Length; i++)
            {
                if (groups[i].Success && groups[i].Value.Length > 0) return groups[i].Value;
            }
            return string.Empty;
        }

        private static NodeShape ShapeFromGroups(Group rounded, Group decision)
        {
            if (decision.Success) return NodeShape.Decision;
            if (rounded.Success) return NodeShape.Rounded;
            return NodeShape.Rectangle;
        }

        private static void AssignLevels(List<Node> nodes, List<Edge> edges)
        {
            Dictionary<Node, int> incoming = nodes.ToDictionary(delegate(Node node) { return node; }, delegate(Node node) { return 0; });
            for (int i = 0; i < edges.Count; i++) incoming[edges[i].To]++;
            Queue<Node> queue = new Queue<Node>(nodes.Where(delegate(Node node) { return incoming[node] == 0; }));
            HashSet<Node> visited = new HashSet<Node>();
            while (queue.Count > 0)
            {
                Node current = queue.Dequeue();
                if (!visited.Add(current)) continue;
                foreach (Edge edge in edges.Where(delegate(Edge candidate) { return ReferenceEquals(candidate.From, current); }))
                {
                    edge.To.Level = Math.Max(edge.To.Level, current.Level + 1);
                    incoming[edge.To]--;
                    if (incoming[edge.To] <= 0) queue.Enqueue(edge.To);
                }
            }
            int fallbackLevel = visited.Count == 0 ? 0 : visited.Max(delegate(Node node) { return node.Level; }) + 1;
            foreach (Node node in nodes)
            {
                if (!visited.Contains(node)) node.Level = fallbackLevel++;
            }
        }

        private static void PositionNodes(List<Node> nodes, int canvasWidth)
        {
            int maxLevel = nodes.Max(delegate(Node node) { return node.Level; });
            for (int level = 0; level <= maxLevel; level++)
            {
                List<Node> row = nodes.Where(delegate(Node node) { return node.Level == level; }).ToList();
                int rowWidth = row.Count * NodeWidth + Math.Max(0, row.Count - 1) * HorizontalGap;
                int x = Math.Max(20, (canvasWidth - rowWidth) / 2);
                for (int i = 0; i < row.Count; i++)
                {
                    row[i].Bounds = new RectangleF(x, 20 + level * (NodeHeight + VerticalGap), NodeWidth, NodeHeight);
                    x += NodeWidth + HorizontalGap;
                }
            }
        }

        private static void DrawEdge(Graphics graphics, Pen pen, Brush textBrush, Font font, Edge edge)
        {
            PointF start = new PointF(edge.From.Bounds.Left + edge.From.Bounds.Width / 2f, edge.From.Bounds.Bottom);
            PointF end = new PointF(edge.To.Bounds.Left + edge.To.Bounds.Width / 2f, edge.To.Bounds.Top);
            float middleY = (start.Y + end.Y) / 2f;
            PointF[] points =
            {
                start,
                new PointF(start.X, middleY),
                new PointF(end.X, middleY),
                end
            };
            graphics.DrawLines(pen, points);
            if (!string.IsNullOrWhiteSpace(edge.Label))
            {
                graphics.DrawString(edge.Label, font, textBrush, new PointF((start.X + end.X) / 2f + 4f, middleY - 18f));
            }
        }

        private static void DrawNode(Graphics graphics, Pen pen, Brush fill, Brush text, Font font, Node node)
        {
            if (node.Shape == NodeShape.Decision)
            {
                PointF[] diamond =
                {
                    new PointF(node.Bounds.Left + node.Bounds.Width / 2f, node.Bounds.Top),
                    new PointF(node.Bounds.Right, node.Bounds.Top + node.Bounds.Height / 2f),
                    new PointF(node.Bounds.Left + node.Bounds.Width / 2f, node.Bounds.Bottom),
                    new PointF(node.Bounds.Left, node.Bounds.Top + node.Bounds.Height / 2f)
                };
                graphics.FillPolygon(fill, diamond);
                graphics.DrawPolygon(pen, diamond);
            }
            else if (node.Shape == NodeShape.Rounded)
            {
                using (GraphicsPath path = RoundedRectangle(node.Bounds, 12f))
                {
                    graphics.FillPath(fill, path);
                    graphics.DrawPath(pen, path);
                }
            }
            else
            {
                graphics.FillRectangle(fill, node.Bounds);
                graphics.DrawRectangle(pen, node.Bounds.X, node.Bounds.Y, node.Bounds.Width, node.Bounds.Height);
            }

            using (StringFormat format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            {
                graphics.DrawString(node.Label, font, text, node.Bounds, format);
            }
        }

        private static GraphicsPath RoundedRectangle(RectangleF bounds, float radius)
        {
            float diameter = radius * 2f;
            GraphicsPath path = new GraphicsPath();
            path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180f, 90f);
            path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270f, 90f);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0f, 90f);
            path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90f, 90f);
            path.CloseFigure();
            return path;
        }

        private static string ComputeKey(string value)
        {
            using (SHA256 sha = SHA256.Create())
            {
                return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty)));
            }
        }

        private sealed class Node
        {
            public string Id;
            public string Label;
            public NodeShape Shape;
            public int Level;
            public RectangleF Bounds;
        }

        private sealed class Edge
        {
            public Node From;
            public Node To;
            public string Label;
        }

        private sealed class CacheEntry
        {
            public CacheEntry(byte[] bytes, int width, int height)
            {
                Bytes = (byte[])bytes.Clone();
                Width = width;
                Height = height;
            }

            public byte[] Bytes;
            public int Width;
            public int Height;
        }

        private enum NodeShape
        {
            Rectangle,
            Rounded,
            Decision
        }
    }
}
