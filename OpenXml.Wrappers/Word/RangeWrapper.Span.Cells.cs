using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class RangeWrapper
{
    /// <summary>
    /// How a range inside a table relates to the edges of cells, which Word treats specially: a range that crosses an edge cannot be deleted
    /// unless it takes in whole cells (then what is in the cells is cleared and the cells stay), and replacing the text of one replaces only
    /// what is in the first cell.
    /// </summary>
    private sealed class CellCrossing
    {
        public bool Crosses { get; private set; }

        public bool WholeCells { get; private set; }

        public int FirstEdge { get; private set; }

        /// <summary>The content of each cell in the range: the places between the edges.</summary>
        public List<(int Start, int End)> Parts { get; } = new List<(int Start, int End)>();

        public static CellCrossing Of(StoryMap map, int start, int end)
        {
            CellCrossing result = new CellCrossing();
            List<Atom> inRange = map.AtomsIn(start, end).ToList();
            if (inRange.Count == 0)
            {
                return result;
            }

            Paragraph first = ParagraphOf(inRange[0]);
            Table table = first?.Ancestors<Table>().FirstOrDefault();
            Atom last = inRange[inRange.Count - 1];
            if (table == null || !IsInside(last, table))
            {
                //A range that starts outside a table, or runs on past it, is not a range within a table.
                return result;
            }

            List<Atom> edges = inRange.Where(IsCellEdge).ToList();
            result.Crosses = edges.Count > 1 || (edges.Count == 1 && !ReferenceEquals(edges[0], last));
            if (!result.Crosses)
            {
                return result;
            }

            result.FirstEdge = edges[0].Start;
            int partStart = start;
            foreach (Atom edge in edges)
            {
                if (edge.Start > partStart)
                {
                    result.Parts.Add((partStart, edge.Start));
                }
                partStart = edge.End;
            }
            if (end > partStart)
            {
                result.Parts.Add((partStart, end));
            }

            result.WholeCells = IsCellEdge(last) && first.Parent is TableCell cell && ReferenceEquals(cell.Elements<Paragraph>().FirstOrDefault(), first)
                && FirstPositionOf(map, first) == start;
            return result;
        }

        private static bool IsInside(Atom atom, Table table)
        {
            OpenXmlElement element = atom.Element is TableRow row ? row.Parent : atom.Element;
            return element != null && (ReferenceEquals(element, table) || element.Ancestors().Any(a => ReferenceEquals(a, table)));
        }

        private static bool IsCellEdge(Atom atom)
        {
            return atom.Kind == AtomKind.RowMark || (atom.Kind == AtomKind.ParagraphMark && IsCellEnd((Paragraph)atom.Element));
        }
    }
}
