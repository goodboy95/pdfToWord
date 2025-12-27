using Pdf2Word.Core.Models;
using Pdf2Word.Core.Models.Ir;

namespace Pdf2Word.Core.Services;

public static class TableIrBuilder
{
    public static TableBlockIr Build(TableDetection detection, IReadOnlyDictionary<string, string>? texts = null)
    {
        var rows = new List<TableRowIr>();
        for (var i = 0; i < detection.NRows; i++)
        {
            rows.Add(new TableRowIr());
        }

        foreach (var cell in detection.Cells)
        {
            var text = string.Empty;
            if (texts != null && texts.TryGetValue(cell.Id, out var value))
            {
                text = value ?? string.Empty;
            }

            if (cell.Row < 0 || cell.Row >= rows.Count)
            {
                continue;
            }

            rows[cell.Row].Cells.Add(new TableCellIr
            {
                Text = text,
                Rowspan = Math.Max(1, cell.Rowspan),
                Colspan = Math.Max(1, cell.Colspan),
                CellBBoxInTableImage = cell.BBoxInTable,
                CellId = cell.Id
            });
        }

        foreach (var row in rows)
        {
            row.Cells = row.Cells.OrderBy(c => c.CellBBoxInTableImage?.X ?? 0).ToList();
        }

        return new TableBlockIr
        {
            TableBBox = detection.TableBBoxInPage,
            NCols = detection.NCols,
            Rows = rows,
            StructureMeta = new TableStructureMeta
            {
                Engine = "OpenCV",
                DetectedCellCount = detection.Cells.Count
            }
        };
    }
}
