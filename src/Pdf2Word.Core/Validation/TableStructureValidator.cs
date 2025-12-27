using Pdf2Word.Core.Models.Ir;

namespace Pdf2Word.Core.Validation;

public sealed class TableStructureValidationResult
{
    public bool IsValid { get; set; }
    public string? ErrorCode { get; set; }
    public string? Message { get; set; }
}

public static class TableStructureValidator
{
    public static TableStructureValidationResult Validate(TableBlockIr table)
    {
        if (table.NCols < 1 || table.Rows.Count < 1)
        {
            return new TableStructureValidationResult { IsValid = false, ErrorCode = "E_TABLE_GRID_INSUFFICIENT_LINES", Message = "表格行列数不足。" };
        }

        var rowCount = table.Rows.Count;
        var owner = new (int r, int c)?[rowCount, table.NCols];
        for (var r = 0; r < rowCount; r++)
        {
            var colCursor = 0;
            foreach (var cell in table.Rows[r].Cells)
            {
                var rowspan = Math.Max(1, cell.Rowspan);
                var colspan = Math.Max(1, cell.Colspan);
                while (colCursor < table.NCols && owner[r, colCursor].HasValue)
                {
                    colCursor++;
                }

                if (colCursor >= table.NCols)
                {
                    return new TableStructureValidationResult { IsValid = false, ErrorCode = "E_TABLE_GRID_CONFLICT", Message = "表格列跨度超出范围。" };
                }

                if (r + rowspan > rowCount || colCursor + colspan > table.NCols)
                {
                    return new TableStructureValidationResult { IsValid = false, ErrorCode = "E_TABLE_GRID_CONFLICT", Message = "表格合并单元格越界。" };
                }

                for (var rr = r; rr < r + rowspan; rr++)
                {
                    for (var cc = colCursor; cc < colCursor + colspan; cc++)
                    {
                        if (owner[rr, cc].HasValue)
                        {
                            return new TableStructureValidationResult { IsValid = false, ErrorCode = "E_TABLE_GRID_CONFLICT", Message = "表格合并单元格冲突。" };
                        }

                        owner[rr, cc] = (r, colCursor);
                    }
                }

                colCursor += colspan;
            }
        }

        for (var r = 0; r < rowCount; r++)
        {
            for (var c = 0; c < table.NCols; c++)
            {
                if (!owner[r, c].HasValue)
                {
                    return new TableStructureValidationResult { IsValid = false, ErrorCode = "E_TABLE_GRID_CONFLICT", Message = "表格存在未覆盖的单元格区域。" };
                }
            }
        }

        return new TableStructureValidationResult { IsValid = true };
    }
}
