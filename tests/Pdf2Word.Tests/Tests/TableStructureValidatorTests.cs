using Pdf2Word.Core.Models.Ir;
using Pdf2Word.Core.Validation;
using Xunit;

namespace Pdf2Word.Tests.Tests;

public class TableStructureValidatorTests
{
    [Fact]
    public void ValidTablePasses()
    {
        var table = new TableBlockIr
        {
            NCols = 2,
            Rows =
            {
                new TableRowIr
                {
                    Cells =
                    {
                        new TableCellIr { Text = "A", Rowspan = 2, Colspan = 1 },
                        new TableCellIr { Text = "B", Rowspan = 1, Colspan = 1 }
                    }
                },
                new TableRowIr
                {
                    Cells =
                    {
                        new TableCellIr { Text = "C", Rowspan = 1, Colspan = 1 }
                    }
                }
            }
        };

        var result = TableStructureValidator.Validate(table);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void ConflictTableFails()
    {
        var table = new TableBlockIr
        {
            NCols = 2,
            Rows =
            {
                new TableRowIr
                {
                    Cells =
                    {
                        new TableCellIr { Text = "A", Rowspan = 2, Colspan = 2 }
                    }
                },
                new TableRowIr
                {
                    Cells =
                    {
                        new TableCellIr { Text = "B", Rowspan = 1, Colspan = 2 }
                    }
                }
            }
        };

        var result = TableStructureValidator.Validate(table);
        Assert.False(result.IsValid);
    }
}
