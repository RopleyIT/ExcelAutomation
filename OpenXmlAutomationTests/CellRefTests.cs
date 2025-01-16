using OpenXmlAutomation;
namespace OpenXmlAutomationTests;

[TestClass]
public class CellRefTests
{
    [DataTestMethod]
    [DataRow("Sheet one!$A$3:$C$7", "Sheet one", "A", 3, "C", 7)]
    [DataRow("A1", null, "A", 1, "A", 1)]
    [DataRow("FF333", null, "FF", 333, "FF", 333)]
    [DataRow("ZZ999", null, "ZZ", 999, "ZZ", 999)]
    [DataRow("Sheet 1!A1", "Sheet 1", "A", 1, "A", 1)]
    [DataRow("Sheet 1!FF333", "Sheet 1", "FF", 333, "FF", 333)]
    [DataRow("Sheet 1!ZZ999", "Sheet 1", "ZZ", 999, "ZZ", 999)]
    public void SimpleCellRefs(string cellRef, string? sheet,
        string col, int row, string lastCol, int lastRow)
    {
        XlCellRef cr = new(cellRef);
        Assert.AreEqual(sheet, cr.SheetName);
        Assert.AreEqual(row, cr.Row);
        Assert.AreEqual(col, cr.Column);
        Assert.AreEqual(lastRow, cr.LastRow);
        Assert.AreEqual(lastCol, cr.LastColumn);
    }
}
