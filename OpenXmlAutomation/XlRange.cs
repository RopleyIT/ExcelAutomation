namespace OpenXmlAutomation;

/// <summary>
/// Represents a range of cells
/// </summary>
public class XlRange
{
    private readonly string range;
    private readonly XlSheet sheet;
    public List<List<XlCell>> Cells { get; private set; }

    public int Width => Cells.Count;

    public int Height => Cells.Count > 0 ? Cells[0].Count : 0;

    public XlRange(XlSheet s, string tlbr)
    {
        sheet = s;
        range = tlbr;
        XlCellRef cellRange = new(sheet, range);
        int topIdx = XlCellRef.Index(cellRange.Row);
        int botIdx = XlCellRef.Index(cellRange.LastRow);
        int leftIdx = XlCellRef.Index(cellRange.Column);
        int rtIdx = XlCellRef.Index(cellRange.LastColumn);

        // Create the two dimensional array of cells as a list of lists
        // with the correct capacity for the number of rows and columns
        // to prevent reallocation and copying on list growth.

        int rowLength = rtIdx - leftIdx + 1;
        int colDepth = botIdx - topIdx + 1;
        Cells = new(rowLength);
        for (int colIdx = 0; colIdx < rowLength; colIdx++)
        {
            Cells.Add(new(colDepth));
            for (int rowIdx = 0; rowIdx < colDepth; rowIdx++)
            {
                XlCell cell = sheet.FindCell(XlCellRef
                    .ToColName(leftIdx + colIdx) + (topIdx + rowIdx + 1));
                Cells.Last().Add(cell);
            }
        }
    }
}
