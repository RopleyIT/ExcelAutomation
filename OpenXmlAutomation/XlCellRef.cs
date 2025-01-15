using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OpenXmlAutomation;

/// <summary>
/// Utility class for translating cell references
/// into component values used for finding cells
/// </summary>

public class XlCellRef
{
    /// <summary>
    /// The name of the sheet part of the cell reference if it
    /// exists in the cell reference. Will be a null string
    /// if not present. The sheet reference would be 'Sheet 1'
    /// in the example: "Sheet 1!A3:B5"
    /// </summary>
    public string? SheetName { get; init; }

    /// <summary>
    /// The column of the cell being referenced, or for the
    /// top left corner cell if a cell range being specified
    /// </summary>
    public string Column { get; init; }

    /// <summary>
    /// The row number for the cell being referenced, or the
    /// row number of the top left corner cell if a cell
    /// range is being specified. Note the top row has value
    /// 1 rather than 0
    /// </summary>
    public int Row { get; init; }

    /// <summary>
    /// The column of the bottom right corner cell of a range.
    /// This is set equal to the top left column if a single
    /// cell.
    /// </summary>
    public string LastColumn { get; init; }

    /// <summary>
    /// The row number of the bottom right corner cell for a
    /// range. This is equal to the top left corner row if
    /// a single cell.
    /// </summary>
    public int LastRow { get; init; }

    /// <summary>
    /// Detect if the cell reference is to a single cell
    /// </summary>
    public bool IsCell => IsRowVector && IsColumnVector;

    /// <summary>
    /// Detect if the range consists of a row of cells
    /// </summary>
    public bool IsRowVector => Row == LastRow;

    /// <summary>
    /// Detect if the range consists of a column of cells
    /// </summary>
    public bool IsColumnVector => Column == LastColumn;

    /// <summary>
    /// Convert the row number into a zero-based index, as
    /// might be used for a two dimensional array
    /// </summary>
    /// <param name="row">The OpenXML row number, which
    /// is one-based</param>
    /// <returns>The zero-based row index</returns>
    /// <exception cref="IndexOutOfRangeException"></exception>
    public static int Index(int row) => row > 0 ? row - 1 
        : throw new IndexOutOfRangeException
            ("Row number must be greater than one");

    private readonly static Regex reCol 
        = new("[A-Z]{1,2}", RegexOptions.Compiled);

    /// <summary>
    /// Convert a column name into a zero-based column
    /// index, as might be used in a two dimensional
    /// array of cells.
    /// </summary>
    /// <param name="col">The column name</param>
    /// <returns>The zero-based column index</returns>
    public static int Index(string col)
    {
        if (!reCol.IsMatch(col))
            throw new ArgumentException
                ($"Column '{col}' should be one or two upper-case letters");
        int result = 0;
        if (col.Length == 2)
            result += 26 * (col[0] - 'A' + 1);
        result += col[^1] - 'A';
        return result;
    }

    /// <summary>
    /// Given a column number in the range 0 .. 701,
    /// generate the equivalent one or two character
    /// column name
    /// </summary>
    /// <param name="idx">The column index</param>
    /// <returns>The column name</returns>
    /// <exception cref="ArgumentException">Thrown if
    /// the index is out of range</exception>

    public static string ToColName(int idx)
    {
        if (idx < 26)
            return "" + (char)('A' + idx);
        if (idx < 27 * 26)
            return "" + (char)(idx / 26 - 1 + 'A') + (char)('A' + idx % 26);
        throw new ArgumentException
            ("Only one or two-letter column names supported");
    }

    /// <summary>
    /// Given a zero-based row number, convert it to
    /// one-based as used in spreadsheets
    /// </summary>
    /// <param name="idx">The zero-based row index</param>
    /// <returns>The one-based equivalent</returns>
    /// <exception cref="ArgumentException">Thrown if argument
    /// is negative</exception>
    public static int ToRowNumber(int idx) => idx >= 0 ? idx + 1 
        : throw new ArgumentException("Row index must be zero or positive");

    private readonly static Regex reCellRef = new
        (@"^((\w[\w ]*\w)!)?\$?([A-Z]{1,2})\$?([1-9]\d*)(:\$?([A-Z]{1,2})\$?([1-9]\d*))?$",
            RegexOptions.Compiled);

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="cellRef">A valid Excel cell reference</param>
    /// <exception cref="ArgumentException">The cell reference 
    /// string was incorrectly formed</exception>
    public XlCellRef(string cellRef)
    {
        Match m = reCellRef.Match(cellRef);
        if (!m.Success)
            throw new ArgumentException
                ($"Invalid cell(s) reference: {cellRef}");
        string? sheetName = m.Groups[2].Value;
        if (string.IsNullOrEmpty(sheetName))
            SheetName = null;
        else
            SheetName = sheetName;
        Column = m.Groups[3].Value;
        Row = int.Parse(m.Groups[4].Value);
        string? lastCol = m.Groups[6].Value;
        if (string.IsNullOrEmpty(lastCol))
            LastColumn = Column;
        else
            LastColumn = lastCol;
        string? lastRow = m.Groups[7].Value;
        if (string.IsNullOrEmpty(lastRow))
            LastRow = Row;
        else
            LastRow = int.Parse(lastRow);
        if (Index(LastColumn) < Index(Column) || LastRow < Row)
            throw new ArgumentException
                ("Row or column numbers the wrong way round");
    }

    public XlCellRef(XlSheet sheet, string cellRef)
        : this(cellRef)
    {
        if (SheetName == null)
            SheetName = sheet.Name;
        else if (SheetName != sheet.Name)
            throw new ArgumentException
                ($"Ambiguous sheet references {sheet.Name} and {SheetName}");
    }
}
