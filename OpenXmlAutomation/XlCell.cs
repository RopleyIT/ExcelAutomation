using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OpenXmlAutomation;

public class XlCell
{
    internal XlSheet sheet;
    internal Cell? cell;
    internal XlCell(XlSheet parentSheet, string col, uint row) 
        : this(parentSheet, col + row) { }

    internal XlCell(XlSheet parentSheet, uint colIdx, uint rowIdx)
    {
        sheet = parentSheet;
        ColumnIndex = colIdx;
        RowIndex = rowIdx;
    }

    internal XlCell(XlSheet parentSheet, string cellName)
    {
        sheet = parentSheet;
        CellName = cellName;
    }

    public string Column { get; private set; } = string.Empty;
    public uint Row { get; private set; } = 0;


    private static readonly Regex reColName
        = new("^[A-Z]{1,2}$", RegexOptions.Compiled);

    /// <summary>
    /// Given a column name, e.g. "C" or "AA", return
    /// its zero-based column index. WARNING: This
    /// method assumes you've already checked the string
    /// 'col' has one or two uppercase letters.
    /// </summary>
    /// <param name="col">The column string</param>
    /// <returns>The index for the column</returns>

    public static int ToColIndex(string col)
    {
        if (!reColName.IsMatch(col))
            throw new ArgumentException
                ("Column names must be one or two upper case letters");
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
            return ('A' + idx).ToString();
        if (idx < 27 * 26)
            return (idx / 26 - 1 + 'A').ToString() + ('A' + idx % 26);
        throw new ArgumentException
            ("Only supports up to two-letter column names");
    }

    public uint ColumnIndex
    {
        get
        {
            return (uint)ToColIndex(Column);
        }
        set
        {
            if (value > 27*26)
                throw new ArgumentException("Column index too high for two letters");
            Column = ToColName((int)value);
        }
    }

    public uint RowIndex
    {
        get => Row - 1;
        set => Row = value + 1;
    }

    private readonly static Regex reCellName 
        = new(@"^([A-Z]{1,2})(\d+)$", RegexOptions.Compiled);
    
    public string CellName
    {
        get => Column + Row;
        set
        {
            Match m = reCellName.Match(value);
            if (!m.Success || m.Groups.Count != 3)
                throw new ArgumentException
                    ($"Cell name badly formed: {value}");
            Row = uint.Parse(m.Groups[2].Value);
            Column = m.Groups[1].Value;
        }
    }

    public string? Value
    {
        get
        {
            if (cell is null)
                return string.Empty;
            if (cell.DataType is not null
                && cell.DataType.Value == CellValues.SharedString)
            {
                string? id = cell.CellValue?.Text;
                if(int.TryParse(id, out int ssi))
                    return sheet.document.LookupSharedString(ssi);
            }
            return null;
        }
        set
        {
            sheet.UpdateCellInSheet(this, value);
        }
    }
}
