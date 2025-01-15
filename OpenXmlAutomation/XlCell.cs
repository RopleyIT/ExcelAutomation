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
    internal XlCellRef cellRef;
    internal XlSheet sheet;
    internal Cell? cell;

    public const int Empty = 0;
    public const int String = 1;
    public const int Number = 2;
    public const int Date = 3;
    public const int Boolean = 4;

    internal XlCell(XlSheet parentSheet, string cellName)
    {
        sheet = parentSheet;
        cellRef = new(parentSheet, cellName);
    }

    public string Column => cellRef.Column;
    public uint Row => (uint)cellRef.Row;

    private int dataType = 1;
    
    /// <summary>
    /// The type of data last placed in the cell
    /// </summary>
    public int DataType
    {
        get => dataType;
        set
        {
            if(value <= XlCell.Boolean && value >= XlCell.Empty)
                dataType = value;
            else 
                dataType = 0;
        }
    }

    private CellValues ToDataType(int dt)
        => dt switch
        {
            XlCell.String => CellValues.SharedString,
            XlCell.Number => CellValues.Number,
            XlCell.Date => CellValues.Date,
            XlCell.Boolean => CellValues.Boolean,
            _ => CellValues.SharedString
        };

    /// <summary>
    /// Given a column name, e.g. "C" or "AA", return
    /// its zero-based column index. WARNING: This
    /// method assumes you've already checked the string
    /// 'col' has one or two uppercase letters.
    /// </summary>
    /// <param name="col">The column string</param>
    /// <returns>The index for the column</returns>

    public static int ToColIndex(string col) 
        => XlCellRef.Index(col);

    public uint ColumnIndex => (uint)ToColIndex(Column);

    public uint RowIndex => Row - 1;
    
    public string CellName
    {
        get => Column + Row;
    }

    public string? Value
    {
        get
        {
            if (cell is null || cell.DataType is null)
                return string.Empty;
            CellValues dt = cell.DataType.Value;
            if (dt == CellValues.SharedString)
            {
                string? id = cell.CellValue?.Text;
                if(int.TryParse(id, out int ssi))
                    return sheet.document.LookupSharedString(ssi);
            }
            else if(dt == CellValues.Number || dt == CellValues.Date
                || dt == CellValues.Boolean || dt == CellValues.InlineString
                || dt == CellValues.String)
            {
                return cell.CellValue?.Text;
            }
            return null;
        }
        set
        {
            sheet.UpdateCellInSheet(this, value, ToDataType(DataType));
        }
    }
}
