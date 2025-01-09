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


    public uint ColumnIndex
    {
        get
        {
            uint idx = 0;
            foreach(char c in Column)
            {
                idx *= 26;
                idx += (uint)(char.ToUpper(c) - 'A');
            }
            return idx;
        }
        set
        {
            if (value > 256)
                throw new ArgumentException("Column index > 256");
            uint leftDigit = value / 26;
            uint rightDigit = value % 26;
            if (leftDigit != 0)
                Column = ('A' + leftDigit).ToString() + (rightDigit + 'A');
            else
                Column = (rightDigit + 'A').ToString();
        }
    }

    public uint RowIndex
    {
        get => Row - 1;
        set => Row = value + 1;
    }

    private readonly static Regex reColumn 
        = new("[A-Z][A-Z]?", RegexOptions.Compiled);
    private readonly static Regex reRow 
        = new("\\d+", RegexOptions.Compiled);
    private readonly static Regex reCellName 
        = new("[A-Z][A-Z]?\\d+", RegexOptions.Compiled);
    
    public string CellName
    {
        get => Column + Row;
        set
        {
            if (reCellName.IsMatch(value))
            {
                Row = uint.Parse(reRow.Match(value).Value);
                Column = reColumn.Match(value).Value;
            }
            else
                throw new ArgumentException($"Invalid cell name: {value}");
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
