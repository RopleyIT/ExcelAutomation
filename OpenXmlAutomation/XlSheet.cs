﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OpenXmlAutomation;

public class XlSheet
{
    internal XlDocument document;
    internal WorksheetPart part;
    internal Sheet sheet;
    private readonly Dictionary<string, XlCell> xlCells;
    internal XlSheet(XlDocument doc, WorksheetPart wsPart, Sheet s)
    {
        document = doc;
        part = wsPart;
        sheet = s;
        xlCells = [];
    }

    public string Name
    {
        get
        {
            return sheet.Name?.Value ?? string.Empty;
        }
        set
        {
            if (document.Sheets.Any(s => s.Name == value))
                throw new ArgumentException
                    ($"A worksheet named {value} already exists");
            sheet.Name = value;
        }
    }

    /// <summary>
    /// Factory for an XlCell. Do not use the
    /// XlCell constructor from outside this
    /// package, use this factory method
    /// instead. That is why the XlCell
    /// constructor is marked internal.
    /// </summary>
    /// <param name="cellName">The cell 
    /// reference, without the sheet name</param>
    /// <returns>The cell reference. Note that
    /// this cell reference will use lazy
    /// updates to the spreadsheet itself, only
    /// creating the real cell if a value
    /// is applied to it.</returns>

    public XlCell FindCell(string cellName)
    {
        // Look for the cell object in the cache,
        // otherwise create it afresh and add
        // it to the cache.

        if (xlCells.TryGetValue(cellName, out XlCell? xlCell))
            return xlCell;

        xlCell = new(this, cellName);
        xlCells.Add(cellName, xlCell);

        // Attach the cell object from the
        // spreadsheet, if it already exists

        Worksheet workSheet = part.Worksheet;
        SheetData? sheetData =
            workSheet.GetFirstChild<SheetData>();
        if (sheetData is null)
            return xlCell;

        // Does the row exist in which this cell
        // would be found?

        var rowQuery = sheetData
            .Elements<Row>()
            .Where(r => r.RowIndex is not null
                && r.RowIndex == xlCell.Row);
        Row? row = rowQuery.FirstOrDefault();
        if (row is null)
            return xlCell;

        // Does the cell with this cell name
        // exist in the row?

        var cellQuery = row
            .Elements<Cell>()
            .Where(c => c.CellReference is not null
                && c.CellReference.Value == xlCell.CellName);
        xlCell.cell = cellQuery.FirstOrDefault();
        return xlCell;
    }

    /// <summary>
    /// Set a value into a cell by cell name
    /// </summary>
    /// <param name="cellName">The cell to set a value into</param>
    /// <param name="value">The string to place in the cell</param>
    /// <returns>The cell object that was set</returns>

    public XlCell SetCell(string cellName, string? value)
    {
        XlCell xlCell = FindCell(cellName);
        xlCell.Set(value);
        return xlCell;
    }

    public XlCell SetCell(string cellName, double value)
    {
        XlCell xlCell = FindCell(cellName);
        xlCell.Set(value);
        return xlCell;
    }

    /// <summary>
    /// Lazy update of the actual XML cell in the sheet
    /// </summary>
    /// <param name="xlCell">The containing XlCell</param>
    /// <param name="cellValue">The value to place in the cell
    /// </param>

    internal void UpdateCellInSheet
        (XlCell xlCell, string? cellValue, CellValues dataType)
    {
        // If the cell previously existed, remove it before
        // updating its contents

        if (xlCell.cell is not null)
        {
            DeleteCell(xlCell.cell);
            xlCell.cell = null;
        }

        // If the cell is being cleared, there is no more to do here

        if (string.IsNullOrEmpty(cellValue))
            return;

        // The cell does not exist, and we need
        // to create it with non-empty contents

        Worksheet workSheet = part.Worksheet;
        SheetData sheetData =
            workSheet.GetFirstChild<SheetData>()
            ?? workSheet.AppendChild(new SheetData());

        // Worksheets are sparsely populated. This means that row
        // and cell objects are only present if they have been added.
        // First, check that the row containing this new cell is already
        // there. If not, add the row.

        var rowQuery = sheetData
            .Elements<Row>()
            .Where(r => r.RowIndex is not null
                && r.RowIndex == xlCell.Row);
        Row? row = rowQuery.FirstOrDefault();
        if (row is null)
        {
            row = new Row() { RowIndex = xlCell.Row };
            sheetData.Append(row);
        }

        // Now look along the row, and see if the
        // cell we are adding is already there. If not,
        // create it and add it to the row.

        var cellQuery = row
            .Elements<Cell>()
            .Where(c => c.CellReference is not null
                && c.CellReference.Value == xlCell.CellName);
        Cell? cell = cellQuery.FirstOrDefault();

        // Use the existing cell if it has been found. Otherwise
        // create a new cell and add it to the spreadsheet.

        if (cell is not null)
            throw new ArgumentException
                ("Cell exists on document, but not in XlCell");

        // Cells must be in order of ascending CellReference.
        // Determine where to insert the new cell.

        Cell? refCell = null;
        refCell = row.Elements<Cell>()
            .FirstOrDefault(c => string.Compare
                (c.CellReference?.Value, xlCell.CellName, true) > 0);
        cell = new Cell { CellReference = xlCell.CellName };
        row.InsertBefore(cell, refCell);

        UpdateCellValue(cell, cellValue, dataType);
        xlCell.cell = cell;
    }

    private void UpdateCellValue
        (Cell cell, string newValue, CellValues dataType)
    {
        if (dataType == CellValues.SharedString)
        {
            int stringIndex = document.InsertSharedStringItem(newValue);
            cell!.CellValue = new CellValue(stringIndex.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }
        else if (dataType == CellValues.Number)
        {
            if (decimal.TryParse(newValue, out decimal value))
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.CellValue = new CellValue(value);
                //cell.StyleIndex = 3;
            }
            else
                throw new ArgumentException
                    ($"Cell of type Number cannot be set to {newValue}");
        }
        else if (dataType == CellValues.Date)
        {
            if (DateTime.TryParse(newValue, out DateTime value))
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                cell.CellValue = new CellValue
                    (value.ToOADate().ToString(CultureInfo.InvariantCulture));
                //cell.StyleIndex = 2;
            }
            throw new ArgumentException
                ($"Cell of date type cannot be set to {newValue}");
        }
        part.Worksheet.Save();
    }

    /// <summary>
    /// Delete the cell specified. Any use of subsequent
    /// XlCell objects referring to this cell will fail.
    /// </summary>
    /// <param name="cellName">The cell to be deleted</param>

    private void DeleteCell(Cell? cell)
    {
        if (cell is not null
            && cell.DataType is not null
            && cell.DataType.Value == CellValues.SharedString)
        {
            string? id = cell.CellValue?.Text;
            if (int.TryParse(id, out int ssi))
                document.UnlinkSharedStringItem(ssi);
        }
        cell?.Remove();
        part.Worksheet.Save();
    }
}
