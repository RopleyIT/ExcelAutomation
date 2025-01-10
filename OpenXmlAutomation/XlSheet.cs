using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace OpenXmlAutomation
{
    public class XlSheet
    {
        internal XlDocument document;
        internal WorksheetPart part;
        internal Sheet sheet;
        private Dictionary<string, XlCell> xlCells;
        internal XlSheet(XlDocument doc, WorksheetPart wsPart, Sheet s)
        {
            document = doc;
            part = wsPart;
            sheet = s;
            xlCells = new();
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

        // Match cell ranges, including single cells. Example:
        // "$A$3:$B$12" for cells in rectangle A3 to B12, or
        // "$A$3" for the single cell A3.

        private static readonly Regex reRange 
            = new(@"\$([A-Z]{1,2})\$(\d+)(::\$([A-Z]{1,2})\$(\d+))?$", 
                RegexOptions.Compiled);
        
        /// <summary>
        /// Given a cell range string in the current worksheet,
        /// find the array of cells matching this range
        /// </summary>
        /// <param name="range">The range string</param>
        /// <returns>The cell range object</returns>
        /// <exception cref="ArgumentException">
        /// Thrown if the range string is badly formed
        /// </exception>
        
        public XlRange FindRange(string range)
        {
            Match m = reRange.Match(range);
            if(!m.Success || m.Groups.Count != 5)
                throw new ArgumentException($"Range {range} is invalid.");
            int topIdx = int.Parse(m.Groups[2].Value)-1;
            int botIdx = int.Parse(m.Groups[4].Value)-1;
            int leftIdx = XlCell.ToColIndex(m.Groups[1].Value);
            int rtIdx = XlCell.ToColIndex(m.Groups[3].Value);
            if (rtIdx < leftIdx || botIdx < topIdx)
                throw new ArgumentException
                    ($"Range {range} has cell references reversed");

            // Create the two dimensional array of cells as a list of lists
            // with the correct capacity for the number of rows and columns
            // to prevent reallocation and copying on list growth.

            int rowLength = rtIdx - leftIdx + 1;
            int colDepth = botIdx - topIdx + 1;
            List<List<XlCell>> cells = new(rowLength);
            for (int colIdx = 0; colIdx < rowLength; colIdx++)
            {
                cells.Add(new(colDepth));
                for(int rowIdx = 0; rowIdx < colDepth; rowIdx++)
                {
                    XlCell cell = FindCell(XlCell
                        .ToColName(leftIdx + colIdx) + (topIdx + rowIdx + 1));
                    cells.Last().Add(cell);
                }
            }
            return new XlRange(this, range, cells);
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
            if(row is null)
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
            xlCell.Value = value;
            return xlCell;
        }

        /// <summary>
        /// Lazy update of the actual XML cell in the sheet
        /// </summary>
        /// <param name="xlCell">The containing XlCell</param>
        /// <param name="cellValue">The value to place in the cell
        /// </param>

        internal void UpdateCellInSheet
            (XlCell xlCell, string? cellValue)
        {
            // Deal with the removal of data from a cell

            if (string.IsNullOrEmpty(cellValue))
            {
                if (xlCell.cell is not null)
                {
                    DeleteCell(xlCell.cell);
                    xlCell.cell = null;
                }
                return;
            }

            // The cell previously existed and we are updating its contents

            if (xlCell.cell is not null)
            {
                UpdateSharedStringsForCell(xlCell.cell, cellValue);
                return;
            }

            // The cell did not previously exist, and we need
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

            UpdateSharedStringsForCell(cell, cellValue);
            xlCell.cell = cell;
        }

        private void UpdateSharedStringsForCell(Cell cell, string newValue)
        {
            if (string.IsNullOrEmpty(newValue))
                throw new ArgumentException
                    ("Can only update shared string table to new value");
            
            if(cell is not null 
                && cell.DataType is not null
                && cell.DataType.Value == CellValues.SharedString)
            {
                string? id = cell.CellValue?.Text;
                if (int.TryParse(id, out int ssi))
                    document.UnlinkSharedStringItem(ssi);
            }

            int stringIndex = document.InsertSharedStringItem(newValue);
            cell!.CellValue = new CellValue(stringIndex.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
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
}
