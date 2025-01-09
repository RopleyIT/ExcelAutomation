using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Office2021.Excel.RichDataWebImage;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
namespace OpenXmlAutomation;

/// <summary>
/// Class representing an Excel document plus its
/// assumed single top-level workbook.
/// </summary>

public class XlDocument : IDisposable
{
    private SpreadsheetDocument document;
    private WorkbookPart workbookPart;
    private SharedStringTablePart sharedStringTablePart;
    
    /// <summary>
    /// Constructor used to open an existing Excel
    /// spreadsheet document
    /// </summary>
    /// <param name="path">The path to the Excel
    /// document file</param>
    /// <exception cref="ArgumentException">
    /// Thrown if the Excel file is missing, badly
    /// named, or badly formed</exception>
    
    public XlDocument(string path)
    {
        ValidatePath(path);
        document = SpreadsheetDocument.Open(path, true);
        if (document == null)
            throw new ArgumentException
                ($"Spreadsheet file {path} is badly formed");
        var wbp = document.WorkbookPart;
        if (wbp == null)
            throw new ArgumentException($"Spreadsheet file {path} has no workbook");
        else
            workbookPart = wbp;

        // Initialise the list of sheets

        IEnumerable<Sheet> sheets 
            = workbookPart.Workbook.Descendants<Sheet>();
        foreach (Sheet sheet in sheets)
        {
            string? id = sheet.Id;
            if (id is null) 
                continue;
            WorksheetPart wsp = (WorksheetPart)workbookPart.GetPartById(id);
            Sheets.Add(new XlSheet(this, wsp, sheet));
        }

        // Find or create the shared string table

        if (workbookPart.GetPartsOfType<SharedStringTablePart>().Any())
            sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
        else
            sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
    }

    /// <summary>
    /// Create a new empty spreadsheet (with one workbook and one
    /// worksheet named "Sheet1" as Excel does)
    /// </summary>
    /// <param name="path">The path to the file location for the
    /// new spreadsheet file</param>
    /// <param name="overwrite">Set true to overwrite any existing
    /// file with the specified path</param>
    /// <returns>The open document object if successful. Null
    /// if not.</returns>
    
    public static XlDocument? Create(string path, bool overwrite = false)
    {
        if (File.Exists(path))
        {
            if (!overwrite)
                return null;
            else
                File.Delete(path);
        }

        // Create an XLSX spreadsheet document by supplying the filepath.
        // By default, AutoSave = true, Editable = true, and Type = xlsx.

        using (SpreadsheetDocument doc = SpreadsheetDocument
            .Create(path, SpreadsheetDocumentType.Workbook))
        {
            WorkbookPart wbp = doc.AddWorkbookPart();
            wbp.Workbook = new();

            // Create the shared string table

            SharedStringTablePart sstPart 
                = wbp.AddNewPart<SharedStringTablePart>();
            sstPart.SharedStringTable = new SharedStringTable();
            sstPart.SharedStringTable.Save();

            // Add the single first sheet to the spreadsheet

            AddSheetToWorkbook(wbp, "Sheet1");

            // Save the new spreadsheet to disk

            wbp.Workbook.Save();
        }

        // Now reopen the document to construct the XlDocument in memory

        XlDocument xlDoc = new (path);
        return xlDoc;
    }

    private static (WorksheetPart, Sheet) AddSheetToWorkbook
        (WorkbookPart wbp, string name)
    {
        WorksheetPart wsp = wbp.AddNewPart<WorksheetPart>();
        wsp.Worksheet = new Worksheet(new SheetData());
        string relationshipId = wbp.GetIdOfPart(wsp);

        Sheets sheets = wbp.Workbook.GetFirstChild<Sheets>()
            ?? wbp.Workbook.AppendChild(new Sheets());

        uint sheetId = 1;
        if (sheets.Elements<Sheet>().Any())
        {
            uint? maxId = sheets.Elements<Sheet>()
                .Select(s => s.SheetId?.Value).Max();
            if (maxId is null)
                sheetId = (uint)sheets.Elements<Sheet>().Count() + 1;
            else
                sheetId = maxId.Value + 1;
        }

        Sheet sheet = new ()
        {
            Id = relationshipId,
            SheetId = sheetId,
            Name = name
        };
        sheets.Append(sheet);
        return (wsp, sheet);
    }

    /// <summary>
    /// The collection of worksheets inside the top level workbook
    /// </summary>

    public List<XlSheet> Sheets { get; private set; } = [];

    /// <summary>
    /// Add another new worksheet to the workbook. Note that we do not
    /// add a sheet directly to the Sheets collection, as the underlying
    /// OpenXml objects would not be created. This is prevented by
    /// making the XlSheet constructor to have internal scope.
    /// </summary>
    /// <param name="sheetName">The name of the new worksheet. The name
    /// must be unique within the document.</param>
    /// <returns>The new worksheet, or null if the name was already
    /// in use for another worksheet</returns>
    
    public XlSheet? AddSheet(string sheetName)
    {
        // Check there is not already a sheet with this name

        if (Sheets.Any(s => s.Name == sheetName))
            return null;

        // Use OpenXml API to create a
        // worksheet and add it to the workbook
        
        WorksheetPart? wsp;
        Sheet? sheet;
        (wsp, sheet) = AddSheetToWorkbook(workbookPart, sheetName);

        // Now create the high-level XlSheet object and add it
        // to the list of sheets within the document

        XlSheet xlSheet = new(this, wsp, sheet);
        Sheets.Add(xlSheet);
        return xlSheet;
    }

    private static Regex r = new (".*\\.[xX][lL][sS][xXmM]?$");

    private static void ValidatePath(string path)
    {
        if (!File.Exists(path))
            throw new ArgumentException($"File {path} does not exist");
        if (!r.IsMatch(path))
            throw new ArgumentException
                ($"File {path} must end with '.xls.', '.xlsx' or '.xlsm'");
    }

    private static readonly Regex cellNameRegex
        = new Regex(@"^([^!]+)!\$([A-Z]+)\$(\d+)(::\$(A-Z)+\$(\d+))?$", 
            RegexOptions.Compiled);

    //public XlRange FindRange(string cellRange)
    //{
    //    Match m = cellNameRegex.Match(cellRange);
    //    if (!m.Success)
    //        throw new ArgumentException($"Badly formed cell reference: {cellRange}");

    //    // TODO: FInish this

    //}

    /// <summary>
    /// Ensure that a text string used in the spreadsheet is
    /// held in the shared string table of the document
    /// </summary>
    /// <param name="text">The string to be stored</param>
    /// <returns>The index in the string table at which
    /// the string was stored</returns>
    
    internal int InsertSharedStringItem(string text)
    {
        // If the part does not yet contain a SharedStringTable, create it.
        
        sharedStringTablePart.SharedStringTable ??= new SharedStringTable();

        // See if the string exists already in the table. If so, return its index.

        var ssiQuery = sharedStringTablePart
            .SharedStringTable.Elements<SharedStringItem>();
        int i = 0;
        foreach (SharedStringItem item in ssiQuery)
        {
            if (item.InnerText == text)
                return i;
            i++;
        }

        // The text does not exist in the table. Create the SharedStringItem.

        SharedStringItem ssi = new (new DocumentFormat.OpenXml.Spreadsheet.Text(text));
        sharedStringTablePart.SharedStringTable.AppendChild(ssi);
        sharedStringTablePart.SharedStringTable.Save();
        return i;
    }

    /// <summary>
    /// Return the list of strings that are held in the
    /// shared string table. Used for debugging and testing.
    /// </summary>
    
    public IEnumerable<string> SharedStringTableContents
    {
        get
        {
            if(sharedStringTablePart.SharedStringTable is null)
                return Enumerable.Empty<string>();
            return sharedStringTablePart.SharedStringTable
                .Elements<SharedStringItem>()
                .Select(ssi => ssi.InnerText);
        }
    }

    internal string? LookupSharedString(int ssi)
    {
        // If the part does not yet contain a SharedStringTable, create it.

        if (sharedStringTablePart.SharedStringTable is null)
        {
            sharedStringTablePart.SharedStringTable = new SharedStringTable();
            return null;
        }

        // See if the string exists already in the table. If so, return it.

        var itemQuery = sharedStringTablePart
            .SharedStringTable.Elements<SharedStringItem>();
        if (itemQuery.Count() <= ssi || ssi < 0)
            return null;
        SharedStringItem? item = itemQuery.ElementAt(ssi);
        return item.InnerText;
    }

    /// <summary>
    /// When a cell is deleted, or the text in a cell is changed,
    /// we need to remove its old text from the shared string table
    /// </summary>
    /// <param name="ssi">The shared string table index for the string</param>

    internal void UnlinkSharedStringItem(int ssi)
    {
        int refCount = 0;

        // Search the entire workbook for other cells that are referencing the
        // shared string item with this ID. If found, we cannot remove it.

        foreach (var part in workbookPart.GetPartsOfType<WorksheetPart>())
        {
            var cells = part.Worksheet
                .GetFirstChild<SheetData>()
                ?.Descendants<Cell>();
            if (cells is null)
                continue;
            foreach (var cell in cells)
            {
                if (cell.DataType is not null
                    && cell.DataType.Value == CellValues.SharedString
                    && cell.CellValue?.Text == ssi.ToString())
                {
                    refCount++;
                    if(refCount > 1)
                        break;
                }
            }
            if (refCount > 1)
                break;
        }

        // If we get here and the reference count is one or less, we
        // can remove the shared string from the shared string table.

        if(refCount <= 1)
        {
            SharedStringItem removee = sharedStringTablePart
                .SharedStringTable
                .Elements<SharedStringItem>()
                .ElementAt(ssi);
            removee?.Remove();

            // As we have removed one item from potentially the middle
            // of a list of shared string items, all the remaining
            // items in the list have a shared string ID that is now
            // off by one. (Who on earth designed this??) We now need
            // to fix these broken references.

            foreach (var part in workbookPart.GetPartsOfType<WorksheetPart>())
            {
                var cells = part.Worksheet
                    .GetFirstChild<SheetData>()
                    ?.Descendants<Cell>();

                if (cells is null)
                    continue;

                foreach (var cell in cells)
                {
                    if (cell.DataType is not null 
                        && cell.DataType.Value == CellValues.SharedString 
                        && int.TryParse(cell.CellValue?.Text, out int itemIndex))
                    {
                        if (itemIndex > ssi)
                            cell.CellValue.Text = (itemIndex - 1).ToString();
                    }
                }
                part.Worksheet.Save();
            }
            workbookPart.SharedStringTablePart?.SharedStringTable.Save();
        }
    }

    private bool disposedValue;

    protected virtual void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
                document?.Dispose();
            disposedValue = true;
        }
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
