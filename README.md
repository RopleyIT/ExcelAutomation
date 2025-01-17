# ExcelAutomation API

This library is designed to provide a less painful access to the OpenXML library
used to create or manipulate Open office documents, such as those created using
Microsoft 365 tools.

In particular, this small library focusses on APIs used to create or edit spreadsheets.
With the library it is possible to create or open existing .XLSX files, to add or alter
the worksheets within the docuemnt's top-level workbook, to manipulate
the content of cells within worksheets, and to create charts linked to cells in those
worksheets.

At the current time, the library supports just bar charts as a chart type, but may in
due course expand to include other chart types if uptake and interest exist.

## A simple example

The following sample code shows how to use the library to create a new empty
spreadsheet, to populate some grids of cells with text and with numbers as tables,
and to add bar charts to the sheet based on those tables as their data sources.

First, the new spreadsheet is created using a static method of the `XlDocument`
class. Since this is intended to be a new spreadsheet, the method is invoked with
its second argument set to true, which causes a pre-existing file with the same 
name to be overwritten if that file exists. The new spreadsheet file will contain 
a workbook with a single worksheet in it named `Sheet 1` that holds no populated
cells.

```
using OpenXmlAutomation;

namespace OpenXmlDemo;

public class Demo
{
    public void CreateDemoSpreadsheet()
    {
        using XlDocument? doc = XlDocument.Create
            ("C:\\tmp\\WithChart.xlsx", true);
      
        . . .
    }
}
```
Next we need to access the worksheet that was created automatically. The `XlDocument`
object contains a `Sheets` property that is a list of `XlSheet` objects, in this
case containing just one `XlSheet`:

```
        . . .

        using XlDocument? doc = XlDocument.Create
            ("C:\\tmp\\WithChart.xlsx", true);
        XlSheet sheet = doc.Sheets[0];
```
We shall now create a small table with text cells for row and column headers, and
number cells for some values:

```
        . . .

        sheet1.SetCell("A1", "Chart title");
        sheet1.SetCell("A2", "Apples");
        sheet1.SetCell("A3", "Pears");
        sheet1.SetCell("A4", "Plums");
        sheet1.SetCell("B1", "2024");
        sheet1.SetCell("C1", "2025");
        sheet1.SetCell("B2", 0.78);
        sheet1.SetCell("B3", 0.56);
        sheet1.SetCell("B4", 0.33);
        sheet1.SetCell("C2", 0.31);
        sheet1.SetCell("C3", 0.49);
        sheet1.SetCell("C4", 0.72);
    }
}
```
Note that the end of the function closes the `using` statement, causing the
`XlDocument` object's `Dispose()` method to be invoked. This in turn causes
the spreadsheet file to be saved to disk.

## Creating a barchart

We shall now add a bar chart to the same spreadsheet. Since the file already
exists, we shall open it rather than create it, then obtain a reference to
the same sheet containing the table we just created. IN the following code,
we assume we are inside some other method of a class in our application:

```
     using XlDocument doc2 = XlDocument.Open("C:\\tmp\\WithChart.xlsx");
     XlSheet sheet2 = doc2.Sheets[0];
```
Now we create a bar chart object and make it a child of the sheet. Notice
that we also set the cell to use when selecting a title for the bar chart,
we set the chart grouping to clustered as there are multiple rows and
columns of data, and we arrange for the bars to be vertical rather than
horizontal. We also arrange for the chart area to have rounded corners,
and we choose the range of cells that determine the area where the chart
will be drawn on the sheet:

```
    XlBarChart barChart = new(sheet2)
    {
        ChartIndex = 1,
        ChartTitle = "A1",
        Grouping = "clustered",
        RoundCorners = true,
        Direction = "col",
        CellArea = "E3:L20"
    };
```
Our table has two columns of values in columns B and C, and has the
category labels for each row alongside those values in column A. We
create a bar chart series for each column of data, since we are going
to show clustered columns on the bar chart:

```
    barChart.SeriesList.Add(new XlBarChartSeries(sheet2)
    {
        CategoryCellRange = "Sheet1!$A$2:$A$4",
        SeriesTitleCell = "Sheet1!B1",
        ValueCellRange = "Sheet1!B2:B4",
        ValueFormat = "##.#%"
    });
    barChart.SeriesList.Add(new XlBarChartSeries(sheet2)
    {
        CategoryCellRange = "Sheet1!$A$2:$A$4",
        SeriesTitleCell = "Sheet1!C1",
        ValueCellRange = "Sheet1!C2:C4",
        ValueFormat = "##.#%"
    });
```
Lastly, we invoke the `Generate` method to create the barchart within 
the spreadsheet, saving the modified contents away into the object tree:

```
    barChart.Generate();
```
As before, the closing brace of the function containing the `using`
statement that opened the spreadsheet file will save the contents of
the file to disk.

Note that if you wish to create multiple charts on the same sheet,
you can use the same code structure as above. Just remember to give
each successive bar chart an increasing value for its `ChartIndex`
property when you create the bar chart object.

## Classes, methods and properties

### The `XlDocument` class

This class represents an entire spreadsheet file while it has been
loaded into memory and is being manipulated. The class implements
the `IDisposable` interface, guaranteeing that the document is
saved to disk securely upon going out of scope.

### Methods and properties

`public static XlDocument Open(string path);`

Used to open an existing spreadsheet file. Files must end with the
extension `.xlsx` and be a valid OpenXml-compliant spreadsheet. If
there is a problem with the file being opened, the method throws
an `ArgumentException` with a message indicating the reason for
failure.

`public static XlDocument? Create(string path, bool overwrite);`

This will create a new spreadsheet file named using the argument
called `path`. The file will contain a single empty worksheet
named `Sheet 1` similar to the behaviour of Excel. If the 
`overwrite` flag is set to true, an existing
spreadsheet with the same path will be overwritten with a new
empty spreadsheet. If `overwrite` is false, the method will return
`null` if the file already exists. Note that for all other problems,
such as the filename extension being invalid or the path being
unwritable, an `ArgumentException` is thrown.

`public void SaveWorkbook();`

Usually the `using` statment can be used to arrange that the
spreadsheet is closed at the appropriate time. However, the top
level workbook can be force-saved using this method. Note that
the spreadsheet remains open, and the `XlDocument` accessible
for further edits after a call to this method.

`public XlSheet? AddSheet(string sheetName);`

Add a new additional worksheet to the workbook in the spreadsheet
file. Returns a reference to the new worksheet in the form of
an `XlSheet` object if successful. The new sheet will have the
name provided in the parameter to the method. Note that if a sheet
with this name already exists, the method returns `null`. It does
not return a reference to the pre-existing sheet.

`public List<XlSheet> Sheets { get; }`

This property gives access to all the worksheets in the document.
Note that you should not create worksheets and add them to this
list yourself. Use the `AddSheet` method to do this.

`public XlRange? FindRange(string cellsRef);`

In Excel, we can create cell range references such as:
`Sheet 1!$A$3:$B$6` where `Sheet 1` is the name of a worksheet,
and the remainder of the reference is describing the rectangle
of cells from A3 to B6. As described below, an `XlRange` object
contains a two dimensional list of the `XlCell` objects found
in the named sheet when the cell range reference is passed to
the `FindRange` method. Note that as this method is in the
`XlDocument` class, the cell range reference must have the sheet
name at the start.

This function will return a `null` if the sheet name is missing
from the cell reference, or if there is no sheet with the name
provided.

`public void Dispose();`

As expected, this method flushes and closes the file, and releases
managed and unmanaged resources associated with the `XlDocument`.

## The `XlSheet` class

For each worksheet in the workbook of a spreadsheet document, there
will be a corresponding `XlSheet` object. A list of these is held
in the `XlDocument` object's `Sheets` property. This class is the
workhorse of the spreadhseet, as it gives acces to all the cells
in the sheet and their values and types, as well as managing any
charts or graphs created on the sheet.

### Methods and properties

`public string Name { get; set; }`

The name of this sheet, as appears on the tab at the bottom of the
sheet in Excel, and as used in the full cell range reference
succeeded by an exclamation mark. Setting this property will throw
an exception if the name is changed to something that is already
the name of another worksheet in the workbook.

`public XlCell FindCell(string cellRef);`

Looks for a single cell in the current worksheet. The `stringRef`
argument does not have the sheet name at the beginning, just the
unadorned cell reference within the sheet, e.g. `A12`. This method
will throw an exception if the cell reference is badly formed.

```
public XlCell SetCell(string cellName, string? value);
public XlCell SetCell(string cellName, double value);
```
In the current sheet, set the value of the cell at cell reference
`cellName` to the specified value. The type of the cell is also
set. For example the string version above sets the cell type to
general text, whereas the version taking a `double` argument sets
the cell to have a number type.

One important thing to note is that calling the string version of
this method with a `null` string argument removes the cell from the
sheet, i.e. sets it to have no value and no type. This can be used to
clear a cell on the sheet.

## The `XlCell` class

This class encapsulates all the attributes and behaviours for a single
cell on a worksheet. Instances of the `XlCell` class are looked up
using the `XlSheet` object's `FindCell` method. If the cell is empty,
that method still finds the cell, even though it contains no value.
Similarly, values can be set into a cell from the parent sheet using
its `SetCell` method as described previously.

### Methods and properties

`public string Column { get; }`

For a cell whose fully qualified cell reference was `Sheet 1!$AA$23`
this would give the value `AA` as a string object.

`public uint Row { get; }`

For the same cell reference above, this would yield the value 23. Note
that row numbers start at 1 for the top row, not 0.

`public int DataType { get; set; }`

This should really be an enum, but for now it is set to one of a number
of named integer constants from the following list:

`XlCell.Empty, XlCell.String, XlCell.Number, XlCell.Date, XlCell.Boolean`

It is important to set the data type before setting the `Value` property
so that valid type conversion can be applied within the underlying
OpenXml SDK. The `XlSheet` class's `SetCell` methods do this automatically.

`public uint ColumnIndex { get; }` and `public uint RowIndex { get; }`

These two properties convert the `Column` and `Row` values into zero-based
integer values. Thus row 1 has row index of zero, and column `"ZZ"` has
a column index of 701.

`public string CellName { get; }`

This property merely fuses the `Column` and `Row` properties to provide
the local cell name for the cell.

```
public bool IsDouble { get; }
public bool IsDecimal { get; }
public bool IsDate { get; }
```
These boolean properties confirm the data type of the current cell, from the
point of view of a variable into which you might assign its value.

```
public double AsDouble { get; }
public decimal AsDecimal { get; }
public DateTime AsDate { get; }
```
These properties give back the value of the cell in the nominated data type.
If the cell cannot be converted to that type, or is empty, these will throw
exceptions. Hence using the `IsXXX` properties first is advised. Note too
that for strings held in the cell, the `Value` property can be used directly.

```
public void Set(double value);
public void Set(decimal value);
public void Set(DateTime value);
public void Set(string? value);
```
These methods apply the appropriate type to the `DataType` property, then
set the value of the cell to the argument.

`public string? Value { get; set; }`

If the cell contains a value, it is returned as a string representation,
regardless of the data type the cell is specified to have. If the cell
is empty, this will return a `null`. When setting the value using this
property, it is important to set the `DataType` first, otherwise the value
will be stored with a string type in the cell.

Setting the value to null or the empty string clears the cell.

## The `XlCellRef` class

Really a helper class, this class encodes and decodes cell references and
cell range references safely. It is used by other classes in this
library, but is made available in case you wish to parse your own cell
references.

### Methods and properties

`public XlCellRef(string cellRef);`

The constructor for this class. Parses the cell reference and extracts
its various fields, making them available through its other properties
and methods. The cell reference argument can have the sheet name on the
front or not.

`public XlCellRef(XlSheet sheet, string cellRef);`

If the cell reference does not have the sheet name at the front, the name
is captured from the `sheet` argument instead. If both the `cellRef` and
`sheet` arguments define a sheet the cell is on, the constructor looks to
check they match. If not, an `ArgumentException` is thrown.

`public string? SheetName { get; }`

Yields the name of the worksheet this cell reference is pointing to. Will
be set to `null` if the `XlCellRef` reference does not contain the sheet name 
at its front.

`public string Column { get; }`

The column name for this cell, or the top left corner cell if the `XlCellRef`
is referencing a range of cells.

`public int Row { get; }`

The one-based row number for the cell, or the top left corner cell if a range.

`public string LastColumn { get; }`

The column name for the cell at the bottom right corner of the cell range. If
the cell reference is to a single cell, this will be the same as the
`Column` property.

`public int LastRow { get; }`

The one-based row number of the bottom right corner cell of the cell range.
If the cell reference was for a single cell, this will have the same value
as the `Row` property.

`public bool IsCell { get; }`

True if the cell reference points to a single cell, even if it has been
described using range syntax (e.g. `Sheet 2!$B$3:$B$3`).

`public bool IsRowVector { get; }`

True if all the cells in the range are in the same row. Also true therefore
or a single cell.

`public bool IsColumnVector { get; }`

True if the cells in the range are in the same column. Also true for a
single cell.

```
public static int Index(int row);
public static int Index(string column);
```
These static helper functions take a one-based row number or a column
identifier as a single letter or letter pair, and generate the corresponding
zero-based index for the cell concerned.

```
public static int ToRowNumber(int rowIndex);
public static string ToColName(int colIndex);
```
These static helper functions perform the reverse operations to the two
`Index` methods, converting the row or column indices to either the one-based
row number or the column names.

